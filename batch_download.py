"""
batch_download.py
=================
batch_submit.py 로 제출한 번역 배치 결과를 수집해 DOCX/PDF/MD 로 저장. (v4)
  - 각주 번역 결과(notes_NNNN) 병합, 비전 모드 각주 id 정리
  - 요약본/누락 의심 청크는 실시간 API 로 1회 재번역 (일반 요금)
  - 저장 후 QA 점검 (쪽 마커 순차, 각주 짝, 한글/원문 분량 비율)
"""

import sys
import json
import traceback
import anthropic
from pathlib import Path
from paper_utils import (
    API_KEY, OUTPUT_FORMAT, DEFAULT_MODEL, EFFORT,
    CostTracker, setup_tee_log, extract_response_text, call_with_retry,
    clean_translated, restore_markers, process_clip_markers,
    namespace_vision_notes, parse_notes_output, notes_to_defs,
    build_system_blocks, build_user_message, get_thinking_kwargs, MAX_OUTPUT_TOKENS,
    looks_summarized, finalize_and_save,
)
from pdf_extract import load_manifest

setup_tee_log("last_download.log")

BATCH_INFO_PATH = Path(__file__).parent / "batch_info.json"

client = anthropic.Anthropic(api_key=API_KEY)

RETRY_NOTE = ("\n\n(IMPORTANT: The previous attempt omitted or summarized content. Translate the ENTIRE "
              "text above sentence by sentence, with nothing left out.)")


def check_and_download():
    if not BATCH_INFO_PATH.exists():
        print("batch_info.json 파일이 없습니다.", flush=True)
        print("먼저 배치를 제출하세요 (bat: 메뉴 2 / GUI: [배치 제출]).", flush=True)
        return

    info = json.loads(BATCH_INFO_PATH.read_text(encoding="utf-8"))
    batch_id     = info["batch_id"]
    input_file   = info["input_file"]
    stem         = info["stem"]
    total_chunks = info["chunks"]
    model        = info.get("model", DEFAULT_MODEL)
    sources      = info.get("chunk_sources", [])
    passthrough  = info.get("passthrough", {})
    fmt          = info.get("output_format") or OUTPUT_FORMAT
    assets_dir   = info.get("assets_dir", "")
    note_groups  = info.get("note_groups") or []
    glossary     = info.get("glossary", "")
    qa_stats     = info.get("qa_stats") or {}

    print(f"Batch ID: {batch_id}", flush=True)
    print(f"파일:     {stem}", flush=True)
    print(f"청크:     {total_chunks}" + (f"  (+ 각주 요청 {len(note_groups)})" if note_groups else ""), flush=True)
    print(flush=True)

    batch  = client.messages.batches.retrieve(batch_id)
    status = batch.processing_status
    print(f"상태: {status}", flush=True)

    if status != "ended":
        print("\n아직 처리 중입니다. 나중에 다시 실행하세요.", flush=True)
        if batch.request_counts:
            rc   = batch.request_counts
            done = rc.succeeded + rc.errored
            print(f"  진행: {done}/{total_chunks - len(passthrough) + len(note_groups)}", flush=True)
        return

    print("결과 다운로드 중...", flush=True)

    results = {}
    tracker = CostTracker(model, batch_discount=True)
    rt_tracker = CostTracker(model)          # 재번역(실시간) 비용
    errors = 0
    truncated = 0
    missing_total = 0

    try:
        for result in client.messages.batches.results(batch_id):
            cid = result.custom_id
            if result.result.type == "succeeded":
                msg = result.result.message
                results[cid] = extract_response_text(msg.content)
                tracker.add(msg.usage)
                if getattr(msg, "stop_reason", None) == "max_tokens":
                    truncated += 1
                    print(f"  [경고] {cid}: max_tokens 도달 — 번역문이 잘렸을 수 있음", flush=True)
            else:
                errors += 1
                print(f"  오류: {cid}", flush=True)

        print(f"\n  결과 수집 완료 — 성공: {len(results)}  실패: {errors}  (잘림 의심: {truncated})", flush=True)

        assets = load_manifest(assets_dir) if assets_dir else {}
        assets = {k: v for k, v in assets.items() if v.get("kind") in ("fig", "tbl", "eq")}

        ordered = []
        vision = info.get("mode") == "vision"
        if vision:
            vision_pages = info.get("vision_pages") or []
            for i in range(total_chunks):
                key = f"group_{i:04d}"
                if key in results:
                    t = clean_translated(results[key])
                    t = namespace_vision_notes(t, i, vision_pages[i*2:(i+1)*2])
                    t = process_clip_markers(t, input_file, vision_pages, assets, assets_dir)
                    ordered.append(t)
                else:
                    labels = ", ".join(str(g["label"]) for g in vision_pages[i*2:(i+1)*2])
                    ordered.append(f"[번역 실패: p.{labels}]")
        else:
            system_blocks = None
            units = info.get("units") or [{"type": "text", "id": f"chunk_{i:04d}"} for i in range(total_chunks)]
            for i, u in enumerate(units):
                if u["type"] == "vision":
                    key = u["id"]
                    if key in results:
                        t = clean_translated(results[key])
                        t = namespace_vision_notes(t, i, u["group"])
                        t = process_clip_markers(t, input_file, u["group"], assets, assets_dir)
                        ordered.append(t)
                    else:
                        labels = ", ".join(str(g["label"]) for g in u["group"])
                        ordered.append(f"[번역 실패: p.{labels}]")
                    continue
                key = u["id"]
                src = sources[i] if i < len(sources) else ""
                if key in passthrough:
                    ordered.append(passthrough[key])
                    continue
                if key in results:
                    t = clean_translated(results[key])
                    if looks_summarized(src, t):
                        print(f"  [재번역] {key}: 원문 대비 분량이 너무 적음(요약/누락 의심) — 실시간 1회 재시도", flush=True)
                        try:
                            if system_blocks is None:
                                system_blocks = build_system_blocks(glossary)
                            user = build_user_message(src, i, total_chunks, "", "") + RETRY_NOTE
                            resp = call_with_retry(lambda: client.messages.create(
                                model=model, max_tokens=MAX_OUTPUT_TOKENS, system=system_blocks,
                                messages=[{"role": "user", "content": user}],
                                **get_thinking_kwargs(model, EFFORT)))
                            rt_tracker.add(resp.usage)
                            t2 = clean_translated(extract_response_text(resp.content))
                            if len(t2) > len(t):
                                t = t2
                        except Exception as e:
                            print(f"    재번역 실패 (배치 결과 유지): {e}", flush=True)
                    t, missing = restore_markers(t, src)
                    missing_total += len(missing)
                    ordered.append(t)
                else:
                    ordered.append(f"[번역 실패: {key}]\n\n" + src)

            # 각주
            if note_groups:
                note_map = {}
                for gi, g in enumerate(note_groups):
                    key = f"notes_{gi:04d}"
                    if key in results:
                        note_map.update(parse_notes_output(results[key], g))
                all_notes = [nt for g in note_groups for nt in g]
                kept = sum(1 for v in note_map.values() if v)
                dropped = sum(1 for v in note_map.values() if v is None)
                print(f"  각주: 번역 {kept}개" + (f", 서지 각주 삭제 {dropped}개" if dropped else ""), flush=True)
                defs = notes_to_defs(note_map, all_notes)
                if defs:
                    ordered.append(defs)

        fmt = fmt if fmt in ("docx", "pdf", "docx+pdf", "md") else "docx"
        print(f"  저장 경로: {Path(input_file).parent}", flush=True)
        print(f"  출력 포맷: {fmt.upper()}  (이미지 {len(assets)}개)", flush=True)

        saved = finalize_and_save(ordered, input_file, fmt, assets, qa_stats, vision=vision,
                                  citation=info.get("citation", ""))
        if saved:
            if missing_total:
                print(f"  (참고) 모델이 놓친 토큰 {missing_total}개는 비슷한 위치에 자동 배치됨", flush=True)
            try:
                BATCH_INFO_PATH.unlink()
                print("  임시 파일 정리 완료 (assets, batch_info.json)", flush=True)
            except Exception:
                pass

    except Exception:
        print("\n[예기치 않은 오류 발생]", flush=True)
        traceback.print_exc(file=sys.stdout)
        sys.stdout.flush()

    finally:
        tracker.report()
        if rt_tracker.calls:
            print("  (재번역 실시간 호출 비용은 별도)", flush=True)
            rt_tracker.report()


if __name__ == "__main__":
    check_and_download()
