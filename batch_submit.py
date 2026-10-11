"""
batch_submit.py
===============
논문 번역 요청을 Anthropic Batch API에 제출한다. (v4)
  - v4: 각주 번역 요청(notes_NNNN)을 같은 배치에 포함, 쪽 마커/QA 정보를 batch_info.json 에 기록
  - 용어집은 제출 전에 동기 호출로 1회 추출(옵션), 모든 배치 요청의 system 에 포함
  - 배치는 병렬 처리라 '직전 번역문' 문맥은 줄 수 없고, 직전 '원문' 꼬리만 문맥으로 제공
  - 이미지 자산(도식/표/수식)은 assets/<stem>/ 에 저장, batch_info.json 에 경로 기록
일반 API 대비 50% 할인.
"""

import sys
import json
import anthropic
from pathlib import Path
from datetime import datetime
from paper_utils import (
    API_KEY, MODEL, EFFORT, USE_GLOSSARY, EXTRACT_FIGURES, OUTPUT_FORMAT, FOOTNOTE_MODE,
    build_notes_system_blocks, group_notes, build_notes_user_message,
    read_document, clean_text, chunk_text, chunk_needs_api,
    build_system_blocks, build_user_message, build_glossary, tail, CONTEXT_SRC_CHARS,
    setup_tee_log, get_thinking_kwargs, MAX_OUTPUT_TOKENS, CostTracker, assets_dir_for,
    build_vision_system_blocks, build_vision_user_content, build_glossary_vision,
    group_vision_pages, build_units, get_citation,
)

setup_tee_log("last_submit.log")

BATCH_INFO_PATH = Path(__file__).parent / "batch_info.json"

client = anthropic.Anthropic(api_key=API_KEY)


def submit_batch(input_file):
    input_file = input_file.strip('"').strip("'")
    print(f"읽는 중: {input_file}", flush=True)
    print(f"모델: {MODEL}  (effort={EFFORT}, glossary={'on' if USE_GLOSSARY else 'off'}, "
          f"figures={'on' if EXTRACT_FIGURES else 'off'}, footnotes={FOOTNOTE_MODE})", flush=True)

    text, assets, stats = read_document(input_file)
    print(f"  원문: {len(text)} chars", flush=True)

    if stats.get("no_text_layer"):
        print("  텍스트 레이어 없는 스캔 PDF -> 페이지 이미지 비전 배치 모드", flush=True)
        return submit_batch_vision(input_file, stats)

    if stats:
        print(f"  추출: 도식 {stats.get('figures',0)}, 표 {stats.get('tables',0)}, "
              f"수식 {stats.get('equations',0)}, 페이지 경계 병합 {stats.get('merged_pages',0)}", flush=True)

    text = clean_text(text)
    print(f"  정제 후: {len(text)} chars", flush=True)

    units = build_units(text, stats)
    chunks = [u.get("src", "") for u in units]          # 단위별 원문 (비전 단위는 빈 문자열)
    nvis = sum(1 for u in units if u["type"] == "vision")
    print(f"  청크 수: {len(units) - nvis}" + (f"  + 스캔 쪽 비전 요청 {nvis}건" if nvis else ""), flush=True)

    glossary = ""
    tracker = CostTracker(MODEL)
    if USE_GLOSSARY and len(text) > 1500:
        print("  용어집 추출 중 (동기 호출 1회)...", flush=True)
        try:
            glossary, usage = build_glossary(client, MODEL, text)
            tracker.add(usage)
            print(f"  용어집 {len([l for l in glossary.splitlines() if l.strip()])}개 항목", flush=True)
        except Exception as e:
            print(f"  [경고] 용어집 추출 실패 (용어집 없이 진행): {e}", flush=True)

    system_blocks = build_system_blocks(glossary)
    vision_blocks = build_vision_system_blocks(glossary) if nvis else None
    requests = []
    passthrough = {}
    unit_desc = []
    prev_src = ""
    for i, u in enumerate(units):
        if u["type"] == "vision":
            gid = f"group_{i:04d}"
            unit_desc.append({"type": "vision", "id": gid, "group": u["group"]})
            requests.append({
                "custom_id": gid,
                "params": {
                    "model": MODEL,
                    "max_tokens": MAX_OUTPUT_TOKENS,
                    "system": vision_blocks,
                    "messages": [{"role": "user",
                                  "content": build_vision_user_content(u["group"], len(u["group"]))}],
                    **get_thinking_kwargs(MODEL, EFFORT),
                },
            })
            prev_src = ""
            continue
        chunk = u["src"]
        cid = f"chunk_{i:04d}"
        unit_desc.append({"type": "text", "id": cid})
        if not chunk_needs_api(chunk):
            passthrough[cid] = chunk
            continue
        requests.append({
            "custom_id": cid,
            "params": {
                "model": MODEL,
                "max_tokens": MAX_OUTPUT_TOKENS,
                "system": system_blocks,
                "messages": [{"role": "user",
                              "content": build_user_message(chunk, i, len(chunks), prev_src, "")}],
                **get_thinking_kwargs(MODEL, EFFORT),
            },
        })
        prev_src = tail(chunk, CONTEXT_SRC_CHARS)

    notes = stats.get("notes") or []
    note_groups = group_notes(notes) if (notes and FOOTNOTE_MODE != "strip") else []
    if note_groups:
        nsys = build_notes_system_blocks(glossary)
        for gi, g in enumerate(note_groups):
            requests.append({
                "custom_id": f"notes_{gi:04d}",
                "params": {
                    "model": MODEL,
                    "max_tokens": MAX_OUTPUT_TOKENS,
                    "system": nsys,
                    "messages": [{"role": "user", "content": build_notes_user_message(g)}],
                    **get_thinking_kwargs(MODEL, "low"),
                },
            })
        print(f"  각주 {len(notes)}개 -> 번역 요청 {len(note_groups)}건 추가 (모드: {FOOTNOTE_MODE})", flush=True)

    if not requests:
        print("\n[중단] 이 문서에서 번역할 텍스트를 찾지 못해 배치를 제출하지 않습니다.", flush=True)
        print("  PDF가 암호화되었거나 내용이 비어 있을 수 있습니다.", flush=True)
        return None

    print(f"  요청 제출 중 ({len(requests)}개, 통과 {len(passthrough)}개)...", flush=True)
    batch = client.messages.batches.create(requests=requests)

    print(flush=True)
    print("=" * 50, flush=True)
    print("  배치 제출 완료!", flush=True)
    print("=" * 50, flush=True)
    print(f"  Batch ID: {batch.id}", flush=True)
    print(f"  상태:     {batch.processing_status}", flush=True)
    print(f"  청크 수:  {len(chunks)}", flush=True)
    print("=" * 50, flush=True)

    info = {
        "batch_id":     batch.id,
        "input_file":   input_file,
        "stem":         Path(input_file).stem,
        "chunks":       len(units),
        "chunk_sources": chunks,            # 마커 복원용 원문 청크 (단위 순서)
        "units":        unit_desc,
        "citation":     get_citation(),
        "passthrough":  passthrough,
        "model":        MODEL,
        "output_format": OUTPUT_FORMAT,
        "assets_dir":   str(assets_dir_for(input_file)) if assets else "",
        "glossary_tokens": {"input": tracker.input_tokens, "output": tracker.output_tokens},
        "glossary":     glossary,
        "note_groups":  note_groups,
        "qa_stats":     {k: stats.get(k) for k in ("page_words", "page_source", "page_numbers", "notes",
                                               "hybrid", "vision_page_count", "endnote_pages",
                                               "endnotes_linked")},
        "submitted_at": datetime.now().isoformat(),
    }
    BATCH_INFO_PATH.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  정보 저장: {BATCH_INFO_PATH}", flush=True)
    if tracker.calls:
        print(f"  (용어집 추출 비용은 일반 요금: input {tracker.input_tokens:,} / output {tracker.output_tokens:,} tokens)", flush=True)
    print(flush=True)
    print("  컴퓨터를 꺼도 됩니다.", flush=True)
    print("  나중에 option 3 으로 결과를 받으세요.", flush=True)
    return batch.id


def submit_batch_vision(input_file, stats):
    """스캔 PDF: 페이지 이미지 요청들을 배치로 제출."""
    vision_pages = stats.get("vision_pages") or []
    if not vision_pages:
        print("[중단] 렌더된 페이지가 없습니다.", flush=True)
        return None
    groups = group_vision_pages(vision_pages)
    print(f"  페이지 수: {len(vision_pages)} (요청 {len(groups)}건, 2쪽씩)", flush=True)

    glossary = ""
    tracker = CostTracker(MODEL)
    if USE_GLOSSARY and len(vision_pages) >= 3:
        print("  용어집 추출 중 (동기 호출 1회)...", flush=True)
        try:
            glossary, usage = build_glossary_vision(client, MODEL, vision_pages)
            tracker.add(usage)
            print(f"  용어집 {len([l for l in glossary.splitlines() if l.strip()])}개 항목", flush=True)
        except Exception as e:
            print(f"  [경고] 용어집 추출 실패 (용어집 없이 진행): {e}", flush=True)

    system_blocks = build_vision_system_blocks(glossary)
    requests = []
    for gi, group in enumerate(groups):
        requests.append({
            "custom_id": f"group_{gi:04d}",
            "params": {
                "model": MODEL,
                "max_tokens": MAX_OUTPUT_TOKENS,
                "system": system_blocks,
                "messages": [{"role": "user",
                              "content": build_vision_user_content(group, len(vision_pages))}],
                **get_thinking_kwargs(MODEL, EFFORT),
            },
        })

    print(f"  요청 제출 중 ({len(requests)}개)...", flush=True)
    batch = client.messages.batches.create(requests=requests)

    print(flush=True)
    print("=" * 50, flush=True)
    print("  배치 제출 완료! (비전 모드)", flush=True)
    print("=" * 50, flush=True)
    print(f"  Batch ID: {batch.id}", flush=True)
    print(f"  상태:     {batch.processing_status}", flush=True)
    print(f"  요청 수:  {len(requests)}", flush=True)
    print("=" * 50, flush=True)

    info = {
        "batch_id":     batch.id,
        "mode":         "vision",
        "input_file":   input_file,
        "stem":         Path(input_file).stem,
        "chunks":       len(groups),
        "vision_pages": vision_pages,
        "glossary":     glossary,
        "citation":     get_citation(),
        "model":        MODEL,
        "output_format": OUTPUT_FORMAT,
        "assets_dir":   str(assets_dir_for(input_file)),
        "submitted_at": datetime.now().isoformat(),
    }
    BATCH_INFO_PATH.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  정보 저장: {BATCH_INFO_PATH}", flush=True)
    print(flush=True)
    print("  컴퓨터를 꺼도 됩니다. (assets 폴더의 페이지 이미지가 필요하니 지우지 마세요)", flush=True)
    print("  나중에 option 3 으로 결과를 받으세요.", flush=True)
    return batch.id


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        submit_batch(sys.argv[1])
    else:
        print("사용법: python batch_submit.py <파일경로>")
        sys.exit(1)
