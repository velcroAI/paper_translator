"""
translate_paper.py
==================
단일 논문 파일을 즉시(실시간) 번역한다.  (v4)
  - 용어집 사전 추출(옵션, 사용자 용어집 우선) -> 청크 순차 번역(직전 청크 문맥 제공)
  - 원서 쪽 마커 / 각주 토큰 보존, 요약본 의심 청크는 1회 재번역
  - 각주(commentary/all 모드) 번역 패스 -> Word 각주
  - 저장(DOCX/PDF/MD) + QA 점검
비용은 배치(batch_submit.py)의 약 2배.
"""

import sys
import time
import anthropic
from pathlib import Path
from paper_utils import (
    API_KEY, MODEL, OUTPUT_FORMAT, EFFORT, USE_GLOSSARY, EXTRACT_FIGURES, FOOTNOTE_MODE,
    read_document, clean_text, chunk_text, chunk_needs_api,
    build_system_blocks, build_user_message, build_glossary,
    tail, CONTEXT_SRC_CHARS, CONTEXT_TR_CHARS,
    CostTracker, setup_tee_log, extract_response_text, call_with_retry,
    get_thinking_kwargs, MAX_OUTPUT_TOKENS, restore_markers, clean_translated,
    build_vision_system_blocks, build_vision_user_content, build_glossary_vision,
    group_vision_pages, process_clip_markers, assets_dir_for, namespace_vision_notes,
    build_notes_system_blocks, group_notes, build_notes_user_message, parse_notes_output,
    notes_to_defs, looks_summarized, finalize_and_save, build_units,
)

setup_tee_log("last_translate.log")

client = anthropic.Anthropic(api_key=API_KEY)

RETRY_NOTE = ("\n\n(IMPORTANT: The previous attempt omitted or summarized content. Translate the ENTIRE "
              "text above sentence by sentence, with nothing left out.)")


def _create(system_blocks, user_content, effort=None):
    return call_with_retry(lambda: client.messages.create(
        model=MODEL,
        max_tokens=MAX_OUTPUT_TOKENS,
        system=system_blocks,
        messages=[{"role": "user", "content": user_content}],
        **get_thinking_kwargs(MODEL, effort or EFFORT),
    ))


def translate_chunk(chunk, idx, total, tracker, system_blocks, prev_src, prev_tr, depth=0):
    print(f"  [{idx+1}/{total}] 번역 중... ({len(chunk)} chars)", flush=True)
    user = build_user_message(chunk, idx, total, prev_src, prev_tr)
    resp = _create(system_blocks, user)
    tracker.add(resp.usage)
    text = extract_response_text(resp.content)
    if resp.stop_reason == "max_tokens" and depth < 2 and len(chunk) > 800:
        # 출력 한도 도달 -> 청크를 반으로 나눠 재번역
        print(f"  [경고] 청크 {idx+1}: max_tokens 도달 — 반으로 나눠 재번역", flush=True)
        paras = chunk.split("\n\n")
        mid = max(1, len(paras) // 2)
        a, b = "\n\n".join(paras[:mid]), "\n\n".join(paras[mid:])
        if not a.strip() or not b.strip():
            return text
        ta = translate_chunk(a, idx, total, tracker, system_blocks, prev_src, prev_tr, depth + 1)
        tb = translate_chunk(b, idx, total, tracker, system_blocks,
                             tail(a, CONTEXT_SRC_CHARS), tail(ta, CONTEXT_TR_CHARS), depth + 1)
        return ta.rstrip() + "\n\n" + tb.lstrip()
    if depth == 0 and looks_summarized(chunk, clean_translated(text)):
        print(f"  [재번역] 청크 {idx+1}: 원문 대비 번역 분량이 너무 적음(요약/누락 의심) — 1회 재시도", flush=True)
        resp2 = _create(system_blocks, user + RETRY_NOTE)
        tracker.add(resp2.usage)
        t2 = extract_response_text(resp2.content)
        if len(clean_translated(t2)) > len(clean_translated(text)):
            text = t2
    print(f"  [{idx+1}/{total}] 완료 ({resp.usage.output_tokens} tokens)", flush=True)
    return text


def translate_notes(notes, glossary, tracker):
    """각주 번역 패스. 반환: '[^id]: ...' 정의 블록 문자열."""
    if not notes or FOOTNOTE_MODE == "strip":
        return ""
    groups = group_notes(notes)
    label = "논평성 각주만" if FOOTNOTE_MODE == "commentary" else "전체 각주"
    print(f"\n  각주 번역 ({label}): {len(notes)}개, 요청 {len(groups)}건", flush=True)
    system_blocks = build_notes_system_blocks(glossary)
    note_map = {}
    for gi, g in enumerate(groups):
        try:
            resp = _create(system_blocks, build_notes_user_message(g), effort="low")
            tracker.add(resp.usage)
            note_map.update(parse_notes_output(extract_response_text(resp.content), g))
        except Exception as e:
            print(f"  [경고] 각주 묶음 {gi+1} 실패 (원문 유지): {e}", flush=True)
    kept = sum(1 for v in note_map.values() if v)
    dropped = sum(1 for v in note_map.values() if v is None)
    print(f"  각주 완료: 번역 {kept}개" + (f", 서지 각주 삭제 {dropped}개" if dropped else ""), flush=True)
    return notes_to_defs(note_map, notes)


def _glossary_text(fn, *args):
    try:
        g, usage = fn(client, MODEL, *args)
        n = len([l for l in g.splitlines() if l.strip()])
        print(f"  용어집 {n}개 항목", flush=True)
        try:
            (Path(__file__).parent / "last_glossary.txt").write_text(g, encoding="utf-8")
        except Exception:
            pass
        return g, usage
    except Exception as e:
        print(f"  [경고] 용어집 추출 실패 (용어집 없이 진행): {e}", flush=True)
        return "", None


def run(input_file, output_format="docx"):
    input_file = input_file.strip('"').strip("'")
    print(f"파일: {input_file}", flush=True)
    print(f"모델: {MODEL}  (effort={EFFORT}, glossary={'on' if USE_GLOSSARY else 'off'}, "
          f"figures={'on' if EXTRACT_FIGURES else 'off'}, footnotes={FOOTNOTE_MODE})", flush=True)

    text, assets, stats = read_document(input_file)
    print(f"  원문: {len(text)} chars", flush=True)

    if stats.get("no_text_layer"):
        print("  텍스트 레이어 없는 스캔 PDF -> 페이지 이미지 비전 번역 모드", flush=True)
        return run_vision(input_file, output_format, assets, stats)

    if stats:
        pn = stats.get("page_numbers") or []
        print(f"  추출: 도식 {stats.get('figures',0)}, 표 {stats.get('tables',0)}, "
              f"수식 {stats.get('equations',0)}, 페이지 경계 병합 {stats.get('merged_pages',0)}"
              + (f", OCR 페이지 {stats['ocr_pages']}" if stats.get('ocr_pages') else ""), flush=True)
        if pn:
            src = {"printed": "인쇄 쪽번호", "manual": "지정값", "pdf": "PDF 쪽번호"}.get(stats.get("page_source"), "")
            print(f"  원서 쪽: p.{pn[0]}–{pn[-1]} ({src})", flush=True)
        if stats.get("notes"):
            print(f"  각주: {len(stats['notes'])}개 수집", flush=True)

    text = clean_text(text)
    print(f"  정제 후: {len(text)} chars", flush=True)

    units = build_units(text, stats)
    chunks = [u["src"] for u in units if u["type"] == "text"]
    nvis = sum(1 for u in units if u["type"] == "vision")
    print(f"  청크 수: {len(chunks)}" + (f"  + 스캔 쪽 비전 요청 {nvis}건 (텍스트 레이어 없는 {stats.get('vision_page_count')}쪽)" if nvis else ""), flush=True)
    if stats.get("endnote_pages"):
        ep = stats["endnote_pages"]
        print(f"  장 끝 미주(Notes) 절: p.{ep[0]}–{ep[-1]} 본문에서 분리", flush=True)
    print(flush=True)

    if not units:
        print("[중단] 이 문서에서 번역할 텍스트를 찾지 못했습니다.", flush=True)
        print("  PDF가 암호화되었거나 내용이 비어 있을 수 있습니다.", flush=True)
        return

    tracker = CostTracker(MODEL)
    translated = []
    glossary = ""

    try:
        if USE_GLOSSARY and len(text) > 1500:
            print("  용어집 추출 중...", flush=True)
            glossary, usage = _glossary_text(build_glossary, text)
            if usage:
                tracker.add(usage)
            print(flush=True)

        system_blocks = build_system_blocks(glossary)
        vision_blocks = build_vision_system_blocks(glossary) if nvis else None
        prev_src, prev_tr = "", ""
        restored = 0
        total = len(units)

        for i, u in enumerate(units):
            if u["type"] == "vision":
                group = u["group"]
                labels = "-".join(str(g["label"]) for g in group)
                print(f"  [{i+1}/{total}] 스캔 p.{labels} 번역 중 (이미지)...", flush=True)
                try:
                    resp = _create(vision_blocks, build_vision_user_content(group, len(group), prev_tr))
                    tracker.add(resp.usage)
                    out = clean_translated(extract_response_text(resp.content))
                    if resp.stop_reason == "max_tokens":
                        print(f"  [경고] p.{labels}: max_tokens 도달 — 잘렸을 수 있음", flush=True)
                    out = namespace_vision_notes(out, i, group)
                    out = process_clip_markers(out, input_file, group, assets, assets_dir_for(input_file))
                    print(f"  [{i+1}/{total}] 완료 ({resp.usage.output_tokens} tokens)", flush=True)
                except Exception as e:
                    print(f"  [오류] p.{labels} 실패: {e}", flush=True)
                    out = f"[번역 실패: p.{labels}]"
                translated.append(out)
                prev_src, prev_tr = "", tail(out, CONTEXT_TR_CHARS)
                time.sleep(0.3)
                continue
            c = u["src"]
            if not chunk_needs_api(c):
                print(f"  [{i+1}/{total}] 마커만 있음 — 통과", flush=True)
                translated.append(c)
                continue
            try:
                result = translate_chunk(c, i, total, tracker, system_blocks, prev_src, prev_tr)
            except Exception as e:
                try:
                    print(f"\n  [오류] 청크 {i+1} 실패: {e}", flush=True)
                except Exception:
                    print(f"\n  [오류] 청크 {i+1} 실패 (오류 메시지 출력 불가)", flush=True)
                result = f"[번역 실패: 청크 {i+1}]\n\n" + c
            result = clean_translated(result)
            result, missing = restore_markers(result, c)
            if missing:
                restored += len(missing)
                print(f"    [복원] 빠진 마커/각주 토큰 {len(missing)}개를 원래 위치 근처에 복원", flush=True)
            translated.append(result)
            prev_src = tail(c, CONTEXT_SRC_CHARS)
            prev_tr  = tail(result, CONTEXT_TR_CHARS)
            time.sleep(0.3)

        defs = translate_notes(stats.get("notes") or [], glossary, tracker)
        if defs:
            translated.append(defs)

        finalize_and_save(translated, input_file, output_format, assets, stats)
        if restored:
            print(f"  (참고) 모델이 놓친 토큰 {restored}개는 비슷한 위치에 자동 배치됨", flush=True)

    finally:
        print(flush=True)
        tracker.report()


def run_vision(input_file, output_format, assets, stats):
    """텍스트 레이어 없는 스캔 PDF: 페이지 이미지를 모델에 직접 보내 번역."""
    vision_pages = stats.get("vision_pages") or []
    if not vision_pages:
        print("[중단] 렌더된 페이지가 없습니다.", flush=True)
        return
    groups = group_vision_pages(vision_pages)
    print(f"  페이지 수: {len(vision_pages)} (요청 {len(groups)}건, 2쪽씩)", flush=True)
    print(flush=True)

    tracker = CostTracker(MODEL)
    translated = []
    glossary = ""
    try:
        if USE_GLOSSARY and len(vision_pages) >= 3:
            print("  용어집 추출 중 (앞부분 페이지 이미지)...", flush=True)
            glossary, usage = _glossary_text(build_glossary_vision, vision_pages)
            if usage:
                tracker.add(usage)
            print(flush=True)

        system_blocks = build_vision_system_blocks(glossary)
        prev_tr = ""
        for gi, group in enumerate(groups):
            labels = "-".join(str(g["label"]) for g in group)
            print(f"  [{gi+1}/{len(groups)}] p.{labels} 번역 중...", flush=True)
            content = build_vision_user_content(group, len(vision_pages), prev_tr)
            try:
                resp = _create(system_blocks, content)
            except Exception as e:
                print(f"  [오류] p.{labels} 실패: {e}", flush=True)
                translated.append(f"[번역 실패: p.{labels}]")
                prev_tr = ""
                continue
            tracker.add(resp.usage)
            out = clean_translated(extract_response_text(resp.content))
            if resp.stop_reason == "max_tokens":
                print(f"  [경고] p.{labels}: max_tokens 도달 — 잘렸을 수 있음", flush=True)
            out = namespace_vision_notes(out, gi, group)
            out = process_clip_markers(out, input_file, vision_pages, assets, assets_dir_for(input_file))
            translated.append(out)
            prev_tr = tail(out, CONTEXT_TR_CHARS)
            print(f"  [{gi+1}/{len(groups)}] 완료 ({resp.usage.output_tokens} tokens)", flush=True)
            time.sleep(0.3)

        finalize_and_save(translated, input_file, output_format, assets, stats, vision=True)
    finally:
        print(flush=True)
        tracker.report()


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        INPUT_FILE = sys.argv[1]
        FMT = sys.argv[2] if len(sys.argv) >= 3 else OUTPUT_FORMAT
    else:
        print("사용법: python translate_paper.py <파일경로> [docx|pdf|docx+pdf|md]")
        sys.exit(1)
    run(INPUT_FILE, FMT)
