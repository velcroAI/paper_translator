"""
pdf_extract.py
==============
PDF -> (본문 텍스트, 이미지 자산) 추출기.  paper_utils.read_document() 가 호출.

핵심 기능
  1) 블록(단락) 단위 추출 + 페이지 경계에서 잘린 단락 병합 + 줄끝 하이픈 결합
     -> 청크가 문장 중간에서 끊기지 않으므로 페이지 경계 누락이 사라진다.
  2) 러닝 헤더/푸터/페이지 번호를 위치·반복 기반으로 제거 (clean_text의 2차 필터 이전)
  3) 도식(래스터 이미지 + 벡터 드로잉), 표, 디스플레이 수식 영역을 감지해
     원본 PDF에서 PNG로 잘라내고, 본문의 해당 위치에 [[IMG:id]] 마커를 넣는다.
     마커는 번역 모델이 그대로 복사하도록 지시되며, 저장 단계에서 이미지로 치환된다.
  4) 각주 제거(작은 글씨 AND 페이지 하단) — 수식/캡션 블록은 각주 판정에서 제외
  5) 스캔+OCR PDF(JSTOR 등, GlyphLessFont)는 폰트 크기 정보가 엉터리이므로
     줄 높이/줄 간격 기반의 '유효 크기'를 대신 사용한다.

v4 추가
  6) 원서 쪽 마커: 각 쪽 시작에 [[PAGE:N]] 토큰 (N = 인쇄된 쪽번호 자동 감지, 실패 시 PDF 쪽).
     쪽 경계에서 이어지는 단락은 문단 중간에 인라인 토큰으로 들어간다.
  7) 각주 수집(footnote_mode = commentary | all): 하단 각주를 번호별로 모아 stats["notes"]로 반환하고,
     본문 콜아웃(위첨자/OCR 숫자)을 [^nK] 토큰으로 바꾼다. strip 모드는 기존처럼 제거만 한다.

반환값
  text   : 단락은 빈 줄(\n\n)로 구분. 헤딩은 '# ' / '## ' 접두어. 마커는 독립 단락.
  assets : {img_id: {"path": png 절대경로, "kind": fig|tbl|eq, "page": n,
                     "width_pt": w, "height_pt": h}}
"""

import re
import json
import shutil
import statistics
from collections import Counter
from pathlib import Path

try:
    import pymupdf as fitz          # PyMuPDF >= 1.24
except ImportError:                 # pragma: no cover
    import fitz                     # 구버전 호환

# ──────────────────────────────────────────────
# 상수 / 정규식
# ──────────────────────────────────────────────
MARKER_RE   = re.compile(r"\[\[\s*IMG\s*:\s*([A-Za-z0-9_\-]+)\s*\]\]")
CAPTION_RE  = re.compile(
    r"^\s*(figure|fig\.?|table|chart|map|exhibit|graph|diagram|plate|scheme|box)\s*"
    r"(\d+[a-z]?|[ivx]+)\b", re.IGNORECASE)
EQNUM_RE    = re.compile(r"^\s*[\(\[]\s*\d{1,3}[a-z]?(\.\d+)?\s*[\)\]]\s*$")
PAGENUM_RE  = re.compile(r"^\s*\d{1,4}\s*$")
FOOTNOTE_START_RE = re.compile(r"^\s*\d{1,3}[\.\)]\s+\S")
NOTE_LEAD_RE = re.compile(r"^\s*(?:\x01(\d{1,3})\x02|(\d{1,3})[\.\)]?(?=\s+\S))\s*")
SUP_RE = re.compile(r"\x01(\d{1,3})\x02")
NOTES_HEAD_RE = re.compile(r"^\s*(notes|endnotes|notes and references|references and notes|notes to chapter.*)\s*$",
                           re.IGNORECASE)
CHAPTER_LABEL_RE = re.compile(r"^\s*[\[\(]?\s*(\d{1,2}|[IVXL]{1,6})\s*[\]\)]?\.?\s*$|^\s*(chapter|part)\s+[\w\-]+\s*$",
                              re.IGNORECASE)
PAGE_TOKEN = "[[PAGE:{}]]"

# 수식 전용 폰트 (TeX Computer Modern, MathType, STIX, Cambria Math 등)
MATH_FONT_RE = re.compile(
    r"(cmsy|cmmi|cmex|cmr\d|cmbx|msam|msbm|euler|eufm|eusm|rsfs|wasy|stmary|"
    r"txsy|txmi|txex|pxsy|pxmi|mtmi|mtsy|mtex|mt-?extra|"
    r"cambria\s*math|stix.*math|xits.*math|latinmodernmath|lmmath|"
    r"asana.*math|tex.*gyre.*math|math(pi|ematical)|symbol$|symbolmt)",
    re.IGNORECASE)

_MATH_CHARS = set(
    "∑∏∫∬∭∮√∞≤≥≠≈≡±∓×÷·∂∇∆∈∉∋⊂⊃⊆⊇∪∩∧∨¬→←⇒⇐⇔↔∀∃∄∅∝∠⊥∥′″‴⟨⟩〈〉"
    "αβγδεζηθικλμνξοπρστυφχψωΓΔΘΛΞΠΣΦΨΩ"
    "⁰¹²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉ⁿ"
    "ℝℕℤℚℂℙℓℏ"
)
_MATH_OPS = set("=+−<>/^_|{}[]")

SENT_END_RE   = re.compile(r"[.!?:;…]\s*[\"”’')\]]*\s*$")
PARA_START_RE = re.compile(r'^[\"“‘\'(\[A-Z0-9#*•\-–—]')

OCR_FONTS = ("glyphlessfont", "code2000", "ocr", "invisible")


# ──────────────────────────────────────────────
# 소형 유틸
# ──────────────────────────────────────────────
def _rect(bbox):
    return fitz.Rect(bbox)


def _span_text(line):
    return "".join(s.get("text", "") for s in line.get("spans", []))


def _block_lines_text(block):
    return [_span_text(l) for l in block.get("lines", [])]


def _span_text_sup(line):
    """위첨자 숫자 span 을 \x01N\x02 로 감싼 줄 텍스트 (각주 콜아웃/각주 번호 식별용)."""
    spans = [s for s in line.get("spans", []) if s.get("text", "")]
    if not spans:
        return ""
    sizes = [s.get("size", 0) for s in spans if s.get("text", "").strip()]
    ref = max(sizes) if sizes else 0
    out = []
    for s in spans:
        t = s.get("text", "")
        ts = t.strip()
        is_sup = bool(re.fullmatch(r"\d{1,3}", ts)) and (
            (s.get("flags", 0) & 1) or (ref > 0 and s.get("size", ref) <= ref * 0.80))
        if is_sup:
            out.append("\x01" + ts + "\x02")
        else:
            out.append(t)
    return "".join(out)


def _block_lines_text_sup(block):
    return [_span_text_sup(l) for l in block.get("lines", [])]


def _block_text(block):
    return "\n".join(_block_lines_text(block))


def _block_is_bold(block):
    n = 0; b = 0
    for l in block.get("lines", []):
        for s in l.get("spans", []):
            t = s.get("text", "").strip()
            if not t:
                continue
            n += len(t)
            if (s.get("flags", 0) & 16) or "bold" in s.get("font", "").lower():
                b += len(t)
    return n > 0 and b / n > 0.7


def _block_is_italic(block):
    n = 0; k = 0
    for l in block.get("lines", []):
        for s in l.get("spans", []):
            t = s.get("text", "").strip()
            if not t:
                continue
            n += len(t)
            if (s.get("flags", 0) & 2) or "italic" in s.get("font", "").lower():
                k += len(t)
    return n > 0 and k / n > 0.7


def _dehyphenate_join(lines):
    """블록 내 줄들을 하나의 문자열로. 줄끝 하이픈은 다음 줄이 소문자로 시작하면 결합."""
    out = ""
    for ln in lines:
        ln = ln.replace("\x00", "").rstrip()
        if not ln.strip():
            continue
        if not out:
            out = ln.strip()
            continue
        nxt = ln.lstrip()
        if out.endswith("-") and nxt[:1].islower():
            out = out[:-1] + nxt                      # 'inter-' + 'national'
        elif out.endswith("-") and nxt[:1].isupper():
            out = out + nxt                           # 'Anglo-' + 'American' (하이픈 유지)
        else:
            out = out + " " + nxt
    return re.sub(r"[ \t]{2,}", " ", out).strip()


def _norm_header(s: str) -> str:
    s = re.sub(r"\d+", "", s.lower())
    return re.sub(r"\s+", " ", s).strip()


def _is_caption_text(text: str, n_lines: int = 1) -> bool:
    """'Figure 3. ...' / 'TABLE 2 ...' 형태의 캡션 블록 (본문 문장 'Figure 3 shows...' 는 길이로 배제)."""
    return bool(CAPTION_RE.match(text)) and len(text) <= 320 and n_lines <= 4


def _wordness(text: str) -> float:
    """토큰 중 정상 영단어(3자 이상 알파벳) 비율. OCR 깨진 도식/표 본문은 낮다."""
    toks = text.split()
    if not toks:
        return 0.0
    good = sum(1 for t in toks if re.match(r"^[A-Za-z][a-z]{2,}[,.;:'\"\)\]]*$", t))
    return good / len(toks)


def _looks_like_table_or_garbage(text: str, n_lines: int) -> bool:
    if _wordness(text) < 0.55:
        return True
    lines = [l for l in text.split("\n") if l.strip()]
    if n_lines >= 3 and lines and max(len(l) for l in lines) <= 45:
        return True
    return False


# ──────────────────────────────────────────────
# 블록 메트릭 (폰트 크기 / OCR 대응 유효 크기)
# ──────────────────────────────────────────────
def _block_metrics(block, ocr: bool):
    """(eff_size, first_x0, last_line_width, n_lines)

    eff_size:
      - 일반 PDF: 글자수 가중 평균 span size
      - OCR PDF : 여러 줄이면 줄 간격(pitch), 한 줄이면 줄 높이*1.18
    """
    lines = [l for l in block.get("lines", []) if _span_text(l).strip()]
    if not lines:
        return 0.0, block["bbox"][0], 0.0, 0
    if not ocr:
        tot = 0; acc = 0.0
        for l in lines:
            for s in l.get("spans", []):
                t = s.get("text", "").strip()
                if t:
                    tot += len(t); acc += s["size"] * len(t)
        eff = acc / tot if tot else 0.0
    else:
        if len(lines) >= 2:
            ys = [l["bbox"][1] for l in lines]
            pitches = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
            pitches = [p for p in pitches if p > 0]
            eff = statistics.median(pitches) if pitches else (lines[0]["bbox"][3] - lines[0]["bbox"][1]) * 1.18
        else:
            eff = (lines[0]["bbox"][3] - lines[0]["bbox"][1]) * 1.18
    first_x0 = lines[0]["bbox"][0]
    last_w = lines[-1]["bbox"][2] - lines[-1]["bbox"][0]
    return eff, first_x0, last_w, len(lines)


def _page_is_ocr(page, d) -> bool:
    """스캔+OCR 페이지 판정: OCR 전용 폰트가 다수이거나, 페이지 대부분을 덮는 이미지 위에 텍스트가 있는 경우."""
    tot = 0; ocr = 0
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for l in b.get("lines", []):
            for s in l.get("spans", []):
                n = len(s.get("text", ""))
                tot += n
                if any(k in s.get("font", "").lower() for k in OCR_FONTS):
                    ocr += n
    if tot == 0:
        return False
    if ocr / tot > 0.5:
        return True
    parea = page.rect.get_area()
    for b in d.get("blocks", []):
        if b.get("type") == 1:
            r = fitz.Rect(b["bbox"]) & page.rect
            if r.get_area() > 0.6 * parea:
                return True
    return False


# ──────────────────────────────────────────────
# 수식 판정
# ──────────────────────────────────────────────
def _math_stats(block):
    total = 0; mfont = 0; msym = 0; mops = 0; italic_single = 0; tokens = 0
    for l in block.get("lines", []):
        for s in l.get("spans", []):
            ts = s.get("text", "").strip()
            if not ts:
                continue
            total += len(ts)
            if MATH_FONT_RE.search(s.get("font", "")):
                mfont += len(ts)
            msym += sum(1 for ch in ts if ch in _MATH_CHARS)
            mops += sum(1 for ch in ts if ch in _MATH_OPS)
            for tok in ts.split():
                tokens += 1
                if len(tok) == 1 and tok.isalpha() and (s.get("flags", 0) & 2):
                    italic_single += 1
    return total, mfont, msym, mops, italic_single, tokens


def _is_display_equation(block, col_width):
    text = _block_text(block).strip()
    if not text:
        return False
    total, mfont, msym, mops, ital1, tokens = _math_stats(block)
    if total == 0 or total > 450:
        return False
    if CAPTION_RE.match(text):
        return False
    words = re.findall(r"[A-Za-z]{4,}", text)
    long_word_ratio = sum(len(w) for w in words) / total
    if long_word_ratio > 0.55:
        return False
    f_font = mfont / total
    f_sym  = msym / total
    f_ops  = mops / total
    f_ital = (ital1 / tokens) if tokens else 0
    has_eqnum = bool(re.search(r"[\(\[]\s*\d{1,3}[a-z]?\s*[\)\]]\s*$", text))
    bw = block["bbox"][2] - block["bbox"][0]
    narrow = col_width > 0 and bw < 0.92 * col_width

    score = 3.0 * f_font + 4.0 * f_sym + 1.5 * f_ops + 1.5 * f_ital
    if has_eqnum:
        score += 0.6
    if narrow:
        score += 0.15
    lines = [t for t in _block_lines_text(block) if t.strip()]
    if len(lines) >= 2 and max(len(t) for t in lines) <= 40:
        score += 0.2
    return score >= 0.75


# ──────────────────────────────────────────────
# 항목
# ──────────────────────────────────────────────
class _Item:
    __slots__ = ("kind", "text", "rect", "heading", "para_end", "para_start",
                 "full_last", "first_indent", "marker", "page", "size", "bold", "italic")
    # kind: "text" | "marker" | "page"

    def __init__(self, kind, text="", rect=None, heading=0, marker=None, page=0):
        self.kind = kind            # "text" | "marker"
        self.text = text
        self.rect = rect
        self.heading = heading
        self.marker = marker
        self.page = page
        self.para_end = True
        self.para_start = True
        self.full_last = False
        self.first_indent = False
        self.size = 0.0
        self.bold = False
        self.italic = False


def _detect_body_size(pairs):
    c = Counter()
    for sz, n in pairs:
        c[round(sz, 1)] += n
    if not c:
        return 10.0
    top = c.most_common(3)
    return max(sz for sz, _ in top)


def _line_y_overlap(a, b):
    return min(a.y1, b.y1) - max(a.y0, b.y0)


class PdfExtractor:
    def __init__(self, path, assets_dir, strip_footnotes=True,
                 extract_figures=True, extract_tables=True, extract_equations=True,
                 fig_dpi=200, eq_dpi=300, footnote_mode=None, page_markers=True,
                 first_page=None):
        self.path = Path(path)
        self.assets_dir = Path(assets_dir)
        # footnote_mode: strip(제거) | commentary(논평성만) | all(전체) — 하단 각주 영역은 항상 본문에서 분리
        if footnote_mode not in ("strip", "commentary", "all"):
            footnote_mode = "strip" if strip_footnotes else "all"
        self.footnote_mode = footnote_mode
        self.collect_notes = footnote_mode != "strip"
        self.strip_footnotes = True
        self.page_markers = page_markers
        self.first_page = first_page
        self.extract_figures = extract_figures
        self.extract_tables = extract_tables
        self.extract_equations = extract_equations
        self.fig_dpi = fig_dpi
        self.eq_dpi = eq_dpi
        self.assets = {}
        self.stats = {"figures": 0, "tables": 0, "equations": 0,
                      "merged_pages": 0, "ocr_pages": 0, "skipped_pages": 0,
                      "notes": [], "page_numbers": [], "page_source": "", "page_words": {}}

    # ── 진입점 ──
    def run(self):
        doc = fitz.open(str(self.path))
        pages = self._select_pages(doc)
        if not pages:
            return "", {}

        if self.assets_dir.exists():
            shutil.rmtree(self.assets_dir, ignore_errors=True)
        self.assets_dir.mkdir(parents=True, exist_ok=True)

        # 텍스트 레이어가 사실상 없는 순수 스캔 → 비전(이미지) 모드용 페이지 렌더
        page_chars = [len(p.get_text().strip()) for p in pages]
        total_chars = sum(page_chars)
        if total_chars < max(200, 8 * len(pages)):
            self.stats["no_text_layer"] = True
            self.stats["vision_pages"] = self._render_pages_for_vision(pages)
            self._write_manifest()
            return "", self.assets

        # 혼합 PDF: 일부 쪽만 텍스트 레이어가 없는 경우(앞 장만 OCR 된 책 등) → 그 쪽들은 비전 처리
        scan_flags = [c < 120 and self._mostly_image(p) for p, c in zip(pages, page_chars)]

        dicts = [p.get_text("dict") for p in pages]
        ocr_flags = [_page_is_ocr(p, d) for p, d in zip(pages, dicts)]
        self.stats["ocr_pages"] = sum(ocr_flags)

        # 문서 전체 본문 크기 (유효 크기 기준, 글자수 가중)
        pairs = []
        indent_votes = []
        for d, ocr in zip(dicts, ocr_flags):
            for b in d.get("blocks", []):
                if b.get("type") != 0:
                    continue
                n = len(_block_text(b).strip())
                if n == 0:
                    continue
                eff, _, _, nl = _block_metrics(b, ocr)
                if eff > 0:
                    pairs.append((eff, n))
        if not pairs:
            return "", {}
        self.body_size = _detect_body_size(pairs)
        self.header_norms = self._collect_running_headers(pages, dicts)
        self.title_done = False
        self.page_nums = self._detect_page_numbers(pages, dicts)
        self.stats["page_numbers"] = list(self.page_nums)
        self.page_notes = []          # [(page_index, [line_sup, ...]), ...]

        page_items = []
        runs = []                     # 비전 처리할 연속 스캔 쪽 묶음
        for pno, (page, d, ocr) in enumerate(zip(pages, dicts, ocr_flags)):
            if scan_flags[pno]:
                if pno > 0 and scan_flags[pno - 1] and runs:
                    runs[-1].append(pno)
                    page_items.append([])
                else:
                    runs.append([pno])
                    page_items.append([_Item("vision", f"[[VISION:{len(runs) - 1}]]", None, page=pno)])
                continue
            page_items.append(self._process_page(page, d, pno, ocr))
        if runs:
            self.stats["hybrid"] = True
            hinted = self.stats.get("page_source") in ("printed", "manual")
            self.stats["vision_runs"] = [
                self._render_pages_for_vision([pages[k] for k in run],
                                              labels=[self.page_nums[k] for k in run] if hinted else None,
                                              start=sum(len(r) for r in runs[:ri]) + 1)
                for ri, run in enumerate(runs)]
            self.stats["vision_page_count"] = sum(len(r) for r in runs)

        self._fix_headings(page_items)
        self._extract_endnotes(page_items)

        # 쪽별 본문 단어 수 (QA: 번역 누락 탐지용)
        for pno, pi in enumerate(page_items):
            if scan_flags[pno]:
                continue
            w = sum(len(SUP_RE.sub("", it.text).split()) for it in pi if it.kind == "text")
            self.stats["page_words"][str(self.page_nums[pno])] = w

        if self.collect_notes:
            self._link_endnotes(page_items)
            self._resolve_notes(page_items, ocr_flags)
        else:
            for pi in page_items:
                for it in pi:
                    if it.kind == "text":
                        it.text = SUP_RE.sub(lambda m: m.group(1), it.text)

        # 문단 들여쓰기 스타일 여부 (전체 텍스트 항목 중 30% 이상 들여쓰기)
        texts = [it for pi in page_items for it in pi if it.kind == "text" and not it.heading]
        self.indent_style = bool(texts) and sum(1 for it in texts if it.first_indent) / len(texts) >= 0.3

        page_items = [self._merge_items(pi, same_page=True) for pi in page_items]
        items = self._merge_across_pages(page_items)
        text = self._render(items)
        self._write_manifest()
        return text, self.assets

    # ── 0) 페이지 선택 (HeinOnline 메타 시트 / JSTOR 표지 제외) ──
    def _select_pages(self, doc):
        keep = []
        for page in doc:
            low = page.get_text().lower()
            if "heinonline" in low and ("citations:" in low or "date downloaded" in low):
                self.stats["skipped_pages"] += 1
                continue
            if "jstor" in low and ("stable url" in low or "accessibility support" in low
                                   or "jstor is a not-for-profit" in low):
                self.stats["skipped_pages"] += 1
                continue
            keep.append(page)
        return keep

    # ── 러닝 헤더/푸터: 상/하단 영역에서 반복되는 정규화 텍스트 ──
    def _collect_running_headers(self, pages, dicts):
        c = Counter()
        for page, d in zip(pages, dicts):
            h = page.rect.height
            seen = set()
            for b in d.get("blocks", []):
                if b.get("type") != 0:
                    continue
                y0, y1 = b["bbox"][1], b["bbox"][3]
                if not (y1 < h * 0.10 or y0 > h * 0.90):
                    continue
                t = _dehyphenate_join(_block_lines_text(b))
                if not t or len(t) > 120:
                    continue
                n = _norm_header(t)
                if n and n not in seen:
                    seen.add(n)
                    c[n] += 1
        thresh = 2 if len(pages) <= 3 else 3
        return {k for k, v in c.items() if v >= thresh}

    def _is_running_header(self, block, page, text):
        h = page.rect.height
        y0, y1 = block["bbox"][1], block["bbox"][3]
        if not ((y1 < h * 0.10) or (y0 > h * 0.90)):
            return False
        if not text or PAGENUM_RE.match(text):
            return True
        if len(text) <= 120 and _norm_header(text) in self.header_norms:
            return True
        if len(text) <= 100 and not re.search(r"[.!?;]", text) \
                and re.search(r"(^\d{1,4}\s|\s\d{1,4}$)", text):
            return True
        low = text.lower()
        if "downloaded from" in low or "jstor" in low or "all use subject" in low:
            return True
        return False

    # ── 1) 페이지 처리 ──
    def _process_page(self, page, d, pno, ocr):
        page_rect = page.rect
        ph = page_rect.height
        pw = page_rect.width
        blocks = d.get("blocks", [])
        text_blocks = [b for b in blocks if b.get("type") == 0 and _block_text(b).strip()]
        metrics = [_block_metrics(b, ocr) for b in text_blocks]

        # 컬럼 폭 / 좌측 경계 추정 (본문 크기 블록 기준)
        body_blocks = [i for i, b in enumerate(text_blocks)
                       if abs(metrics[i][0] - self.body_size) < max(1.0, self.body_size * 0.1)
                       and len(_block_text(b)) > 150]
        widths = [text_blocks[i]["bbox"][2] - text_blocks[i]["bbox"][0] for i in body_blocks]
        col_width = max(widths) if widths else pw * 0.8
        x0s = [round(l["bbox"][0] / 2) * 2 for i in body_blocks
               for l in text_blocks[i]["lines"][1:]]      # 첫 줄 제외(들여쓰기)
        if not x0s:
            x0s = [round(l["bbox"][0] / 2) * 2 for i in body_blocks for l in text_blocks[i]["lines"]]
        col_x0 = Counter(x0s).most_common(1)[0][0] if x0s else None

        # ── 도식/표 영역 ──
        regions = []
        if self.extract_figures:
            for r in self._image_rects(page, page_rect):
                regions.append([r, "fig"])
            for r in self._drawing_rects(page, page_rect, text_blocks, metrics):
                regions.append([r, "fig"])
        if self.extract_tables:
            for r in self._table_rects(page, page_rect):
                regions.append([r, "tbl"])
        regions = self._merge_regions(regions)

        # 영역 내부 텍스트 블록(축 레이블, 표 본문 등) 흡수 — 캡션은 제외
        for reg in regions:
            for b in text_blocks:
                br = _rect(b["bbox"])
                t = _dehyphenate_join(_block_lines_text(b))
                if _is_caption_text(t, 4) and len(t) > 12:
                    continue
                if br.get_area() > 0 and (br & reg[0]).get_area() / br.get_area() > 0.5:
                    reg[0] = reg[0] | br
        regions = self._merge_regions(regions)
        # 영역 바로 위/아래에 붙은 짧은 제목 줄(서브피겨 제목, 축 제목)도 흡수
        for reg in regions:
            for b in text_blocks:
                br = _rect(b["bbox"])
                t = _dehyphenate_join(_block_lines_text(b))
                if not t or len(t) > 60 or _is_caption_text(t, 4) or br.width <= 0:
                    continue
                hx = min(br.x1, reg[0].x1) - max(br.x0, reg[0].x0)
                if hx < 0.8 * br.width:
                    continue
                gap_above = reg[0].y0 - br.y1
                gap_below = br.y0 - reg[0].y1
                if (0 <= gap_above <= 14) or (0 <= gap_below <= 14):
                    reg[0] = reg[0] | br
        regions = [r for r in regions
                   if r[0].get_area() < 0.75 * page_rect.get_area()
                   and r[0].width >= 40 and r[0].height >= 25]

        # 캡션 기준 영역 보완: 스캔(OCR) 도식, 괘선 없는 표 등
        if self.extract_figures or self.extract_tables:
            extra = self._caption_anchored_regions(page, text_blocks, metrics, regions,
                                                   ocr, col_x0, col_width)
            if extra:
                regions = self._merge_regions([(r, k) for r, k in regions] + extra)
                regions = [r for r in regions if r[0].get_area() < 0.75 * page_rect.get_area()]

        def in_region(b):
            br = _rect(b["bbox"])
            if br.get_area() <= 0:
                return None
            for reg in regions:
                if (br & reg[0]).get_area() / br.get_area() > 0.5:
                    return reg
            return None

        self._col_x_right = max((b["bbox"][2] for b in text_blocks), default=pw)
        self._page_ocr = ocr

        # ── 수식 / 캡션 플래그 ──
        eq_flags = {}
        if self.extract_equations and not ocr:
            for i, b in enumerate(text_blocks):
                if in_region(b) is None:
                    eq_flags[i] = _is_display_equation(b, col_width)
        elif self.extract_equations and ocr:
            # OCR: 폰트 정보가 없으므로 '정상 단어가 거의 없고 연산자가 있는 블록'을 수식으로 판정
            for i, b in enumerate(text_blocks):
                if in_region(b) is None:
                    t = _dehyphenate_join(_block_lines_text(b))
                    total = len(t)
                    if 8 <= total <= 600 and not self._is_running_header(b, page, t):
                        ops = sum(1 for ch in t if ch in "=+−<>/()[]{}^_|≤≥≠∑∫")
                        f_sym = sum(1 for ch in t if ch in _MATH_CHARS) / total
                        has_stop = bool(re.search(r"\b(if|and|or|for|where|the|then|is|of|in|to|we|by|with|that|let)\b", t))
                        eq_flags[i] = (_wordness(t) <= 0.06 and ops >= 2 and not has_stop) or f_sym >= 0.15
        caption_flags = {i: _is_caption_text(_dehyphenate_join(_block_lines_text(b)), metrics[i][3])
                         for i, b in enumerate(text_blocks)}

        # 클립 시 이웃 텍스트 침범 방지용: 영역/수식 밖 텍스트의 '줄' 단위 사각형
        self._clip_neighbors = [
            fitz.Rect(l["bbox"])
            for i, b in enumerate(text_blocks)
            if in_region(b) is None and not eq_flags.get(i)
            for l in b.get("lines", []) if _span_text(l).strip()
        ]

        # ── 각주 경계 ──
        footnote_line = None
        if self.strip_footnotes:
            body = self.body_size
            cand_info = []
            for i, b in enumerate(text_blocks):
                if eq_flags.get(i) or caption_flags.get(i) or in_region(b):
                    continue
                t = _dehyphenate_join(_block_lines_text(b))
                if self._is_running_header(b, page, t):
                    continue
                cand_info.append((i, b["bbox"][1], metrics[i][0], metrics[i][3], len(t), t))

            def mean_eff_below(y):
                tot = 0; acc = 0.0
                for _, y0, eff, _, n, _ in cand_info:
                    if y0 >= y - 0.5 and n > 0:
                        tot += n; acc += eff * n
                return (acc / tot) if tot else body

            ys = []
            for i, y0, eff, nl, n, t in cand_info:
                if y0 <= ph * 0.5:
                    continue
                if not ocr:
                    # 일반 PDF: 폰트 크기가 신뢰 가능 → 본문보다 15% 이상 작은 블록
                    if eff < body * 0.85:
                        ys.append(y0)
                    elif FOOTNOTE_START_RE.match(t) and eff <= body * 0.97 and y0 > ph * 0.55:
                        ys.append(y0)
                else:
                    # OCR: 크기 대신 '번호. 로 시작 + 그 아래 전체가 본문보다 작음' 을 기준으로 판정
                    if FOOTNOTE_START_RE.match(t) and eff <= body * 1.02 and y0 > ph * 0.5 \
                            and not re.match(r"^\d{1,3}\.\s+[A-Z][a-z]+\s*$", t) \
                            and mean_eff_below(y0) < body * 0.93:
                        ys.append(y0)
                    elif eff < body * 0.80 and nl >= 2:
                        ys.append(y0)            # 여러 줄인데 줄 간격이 확연히 좁음
            footnote_line = min(ys) if ys else None
            fn_small = body * (0.97 if ocr else 0.85)

        # ── 항목 구성 (원본 블록 순서 유지) ──
        items = []
        region_marker_done = set()
        used_eqnum = set()
        eqnum_blocks = [i for i, b in enumerate(text_blocks)
                        if EQNUM_RE.match(_block_text(b).strip())]

        i = 0
        while i < len(text_blocks):
            b = text_blocks[i]
            br = _rect(b["bbox"])
            eff, first_x0, last_w, n_lines = metrics[i]
            text = _dehyphenate_join(_block_lines_text(b))

            reg = in_region(b)
            if reg is not None:
                if id(reg) not in region_marker_done:
                    region_marker_done.add(id(reg))
                    items.append(self._make_region_item(page, pno, reg))
                i += 1
                continue

            if i in used_eqnum or not text:
                i += 1
                continue

            if self._is_running_header(b, page, text):
                i += 1
                continue

            if footnote_line is not None and br.y0 >= footnote_line - 0.5 \
                    and not eq_flags.get(i) and not caption_flags.get(i) \
                    and (eff < fn_small or FOOTNOTE_START_RE.match(text)):
                if self.collect_notes:
                    self.page_notes.append((pno, [l for l in _block_lines_text_sup(b) if l.strip()]))
                i += 1
                continue

            # ── 디스플레이 수식 ──
            if eq_flags.get(i):
                rect = fitz.Rect(br)
                j = i
                while j + 1 < len(text_blocks) and eq_flags.get(j + 1):
                    nb = _rect(text_blocks[j + 1]["bbox"])
                    if nb.y0 - rect.y1 <= 10 and in_region(text_blocks[j + 1]) is None:
                        rect |= nb
                        j += 1
                    else:
                        break
                for k in eqnum_blocks:
                    if k in used_eqnum or i <= k <= j:
                        continue
                    nr = _rect(text_blocks[k]["bbox"])
                    if _line_y_overlap(nr, rect) > 0.4 * min(nr.height, rect.height):
                        rect |= nr
                        used_eqnum.add(k)
                if ocr:
                    # OCR bbox 는 수식 폭을 과소평가 → 컬럼 우측 끝까지 확장 (수식 번호 포함)
                    rect = fitz.Rect(min(rect.x0, (col_x0 or rect.x0) - 2), rect.y0 - 3,
                                     max(rect.x1, self._col_x_right + 2), rect.y1 + 3)
                items.append(self._make_equation_item(page, pno, rect))
                i = j + 1
                continue

            # ── 일반 텍스트 ──
            item = _Item("text", text, br, page=pno)
            item.size = eff
            item.bold = _block_is_bold(b)
            item.italic = _block_is_italic(b)
            nwords = len(text.split())
            is_short = len(text) <= 160 and nwords <= 20 and not SENT_END_RE.search(text)
            not_full = col_width > 0 and (br.width < 0.80 * col_width)
            if _is_caption_text(text, n_lines):
                pass                                   # 'Table 2.1 ...' 캡션은 헤딩이 아니다
            elif ((not self.title_done and pno == 0) or (eff >= self.body_size * 1.8 and br.y0 < ph * 0.45)) \
                    and eff >= self.body_size * 1.3 and len(text) <= 250 and n_lines <= 4:
                item.heading = 1
                self.title_done = True
            elif is_short and not_full and n_lines <= 3 and re.match(r"^[A-Z0-9\[\(]", text) and (
                    eff >= self.body_size * 1.10
                    or (_block_is_bold(b) and eff >= self.body_size * 0.95)
                    or re.match(r"^(\d+(\.\d+)*\.?|[IVX]+\.)\s+[A-Z]", text) and nwords <= 12):
                item.heading = 2
            item.para_end = bool(SENT_END_RE.search(text)) or item.heading > 0
            item.para_start = bool(PARA_START_RE.match(text))
            item.full_last = col_width > 0 and last_w >= 0.85 * col_width
            item.first_indent = col_x0 is not None and (first_x0 - col_x0) > 6
            if self.collect_notes and not item.heading:
                # 판정은 일반 텍스트로 끝냈으므로, 콜아웃 식별용 위첨자 표시 텍스트로 교체
                item.text = _dehyphenate_join(_block_lines_text_sup(b))
            items.append(item)
            i += 1

        # 텍스트 블록 없이 존재하는 영역 → 위치 기반 삽입
        for reg in regions:
            if id(reg) in region_marker_done:
                continue
            mi = self._make_region_item(page, pno, reg)
            idx = len(items)
            for k, it in enumerate(items):
                if it.rect is not None and it.rect.y0 >= reg[0].y1 - 5:
                    idx = k
                    break
            items.insert(idx, mi)

        items = self._attach_markers_to_captions(items)
        if self.page_markers:
            items.insert(0, _Item("page", PAGE_TOKEN.format(self.page_nums[pno]), None, page=pno))
        return items

    # ── 캡션 기준 영역 (스캔 도식 / 괘선 없는 표) ──
    def _caption_anchored_regions(self, page, text_blocks, metrics, regions, ocr, col_x0, col_width):
        ph = page.rect.height
        pw = page.rect.width
        out = []
        infos = []
        for i, b in enumerate(text_blocks):
            t = _dehyphenate_join(_block_lines_text(b))
            r = _rect(b["bbox"])
            hdr = self._is_running_header(b, page, t)
            infos.append((i, b, t, r, metrics[i], hdr))
        content = [x for x in infos if not x[5]]
        if not content:
            return out
        top_y = min(x[3].y0 for x in content)
        bot_y = max(x[3].y1 for x in content)
        x_left = (col_x0 - 6) if col_x0 is not None else min(x[3].x0 for x in content) - 4
        x_right = max(x[3].x1 for x in content) + 6
        x_left = max(0, x_left); x_right = min(pw, x_right)

        def near_existing(r):
            for reg in regions:
                if (fitz.Rect(reg[0].x0, reg[0].y0 - 45, reg[0].x1, reg[0].y1 + 45)).intersects(r):
                    return True
            return False

        def is_paragraph(t, m):
            return (_wordness(t) >= 0.6 and len(t) >= 100 and m[3] >= 2
                    and abs(m[0] - self.body_size) < self.body_size * 0.25)

        def is_boundary(x):
            t, m, b = x[2], x[4], x[1]
            if is_paragraph(t, m) or FOOTNOTE_START_RE.match(t):
                return True
            # 헤딩(큰 글씨/볼드, 짧은 줄)도 경계
            if ocr:
                # OCR: 크기 정보가 불안정 → 매우 헤딩다운 짧은 줄만 경계로 인정
                return (len(t.split()) <= 5 and _wordness(t) >= 0.75 and m[3] == 1
                        and not re.search(r"[\d'’]", t) and m[0] >= self.body_size * 1.1)
            if len(t) <= 160 and m[3] <= 3 and _wordness(t) >= 0.6 and (
                    m[0] >= self.body_size * 1.1 or _block_is_bold(b)):
                return True
            return False

        for i, b, t, r, m, hdr in content:
            if hdr or not _is_caption_text(t, m[3]):
                continue
            if near_existing(r):
                continue
            is_table = t.strip().lower().startswith("table")

            def build(direction):
                # direction 'above': 캡션 위쪽 영역, 'below': 아래쪽
                if direction == "above":
                    cands = sorted([x for x in content if x[3].y1 <= r.y0 + 2 and x[0] != i],
                                   key=lambda x: -x[3].y1)
                    bound = top_y - 2
                    for x in cands:
                        if is_boundary(x):
                            bound = x[3].y1 + 2
                            break
                    reg = fitz.Rect(x_left, bound + 4, x_right, r.y0 - 6)
                else:
                    cands = sorted([x for x in content if x[3].y0 >= r.y1 - 2 and x[0] != i],
                                   key=lambda x: x[3].y0)
                    bound = bot_y + 2
                    for x in cands:
                        if is_boundary(x):
                            bound = x[3].y0 - 2
                            break
                    reg = fitz.Rect(x_left, r.y1 + 6, x_right, bound - 4)
                if reg.is_empty or reg.height < 40:
                    return None, 0
                inside = [x for x in content if x[0] != i and x[3].get_area() > 0
                          and (x[3] & reg).get_area() / x[3].get_area() > 0.5]
                garbage = sum(1 for x in inside if _looks_like_table_or_garbage(x[2], x[4][3]))
                # 캡션이 아닌 다른 캡션이 영역에 들어오면 무효
                if any(_is_caption_text(x[2], x[4][3]) for x in inside):
                    return None, 0
                text_h = sum(x[3].height for x in inside)
                empty_h = reg.height - text_h
                score = garbage * 2 + (1 if (ocr and empty_h >= 50) else 0)
                if not ocr and garbage == 0:
                    return None, 0
                if ocr and garbage == 0 and empty_h < 50:
                    return None, 0
                return reg, score

            first, second = ("below", "above") if is_table else ("above", "below")
            r1, s1 = build(first)
            r2, s2 = build(second)
            chosen = None
            if r1 is not None and (r2 is None or s1 >= s2):
                chosen = r1
            elif r2 is not None:
                chosen = r2
            if chosen is not None:
                out.append((chosen, "tbl" if is_table else "fig"))
        return out

    # ── 영역 감지 ──
    def _image_rects(self, page, page_rect):
        rects = []
        try:
            infos = page.get_image_info()
        except Exception:
            infos = []
        ph = page_rect.height
        for info in infos:
            r = fitz.Rect(info["bbox"]) & page_rect
            if r.is_empty or r.width < 30 or r.height < 20:
                continue
            if r.get_area() > 0.65 * page_rect.get_area():
                continue                                    # 스캔 배경
            if r.y1 < ph * 0.08 or r.y0 > ph * 0.92:
                continue                                    # 헤더/푸터 로고
            rects.append(r)
        return rects

    def _drawing_rects(self, page, page_rect, text_blocks, metrics):
        rects = []
        try:
            clusters = page.cluster_drawings(drop_size=3, x_tolerance=6, y_tolerance=6)
        except Exception:
            return rects
        ph = page_rect.height
        for r in clusters:
            r = fitz.Rect(r) & page_rect
            if r.is_empty or r.width < 60 or r.height < 40:
                continue
            if r.get_area() > 0.80 * page_rect.get_area():
                continue
            if r.y1 < ph * 0.10 or r.y0 > ph * 0.90:
                continue                                    # 헤더/푸터 장식
            body_chars = 0
            for b, m in zip(text_blocks, metrics):
                br = _rect(b["bbox"])
                if br.get_area() > 0 and (br & r).get_area() > 0.5 * br.get_area() \
                        and abs(m[0] - self.body_size) < 0.8:
                    body_chars += len(_block_text(b))
            if body_chars > 400:
                continue                                    # 텍스트 박스 배경
            rects.append(r)
        return rects

    def _table_rects(self, page, page_rect):
        rects = []
        try:
            tabs = page.find_tables()
        except Exception:
            return rects
        for t in getattr(tabs, "tables", []):
            try:
                if t.row_count < 2 or t.col_count < 2:
                    continue
                r = fitz.Rect(t.bbox) & page_rect
            except Exception:
                continue
            if r.is_empty or r.width < 60 or r.height < 25:
                continue
            if r.get_area() > 0.85 * page_rect.get_area():
                continue
            rects.append(r)
        return rects

    def _merge_regions(self, regions):
        out = []
        for rect, kind in regions:
            rect = fitz.Rect(rect)
            merged = False
            for o in out:
                grown = fitz.Rect(o[0].x0 - 12, o[0].y0 - 12, o[0].x1 + 12, o[0].y1 + 12)
                if grown.intersects(rect):
                    o[0] = o[0] | rect
                    if kind == "tbl":
                        o[1] = "tbl"
                    merged = True
                    break
            if not merged:
                out.append([rect, kind])
        if len(out) < len(regions):
            return self._merge_regions([(r, k) for r, k in out])
        return out

    def _attach_markers_to_captions(self, items):
        markers = [it for it in items if it.kind == "marker" and it.marker["kind"] != "eq"]
        result = list(items)
        for m in markers:
            mr = m.rect
            best = None; best_d = 1e9
            for it in result:
                if it.kind != "text" or it.rect is None or not _is_caption_text(it.text, 4):
                    continue
                dy = min(abs(it.rect.y0 - mr.y1), abs(mr.y0 - it.rect.y1))
                hx = min(it.rect.x1, mr.x1) - max(it.rect.x0, mr.x0)
                if dy <= 40 and hx > 0 and dy < best_d:
                    best, best_d = it, dy
            if best is None:
                continue
            result.remove(m)
            idx = result.index(best)
            if best.rect.y0 >= mr.y1 - 5:
                result.insert(idx, m)          # 캡션이 아래 → 마커를 캡션 앞
            else:
                result.insert(idx + 1, m)      # 캡션이 위 → 마커를 캡션 뒤
        return result

    # ── 이미지 렌더링 ──
    def _clip_png(self, page, rect, dpi, pad):
        """rect 주변 pad 만큼 여유를 두되, 위/아래 이웃 텍스트 블록과 겹치지 않도록 중간점에서 자른다."""
        top = rect.y0 - pad
        bottom = rect.y1 + pad
        rc = (rect.y0 + rect.y1) / 2.0
        for nb in getattr(self, "_clip_neighbors", []):
            if nb.x1 <= rect.x0 or nb.x0 >= rect.x1:
                continue                                  # 수평으로 안 겹침
            if nb.y1 <= rect.y0 + 1.5:                    # 위쪽 이웃 줄
                top = max(top, (nb.y1 + rect.y0) / 2.0)
            elif nb.y0 >= rect.y1 - 1.5:                  # 아래쪽 이웃 줄
                bottom = min(bottom, (rect.y1 + nb.y0) / 2.0)
            else:                                         # 부분 겹침 (OCR bbox 오차)
                ov = min(nb.y1, rect.y1) - max(nb.y0, rect.y0)
                if ov < 0.5 * nb.height:
                    if (nb.y0 + nb.y1) / 2.0 < rc:
                        top = max(top, nb.y1 - 0.5)
                    else:
                        bottom = min(bottom, nb.y0 + 0.5)
        if bottom - top < 4:
            top, bottom = rect.y0 - 1, rect.y1 + 1
        clip = fitz.Rect(rect.x0 - pad, top, rect.x1 + pad, bottom) & page.rect
        pix = page.get_pixmap(clip=clip, dpi=dpi, alpha=False)
        return pix, clip

    def _save_asset(self, page, pno, rect, kind, dpi, pad):
        key = {"fig": "figures", "tbl": "tables", "eq": "equations"}[kind]
        self.stats[key] += 1
        img_id = f"{kind}_p{pno+1:02d}_{self.stats[key]:02d}"
        pix, clip = self._clip_png(page, rect, dpi, pad)
        path = self.assets_dir / f"{img_id}.png"
        pix.save(str(path))
        self.assets[img_id] = {"path": str(path), "kind": kind, "page": pno + 1,
                               "width_pt": round(clip.width, 1), "height_pt": round(clip.height, 1)}
        return _Item("marker", f"[[IMG:{img_id}]]", fitz.Rect(rect),
                     marker=self.assets[img_id], page=pno)

    def _make_region_item(self, page, pno, reg):
        return self._save_asset(page, pno, reg[0], reg[1], self.fig_dpi, pad=4)

    def _make_equation_item(self, page, pno, rect):
        return self._save_asset(page, pno, rect, "eq", self.eq_dpi, pad=3)

    # ── 단락 병합 ──
    def _should_merge(self, prev, nxt, same_page):
        if prev.kind != "text" or nxt.kind != "text":
            return False
        if prev.heading or nxt.heading:
            return False
        if _is_caption_text(nxt.text, 4) or FOOTNOTE_START_RE.match(nxt.text):
            return False
        if _is_caption_text(prev.text, 4):
            return False
        if prev.text.endswith("-"):
            return True
        # 문장 중간에서 끊김
        if not prev.para_end and (not nxt.para_start or prev.full_last):
            return True
        # 문장은 끝났지만 마지막 줄이 꽉 차고 다음 블록이 소문자로 시작 (OCR 블록 분절)
        if prev.full_last and re.match(r"^[a-z]", nxt.text):
            return True
        # 문장 끝 + 꽉 찬 마지막 줄 + 다음 블록 들여쓰기 없음 (들여쓰기 스타일 문서)
        if prev.para_end and prev.full_last and not nxt.first_indent \
                and getattr(self, "indent_style", False):
            if same_page:
                gap = nxt.rect.y0 - prev.rect.y1 if (prev.rect and nxt.rect) else 0
                return gap < self.body_size * 0.6
            return True
        return False

    @staticmethod
    def _join(prev, nxt):
        a, b = prev.text, nxt.text
        if a.endswith("-") and b[:1].islower():
            prev.text = a[:-1] + b
        elif a.endswith("-"):
            prev.text = a + b
        else:
            prev.text = a + " " + b
        prev.para_end = nxt.para_end
        prev.full_last = nxt.full_last
        if prev.rect is not None and nxt.rect is not None and prev.page == nxt.page:
            prev.rect = prev.rect | nxt.rect
        else:
            prev.rect = nxt.rect
            prev.page = nxt.page

    def _merge_items(self, items, same_page):
        out = []
        for it in items:
            if out and self._should_merge(out[-1], it, same_page):
                self._join(out[-1], it)
            else:
                out.append(it)
        return out

    def _merge_across_pages(self, page_items):
        all_items = []
        for pitems in page_items:
            if not pitems:
                continue
            pitems = list(pitems)
            page_tok = None
            if pitems[0].kind == "page":
                page_tok = pitems.pop(0)
            if not pitems:
                if page_tok is not None:
                    all_items.append(page_tok)
                continue
            if all_items and self._should_merge(all_items[-1], pitems[0], same_page=False):
                nxt = pitems[0]
                if page_tok is not None:
                    # 쪽 경계에서 이어지는 단락: 문단 중간 인라인 쪽 마커
                    if all_items[-1].text.endswith("-"):
                        parts = nxt.text.split(" ", 1)
                        nxt.text = parts[0] + " " + page_tok.text + ((" " + parts[1]) if len(parts) > 1 else "")
                    else:
                        nxt.text = page_tok.text + " " + nxt.text
                self._join(all_items[-1], nxt)
                self.stats["merged_pages"] += 1
                all_items.extend(pitems[1:])
            else:
                if page_tok is not None:
                    all_items.append(page_tok)
                all_items.extend(pitems)
        return all_items

    # ── 원서 쪽번호 감지 ──
    def _detect_page_numbers(self, pages, dicts):
        """각 쪽의 인쇄된 쪽번호. 상/하단 영역 숫자로 (인쇄번호 - 쪽순번) 오프셋을 투표.
        first_page 가 지정되면 그 값을 첫 쪽 번호로 사용. 실패 시 PDF 쪽번호(1부터)."""
        n = len(pages)
        if isinstance(self.first_page, int) and self.first_page > 0:
            self.stats["page_source"] = "manual"
            return [self.first_page + k for k in range(n)]
        votes = Counter()
        for k, (page, d) in enumerate(zip(pages, dicts)):
            h = page.rect.height
            seen = set()
            for b in d.get("blocks", []):
                if b.get("type") != 0:
                    continue
                y0, y1 = b["bbox"][1], b["bbox"][3]
                if not (y1 < h * 0.12 or y0 > h * 0.88):
                    continue
                for ln in _block_lines_text(b):
                    t = ln.strip()
                    if not t or len(t) > 120:
                        continue
                    cands = []
                    if PAGENUM_RE.match(t):
                        cands.append(int(t))
                    m1 = re.match(r"^(\d{1,4})\s+\S", t)
                    m2 = re.search(r"\S\s+(\d{1,4})$", t)
                    if m1: cands.append(int(m1.group(1)))
                    if m2: cands.append(int(m2.group(1)))
                    for c in cands:
                        off = c - k
                        if c > 0 and off not in seen and not (1900 <= c <= 2100):
                            seen.add(off)
                            votes[off] += 1
        if votes:
            off, cnt = votes.most_common(1)[0]
            if cnt >= max(2, int(n * 0.3 + 0.5)) and off + 0 >= 1:
                self.stats["page_source"] = "printed"
                return [off + k for k in range(n)]
        self.stats["page_source"] = "pdf"
        return [p.number + 1 for p in pages]

    # ── 각주 수집 + 콜아웃 연결 ──
    def _resolve_notes(self, page_items, ocr_flags):
        notes = []                 # dict(id, num, page_idx, text)
        last_num = 0
        for pno, lines in self.page_notes:
            for ln in lines:
                m = NOTE_LEAD_RE.match(ln)
                num = None
                if m:
                    num = int(m.group(1) or m.group(2))
                    sup_lead = m.group(1) is not None
                    plausible = sup_lead or FOOTNOTE_START_RE.match(ln) or (0 < num - last_num <= 3)
                    if not plausible or not (num == 1 or 0 < num - last_num <= 5 or sup_lead):
                        num = None
                if num is not None:
                    body = ln[m.end():]
                    notes.append({"num": num, "page_idx": pno, "lines": [body]})
                    last_num = num
                elif notes:
                    notes[-1]["lines"].append(ln)            # 이전 각주(또는 이전 쪽 각주)의 연속
        for k, nt in enumerate(notes, start=1):
            nt["id"] = f"n{k}"
            txt = _dehyphenate_join(nt.pop("lines"))
            nt["text"] = SUP_RE.sub(lambda m: m.group(1), txt).strip()

        by_page = {}
        for nt in notes:
            by_page.setdefault(nt["page_idx"], {})[nt["num"]] = nt
        used = set()

        # 1) 위첨자 콜아웃 → 같은 쪽(없으면 인접 쪽) 각주
        for pno, pi in enumerate(page_items):
            for it in pi:
                if it.kind != "text":
                    continue
                def _sub(m, pno=pno):
                    num = int(m.group(1))
                    for q in (pno, pno - 1, pno + 1):
                        nt = by_page.get(q, {}).get(num)
                        if nt is not None and nt["id"] not in used:
                            used.add(nt["id"])
                            return "[^" + nt["id"] + "]"
                    return ""
                it.text = SUP_RE.sub(_sub, it.text)

        # 2) OCR 등 위첨자 정보가 없는 경우: 'word.15' 형태의 숫자 콜아웃 탐색
        for nt in notes:
            if nt["id"] in used:
                continue
            pi = page_items[nt["page_idx"]] if nt["page_idx"] < len(page_items) else []
            pat = re.compile(r"(?<=[A-Za-z\.,;:\"\u201d\u2019\)])" + str(nt["num"]) + r"(?=[\s\.,;:]|$)")
            for it in pi:
                if it.kind != "text" or it.heading:
                    continue
                m = pat.search(it.text)
                if m:
                    it.text = it.text[:m.start()] + "[^" + nt["id"] + "]" + it.text[m.end():]
                    used.add(nt["id"])
                    break

        # 3) 그래도 못 찾으면 그 쪽 마지막 본문 단락 끝에 붙인다
        for nt in notes:
            if nt["id"] in used:
                continue
            pi = page_items[nt["page_idx"]] if nt["page_idx"] < len(page_items) else []
            texts = [it for it in pi if it.kind == "text" and not it.heading]
            if texts:
                texts[-1].text = texts[-1].text.rstrip() + "[^" + nt["id"] + "]"
                used.add(nt["id"])

        self.stats["notes"] = (self.stats.get("notes") or []) + [
            {"id": nt["id"], "num": nt["num"], "page": self.page_nums[nt["page_idx"]], "text": nt["text"]}
            for nt in notes if nt["id"] in used and nt["text"]]

    # ── 출력 ──
    def _render(self, items):
        paras = []
        for it in items:
            if it.kind in ("marker", "page", "vision"):
                paras.append(it.text)
            elif it.heading == 1:
                paras.append("# " + it.text)
            elif it.heading == 2:
                paras.append("## " + it.text)
            elif it.heading >= 3:
                paras.append("### " + it.text)
            else:
                paras.append(it.text)
        text = "\n\n".join(p for p in paras if p.strip())
        return text.encode("utf-8", errors="ignore").decode("utf-8")

    @staticmethod
    def _mostly_image(page):
        try:
            parea = page.rect.get_area()
            for info in page.get_image_info():
                r = fitz.Rect(info["bbox"]) & page.rect
                if r.get_area() > 0.5 * parea:
                    return True
        except Exception:
            pass
        return False

    # ── 제목 블록 / 헤딩 수준 정리 ──
    def _fix_headings(self, page_items):
        body = self.body_size
        for pi in page_items:
            k = 0
            while k < len(pi) - 1:
                a, b = pi[k], pi[k + 1]
                if a.kind == "text" and b.kind == "text" and a.heading and b.heading:
                    # 장 번호 라벨('[2]', 'Chapter 3') + 제목 → 하나로
                    if CHAPTER_LABEL_RE.match(a.text) and not CHAPTER_LABEL_RE.match(b.text):
                        lab = re.sub(r"[\[\]\(\)\.]", "", a.text).strip()
                        b.text = (lab + ". " if re.fullmatch(r"\d{1,2}|[IVXL]{1,6}", lab) else lab + ": ") + b.text
                        b.heading = 1 if a.heading == 1 or b.size >= body * 1.3 else b.heading
                        if b.heading == 1:
                            self.title_done = True
                        pi.pop(k)
                        continue
                    # 여러 줄로 쪼개진 제목/헤딩 → 하나로
                    big = a.size >= body * 1.25 and b.size >= body * 1.25
                    similar = a.size > 0 and abs(a.size - b.size) / a.size < 0.15
                    gap = (b.rect.y0 - a.rect.y1) if (a.rect is not None and b.rect is not None) else 99
                    if big and similar and not SENT_END_RE.search(a.text) and gap < body * 1.2:
                        a.text = a.text.rstrip() + " " + b.text.lstrip()
                        a.rect = a.rect | b.rect
                        pi.pop(k + 1)
                        continue
                k += 1
            # 제목 바로 뒤의 저자 표시줄(대문자 이름 등)은 헤딩이 아니라 저자 줄
            for k, it in enumerate(pi):
                if it.kind == "text" and it.heading == 1 and k + 1 < len(pi):
                    nx = pi[k + 1]
                    if nx.kind == "text" and nx.heading and len(nx.text) <= 80 and len(nx.text.split()) <= 8 \
                            and not re.search(r"[\d\.!?:;]", nx.text) and not NOTES_HEAD_RE.match(nx.text) \
                            and (nx.text.isupper() or nx.size < it.size * 0.7):
                        nx.heading = 0
                        nx.para_end = True
                        nx.para_start = True
        # 헤딩 수준: 크기·굵기 기준으로 상위 스타일 = ##, 그 아래 = ###
        #   OCR 쪽의 유효 크기는 줄 간격 추정치라 들쭉날쭉하므로, 확연히 작을 때(<85%)만 ### 로 내린다.
        heads = [it for pi in page_items for it in pi if it.kind == "text" and it.heading == 2]
        if len(heads) >= 3:
            sizes = sorted(it.size for it in heads)
            top = sizes[int(len(sizes) * 0.75)]          # 상위 스타일 대표 크기(이상치 방지)
            any_bold = sum(1 for it in heads if it.bold) >= 2
            for it in heads:
                if it.size < top * 0.85 or (any_bold and not it.bold and it.size < top * 0.97):
                    it.heading = 3

    # ── 장 끝 미주(Notes) 절 ──
    def _extract_endnotes(self, page_items):
        """'Notes' 헤딩 이후 다음 장 제목 전까지를 본문에서 떼어 미주로 보관(strip 모드면 버림)."""
        self.endnote_sections = []          # [(chapter_seq, [line_sup...])]
        in_notes = False
        chapter = 0
        cur = None
        removed_pages = set()
        for pno, pi in enumerate(page_items):
            keep = []
            for it in pi:
                if it.kind == "vision":
                    in_notes = False
                    chapter += 1
                    keep.append(it); continue
                if it.kind == "text" and it.heading == 1:
                    in_notes = False
                    chapter += 1
                if it.kind == "text" and it.heading and NOTES_HEAD_RE.match(it.text):
                    in_notes = True
                    cur = (chapter, [])
                    self.endnote_sections.append(cur)
                    removed_pages.add(pno)
                    continue
                if in_notes and it.kind == "text":
                    cur[1].append(it.text)
                    removed_pages.add(pno)
                    continue
                if in_notes and it.kind == "marker":
                    continue
                keep.append(it)
            pi[:] = keep
        self.stats["endnote_pages"] = sorted(self.page_nums[p] for p in removed_pages)
        n = sum(len(sec[1]) for sec in self.endnote_sections)
        self.stats["endnote_blocks"] = n

    def _link_endnotes(self, page_items):
        """미주 항목을 번호로 파싱해 같은 장 본문의 콜아웃(위첨자 또는 'word.21' 꼴)과 연결."""
        if not getattr(self, "endnote_sections", None):
            return
        # 장별 본문 항목 목록
        chapters = {}
        chapter = 0
        for pi in page_items:
            for it in pi:
                if it.kind == "vision" or (it.kind == "text" and it.heading == 1):
                    chapter += 1
                if it.kind == "text":
                    chapters.setdefault(chapter, []).append(it)
        seq = len(self.stats.get("notes") or [])
        added = []
        for ch, blocks in self.endnote_sections:
            entries = []
            last = 0
            for blk in blocks:
                for piece in re.split(r"\s(?=(?:\x01\d{1,3}\x02|\d{1,3}\s+[A-Z\'\"\u2018\u201c(]))", blk):
                    m = NOTE_LEAD_RE.match(piece)
                    num = int(m.group(1) or m.group(2)) if m else None
                    if num is not None and 0 < num - last <= 5:
                        entries.append([num, piece[m.end():]])
                        last = num
                    elif entries:
                        entries[-1][1] += " " + piece
            body = chapters.get(ch, [])
            pos_item, pos_char = 0, 0
            for num, txt in entries:
                seq += 1
                nid = f"e{seq}"
                linked = False
                # 1) 위첨자 콜아웃
                for ii in range(pos_item, len(body)):
                    it = body[ii]
                    if it.heading:
                        continue
                    start = pos_char if ii == pos_item else 0
                    mm = re.compile("\x01" + str(num) + "\x02").search(it.text, start)
                    if not mm:
                        mm = re.compile(r"(?<=[A-Za-z\.,;:\"\u201d\u2019\)])" + str(num) + r"(?=[\s\.,;:]|$)").search(it.text, start)
                    if mm:
                        it.text = it.text[:mm.start()] + "[^" + nid + "]" + it.text[mm.end():]
                        pos_item, pos_char = ii, mm.start() + len(nid) + 3
                        linked = True
                        break
                if linked:
                    t = _dehyphenate_join([SUP_RE.sub(lambda m: m.group(1), txt)]).strip()
                    if t:
                        added.append({"id": nid, "num": num, "page": None, "text": t})
        self.stats["notes"] = (self.stats.get("notes") or []) + added
        self.stats["endnotes_linked"] = len(added)

    def _render_pages_for_vision(self, pages, dpi=150, labels=None, start=1):
        """스캔 PDF 페이지를 JPEG로 렌더 (비전 번역용). [{'label': n, 'pdf_page': i, 'path': ...}]
        labels 가 주어지면(혼합 PDF + 인쇄 쪽번호 확인) label = 원서 쪽번호, printed=True."""
        out = []
        pdir = self.assets_dir / "pages"
        pdir.mkdir(parents=True, exist_ok=True)
        for n, page in enumerate(pages, start=start):
            # 긴 변이 약 1600px 이하가 되도록 dpi 조정 (API 이미지 한도/비용 최적)
            long_pt = max(page.rect.width, page.rect.height)
            eff_dpi = min(dpi, int(1568 * 72 / long_pt)) if long_pt > 0 else dpi
            pix = page.get_pixmap(dpi=max(96, eff_dpi), alpha=False)
            path = pdir / f"page_{n:03d}.jpg"
            try:
                pix.save(str(path), jpg_quality=85)
            except TypeError:
                pix.save(str(path))
            rec = {"label": n, "pdf_page": page.number, "path": str(path)}
            if labels:
                rec["label"] = labels[n - start]
                rec["printed"] = True
            out.append(rec)
        return out

    def _write_manifest(self):
        try:
            (self.assets_dir / "manifest.json").write_text(
                json.dumps(self.assets, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass


def extract_pdf(path, assets_dir, **kw):
    ex = PdfExtractor(path, assets_dir, **kw)
    text, assets = ex.run()
    return text, assets, ex.stats


def load_manifest(assets_dir):
    p = Path(assets_dir) / "manifest.json"
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def clip_page_region(pdf_path, pdf_page, top_pct, bottom_pct, out_path, dpi=200):
    """스캔 PDF의 페이지에서 세로 퍼센트 구간을 전체 폭으로 잘라 저장 (비전 모드 도식 클립)."""
    doc = fitz.open(str(pdf_path))
    page = doc[pdf_page]
    r = page.rect
    y0 = r.height * max(0.0, (top_pct - 1.0)) / 100.0
    y1 = r.height * min(100.0, (bottom_pct + 1.0)) / 100.0
    if y1 - y0 < r.height * 0.02:
        y1 = min(r.height, y0 + r.height * 0.02)
    clip = fitz.Rect(r.width * 0.02, y0, r.width * 0.98, y1)
    pix = page.get_pixmap(clip=clip, dpi=dpi, alpha=False)
    pix.save(str(out_path))
    return clip
