# -*- coding: utf-8 -*-
r"""
translator_gui.pyw — Paper Translator GUI (v4)
==============================================
bat\Translator.bat 의 메뉴를 창(GUI)으로 옮긴 프런트엔드.

폴더 구조 (저장소 루트 = 번역 스크립트·설정 파일이 있는 곳)
  <루트>\*.py, extract.cfg, model.cfg, api_key.cfg, glossaries\   공통
  <루트>\bat\Translator.bat                                   콘솔 메뉴
  <루트>\gui\translator_gui.pyw, PaperTranslator.exe, ptworker.exe, _internal\   GUI

설계 원칙
  - 번역 파이프라인(translate_paper.py / batch_submit.py / batch_download.py)은 건드리지 않는다.
    GUI 는 설정 파일(extract.cfg / model.cfg / api_key.cfg)을 bat 과 같은 형식으로 쓰고,
    스크립트를 별도 프로세스로 실행해 출력(로그)을 실시간으로 보여 준다.
  - 따라서 bat 과 GUI 를 섞어 써도 설정·대기 중 배치가 서로 공유된다.
  - 실행 직전에 화면의 설정을 파일에 다시 써서, 화면에 보이는 값 = 실제 실행 값이 되도록 한다.

필요 패키지: customtkinter (필수), tkinterdnd2 (선택: 파일 끌어다 놓기)
실행: PaperTranslator.exe (build_exe.bat 으로 빌드)  또는  pythonw translator_gui.pyw
"""

import os
import re
import sys
import json
import queue
import shutil
import threading
import subprocess
import webbrowser
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont

try:
    import customtkinter as ctk
except ImportError:  # pragma: no cover - 설치 안내
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror(
        "Paper Translator",
        "customtkinter 패키지가 없습니다.\n\n"
        "build_exe.bat 을 실행하면 함께 설치됩니다.\n"
        "수동 설치: python -m pip install customtkinter tkinterdnd2")
    sys.exit(1)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    TkinterDnD = None
    DND_FILES = None
    HAS_DND = False


APP_VERSION = "v4 GUI 1.2"
FROZEN = bool(getattr(sys, "frozen", False))      # PyInstaller exe 로 실행 중인가
# APP_DIR : GUI 자신이 있는 폴더 (gui\ — exe·ptworker·아이콘·gui.cfg)
# BASE_DIR: 번역 스크립트(paper_utils.py)와 공통 설정이 있는 루트. 자기 폴더 -> 한 단계 위 순으로 찾는다.
APP_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent


def _find_root(start: Path) -> Path:
    for cand in (start, start.parent):
        if (cand / "paper_utils.py").exists():
            return cand
    return start.parent if start.name.lower() == "gui" else start


BASE_DIR = _find_root(APP_DIR)
WORKER_EXE = APP_DIR / "ptworker.exe"
APP_ICON = APP_DIR / "app.ico"
APP_ID = "PaperTranslator.GUI"
MODEL_CFG = BASE_DIR / "model.cfg"
EXTRACT_CFG = BASE_DIR / "extract.cfg"
KEY_CFG = BASE_DIR / "api_key.cfg"
GUI_CFG = APP_DIR / "gui.cfg"
BATCH_INFO = BASE_DIR / "batch_info.json"
GLOSS_DIR = BASE_DIR / "glossaries"
GLOSS_EXAMPLE = BASE_DIR / "glossary_example.txt"
QA_FILE = BASE_DIR / "last_qa.txt"

IS_WIN = os.name == "nt"
SUPPORTED_EXT = (".pdf", ".docx", ".md", ".txt")

# ──────────────────────────────────────────────
# 모델 / 옵션 표
# ──────────────────────────────────────────────
DEFAULT_MODEL = "claude-sonnet-5-5"
MODELS = [  # (model id, 표시 이름, 가격 in/out per 1M, 설명)
    ("claude-fable-5-1", "Fable 5.1", "$10 / $50", "최고 품질 · thinking 항상 켜짐"),
    ("claude-opus-5-5", "Opus 5.5", "$4 / $20", "매우 높은 품질 · thinking 항상 켜짐"),
    ("claude-sonnet-5-5", "Sonnet 5.5", "$2 / $10", "균형 · 기본값"),
    ("claude-haiku-4-5-20251001", "Haiku 4.5", "$1 / $5", "빠름·저렴 · effort 조절 없음"),
]
CUSTOM_LABEL = "직접 입력…"
# Translator.bat 안내문의 대략치: 30쪽(약 2.5만 단어) 논문 1편, 실시간 기준
COST_PER_30P = {"fable": 3.0, "opus": 1.2, "sonnet": 0.6, "haiku": 0.3}

EFFORTS = ["off", "low", "medium", "high", "xhigh"]
FOOTNOTE_OPTS = [("strip", "삭제"), ("commentary", "논평만"), ("all", "전체")]
FOOTNOTE_HELP = {
    "strip": "각주와 각주 번호를 모두 삭제",
    "commentary": "설명·논증이 있는 각주만 Word 각주로 번역\n(서지 정보는 영문 유지, 순수 서지 각주는 삭제)",
    "all": "모든 각주를 Word 각주로 번역",
}
OUTPUT_OPTS = [("docx", "DOCX"), ("pdf", "PDF"), ("docx+pdf", "DOCX+PDF"), ("md", "MD")]

# extract.cfg — Translator.bat 의 :savecfg 와 같은 키·순서
CFG_ORDER = ["footnotes", "strip_references", "output_format", "effort", "figures",
             "glossary", "page_markers", "first_page", "glossary_file"]
CFG_DEFAULTS = {"footnotes": "commentary", "strip_references": "true", "output_format": "docx",
                "effort": "low", "figures": "true", "glossary": "true", "page_markers": "true",
                "first_page": "auto", "glossary_file": "none"}
CFG_VALID = {"footnotes": ("strip", "commentary", "all"),
             "output_format": ("docx", "pdf", "docx+pdf", "md"),
             "effort": tuple(EFFORTS)}

ACTION_LABEL = {"translate": "즉시 번역", "submit": "배치 제출", "download": "배치 결과 받기"}

# 로그 해석용 정규식 (스크립트 출력 형식에 맞춤)
RE_STEP = re.compile(r"\[(\d+)/(\d+)\]")
RE_BATCH_PROGRESS = re.compile(r"진행:\s*(\d+)/(\d+)")
RE_SAVED = re.compile(r"^\s*(?:저장 완료|대신 Markdown 으로 저장):\s*(.+?)\s*$")
RE_TOTAL = re.compile(r"Total:\s*\$([\d.]+)")
RE_QA = re.compile(r"QA 점검:\s*(.+)$")
RE_BATCH_ID = re.compile(r"Batch ID:\s*(\S+)")


# ──────────────────────────────────────────────
# 설정 파일 입출력 (bat 호환)
# ──────────────────────────────────────────────
def _read_text(p: Path) -> str:
    for enc in ("utf-8-sig", "cp949"):
        try:
            return p.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    return p.read_text(encoding="utf-8", errors="replace")


def load_extract_cfg():
    """-> (values, extra_lines). 모르는 키·주석은 extra 로 보존한다."""
    vals = dict(CFG_DEFAULTS)
    extra = []
    legacy_strip = None
    seen = set()
    if EXTRACT_CFG.exists():
        for line in _read_text(EXTRACT_CFG).splitlines():
            s = line.strip()
            if not s:
                continue
            if s.startswith("#") or "=" not in s:
                extra.append(s)
                continue
            k, v = s.split("=", 1)
            k = k.strip().lower()
            v = v.strip()
            if k == "strip_footnotes":
                legacy_strip = v.lower() == "true"
                continue
            if k not in vals:
                extra.append(s)
                continue
            seen.add(k)
            if k == "glossary_file":
                vals[k] = v or "none"                 # 대소문자 보존
            elif k == "first_page":
                vals[k] = v if v.isdigit() else "auto"
            elif k in CFG_VALID:
                if v.lower() in CFG_VALID[k]:
                    vals[k] = v.lower()
            else:
                vals[k] = "true" if v.lower() == "true" else "false"
    if "footnotes" not in seen and legacy_strip is not None:
        vals["footnotes"] = "strip" if legacy_strip else "all"
    if "glossary_file" not in seen and (BASE_DIR / "user_glossary.txt").exists():
        vals["glossary_file"] = "user_glossary.txt"   # bat 과 같은 하위 호환
    return vals, extra


def save_extract_cfg(vals: dict, extra: list):
    lines = [f"{k}={vals[k]}" for k in CFG_ORDER] + list(extra)
    data = "\r\n".join(lines) + "\r\n"
    EXTRACT_CFG.write_bytes(data.encode("utf-8"))


def load_model() -> str:
    if MODEL_CFG.exists():
        m = _read_text(MODEL_CFG).strip().splitlines()
        if m and m[0].strip():
            return m[0].strip()
    return DEFAULT_MODEL


def save_model(model_id: str):
    MODEL_CFG.write_bytes((model_id.strip() + "\r\n").encode("ascii", errors="ignore"))


def load_key():
    """-> (key or None, from_env)"""
    env = os.environ.get("ANTHROPIC_API_KEY")
    if env:
        return env, True
    if KEY_CFG.exists():
        k = _read_text(KEY_CFG).strip()
        if k:
            return k, False
    return None, False


def mask_key(k):
    if not k:
        return "설정 안 됨"
    return "sk-ant-***" + k[-4:]


def load_gui_cfg() -> dict:
    try:
        return json.loads(GUI_CFG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_gui_cfg(d: dict):
    try:
        GUI_CFG.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def list_glossaries():
    GLOSS_DIR.mkdir(exist_ok=True)
    return sorted(p.name for p in GLOSS_DIR.glob("*.txt"))


def model_family(model_id: str):
    m = (model_id or "").lower()
    for fam in ("fable", "opus", "sonnet", "haiku"):
        if fam in m:
            return fam
    return None


def worker_command(script: str, args: list) -> list:
    """번역 스크립트 실행 명령.
    exe: ptworker.exe <script> ...  (같은 _internal 라이브러리 사용, 콘솔 exe 라 출력 파이프 안정)
    소스: python.exe -u <script> ..."""
    if FROZEN:
        return [str(WORKER_EXE), script] + list(args)
    return [console_python(), "-u", str(BASE_DIR / script)] + list(args)


def console_python() -> str:
    """pythonw 로 GUI 를 띄웠어도 자식 프로세스는 python.exe 로 (출력 파이프 안정)."""
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        cand = exe.with_name("python.exe")
        if cand.exists():
            return str(cand)
    return str(exe)


def open_path(p):
    p = str(p)
    try:
        if IS_WIN:
            os.startfile(p)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except Exception as e:
        messagebox.showerror("열기 실패", f"{p}\n\n{e}")


def reveal_in_folder(p):
    p = Path(p)
    if IS_WIN and p.exists():
        subprocess.Popen(["explorer", "/select,", str(p)])
    else:
        open_path(p.parent if p.is_file() else p)


def pick_font_family(candidates, fallback):
    try:
        fams = set(tkfont.families())
    except Exception:
        return fallback
    for c in candidates:
        if c in fams:
            return c
    return fallback


# ──────────────────────────────────────────────
# 창 클래스 (드래그 앤 드롭 가능하면 DnDWrapper 혼합)
# ──────────────────────────────────────────────
if HAS_DND:
    class _Base(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self.dnd_ok = True
            except Exception:
                self.dnd_ok = False
else:
    class _Base(ctk.CTk):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self.dnd_ok = False


class App(_Base):
    def __init__(self):
        self.gui_cfg = load_gui_cfg()
        ctk.set_appearance_mode(self.gui_cfg.get("appearance", "system"))
        ctk.set_default_color_theme("blue")
        super().__init__()

        self.title(f"Paper Translator  ·  {APP_VERSION}")
        if IS_WIN and APP_ICON.exists():
            try:
                self.iconbitmap(default=str(APP_ICON))
                # customtkinter 가 시작 직후 기본 아이콘으로 덮어쓰는 경우 대비
                self.after(250, lambda: self.iconbitmap(str(APP_ICON)))
            except Exception:
                pass
        self.geometry(self.gui_cfg.get("geometry", "1180x780"))
        self.minsize(980, 640)

        fam = pick_font_family(["Malgun Gothic", "맑은 고딕", "NanumGothic", "Noto Sans CJK KR"], "TkDefaultFont")
        mono = pick_font_family(["D2Coding", "NanumGothicCoding", "Malgun Gothic", "맑은 고딕",
                                 "NanumGothic", "Noto Sans Mono CJK KR"], "TkFixedFont")
        self.f_body = ctk.CTkFont(family=fam, size=13)
        self.f_small = ctk.CTkFont(family=fam, size=11)
        self.f_head = ctk.CTkFont(family=fam, size=13, weight="bold")
        self.f_title = ctk.CTkFont(family=fam, size=20, weight="bold")
        self.f_btn = ctk.CTkFont(family=fam, size=14, weight="bold")
        self.f_log = ctk.CTkFont(family=mono, size=12)

        # 상태
        self.cfg, self.cfg_extra = load_extract_cfg()
        self.model = load_model()
        self.proc = None
        self.proc_kind = None
        self.q = queue.Queue()
        self.outputs = []
        self.saw_pending = False
        self.auto_job = None
        self.fatal_stop = False
        self.last_dir = self.gui_cfg.get("last_dir", "")

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()
        self._load_into_widgets()
        self._refresh_batch_info()
        self._refresh_key()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-o>", lambda e: self._browse())
        self.after(100, self._poll)

        self._log(f"Paper Translator {APP_VERSION} ({'exe' if FROZEN else 'python'}) — 폴더: {BASE_DIR}", "dim")
        if self.dnd_ok:
            self._log("PDF·DOCX·MD 파일을 창에 끌어다 놓거나 [찾아보기]로 선택하세요.", "dim")
        else:
            self._log("파일 끌어다 놓기 비활성 (tkinterdnd2 없음) — [찾아보기] 또는 경로 붙여넣기를 사용하세요.", "dim")

    # ──────────────── 사이드바 (설정) ────────────────
    def _section(self, parent, title):
        lab = ctk.CTkLabel(parent, text=title, font=self.f_head, anchor="w")
        lab.pack(fill="x", padx=14, pady=(16, 4))
        return lab

    def _hint(self, parent, text="", **kw):
        lab = ctk.CTkLabel(parent, text=text, font=self.f_small, anchor="w", justify="left",
                           text_color=("gray40", "gray65"), wraplength=290, **kw)
        lab.pack(fill="x", padx=16, pady=(0, 2))
        return lab

    def _switch(self, parent, text, key):
        var = tk.BooleanVar(value=False)
        sw = ctk.CTkSwitch(parent, text=text, variable=var, font=self.f_body,
                           command=lambda: self._set_cfg(key, "true" if var.get() else "false"))
        sw.pack(fill="x", padx=16, pady=3)
        return var

    def _build_sidebar(self):
        sb = ctk.CTkScrollableFrame(self, width=320, corner_radius=0,
                                    label_text="설정", label_font=self.f_head)
        sb.grid(row=0, column=0, sticky="nsew")
        self.sidebar = sb

        # 모델
        self._section(sb, "모델")
        self.model_labels = [f"{name}   {price}" for _, name, price, _ in MODELS] + [CUSTOM_LABEL]
        self.model_menu = ctk.CTkOptionMenu(sb, values=self.model_labels, font=self.f_body,
                                            dropdown_font=self.f_body, command=self._on_model)
        self.model_menu.pack(fill="x", padx=16, pady=3)
        self.model_hint = self._hint(sb)

        # effort
        self._section(sb, "Thinking effort")
        self.effort_seg = ctk.CTkSegmentedButton(sb, values=EFFORTS, font=self.f_body,
                                                 command=lambda v: self._set_cfg("effort", v))
        self.effort_seg.pack(fill="x", padx=16, pady=3)
        self.effort_hint = self._hint(sb)

        # 각주
        self._section(sb, "각주")
        self.fn_seg = ctk.CTkSegmentedButton(sb, values=[l for _, l in FOOTNOTE_OPTS], font=self.f_body,
                                             command=self._on_footnote)
        self.fn_seg.pack(fill="x", padx=16, pady=3)
        self.fn_hint = self._hint(sb)

        # 쪽 마커
        self._section(sb, "원서 쪽 마커  [원서 p. N]")
        self.pm_var = self._switch(sb, "쪽 마커 넣기", "page_markers")
        row = ctk.CTkFrame(sb, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=3)
        ctk.CTkLabel(row, text="첫 쪽 번호", font=self.f_body).pack(side="left")
        self.fp_entry = ctk.CTkEntry(row, width=90, font=self.f_body, placeholder_text="auto")
        self.fp_entry.pack(side="left", padx=8)
        self.fp_entry.bind("<Return>", lambda e: self._commit_first_page())
        self.fp_entry.bind("<FocusOut>", lambda e: self._commit_first_page())
        ctk.CTkButton(row, text="자동", width=50, font=self.f_small,
                      command=lambda: self._set_first_page("auto")).pack(side="left")
        self._hint(sb, "비워 두거나 auto = 인쇄 쪽번호 자동 감지.\n자동 감지가 틀리면 숫자로 지정 (예: 101).")

        # 추출
        self._section(sb, "추출")
        self.ref_var = self._switch(sb, "참고문헌 이후 잘라내기", "strip_references")
        self.fig_var = self._switch(sb, "도식·표·수식 이미지 삽입", "figures")

        # 용어집
        self._section(sb, "용어집")
        self.gl_var = self._switch(sb, "자동 용어집 추출 (API 1회 추가)", "glossary")
        ctk.CTkLabel(sb, text="사용자 용어집 (자동 용어집보다 우선)", font=self.f_small,
                     anchor="w").pack(fill="x", padx=16, pady=(6, 0))
        self.gloss_menu = ctk.CTkOptionMenu(sb, values=["사용 안 함"], font=self.f_body,
                                            dropdown_font=self.f_body, command=self._on_glossary)
        self.gloss_menu.pack(fill="x", padx=16, pady=3)
        grow = ctk.CTkFrame(sb, fg_color="transparent")
        grow.pack(fill="x", padx=16, pady=3)
        for txt, cmd in (("새로 만들기", self._gloss_new), ("편집", self._gloss_edit), ("폴더", self._gloss_folder)):
            ctk.CTkButton(grow, text=txt, width=86, font=self.f_small, fg_color="transparent",
                          border_width=1, text_color=("gray10", "gray90"),
                          command=cmd).pack(side="left", padx=(0, 6))

        # 출력
        self._section(sb, "출력 형식")
        self.out_seg = ctk.CTkSegmentedButton(sb, values=[l for _, l in OUTPUT_OPTS], font=self.f_body,
                                              command=self._on_output)
        self.out_seg.pack(fill="x", padx=16, pady=3)
        self._hint(sb, "PDF는 DOCX를 Word로 변환 (실패 시 LibreOffice → 내장 렌더러).")

        # API 키
        self._section(sb, "API 키")
        self.key_label = ctk.CTkLabel(sb, text="", font=self.f_body, anchor="w")
        self.key_label.pack(fill="x", padx=16)
        krow = ctk.CTkFrame(sb, fg_color="transparent")
        krow.pack(fill="x", padx=16, pady=3)
        ctk.CTkButton(krow, text="설정/변경", width=86, font=self.f_small, command=self._key_set).pack(side="left", padx=(0, 6))
        ctk.CTkButton(krow, text="삭제", width=60, font=self.f_small, fg_color="transparent", border_width=1,
                      text_color=("gray10", "gray90"), command=self._key_clear).pack(side="left", padx=(0, 6))
        ctk.CTkButton(krow, text="크레딧 확인", width=86, font=self.f_small, fg_color="transparent", border_width=1,
                      text_color=("gray10", "gray90"),
                      command=lambda: webbrowser.open("https://console.anthropic.com/settings/billing")
                      ).pack(side="left")

        # 화면
        self._section(sb, "화면")
        self.appearance = ctk.CTkSegmentedButton(sb, values=["시스템", "라이트", "다크"], font=self.f_body,
                                                 command=self._on_appearance)
        self.appearance.pack(fill="x", padx=16, pady=(3, 20))

    # ──────────────── 메인 영역 ────────────────
    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=18, pady=14)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=1)

        # 헤더
        head = ctk.CTkFrame(main, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew")
        head.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(head, text="Paper Translator", font=self.f_title, anchor="w").grid(row=0, column=0, sticky="w")
        self.summary_label = ctk.CTkLabel(head, text="", font=self.f_small, anchor="w",
                                          text_color=("gray40", "gray65"))
        self.summary_label.grid(row=1, column=0, sticky="w")

        # 파일 영역 (드롭 존)
        drop = ctk.CTkFrame(main, corner_radius=12, border_width=2,
                            border_color=("gray70", "gray35"))
        drop.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        drop.grid_columnconfigure(0, weight=1)
        self.drop_frame = drop
        self.drop_label = ctk.CTkLabel(
            drop, font=self.f_body, text_color=("gray35", "gray70"),
            text=("여기에 PDF · DOCX · MD 파일을 끌어다 놓으세요" if self.dnd_ok
                  else "파일 경로를 붙여넣거나 [찾아보기]로 선택하세요"))
        self.drop_label.grid(row=0, column=0, columnspan=2, pady=(14, 6))
        self.path_var = tk.StringVar()
        self.path_entry = ctk.CTkEntry(drop, textvariable=self.path_var, font=self.f_body, height=34,
                                       placeholder_text="C:\\...\\paper.pdf")
        self.path_entry.grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(0, 6))
        self.path_entry.bind("<Return>", lambda e: self._set_file(self.path_var.get()))
        self.path_entry.bind("<FocusOut>", lambda e: self._set_file(self.path_var.get(), quiet=True))
        ctk.CTkButton(drop, text="찾아보기", width=96, height=34, font=self.f_body,
                      command=self._browse).grid(row=1, column=1, padx=(0, 14), pady=(0, 6))
        self.file_info = ctk.CTkLabel(drop, text="", font=self.f_small, anchor="w",
                                      text_color=("gray35", "gray70"))
        self.file_info.grid(row=2, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 12))

        if self.dnd_ok:
            for w in (self, drop, self.drop_label, self.file_info):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                    w.dnd_bind("<<DropEnter>>", lambda e: self._drop_highlight(True))
                    w.dnd_bind("<<DropLeave>>", lambda e: self._drop_highlight(False))
                except Exception:
                    pass

        # 실행 버튼
        acts = ctk.CTkFrame(main, fg_color="transparent")
        acts.grid(row=2, column=0, sticky="ew", pady=(2, 6))
        for i in range(4):
            acts.grid_columnconfigure(i, weight=1 if i < 3 else 0)
        self.btn_translate = ctk.CTkButton(acts, text="즉시 번역", height=44, font=self.f_btn,
                                           command=self._start_translate)
        self.btn_translate.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.btn_submit = ctk.CTkButton(acts, text="배치 제출  (50% 할인)", height=44, font=self.f_btn,
                                        fg_color=("#2f7d4f", "#2a6e46"), hover_color=("#286a43", "#235c3b"),
                                        command=self._start_submit)
        self.btn_submit.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.btn_download = ctk.CTkButton(acts, text="배치 결과 받기", height=44, font=self.f_btn,
                                          fg_color=("#6b5bb5", "#5a4c9c"), hover_color=("#5a4c9c", "#4b3f82"),
                                          command=self._start_download)
        self.btn_download.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.btn_stop = ctk.CTkButton(acts, text="중지", width=80, height=44, font=self.f_btn,
                                      fg_color=("#b5473f", "#9c3c35"), hover_color=("#9c3c35", "#83322c"),
                                      state="disabled", command=self._stop)
        self.btn_stop.grid(row=0, column=3, sticky="e")

        # 대기 배치 + 진행
        stat = ctk.CTkFrame(main, corner_radius=10)
        stat.grid(row=3, column=0, sticky="ew", pady=(2, 8))
        stat.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(stat, text="대기 중 배치", font=self.f_small, width=80, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(14, 6), pady=(8, 2))
        self.batch_label = ctk.CTkLabel(stat, text="", font=self.f_small, anchor="w")
        self.batch_label.grid(row=0, column=1, sticky="ew", pady=(8, 2))
        self.auto_var = tk.BooleanVar(value=bool(self.gui_cfg.get("auto_check", False)))
        ctk.CTkCheckBox(stat, text="처리 중이면 5분마다 자동 확인", variable=self.auto_var, font=self.f_small,
                        checkbox_width=18, checkbox_height=18,
                        command=self._on_auto_toggle).grid(row=0, column=2, sticky="e", padx=14, pady=(8, 2))
        self.progress = ctk.CTkProgressBar(stat, height=10)
        self.progress.grid(row=1, column=0, columnspan=3, sticky="ew", padx=14, pady=(6, 4))
        self.progress.set(0)
        self.status_label = ctk.CTkLabel(stat, text="대기", font=self.f_small, anchor="w")
        self.status_label.grid(row=2, column=0, columnspan=3, sticky="ew", padx=14, pady=(0, 8))

        # 로그
        self.log = ctk.CTkTextbox(main, font=self.f_log, wrap="word", corner_radius=10)
        self.log.grid(row=4, column=0, sticky="nsew")
        self.log.configure(state="disabled")
        self.log.tag_config("err", foreground="#e5534b")
        self.log.tag_config("warn", foreground="#d4a72c")
        self.log.tag_config("ok", foreground="#3fb950")
        self.log.tag_config("dim", foreground="#8b949e")
        self.log.tag_config("head", foreground="#58a6ff")

        # 결과 버튼
        res = ctk.CTkFrame(main, fg_color="transparent")
        res.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        self.btn_open = ctk.CTkButton(res, text="결과 열기", width=110, font=self.f_body,
                                      state="disabled", command=self._open_result)
        self.btn_open.pack(side="left", padx=(0, 8))
        self.btn_reveal = ctk.CTkButton(res, text="폴더에서 보기", width=110, font=self.f_body,
                                        state="disabled", command=self._reveal_result)
        self.btn_reveal.pack(side="left", padx=(0, 8))
        self.btn_qa = ctk.CTkButton(res, text="QA 보고서", width=100, font=self.f_body,
                                    fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
                                    command=self._open_qa)
        self.btn_qa.pack(side="left", padx=(0, 8))
        ctk.CTkButton(res, text="로그 지우기", width=100, font=self.f_body, fg_color="transparent",
                      border_width=1, text_color=("gray10", "gray90"),
                      command=self._clear_log).pack(side="right")
        ctk.CTkButton(res, text="번역기 폴더", width=100, font=self.f_body, fg_color="transparent",
                      border_width=1, text_color=("gray10", "gray90"),
                      command=lambda: open_path(BASE_DIR)).pack(side="right", padx=(0, 8))

    # ──────────────── 설정 ↔ 위젯 ────────────────
    def _load_into_widgets(self):
        c = self.cfg
        # 모델
        idx = next((i for i, m in enumerate(MODELS) if m[0] == self.model), None)
        if idx is None:
            self.model_menu.set(f"직접: {self.model}")
        else:
            self.model_menu.set(self.model_labels[idx])
        self._update_model_hint()
        self.effort_seg.set(c["effort"])
        self.fn_seg.set(dict(FOOTNOTE_OPTS)[c["footnotes"]])
        self.fn_hint.configure(text=FOOTNOTE_HELP[c["footnotes"]])
        self.pm_var.set(c["page_markers"] == "true")
        self.fp_entry.delete(0, "end")
        if c["first_page"] != "auto":
            self.fp_entry.insert(0, c["first_page"])
        self.ref_var.set(c["strip_references"] == "true")
        self.fig_var.set(c["figures"] == "true")
        self.gl_var.set(c["glossary"] == "true")
        self._refresh_glossaries()
        self.out_seg.set(dict(OUTPUT_OPTS)[c["output_format"]])
        self.appearance.set({"system": "시스템", "light": "라이트", "dark": "다크"}.get(
            self.gui_cfg.get("appearance", "system"), "시스템"))
        self._update_summary()

    def _set_cfg(self, key, value, save=True):
        self.cfg[key] = value
        if save:
            try:
                save_extract_cfg(self.cfg, self.cfg_extra)
            except Exception as e:
                messagebox.showerror("설정 저장 실패", f"extract.cfg 를 쓸 수 없습니다.\n{e}")
        if key == "effort":
            self._update_model_hint()
        self._update_summary()

    def _update_summary(self):
        c = self.cfg
        fn = dict(FOOTNOTE_OPTS)[c["footnotes"]]
        gl = c["glossary_file"] if c["glossary_file"].lower() != "none" else "없음"
        pm = ("쪽 마커 " + ("자동" if c["first_page"] == "auto" else f"p.{c['first_page']}~")) \
            if c["page_markers"] == "true" else "쪽 마커 끔"
        self.summary_label.configure(
            text=f"{self.model}  ·  effort {c['effort']}  ·  각주 {fn}  ·  {pm}  ·  "
                 f"출력 {dict(OUTPUT_OPTS)[c['output_format']]}  ·  사용자 용어집 {gl}")

    def _update_model_hint(self):
        m = next((x for x in MODELS if x[0] == self.model), None)
        fam = model_family(self.model)
        txt = f"{self.model}\n" + (f"{m[3]}  ·  입력/출력 1M 토큰당 {m[2]}" if m else "직접 입력한 모델 ID")
        self.model_hint.configure(text=txt)
        e = self.cfg.get("effort", "low")
        if fam == "haiku":
            eh = "Haiku 는 effort 조절을 쓰지 않습니다 (설정 무시)."
        elif any(k in self.model.lower() for k in ("fable", "mythos", "opus-5-5")) and e == "off":
            eh = "이 모델은 thinking 을 끌 수 없어 off → low 로 실행됩니다."
        else:
            eh = {"off": "thinking 끔", "low": "최소 thinking — 번역에 권장",
                  "medium": "어려운 문장에 더 신중, 출력 비용 약 1.5배",
                  "high": "깊은 추론, 느리고 비쌈", "xhigh": "가장 깊은 추론, 매우 느리고 비쌈"}[e]
        self.effort_hint.configure(text=eh + "\nthinking 토큰은 출력 토큰으로 과금.")

    def _on_model(self, label):
        if label == CUSTOM_LABEL:
            dlg = ctk.CTkInputDialog(font=self.f_body, text="모델 ID (예: claude-opus-5, claude-sonnet-5)", title="모델 직접 입력")
            mid = (dlg.get_input() or "").strip()
            if not mid or not re.fullmatch(r"[A-Za-z0-9.\-_]+", mid):
                if mid:
                    messagebox.showwarning("모델 ID", "영문·숫자·'-'·'.'·'_' 만 쓸 수 있습니다.")
                self._load_into_widgets()
                return
            self.model = mid
        else:
            self.model = MODELS[self.model_labels.index(label)][0]
        save_model(self.model)
        self._load_into_widgets()
        self._update_file_info()

    def _on_footnote(self, label):
        key = next(k for k, l in FOOTNOTE_OPTS if l == label)
        self.fn_hint.configure(text=FOOTNOTE_HELP[key])
        self._set_cfg("footnotes", key)

    def _on_output(self, label):
        self._set_cfg("output_format", next(k for k, l in OUTPUT_OPTS if l == label))

    def _set_first_page(self, v):
        self.fp_entry.delete(0, "end")
        if v != "auto":
            self.fp_entry.insert(0, v)
        if self.cfg["first_page"] != v:
            self._set_cfg("first_page", v)

    def _commit_first_page(self):
        raw = self.fp_entry.get().strip()
        if raw == "" or raw.lower() == "auto":
            self._set_first_page("auto")
        elif raw.isdigit() and int(raw) > 0:
            self._set_first_page(str(int(raw)))
        else:
            messagebox.showwarning("첫 쪽 번호", "양의 정수 또는 auto 를 입력하세요.")
            self._set_first_page(self.cfg["first_page"])

    def _on_appearance(self, label):
        mode = {"시스템": "system", "라이트": "light", "다크": "dark"}[label]
        ctk.set_appearance_mode(mode)
        self.gui_cfg["appearance"] = mode
        save_gui_cfg(self.gui_cfg)

    # ──────────────── 용어집 ────────────────
    def _refresh_glossaries(self):
        names = list_glossaries()
        active = self.cfg["glossary_file"]
        if active.lower() != "none" and active not in names:
            names.append(active)            # 예전 위치(user_glossary.txt) 등
        self.gloss_menu.configure(values=["사용 안 함"] + names)
        self.gloss_menu.set(active if active.lower() != "none" else "사용 안 함")

    def _on_glossary(self, label):
        self._set_cfg("glossary_file", "none" if label == "사용 안 함" else label)

    def _gloss_path(self, name):
        for cand in (GLOSS_DIR / name, BASE_DIR / name):
            if cand.exists():
                return cand
        return GLOSS_DIR / name

    def _edit_file(self, p):
        if IS_WIN:
            subprocess.Popen(["notepad", str(p)])
        else:
            open_path(p)

    def _gloss_new(self):
        dlg = ctk.CTkInputDialog(font=self.f_body, text="새 용어집 이름 (영문·숫자, 예: POL6456)", title="새 용어집")
        name = (dlg.get_input() or "").strip()
        if not name:
            return
        if not re.fullmatch(r"[A-Za-z0-9 _\-.()]+", name):
            messagebox.showwarning("새 용어집", "영문·숫자·공백·_-.() 만 써 주세요 (설정 파일 호환).")
            return
        if not name.lower().endswith(".txt"):
            name += ".txt"
        GLOSS_DIR.mkdir(exist_ok=True)
        p = GLOSS_DIR / name
        if not p.exists():
            if GLOSS_EXAMPLE.exists():
                shutil.copyfile(GLOSS_EXAMPLE, p)
            else:
                p.write_text("# English term = 한국어\n", encoding="utf-8")
        self._set_cfg("glossary_file", name)
        self._refresh_glossaries()
        self._edit_file(p)

    def _gloss_edit(self):
        name = self.cfg["glossary_file"]
        if name.lower() == "none":
            messagebox.showinfo("용어집 편집", "선택된 사용자 용어집이 없습니다. 먼저 고르거나 새로 만드세요.")
            return
        self._edit_file(self._gloss_path(name))

    def _gloss_folder(self):
        GLOSS_DIR.mkdir(exist_ok=True)
        open_path(GLOSS_DIR)
        self.after(1500, self._refresh_glossaries)

    # ──────────────── API 키 ────────────────
    def _refresh_key(self):
        k, env = load_key()
        self.key_label.configure(text=mask_key(k) + ("   (환경변수 ANTHROPIC_API_KEY 사용 중)" if env else ""))

    def _key_set(self):
        dlg = ctk.CTkInputDialog(font=self.f_body, text="Anthropic API 키 (sk-ant-로 시작)\napi_key.cfg 에 평문으로 저장됩니다.",
                                 title="API 키")
        k = (dlg.get_input() or "").strip().strip('"')
        if not k:
            return
        if not k.startswith("sk-ant-"):
            if not messagebox.askyesno("API 키", "sk-ant- 로 시작하지 않습니다. 그래도 저장할까요?"):
                return
        KEY_CFG.write_bytes(k.encode("ascii", errors="ignore"))
        self._refresh_key()
        self._log("API 키 저장됨.", "ok")

    def _key_clear(self):
        if KEY_CFG.exists() and messagebox.askyesno("API 키", "api_key.cfg 를 삭제할까요?"):
            KEY_CFG.unlink()
            self._refresh_key()

    # ──────────────── 파일 ────────────────
    def _drop_highlight(self, on):
        self.drop_frame.configure(border_color=("#3b8ed0", "#1f6aa5") if on else ("gray70", "gray35"))
        return getattr(tk, "COPY", "copy")

    def _on_drop(self, event):
        self._drop_highlight(False)
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            files = [event.data]
        if files:
            if len(files) > 1:
                self._log(f"여러 파일이 놓였습니다 — 첫 파일만 사용합니다: {Path(files[0]).name}", "warn")
            self._set_file(files[0])
        return getattr(tk, "COPY", "copy")

    def _browse(self):
        p = filedialog.askopenfilename(
            title="번역할 파일 선택", initialdir=self.last_dir or str(Path.home()),
            filetypes=[("논문 파일", "*.pdf *.docx *.md *.txt"), ("PDF", "*.pdf"), ("모든 파일", "*.*")])
        if p:
            self._set_file(p)

    def _set_file(self, p, quiet=False):
        p = (p or "").strip().strip('"').strip("'").strip("{}")
        if not p:
            self.file_info.configure(text="")
            return
        self.path_var.set(str(Path(p)))
        fp = Path(p)
        if not fp.exists():
            if not quiet:
                self.file_info.configure(text="⚠ 파일이 없습니다.", text_color="#e5534b")
            return
        self.last_dir = str(fp.parent)
        self.gui_cfg["last_dir"] = self.last_dir
        self._update_file_info()

    def _update_file_info(self):
        p = self.path_var.get().strip().strip('"')
        if not p:
            return
        fp = Path(p)
        if not fp.exists():
            return
        if fp.suffix.lower() not in SUPPORTED_EXT:
            self.file_info.configure(text=f"⚠ 지원하지 않는 형식: {fp.suffix}  (PDF / DOCX / MD / TXT)",
                                     text_color="#e5534b")
            return
        size = fp.stat().st_size / 1024 / 1024
        info = f"{fp.name}  ·  {size:.1f} MB"
        if fp.suffix.lower() == ".pdf":
            pages, scanned = self._pdf_pages(fp)
            if pages:
                info += f"  ·  {pages}쪽"
                if scanned is not None:
                    info += f"  ·  텍스트 레이어 없는 쪽 {scanned}" if scanned else "  ·  텍스트 PDF"
                fam = model_family(self.model)
                if fam in COST_PER_30P:
                    est = COST_PER_30P[fam] * pages / 30
                    info += f"  ·  예상 비용(대략) 실시간 ~${est:.2f} / 배치 ~${est / 2:.2f}"
        self.file_info.configure(text=info, text_color=("gray35", "gray70"))

    def _pdf_pages(self, fp):
        try:
            import pymupdf as fitz
        except Exception:
            try:
                import fitz
            except Exception:
                return None, None
        try:
            doc = fitz.open(str(fp))
            n = doc.page_count
            scanned = 0
            sample = range(n) if n <= 60 else range(0, n, max(1, n // 60))
            checked = 0
            for i in sample:
                checked += 1
                if len(doc[i].get_text("text").strip()) < 30:
                    scanned += 1
            doc.close()
            if n > 60 and checked:
                scanned = round(scanned * n / checked)
            return n, scanned
        except Exception:
            return None, None

    # ──────────────── 배치 정보 ────────────────
    def _batch_info(self):
        if not BATCH_INFO.exists():
            return None
        try:
            return json.loads(BATCH_INFO.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _refresh_batch_info(self):
        info = self._batch_info()
        if info is None:
            self.batch_label.configure(text="없음")
            return
        if not info:
            self.batch_label.configure(text="batch_info.json 을 읽을 수 없음")
            return
        when = info.get("submitted_at", "")
        try:
            when = datetime.fromisoformat(when).strftime("%m-%d %H:%M 제출")
        except Exception:
            pass
        bid = info.get("batch_id", "")
        self.batch_label.configure(
            text=f"{info.get('stem', '?')}   ·   {when}   ·   {info.get('model', '')}   ·   …{bid[-8:]}")

    # ──────────────── 실행 ────────────────
    def _check_ready(self, need_file=True):
        if self.proc:
            return None
        k, _ = load_key()
        if not k:
            messagebox.showwarning("API 키", "API 키가 설정되지 않았습니다. 왼쪽 [API 키 > 설정/변경]에서 입력하세요.")
            return None
        if not need_file:
            return ""
        p = self.path_var.get().strip().strip('"')
        if not p:
            messagebox.showinfo("파일", "번역할 파일을 먼저 선택하세요.")
            return None
        fp = Path(p)
        if not fp.exists():
            messagebox.showwarning("파일", f"파일이 없습니다:\n{p}")
            return None
        if fp.suffix.lower() not in SUPPORTED_EXT:
            messagebox.showwarning("파일", "PDF / DOCX / MD / TXT 파일만 번역할 수 있습니다.")
            return None
        return str(fp)

    def _sync_cfg_before_run(self):
        """화면 값 = 실행 값 보장 (bat 등 다른 곳에서 파일을 바꿨을 수도 있으므로)."""
        self._commit_first_page()
        save_extract_cfg(self.cfg, self.cfg_extra)
        save_model(self.model)

    def _start_translate(self):
        fp = self._check_ready()
        if fp is None:
            return
        self._run("translate", "translate_paper.py", [fp])

    def _start_submit(self):
        fp = self._check_ready()
        if fp is None:
            return
        info = self._batch_info()
        if info is not None:
            if not messagebox.askyesno(
                    "배치 제출",
                    f"아직 받지 않은 배치가 있습니다:\n  {(info or {}).get('stem', '?')}\n\n"
                    "새로 제출하면 batch_info.json 이 덮어써져, 이전 배치는 이 도구로 받을 수 없게 됩니다.\n"
                    "(결과는 서버에 29일 남아 있어 수동 복구는 가능)\n\n계속할까요?", icon="warning"):
                return
        self._run("submit", "batch_submit.py", [fp])

    def _start_download(self, auto=False):
        if self.proc:
            return
        if self._batch_info() is None:
            if not auto:
                messagebox.showinfo("배치 결과 받기", "대기 중인 배치가 없습니다 (batch_info.json 없음).")
            return
        if self._check_ready(need_file=False) is None:
            return
        self._run("download", "batch_download.py", [])

    def _run(self, kind, script, args):
        self._cancel_auto()
        try:
            self._sync_cfg_before_run()
        except Exception as e:
            messagebox.showerror("설정 저장 실패", str(e))
            return
        if FROZEN and not WORKER_EXE.exists():
            messagebox.showerror("실행 실패", f"ptworker.exe 가 없습니다:\n{WORKER_EXE}\n\n"
                                 "build_exe.bat 으로 다시 빌드하세요.")
            return
        if not (BASE_DIR / script).exists():
            messagebox.showerror("실행 실패", f"번역 스크립트가 없습니다:\n{BASE_DIR / script}\n\n"
                                 "gui 폴더는 번역 스크립트(.py)가 있는 폴더 바로 아래에 있어야 합니다.")
            return
        cmd = worker_command(script, args)
        env = os.environ.copy()
        env.update({"PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1"})
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if IS_WIN else 0
        try:
            self.proc = subprocess.Popen(cmd, cwd=str(BASE_DIR), stdout=subprocess.PIPE,
                                         stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                                         env=env, creationflags=flags)
        except Exception as e:
            self.proc = None
            messagebox.showerror("실행 실패", f"{script} 를 실행할 수 없습니다.\n{e}")
            return
        self.proc_kind = kind
        self.outputs = []
        self.saw_pending = False
        self.last_cost = None
        self.fatal_stop = False
        self.run_started = datetime.now()
        self._set_running(True)
        self.progress.set(0)
        self._log("")
        self._log("─" * 60, "dim")
        self._log(f"▶ {ACTION_LABEL[kind]}   {datetime.now():%H:%M:%S}   ({self.model}, effort {self.cfg['effort']})", "head")
        self._log("─" * 60, "dim")
        self._status(f"{ACTION_LABEL[kind]} 실행 중…")
        threading.Thread(target=self._reader, args=(self.proc,), daemon=True).start()

    def _reader(self, proc):
        try:
            for raw in iter(proc.stdout.readline, b""):
                self.q.put(("line", raw.decode("utf-8", errors="replace").rstrip("\r\n")))
        except Exception as e:
            self.q.put(("line", f"[GUI] 출력 읽기 오류: {e}"))
        rc = proc.wait()
        self.q.put(("exit", rc))

    def _poll(self):
        try:
            for _ in range(400):
                kind, data = self.q.get_nowait()
                if kind == "line":
                    self._handle_line(data)
                else:
                    self._on_exit(data)
        except queue.Empty:
            pass
        self.after(80, self._poll)

    def _handle_line(self, line):
        tag = None
        if "[오류]" in line or "[저장 오류]" in line or line.startswith("Traceback") or "Error" in line[:40]:
            tag = "err"
        elif "[경고]" in line or "[재번역]" in line or "  ! " in line or "[안내]" in line:
            tag = "warn"
        elif RE_SAVED.match(line) or "배치 제출 완료" in line or "QA 점검: 이상 없음" in line:
            tag = "ok"
        self._log(line, tag)

        if self.fatal_stop:      # 중지 후 남은 출력으로 상태 표시를 덮어쓰지 않음
            return
        m = RE_STEP.search(line)
        if m:
            i, total = int(m.group(1)), int(m.group(2))
            if total:
                done = i if "완료" in line or "통과" in line else i - 1
                self.progress.set(max(0.0, min(1.0, done / total)))
                self._status(f"{ACTION_LABEL.get(self.proc_kind, '')}  {done}/{total}")
        m = RE_BATCH_PROGRESS.search(line)
        if m and int(m.group(2)):
            d, t = int(m.group(1)), int(m.group(2))
            self.progress.set(d / t)
            self._status(f"배치 처리 중  {d}/{t}")
        if "아직 처리 중입니다" in line:
            self.saw_pending = True
        # 키 오류·잔액 부족이면 남은 청크도 전부 실패하므로 즉시 중지 (파이프라인은 청크마다 재시도함)
        low = line.lower()
        if self.proc and not self.fatal_stop and (
                "authentication_error" in low or "invalid x-api-key" in low
                or "credit balance is too low" in low or "permission_error" in low):
            self.fatal_stop = True
            try:
                self.proc.terminate()
            except Exception:
                pass
            why = "API 키 오류" if ("authentication" in low or "x-api-key" in low) else \
                ("크레딧 잔액 부족" if "credit" in low else "API 권한 오류")
            self._log(f"[GUI] {why} — 나머지 요청도 실패하므로 작업을 중지했습니다.", "err")
            self.fatal_msg = f"{why} — 왼쪽 [API 키]를 확인하거나 크레딧을 충전하세요"
            self._status(self.fatal_msg)
        m = RE_SAVED.match(line)
        if m:
            self.outputs.append(m.group(1))
        m = RE_TOTAL.search(line)
        if m:
            self.last_cost = m.group(1)

    def _on_exit(self, rc):
        kind = self.proc_kind
        self.proc = None
        self.proc_kind = None
        self._set_running(False)
        self._refresh_batch_info()
        elapsed = datetime.now() - getattr(self, "run_started", datetime.now())
        mins, secs = divmod(int(elapsed.total_seconds()), 60)
        if rc != 0 or self.fatal_stop:
            self._log(f"■ 종료 코드 {rc}  ({mins}분 {secs}초)", "err")
            if self.fatal_stop:
                self._status(getattr(self, "fatal_msg", "API 오류로 중지됨"))
            else:
                self._status(f"{ACTION_LABEL.get(kind, '')} 실패/중지 (코드 {rc}) — 로그를 확인하세요")
            return
        if kind == "download" and self.saw_pending:
            self._status("배치가 아직 처리 중입니다" + (" — 5분 뒤 자동 확인" if self.auto_var.get() else ""))
            if self.auto_var.get():
                self.auto_job = self.after(5 * 60 * 1000, lambda: self._start_download(auto=True))
        elif kind == "submit":
            self.progress.set(1)
            self._status("배치 제출 완료 — 나중에 [배치 결과 받기]를 누르세요 (컴퓨터를 꺼도 됩니다)")
        elif self.outputs:
            self.progress.set(1)
            cost = getattr(self, "last_cost", None)
            self._status(f"완료 ({mins}분 {secs}초)" + (f"  ·  비용 ${cost}" if cost else "")
                         + f"  ·  {Path(self.outputs[-1]).name}")
        else:
            self._status(f"{ACTION_LABEL.get(kind, '')} 종료 — 저장된 파일 없음, 로그를 확인하세요")
        self._log(f"■ 끝  ({mins}분 {secs}초)", "dim")
        self.btn_open.configure(state="normal" if self.outputs else "disabled")
        self.btn_reveal.configure(state="normal" if self.outputs else "disabled")
        if self.outputs:
            self.bell()

    def _set_running(self, on):
        st = "disabled" if on else "normal"
        for b in (self.btn_translate, self.btn_submit, self.btn_download):
            b.configure(state=st)
        self.btn_stop.configure(state="normal" if on else "disabled")

    def _stop(self):
        if not self.proc:
            return
        extra = ""
        if self.proc_kind == "submit":
            extra = "\n\n제출 도중 중지하면 서버에는 배치가 생겼지만 batch_info.json 이 저장되지 않을 수 있습니다."
        elif self.proc_kind == "translate":
            extra = "\n\n이미 처리된 청크의 비용은 청구되며, 결과 파일은 저장되지 않습니다."
        if messagebox.askyesno("중지", f"{ACTION_LABEL.get(self.proc_kind, '')} 을(를) 중지할까요?{extra}", icon="warning"):
            try:
                self.proc.terminate()
            except Exception:
                pass
            self._log("[GUI] 사용자가 중지함", "warn")

    def _on_auto_toggle(self):
        self.gui_cfg["auto_check"] = bool(self.auto_var.get())
        save_gui_cfg(self.gui_cfg)
        if not self.auto_var.get():
            self._cancel_auto()

    def _cancel_auto(self):
        if self.auto_job:
            try:
                self.after_cancel(self.auto_job)
            except Exception:
                pass
            self.auto_job = None

    # ──────────────── 결과 ────────────────
    def _open_result(self):
        if self.outputs:
            p = Path(self.outputs[-1])
            # docx+pdf 면 PDF 우선
            pdfs = [o for o in self.outputs if o.lower().endswith(".pdf")]
            open_path(pdfs[-1] if pdfs else p)

    def _reveal_result(self):
        if self.outputs:
            reveal_in_folder(self.outputs[-1])

    def _open_qa(self):
        if QA_FILE.exists():
            open_path(QA_FILE)
        else:
            messagebox.showinfo("QA 보고서", "아직 QA 보고서(last_qa.txt)가 없습니다.")

    # ──────────────── 로그 ────────────────
    def _log(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag if tag else ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _status(self, text):
        self.status_label.configure(text=text)

    # ──────────────── 종료 ────────────────
    def _on_close(self):
        if self.proc:
            if not messagebox.askyesno("종료", "작업이 실행 중입니다. 중지하고 종료할까요?", icon="warning"):
                return
            try:
                self.proc.terminate()
            except Exception:
                pass
        try:
            self.gui_cfg["geometry"] = self.geometry()
        except Exception:
            pass
        save_gui_cfg(self.gui_cfg)
        self.destroy()


def _report_crash(exc_text):
    try:
        (APP_DIR / "gui_error.log").write_text(exc_text, encoding="utf-8")
    except Exception:
        pass
    try:
        r = tk.Tk(); r.withdraw()
        messagebox.showerror("Paper Translator", "GUI 실행 중 오류가 발생했습니다.\n"
                             "자세한 내용: gui_error.log\n\n" + exc_text[-1200:])
        r.destroy()
    except Exception:
        pass


def main():
    if IS_WIN:
        try:  # 고해상도 화면에서 흐릿함 방지 (customtkinter 도 처리하지만 이중 안전)
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        try:  # 작업 표시줄에서 python 아이콘이 아닌 번역기 아이콘으로 묶이게
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception:
            pass
    app = App()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        _report_crash(traceback.format_exc())
