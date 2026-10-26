# -*- mode: python ; coding: utf-8 -*-
r"""
PyInstaller spec — Paper Translator (폴더형 exe, 실행 파일 2개가 _internal 라이브러리를 공유)

  PaperTranslator.exe : GUI (창 전용, 콘솔 없음)            <- translator_gui.pyw
  ptworker.exe        : 번역 작업 실행기 (콘솔 exe, GUI 가 창 없이 호출)  <- worker_entry.py

번역 스크립트(translate_paper.py 등)는 exe 에 넣지 않고 저장소 루트에 그대로 둔다(외부 .py).
exe 와 _internal 은 이 spec 이 있는 gui\ 폴더에 놓인다 (build_exe.bat).
대신 그 스크립트들이 import 하는 모듈을 AST 로 읽어 hiddenimports 로 넣는다.
-> 스크립트가 새 라이브러리를 import 하게 되면 build_exe.bat 을 다시 실행하면 된다.

빌드: gui\build_exe.bat  (python -m PyInstaller --noconfirm --clean translator.spec)
"""
import ast
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_dynamic_libs

APP = Path(SPECPATH)          # gui\
ROOT = APP.parent             # 저장소 루트 (번역 스크립트)
PIPELINE = ["translate_paper.py", "batch_submit.py", "batch_download.py",
            "paper_utils.py", "pdf_extract.py", "doc_render.py"]
LOCAL = {Path(f).stem for f in PIPELINE}
ICON = str(APP / "app.ico") if (APP / "app.ico").exists() else None


def scan_imports(files):
    names = set()
    for f in files:
        p = ROOT / f
        if not p.exists():
            continue
        tree = ast.parse(p.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names.update(a.name for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                names.add(node.module)
    return sorted(n for n in names if n.split(".")[0] not in LOCAL)


pipeline_imports = scan_imports(PIPELINE)
print("[spec] pipeline imports:", ", ".join(pipeline_imports))

worker_hidden = list(pipeline_imports)
# anthropic 은 하위 모듈을 지연 import 하므로 통째로
worker_hidden += collect_submodules("anthropic")
worker_hidden += ["docx", "lxml.etree", "lxml._elementpath", "pymupdf", "fitz"]
if sys.platform == "win32":
    worker_hidden += ["win32com.client", "pythoncom", "pywintypes", "win32timezone"]
worker_hidden = sorted(set(worker_hidden))

worker_datas = []
worker_datas += collect_data_files("docx")         # python-docx 기본 템플릿(default.docx 등) — 필수
worker_datas += collect_data_files("reportlab")    # 내장 폰트·설정
worker_datas += collect_data_files("anthropic")
import importlib.util
for _opt in ("certifi", "httpx", "truststore"):   # anthropic 버전에 따라 HTTPS 스택이 다름 — 설치돼 있으면 포함
    if importlib.util.find_spec(_opt):
        worker_hidden.append(_opt)
        worker_datas += collect_data_files(_opt)
worker_hidden = sorted(set(worker_hidden))

gui_datas = collect_data_files("customtkinter")    # 테마 json·폰트
gui_bins = []
try:
    gui_datas += collect_data_files("tkinterdnd2")    # tkdnd 네이티브 라이브러리(파일 끌어다 놓기)
    gui_bins += collect_dynamic_libs("tkinterdnd2")
    gui_hidden = ["tkinterdnd2"]
except Exception:
    gui_hidden = []
gui_hidden += ["customtkinter", "pymupdf"]
if ICON:
    gui_datas += [(ICON, ".")]

# 파이프라인 모듈이 exe 안으로 들어가면 BASE_DIR 이 _internal 로 바뀌어 설정을 못 찾는다 -> 반드시 제외
EXCLUDE = sorted(LOCAL)
# PyMuPDF 2.0 의 layout(onnx 모델 50MB + onnxruntime) 은 쓰지 않는다 — 들어가면 데이터 누락으로 import 가 실패함
EXCLUDE += ["pymupdf.layout", "pymupdf_layout", "pymupdf4llm", "onnxruntime"]

a_gui = Analysis(
    [str(APP / "translator_gui.pyw")],
    pathex=[], binaries=gui_bins, datas=gui_datas, hiddenimports=gui_hidden,
    hookspath=[], runtime_hooks=[], excludes=EXCLUDE, noarchive=False,
)
a_wk = Analysis(
    [str(APP / "worker_entry.py")],
    pathex=[], binaries=[], datas=worker_datas, hiddenimports=worker_hidden,
    hookspath=[], runtime_hooks=[], excludes=EXCLUDE + ["tkinter", "customtkinter", "tkinterdnd2"],
    noarchive=False,
)

pyz_gui = PYZ(a_gui.pure)
pyz_wk = PYZ(a_wk.pure)

exe_gui = EXE(
    pyz_gui, a_gui.scripts, [],
    exclude_binaries=True,
    name="PaperTranslator",
    console=False,
    icon=ICON,
    upx=False,
)
exe_wk = EXE(
    pyz_wk, a_wk.scripts,
    [("X utf8", None, "OPTION"), ("u", None, "OPTION")],   # UTF-8 모드 + 출력 버퍼 없음
    exclude_binaries=True,
    name="ptworker",
    console=True,
    icon=ICON,
    upx=False,
)

coll = COLLECT(
    exe_gui, a_gui.binaries, a_gui.datas,
    exe_wk, a_wk.binaries, a_wk.datas,
    name="PaperTranslator",
    upx=False,
)
