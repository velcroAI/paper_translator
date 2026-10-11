# -*- coding: utf-8 -*-
r"""
worker_entry.py — exe 빌드용 작업 실행기 (ptworker.exe 의 진입점)
=================================================================
GUI(PaperTranslator.exe)가 번역 작업을 실행할 때 호출한다:

    ptworker.exe translate_paper.py <파일>
    ptworker.exe batch_submit.py <파일>
    ptworker.exe batch_download.py

번역 스크립트(.py)는 exe 안에 넣지 않고 저장소 루트에 둔다 (exe 는 루트\gui\ 안).
  - paper_utils 등이 Path(__file__).parent 기준으로 설정·로그·batch_info.json·assets 를 찾기 때문
  - 번역 로직을 고쳐도 exe 를 다시 빌드할 필요가 없다 (새 라이브러리를 import 할 때만 재빌드)
exe 에는 파이썬 런타임과 라이브러리(anthropic, pymupdf, python-docx, reportlab, pywin32 …)만 들어 있다.
"""

import os
import sys
import runpy
import traceback
from pathlib import Path

ALLOWED = ("translate_paper.py", "batch_submit.py", "batch_download.py")


def _base_dir() -> Path:
    """번역 스크립트(paper_utils.py)가 있는 루트: 자기 폴더 -> 한 단계 위 순으로 찾는다."""
    here = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
    for cand in (here, here.parent):
        if (cand / "paper_utils.py").exists():
            return cand
    return here.parent


def _selftest(base: Path) -> int:
    """빌드 직후 점검: 번역 스크립트가 쓰는 라이브러리가 exe 에 들어 있는지 (API 호출 없음)."""
    ok = True
    mods = ["anthropic", "docx", "pymupdf", "reportlab.platypus", "lxml.etree"]
    if sys.platform == "win32":
        mods += ["win32com.client", "pythoncom"]
    for m in mods:
        try:
            __import__(m)
            print(f"  ok   {m}")
        except Exception as e:
            ok = False
            print(f"  FAIL {m}: {e}")
    for f in ALLOWED + ("paper_utils.py", "pdf_extract.py", "doc_render.py"):
        if not (base / f).exists():
            ok = False
            print(f"  FAIL 스크립트 없음: {f}")
    print("selftest:", "OK" if ok else "FAIL")
    return 0 if ok else 1


def main() -> int:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        except Exception:
            pass
    if len(sys.argv) < 2:
        print("사용법: ptworker.exe <translate_paper.py|batch_submit.py|batch_download.py> [인자...]")
        return 2
    base = _base_dir()
    if sys.argv[1] == "--selftest":
        return _selftest(base)
    name = Path(sys.argv[1]).name
    if name not in ALLOWED:
        print(f"[ptworker] 허용되지 않은 스크립트: {name}")
        return 2
    script = base / name
    if not script.exists():
        print(f"[ptworker] 스크립트가 없습니다: {script}")
        print("  gui 폴더는 번역 스크립트(.py)가 있는 폴더 바로 아래에 있어야 합니다.")
        return 2
    # 작업 폴더를 루트로 바꾸기 전에, 상대 경로로 준 입력 파일을 절대 경로로 고정
    args = [str(Path(a).resolve()) if (not Path(a).is_absolute() and Path(a).exists()) else a
            for a in sys.argv[2:]]
    os.chdir(base)
    sys.path.insert(0, str(base))
    sys.argv = [str(script)] + args
    try:
        runpy.run_path(str(script), run_name="__main__")
    except SystemExit as e:
        code = e.code
        return code if isinstance(code, int) else (0 if code is None else 1)
    except ModuleNotFoundError as e:
        traceback.print_exc()
        print(f"\n[ptworker] exe 에 포함되지 않은 모듈: {e.name}")
        print("  번역 스크립트가 새 라이브러리를 쓰게 되었다면 build_exe.bat 으로 exe 를 다시 빌드하세요.")
        return 1
    except Exception:
        traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
