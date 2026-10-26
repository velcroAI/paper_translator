@echo off
setlocal EnableDelayedExpansion
title Paper Translation Tool

:: -- Scripts and settings live in the repository root (one folder up) --
for %%I in ("%~dp0..") do set "SCRIPT_DIR=%%~fI\"
set "MODEL_CFG=%SCRIPT_DIR%model.cfg"
set "EXTRACT_CFG=%SCRIPT_DIR%extract.cfg"
set "KEY_CFG=%SCRIPT_DIR%api_key.cfg"

:: -- One-time dependency check (silent if already installed) --
python -c "import anthropic, docx, pymupdf, reportlab" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [Setup] Installing / upgrading required Python packages...
    echo  ------------------------------------------
    python -m pip --version >nul 2>&1
    if errorlevel 1 (
        echo  pip not found - bootstrapping via ensurepip...
        python -m ensurepip --default-pip
    )
    python -m pip install --upgrade anthropic python-docx pymupdf reportlab
    set "INSTALL_RC=!errorlevel!"
    python -m pip install --upgrade pywin32 >nul 2>&1
    echo  ------------------------------------------
    if not "!INSTALL_RC!"=="0" (
        echo.
        echo  [ERROR] Package install failed.
        echo  Manual: python -m pip install --upgrade anthropic python-docx pymupdf reportlab
        pause
    )
)

:: -- Optional: pywin32 for Word-based PDF export (real page-bottom footnotes) --
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo  [Setup] Installing pywin32 for Word PDF export...
    python -m pip install --upgrade pywin32 >nul 2>&1
)

:menu
:: -- Read current model (default: sonnet-5-5) --
set "CURRENT_MODEL=claude-sonnet-5-5"
if exist "!MODEL_CFG!" (
    set /p CURRENT_MODEL=<"!MODEL_CFG!"
)

:: -- Read current options (defaults) --
set "FN_MODE="
set "LEGACY_FN="
set "STRIP_REF_RAW=true"
set "OUT_FMT=docx"
set "EFFORT=low"
set "FIG_RAW=true"
set "GLOSS_RAW=true"
set "PM_RAW=true"
set "FIRST_PAGE=auto"
set "GLOSS_FILE="
set "STYLE_FILE="
if exist "!EXTRACT_CFG!" (
    for /f "usebackq tokens=1,2 delims==" %%a in ("!EXTRACT_CFG!") do (
        if /i "%%a"=="footnotes"        set "FN_MODE=%%b"
        if /i "%%a"=="strip_footnotes"  set "LEGACY_FN=%%b"
        if /i "%%a"=="strip_references" set "STRIP_REF_RAW=%%b"
        if /i "%%a"=="output_format"    set "OUT_FMT=%%b"
        if /i "%%a"=="effort"           set "EFFORT=%%b"
        if /i "%%a"=="figures"          set "FIG_RAW=%%b"
        if /i "%%a"=="glossary"         set "GLOSS_RAW=%%b"
        if /i "%%a"=="page_markers"     set "PM_RAW=%%b"
        if /i "%%a"=="first_page"       set "FIRST_PAGE=%%b"
        if /i "%%a"=="glossary_file"    set "GLOSS_FILE=%%b"
        if /i "%%a"=="style_file"       set "STYLE_FILE=%%b"
    )
)
if not defined FN_MODE (
    if /i "!LEGACY_FN!"=="true"  set "FN_MODE=strip"
    if /i "!LEGACY_FN!"=="false" set "FN_MODE=all"
)
if not defined FN_MODE set "FN_MODE=commentary"
set "FN_DISP=!FN_MODE!"
if /i "!FN_MODE!"=="strip"      set "FN_DISP=strip (delete all)"
if /i "!FN_MODE!"=="commentary" set "FN_DISP=commentary only"
if /i "!FN_MODE!"=="all"        set "FN_DISP=translate all"
if /i "!STRIP_REF_RAW!"=="true" (set "STRIP_REF=ON") else (set "STRIP_REF=OFF")
if /i "!PM_RAW!"=="true" (set "PM_DISP=ON, first page !FIRST_PAGE!") else (set "PM_DISP=OFF")
if /i "!FIG_RAW!"=="true"       (set "FIG_DISP=ON") else (set "FIG_DISP=OFF")
if /i "!GLOSS_RAW!"=="true"     (set "GLOSS_DISP=ON") else (set "GLOSS_DISP=OFF")
if not defined GLOSS_FILE (
    if exist "%SCRIPT_DIR%user_glossary.txt" (set "GLOSS_FILE=user_glossary.txt") else (set "GLOSS_FILE=none")
)
set "GLOSS_FILE_DISP=!GLOSS_FILE!"
if /i "!GLOSS_FILE!"=="none" set "GLOSS_FILE_DISP=OFF"
if not defined STYLE_FILE set "STYLE_FILE=style_ko.txt"
set "STYLE_DISP=ON (!STYLE_FILE!)"
if /i "!STYLE_FILE!"=="none" set "STYLE_DISP=OFF"
if /i not "!STYLE_FILE!"=="none" if not exist "!SCRIPT_DIR!!STYLE_FILE!" set "STYLE_DISP=OFF (file missing)"
set "OUT_FMT_DISP=DOCX"
if /i "!OUT_FMT!"=="pdf"      set "OUT_FMT_DISP=PDF"
if /i "!OUT_FMT!"=="docx+pdf" set "OUT_FMT_DISP=DOCX + PDF"
if /i "!OUT_FMT!"=="md"       set "OUT_FMT_DISP=MD"

:: -- Read API key status (masked display) --
set "KEY_DISP=[Not set]"
if exist "!KEY_CFG!" (
    for /f "usebackq delims=" %%k in ("!KEY_CFG!") do set "API_KEY_VAL=%%k"
    if defined API_KEY_VAL (
        set "KEY_TAIL=!API_KEY_VAL:~-4!"
        set "KEY_DISP=sk-ant-***!KEY_TAIL!"
    )
)

echo.
echo  ==========================================
echo    Paper Translation Tool  v4.0
echo  ==========================================
echo.
echo    1  Translate now       (immediate)
echo    2  Batch submit        (50%% off)
echo    3  Batch download
echo    4  Options
echo    5  Check credits        (opens browser)
echo.
echo    q  Quit
echo  ==========================================
echo    Model     : !CURRENT_MODEL!
echo    Effort    : !EFFORT!
echo    Footnotes : !FN_DISP!    Refs cut: !STRIP_REF!
echo    Page mark : !PM_DISP!
echo    Figures   : !FIG_DISP!    Auto glossary: !GLOSS_DISP!
echo    User gloss: !GLOSS_FILE_DISP!
echo    KR style  : !STYLE_DISP!
echo    Output    : !OUT_FMT_DISP!
echo    API key   : !KEY_DISP!
echo  ==========================================
echo.

:prompt
set "choice="
set /p "choice=  Select [1/2/3/4/5/q]: "

if /i "!choice!"=="1" goto translate
if /i "!choice!"=="2" goto submit
if /i "!choice!"=="3" goto download
if /i "!choice!"=="4" goto options
if /i "!choice!"=="5" goto credits
if /i "!choice!"=="q" goto quit

echo  Invalid choice. Try again.
goto prompt

:translate
echo.
echo  [ Translate Now ]
echo  Supports PDF / DOCX / MD files.
echo  Output saved to the same folder as input.
echo  Figures / tables / equations are clipped from the PDF
echo  and inserted at their original positions.
echo  Output name: [file]_ppA-B_[Korean suffix].docx/pdf  + QA report
echo.
set "fp="
set /p "fp=  File path: "
set "fp=!fp:"=!"
if "!fp!"=="" (
    echo  No path entered.
    goto prompt
)
echo.
echo  Starting translation...
echo  ------------------------------------------
chcp 65001 > nul
set PYTHONUTF8=1
python "%SCRIPT_DIR%translate_paper.py" "!fp!"
set PYTHONUTF8=
echo  ------------------------------------------
echo.
echo  Press any key to back to main menu...
pause > nul
cls
goto menu

:submit
echo.
echo  [ Batch Submit  -  50%% off ]
echo  Submits job to server. Results come later.
echo  Run option 3 to download when ready.
echo.
set "fp="
set /p "fp=  File path: "
set "fp=!fp:"=!"
if "!fp!"=="" (
    echo  No path entered.
    goto prompt
)
echo.
echo  Submitting batch job...
echo  ------------------------------------------
chcp 65001 > nul
set PYTHONUTF8=1
python "%SCRIPT_DIR%batch_submit.py" "!fp!"
set PYTHONUTF8=
echo  ------------------------------------------
echo.
echo  Press any key to back to main menu...
pause > nul
cls
goto menu

:download
echo.
echo  [ Batch Download ]
echo  Checks status and downloads result.
echo.
echo  ------------------------------------------
chcp 65001 > nul
set PYTHONUTF8=1
python "%SCRIPT_DIR%batch_download.py"
set PYTHONUTF8=
echo  ------------------------------------------
echo.
echo  Press any key to back to main menu...
pause > nul
cls
goto menu

:credits
echo.
echo  Opening Anthropic console in browser...
start https://console.anthropic.com/settings/billing
timeout /t 1 > nul
cls
goto menu

:: -- Write extract.cfg from current variables --
:savecfg
(
    echo footnotes=!FN_MODE!
    echo strip_references=!STRIP_REF_RAW!
    echo output_format=!OUT_FMT!
    echo effort=!EFFORT!
    echo figures=!FIG_RAW!
    echo glossary=!GLOSS_RAW!
    echo page_markers=!PM_RAW!
    echo first_page=!FIRST_PAGE!
    echo glossary_file=!GLOSS_FILE!
    echo style_file=!STYLE_FILE!
) > "!EXTRACT_CFG!"
goto :eof

:options
cls
echo.
echo  ==========================================
echo    Options
echo  ==========================================
echo.
echo    1  Model select         (current: !CURRENT_MODEL!)
echo    2  Thinking effort      (current: !EFFORT!)
echo    3  Footnotes            (current: !FN_DISP!)
echo    4  Page markers         (current: !PM_DISP!)
echo    5  Extract options      (refs / figures)
echo    6  Auto glossary pass   (current: !GLOSS_DISP!)
echo    7  Output format        (current: !OUT_FMT_DISP!)
echo    8  API key              (current: !KEY_DISP!)
echo    9  User glossary        (current: !GLOSS_FILE_DISP!)
echo    s  Korean style guide   (current: !STYLE_DISP!)
echo.
echo    b  Back
echo  ==========================================
echo.
set "oc="
set /p "oc=  Select [1-9/s/b]: "

if /i "!oc!"=="1" goto modelselect
if /i "!oc!"=="2" goto effortselect
if /i "!oc!"=="3" goto footnoteselect
if /i "!oc!"=="4" goto pagemarkers
if /i "!oc!"=="5" goto extractoptions
if /i "!oc!"=="6" goto glossarytoggle
if /i "!oc!"=="7" goto outputformat
if /i "!oc!"=="8" goto apikey
if /i "!oc!"=="9" goto glossarymenu
if /i "!oc!"=="s" goto stylemenu
if /i "!oc!"=="b" (cls & goto menu)

echo  Invalid choice.
timeout /t 1 > nul
goto options

:modelselect
cls
echo.
echo  ==========================================
echo    Model Select        (price: input / output per 1M tokens)
echo  ==========================================
echo.
echo    1  claude-fable-5-1          $10 / $50   best quality, thinking always on
echo    2  claude-opus-5-5           $4  / $20   very high quality, thinking always on
echo    3  claude-sonnet-5-5         $2  / $10   balanced / default
echo    4  claude-haiku-4-5          $1  / $5    fast / cheap (no effort control)
echo    5  Enter model ID manually   (e.g. claude-opus-5, claude-sonnet-5)
echo.
echo    b  Back
echo  ==========================================
echo    Current : !CURRENT_MODEL!
echo    Rough cost for a 30-page paper (about 25k words):
echo      sonnet-5-5 ~ $0.6   opus-5-5 ~ $1.2   fable-5-1 ~ $3   (batch: half)
echo  ==========================================
echo.
set "mc="
set /p "mc=  Select [1/2/3/4/5/b]: "

if "!mc!"=="1" (
    echo claude-fable-5-1>"!MODEL_CFG!"
    cls
    goto menu
)
if "!mc!"=="2" (
    echo claude-opus-5-5>"!MODEL_CFG!"
    cls
    goto menu
)
if "!mc!"=="3" (
    echo claude-sonnet-5-5>"!MODEL_CFG!"
    cls
    goto menu
)
if "!mc!"=="4" (
    echo claude-haiku-4-5-20251001>"!MODEL_CFG!"
    cls
    goto menu
)
if "!mc!"=="5" (
    set "mid="
    set /p "mid=  Model ID: "
    if not "!mid!"=="" (
        echo !mid!>"!MODEL_CFG!"
    )
    cls
    goto menu
)
if /i "!mc!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto modelselect

:effortselect
cls
echo.
echo  ==========================================
echo    Thinking Effort
echo  ==========================================
echo.
echo    1  off      no thinking (fable / opus-5-5 cannot turn off: uses low)
echo    2  low      minimal thinking - recommended for translation
echo    3  medium   more careful with hard sentences, ~1.5x output cost
echo    4  high     deep reasoning, slow and expensive
echo    5  xhigh    deepest reasoning, very slow / expensive
echo.
echo    b  Back
echo  ==========================================
echo    Current : !EFFORT!
echo    Thinking tokens are billed as output tokens.
echo  ==========================================
echo.
set "ef="
set /p "ef=  Select [1/2/3/4/5/b]: "

if "!ef!"=="1" (set "EFFORT=off" & call :savecfg & cls & goto menu)
if "!ef!"=="2" (set "EFFORT=low" & call :savecfg & cls & goto menu)
if "!ef!"=="3" (set "EFFORT=medium" & call :savecfg & cls & goto menu)
if "!ef!"=="4" (set "EFFORT=high" & call :savecfg & cls & goto menu)
if "!ef!"=="5" (set "EFFORT=xhigh" & call :savecfg & cls & goto menu)
if /i "!ef!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto effortselect

:extractoptions
cls
echo.
echo  ==========================================
echo    Extract Options
echo  ==========================================
echo.
echo    1  Cut references    : !STRIP_REF!
echo    2  Figures/tables/eq : !FIG_DISP!
echo.
echo    b  Back
echo  ==========================================
echo    Figures  : images, drawings, tables and display
echo               equations are clipped from the PDF and
echo               inserted into the translation in place.
echo  ==========================================
echo.
set "ec="
set /p "ec=  Select [1/2/b]: "

if "!ec!"=="1" (
    if /i "!STRIP_REF_RAW!"=="true" (set "STRIP_REF_RAW=false") else (set "STRIP_REF_RAW=true")
    call :savecfg
    cls
    goto menu
)
if "!ec!"=="2" (
    if /i "!FIG_RAW!"=="true" (set "FIG_RAW=false") else (set "FIG_RAW=true")
    call :savecfg
    cls
    goto menu
)
if /i "!ec!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto extractoptions

:footnoteselect
cls
echo.
echo  ==========================================
echo    Footnotes
echo  ==========================================
echo.
echo    1  strip        delete all footnotes and callout numbers
echo    2  commentary   translate notes with explanation/argument only,
echo                    as real Word footnotes; citations stay in English;
echo                    pure bibliographic notes are deleted
echo    3  all          translate every note as a Word footnote
echo.
echo    b  Back
echo  ==========================================
echo    Current : !FN_DISP!
echo    Tip: if notes are very long compared to the body,
echo         decide first whether you need them (cost).
echo  ==========================================
echo.
set "fc="
set /p "fc=  Select [1/2/3/b]: "

if "!fc!"=="1" (set "FN_MODE=strip" & call :savecfg & cls & goto menu)
if "!fc!"=="2" (set "FN_MODE=commentary" & call :savecfg & cls & goto menu)
if "!fc!"=="3" (set "FN_MODE=all" & call :savecfg & cls & goto menu)
if /i "!fc!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto footnoteselect

:pagemarkers
cls
echo.
echo  ==========================================
echo    Page Markers   [p. N] of the original
echo  ==========================================
echo.
echo    1  Toggle ON/OFF        (current: !PM_RAW!)
echo    2  Auto page numbers    (printed numbers on the pages)
echo    3  Set first page number manually
echo.
echo    b  Back
echo  ==========================================
echo    Current first page : !FIRST_PAGE!
echo    Use 3 if auto detection fails, e.g. the PDF starts
echo    at book page 101 and has no readable page numbers.
echo    Scanned PDFs (vision mode) read the printed numbers.
echo  ==========================================
echo.
set "pc="
set /p "pc=  Select [1/2/3/b]: "

if "!pc!"=="1" (
    if /i "!PM_RAW!"=="true" (set "PM_RAW=false") else (set "PM_RAW=true")
    call :savecfg
    cls
    goto menu
)
if "!pc!"=="2" (set "FIRST_PAGE=auto" & call :savecfg & cls & goto menu)
if "!pc!"=="3" (
    set "fpn="
    set /p "fpn=  First page number: "
    if not "!fpn!"=="" set "FIRST_PAGE=!fpn!"
    call :savecfg
    cls
    goto menu
)
if /i "!pc!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto pagemarkers

:glossarymenu
cls
echo.
echo  ==========================================
echo    User Glossary   (glossaries folder)
echo  ==========================================
echo    Applied to every translation with priority
echo    over the auto glossary. One per line:
echo        English term = Korean
echo  ==========================================
echo    Active : !GLOSS_FILE_DISP!
echo.
if not exist "%SCRIPT_DIR%glossaries" mkdir "%SCRIPT_DIR%glossaries"
set "GN=0"
for %%f in ("%SCRIPT_DIR%glossaries\*.txt") do (
    set /a GN+=1
    set "GL_!GN!=%%~nxf"
    echo    !GN!  %%~nxf
)
if "!GN!"=="0" echo    - no glossary files yet -
echo.
echo    n  New glossary (ASCII file name, e.g. POL6456)
echo    e  Edit active glossary in Notepad
echo    o  Turn user glossary OFF
echo    f  Open glossaries folder
echo    b  Back
echo  ==========================================
echo.
set "gc="
set /p "gc=  Select number or [n/e/o/f/b]: "
if "!gc!"=="" goto glossarymenu
if /i "!gc!"=="b" (cls & goto options)
if /i "!gc!"=="o" (set "GLOSS_FILE=none" & call :savecfg & cls & goto menu)
if /i "!gc!"=="f" (start "" explorer "%SCRIPT_DIR%glossaries" & goto glossarymenu)
if /i "!gc!"=="e" goto glossary_edit
if /i "!gc!"=="n" goto glossary_new
set "pick="
for %%n in (!gc!) do set "pick=!GL_%%n!"
if defined pick (
    set "GLOSS_FILE=!pick!"
    call :savecfg
    cls
    goto menu
)
echo  Invalid choice.
timeout /t 1 > nul
goto glossarymenu

:glossary_edit
if /i "!GLOSS_FILE!"=="none" (
    echo  No active glossary. Select or create one first.
    timeout /t 2 > nul
    goto glossarymenu
)
set "GPATH=%SCRIPT_DIR%glossaries\!GLOSS_FILE!"
if not exist "!GPATH!" set "GPATH=%SCRIPT_DIR%!GLOSS_FILE!"
start "" notepad "!GPATH!"
goto glossarymenu

:glossary_new
set "gname="
set /p "gname=  New glossary name (no spaces needed): "
if "!gname!"=="" goto glossarymenu
if /i not "!gname:~-4!"==".txt" set "gname=!gname!.txt"
if exist "%SCRIPT_DIR%glossaries\!gname!" (
    echo  Already exists - selected.
) else (
    if exist "%SCRIPT_DIR%glossary_example.txt" (
        copy /y "%SCRIPT_DIR%glossary_example.txt" "%SCRIPT_DIR%glossaries\!gname!" > nul
    ) else (
        echo # English term = Korean> "%SCRIPT_DIR%glossaries\!gname!"
    )
)
set "GLOSS_FILE=!gname!"
call :savecfg
start "" notepad "%SCRIPT_DIR%glossaries\!gname!"
cls
goto menu

:stylemenu
cls
echo.
echo  ==========================================
echo    Korean Style Guide   (style_ko.txt)
echo  ==========================================
echo    Rules that remove translationese from the
echo    Korean wording. Fidelity to the source always
echo    comes first. Leftover expressions are counted
echo    in the QA report.
echo  ==========================================
echo    Current : !STYLE_DISP!
echo.
echo    1  Turn ON
echo    2  Turn OFF
echo    3  Edit style_ko.txt in Notepad
echo    b  Back
echo  ==========================================
echo.
set "sc="
set /p "sc=  Select [1/2/3/b]: "
if "!sc!"=="1" (
    if /i "!STYLE_FILE!"=="none" set "STYLE_FILE=style_ko.txt"
    call :savecfg
    cls
    goto menu
)
if "!sc!"=="2" (set "STYLE_FILE=none" & call :savecfg & cls & goto menu)
if "!sc!"=="3" (
    set "SPATH=!SCRIPT_DIR!!STYLE_FILE!"
    if /i "!STYLE_FILE!"=="none" set "SPATH=!SCRIPT_DIR!style_ko.txt"
    start "" notepad "!SPATH!"
    goto stylemenu
)
if /i "!sc!"=="b" (cls & goto options)
echo  Invalid choice.
timeout /t 1 > nul
goto stylemenu

:glossarytoggle
if /i "!GLOSS_RAW!"=="true" (set "GLOSS_RAW=false") else (set "GLOSS_RAW=true")
call :savecfg
cls
goto menu

:outputformat
cls
echo.
echo  ==========================================
echo    Output Format
echo  ==========================================
echo.
echo    1  DOCX        (Word document, default)
echo    2  PDF         (made from the DOCX via Word; fallback: built-in)
echo    3  DOCX + PDF  (both)
echo    4  MD          (Markdown, images linked)
echo.
echo    b  Back
echo  ==========================================
echo    Current : !OUT_FMT_DISP!
echo  ==========================================
echo.
set "of="
set /p "of=  Select [1/2/3/4/b]: "

if "!of!"=="1" (set "OUT_FMT=docx" & call :savecfg & cls & goto menu)
if "!of!"=="2" (set "OUT_FMT=pdf" & call :savecfg & cls & goto menu)
if "!of!"=="3" (set "OUT_FMT=docx+pdf" & call :savecfg & cls & goto menu)
if "!of!"=="4" (set "OUT_FMT=md" & call :savecfg & cls & goto menu)
if /i "!of!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto outputformat

:apikey
cls
echo.
echo  ==========================================
echo    API Key
echo  ==========================================
echo.
echo    1  Set / Update key
echo    2  Clear key
echo.
echo    b  Back
echo  ==========================================
echo    Current : !KEY_DISP!
echo  ==========================================
echo    Key is saved in api_key.cfg (plain text).
echo    Env var ANTHROPIC_API_KEY overrides this.
echo  ==========================================
echo.
set "ak="
set /p "ak=  Select [1/2/b]: "

if "!ak!"=="1" goto apikey_set
if "!ak!"=="2" goto apikey_clear
if /i "!ak!"=="b" (cls & goto options)

echo  Invalid choice.
timeout /t 1 > nul
goto apikey

:apikey_set
echo.
echo  Paste your Anthropic API key (starts with sk-ant-).
echo  Input will be visible.
echo.
set "newkey="
set /p "newkey=  Key: "
if "!newkey!"=="" (
    echo  No input.
    timeout /t 1 > nul
    goto apikey
)
<nul set /p ="!newkey!" > "!KEY_CFG!"
echo.
echo  Saved.
timeout /t 1 > nul
cls
goto menu

:apikey_clear
if exist "!KEY_CFG!" del "!KEY_CFG!"
echo.
echo  Key cleared.
timeout /t 1 > nul
cls
goto menu

:quit
echo.
echo  Goodbye.
timeout /t 1 > nul
exit /b 0
