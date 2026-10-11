@echo off
setlocal
title Build Paper Translator exe
cd /d "%~dp0"

echo.
echo  ==========================================
echo    Build Paper Translator  (folder-type exe)
echo  ==========================================
echo    PaperTranslator.exe : GUI
echo    ptworker.exe        : translation worker (called by the GUI)
echo    _internal\          : Python runtime + libraries
echo    Built into this gui folder. The *.py scripts and settings
echo    stay in the parent folder - the repository root.
echo  ==========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found in PATH.
    pause
    exit /b 1
)
if not exist "..\paper_utils.py" (
    echo  [ERROR] paper_utils.py not found in the parent folder.
    echo  This gui folder must sit directly inside the translator folder.
    pause
    exit /b 1
)

:: -- exe must not be running while its files are replaced --
tasklist /FI "IMAGENAME eq PaperTranslator.exe" 2>nul | find /I "PaperTranslator.exe" >nul
if not errorlevel 1 (
    echo  [ERROR] PaperTranslator.exe is running. Close it and run this again.
    pause
    exit /b 1
)
tasklist /FI "IMAGENAME eq ptworker.exe" 2>nul | find /I "ptworker.exe" >nul
if not errorlevel 1 (
    echo  [ERROR] A translation job - ptworker.exe - is running. Wait or stop it first.
    pause
    exit /b 1
)

echo  [1/4] Installing / updating build tools and libraries...
python -m pip install --upgrade -r ..\requirements.txt -r requirements-gui.txt
if errorlevel 1 (
    echo  [ERROR] pip install failed.
    pause
    exit /b 1
)

echo.
echo  [2/4] Building - this takes 1-3 minutes...
python -m PyInstaller --noconfirm --clean --distpath build\dist --workpath build\work translator.spec > build_exe.log 2>&1
if errorlevel 1 (
    echo  [ERROR] PyInstaller failed. See build_exe.log
    pause
    exit /b 1
)
if not exist "build\dist\PaperTranslator\PaperTranslator.exe" (
    echo  [ERROR] exe not produced. See build_exe.log
    pause
    exit /b 1
)

echo  [3/4] Placing exe next to the scripts...
if exist "_internal" rmdir /s /q "_internal"
robocopy "build\dist\PaperTranslator" "." /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (
    echo  [ERROR] Copy failed.
    pause
    exit /b 1
)
rmdir /s /q build

echo  [4/4] Quick check of the worker...
ptworker.exe --selftest > build_selftest.log 2>&1
if errorlevel 1 (
    echo  [WARN] Worker self-test failed. See build_selftest.log
    type build_selftest.log
) else (
    echo  Worker OK.
)

echo.
echo  Done:  %~dp0PaperTranslator.exe
echo.
choice /C YN /N /M "  Create a desktop shortcut? [Y/N]: "
if errorlevel 2 goto end
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Paper Translator.lnk'); $s.TargetPath='%~dp0PaperTranslator.exe'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%~dp0PaperTranslator.exe,0'; $s.Save()"
if errorlevel 1 (
    echo  Shortcut failed - you can pin PaperTranslator.exe manually.
) else (
    echo  Desktop shortcut created.
)

:end
echo.
pause
exit /b 0
