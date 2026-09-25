@echo off
chcp 936 >nul
setlocal
rem ============================================================
rem  Build the standalone single-file EXE (no Python needed on the
rem  target machine).  Output: dist\DocFormatTool.exe
rem
rem  The build venv intentionally lives OUTSIDE the project so that
rem  the 360 cloud-sync client does not upload it.  It is created
rem  automatically when missing: uv if available, else python -m venv.
rem
rem  This file is kept ASCII-only on purpose.  cmd.exe parses batch
rem  files with the console code page, so a UTF-8 BOM or non-ASCII
rem  bytes shift line boundaries and get executed as commands.
rem
rem  NOTE ON RESTRICTED ENVIRONMENTS (sandboxed agents/CI):
rem  PyInstaller 6.x runs its hook-discovery step in an "isolated
rem  subprocess" that talks over NAMED PIPES.  A sandbox that blocks
rem  named-pipe creation aborts the build with
rem      PermissionError: [WinError 5] Access is denied.
rem  raised from PyInstaller\isolated\_parent.py.  On some setups the
rem  failure produces NO output at all and the command appears to
rem  "succeed" while dist\DocFormatTool.exe is left untouched.  If the
rem  EXE looks stale, build with full/unsandboxed access.  Always
rem  confirm success by checking the timestamp of the produced EXE,
rem  not just the exit code.
rem ============================================================

set "VENV=%USERPROFILE%\.venvs\docformattool-build"
set "PY=%VENV%\Scripts\python.exe"
cd /d "%~dp0"

if exist "%PY%" goto :have_env

echo [1/4] Build environment not found, creating: %VENV%
where uv >nul 2>nul
if errorlevel 1 goto :venv_python

uv venv --python 3.12 "%VENV%"
if errorlevel 1 goto :fail
uv pip install --python "%PY%" pyinstaller pywin32
if errorlevel 1 goto :fail
goto :have_env

:venv_python
rem Prefer the py launcher to pick a 64-bit Python interpreter.
where py >nul 2>nul
if errorlevel 1 goto :venv_plain
py -3 -m venv "%VENV%"
if errorlevel 1 goto :fail
goto :venv_deps

:venv_plain
where python >nul 2>nul
if errorlevel 1 goto :no_python
python -m venv "%VENV%"
if errorlevel 1 goto :fail

:venv_deps
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install pyinstaller pywin32
if errorlevel 1 goto :fail

:have_env
echo [1/4] Build environment ready: %VENV%
echo [2/4] Checking dependencies ...
"%PY%" -c "import PyInstaller, win32com.client" || goto :fail

rem Guard against a silent no-op build (see the sandbox note above).
rem Record the current time, then require the EXE to be newer than it.
rem A no-op run leaves the old EXE (or none) behind and is caught here.
rem cmd's "if ... OT ..." file comparison is NOT supported on modern
rem Windows (it errors out), so the check is done with PowerShell.
rem This never deletes an existing good EXE.
for /f "delims=" %%T in ('powershell -NoProfile -Command "(Get-Date).Ticks"') do set "STAMP=%%T"

echo [3/4] Packaging (takes 1-3 minutes) ...
"%PY%" -m PyInstaller build_onefile.spec --noconfirm --clean || goto :fail

if not exist "dist\DocFormatTool.exe" goto :stale
for %%F in ("dist\DocFormatTool.exe") do if %%~zF LSS 1000000 goto :stale
powershell -NoProfile -Command "if ((Get-Item 'dist\DocFormatTool.exe').LastWriteTime.Ticks -le %STAMP%) { exit 1 }" || goto :stale

echo [4/4] Done. Output:
echo        dist\DocFormatTool.exe
echo.
echo Verify the EXE on the target machine (no Python required):
echo        dist\DocFormatTool.exe --selftest
echo.
echo NOTE: the target machine needs Microsoft Word or WPS Writer
echo       installed - the tool drives it through COM automation.
exit /b 0

:stale
echo.
echo *** BUILD PRODUCED NO EXE - dist\DocFormatTool.exe was not regenerated ***
echo     Any previously built EXE (if present) was left untouched.
echo     If PyInstaller reported a named-pipe / WinError 5 failure, run this
echo     script with full (unsandboxed) access - see the note at the top.
exit /b 1

:no_python
echo.
echo *** No Python found (no uv, py or python on PATH) ***
echo     Install 64-bit Python 3.10+ or uv, then run this script again.
exit /b 1

:fail
echo.
echo *** BUILD FAILED - see the error message above ***
exit /b 1
