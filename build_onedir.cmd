@echo off
chcp 936 >nul
setlocal
rem ============================================================
rem  Build the FOLDER (onedir) distribution and flatten it, so the
rem  executable ends up at
rem        dist\DocFormatTool.exe
rem  overwriting whatever was there, with its support files in
rem        dist\_internal\
rem  next to it.
rem
rem  WHY THE FLATTEN STEP EXISTS
rem  PyInstaller only ever writes dist\DocFormatTool\DocFormatTool.exe.
rem  It will never touch dist\DocFormatTool.exe. Without the move below
rem  the exe at dist\DocFormatTool.exe silently stays at the PREVIOUS
rem  build forever - it looks fine but is stale. The move is what makes
rem  "the new build overwrites the existing exe" actually true.
rem
rem  KEEP THIS FILE ASCII-ONLY.  cmd.exe parses batch files with the
rem  console code page, so a UTF-8 BOM or any non-ASCII byte shifts
rem  line boundaries and gets executed as commands. (That is also why
rem  there are no Chinese comments in here.)
rem
rem  NOTE: build_exe.cmd (the onefile build) writes to the SAME path,
rem  dist\DocFormatTool.exe.  The two builds overwrite each other - a
rem  onedir build followed by a onefile build leaves the onefile exe,
rem  and vice versa.  Build whichever one you are going to ship last.
rem ============================================================

set "VENV=%USERPROFILE%\.venvs\docformattool-build"
set "PY=%VENV%\Scripts\python.exe"
cd /d "%~dp0"

if not exist "%PY%" goto :no_env

echo [1/5] Build environment ready: %VENV%
"%PY%" -c "import PyInstaller, win32com.client" || goto :fail

rem Guard against a silent no-op build (see the named-pipe note in
rem build_exe.cmd).  Record the time, then require the produced EXE to
rem be newer than it.
for /f "delims=" %%T in ('powershell -NoProfile -Command "(Get-Date).Ticks"') do set "STAMP=%%T"

echo [2/5] Packaging folder build (takes 1-3 minutes) ...
"%PY%" -m PyInstaller build_onedir.spec --noconfirm --clean || goto :fail
if not exist "dist\DocFormatTool\DocFormatTool.exe" goto :fail

echo [3/5] Flattening dist\DocFormatTool\* up to dist\ ...
rem The exe goes up one level and replaces the previous one. This is
rem the "overwrite the existing exe" step.
rem
rem These go through :movretry / :rmtree rather than plain move/rmdir.
rem When dist lives inside a cloud sync folder, the sync client can hold
rem handles on the files PyInstaller has just written, so the rename
rem fails with "access denied" and the build dies with a half-flattened
rem dist - new exe next to a stale _internal, which crashes at startup.
rem The retries wait the client out.
call :movretry "dist\DocFormatTool\DocFormatTool.exe" "dist\DocFormatTool.exe" || goto :fail

rem Support files must not be merged into a stale _internal - drop the
rem old tree first, then move the new one into place.
call :rmtree "dist\_internal" || goto :fail
call :movretry "dist\DocFormatTool\_internal" "dist\_internal" || goto :fail

rem Never ship personal settings.  The tool writes a ParaLib.json next to
rem the exe on every save (doc_parameters_manager.py, legacy backup), and
rem it contains history folder paths and window position. The authoritative
rem copy lives in %APPDATA%\AutoAdjustWordFormatV2 and is untouched here;
rem the tool recreates this file the next time settings are saved.
if exist "dist\DocFormatTool\ParaLib.json" del /q "dist\DocFormatTool\ParaLib.json"
if exist "dist\ParaLib.json" del /q "dist\ParaLib.json"

rem Ship the end-user readme alongside the exe. Its filename is Chinese,
rem which cannot appear in this file (see the ASCII-only note at the
rem top), so it lives in dist_extra\ and is copied in wholesale.
if exist "dist_extra" xcopy /y /q "dist_extra\*" "dist\" >nul

rem COLLECT should be empty now. rmdir without /s only removes an EMPTY
rem directory, so if the directory survives it means the cleanup above
rem missed something - fail loudly rather than shipping a directory the
rem tests never covered.
rem
rem DO NOT use a bare `find` / `findstr` / any external tool here. When
rem this script is launched from Git Bash, PATH puts /usr/bin ahead of
rem System32, so `find` resolves to MSYS find, which reads `/c` as the C:
rem drive and silently scans the entire disk. That hangs the build with
rem no output at all. rmdir is a cmd builtin, so it cannot be shadowed.
rmdir "dist\DocFormatTool" 2>nul
if exist "dist\DocFormatTool" goto :leftover

echo [4/5] Verifying ...
if not exist "dist\DocFormatTool.exe" goto :fail
if not exist "dist\_internal" goto :fail
rem The onedir stub is ~2.2 MB; anything much smaller means a bad build.
for %%F in ("dist\DocFormatTool.exe") do if %%~zF LSS 100000 goto :fail
powershell -NoProfile -Command "if ((Get-Item 'dist\DocFormatTool.exe').LastWriteTime.Ticks -le %STAMP%) { exit 1 }" || goto :stale

echo [5/5] Done. Output:
echo        dist\DocFormatTool.exe
echo        dist\_internal\
echo        dist\   (plus the readme copied from dist_extra\)
echo.
echo WARNING: ship BOTH together. The exe will not start without the
echo          _internal folder sitting next to it.
echo.
echo Verify on the target machine (no Python required):
echo        dist\DocFormatTool.exe --selftest
echo.
echo NOTE: the target machine needs Microsoft Word or WPS Writer
echo       installed - the tool drives it through COM automation.
exit /b 0

:leftover
echo.
echo *** dist\DocFormatTool\ was not empty after flattening ***
echo     Something is in there that this script does not know about.
echo     Refusing to delete it. Contents:
dir /b "dist\DocFormatTool"
exit /b 1

:stale
echo.
echo *** dist\DocFormatTool.exe was NOT regenerated (it is stale) ***
echo     If PyInstaller reported a named-pipe / WinError 5 failure, run
echo     this script with full (unsandboxed) access - see build_exe.cmd.
exit /b 1

:no_env
echo.
echo *** No build environment found at %VENV% ***
echo     Run build_exe.cmd once - it creates that same venv.
exit /b 1

:rmtree
rem %1 = directory to remove. Returns 0 once it is gone.
if not exist "%~1" exit /b 0
for /l %%N in (1,1,8) do (
    rmdir /s /q "%~1" 2>nul
    if not exist "%~1" exit /b 0
    powershell -NoProfile -Command "Start-Sleep -Milliseconds 1500" >nul 2>&1
)
echo *** could not remove %~1 - still locked ***
exit /b 1

:movretry
rem %1 = source, %2 = destination. Returns 0 once the source is gone.
rem The move is treated as successful only when the source disappears;
rem "access denied" just means a sync client still holds a handle.
for /l %%N in (1,1,8) do (
    move /y "%~1" "%~2" >nul 2>&1
    if not exist "%~1" exit /b 0
    powershell -NoProfile -Command "Start-Sleep -Milliseconds 1500" >nul 2>&1
)
echo *** could not move %~1 to %~2 - still locked ***
exit /b 1

:fail
echo.
echo *** BUILD FAILED - see the error message above ***
exit /b 1
