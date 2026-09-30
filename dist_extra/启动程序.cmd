@echo off
chcp 936 >nul
setlocal
set "SRC=%~dp0"
set "DST=%LOCALAPPDATA%\DocFormatTool"

if not exist "%SRC%DocFormatTool.exe" goto :no_exe
if not exist "%SRC%_internal\" goto :no_internal

echo.
echo   正在准备运行环境 ...
echo   第一次运行需要把程序复制到本机目录，大约几秒钟，请稍等。
echo.
robocopy "%SRC%." "%DST%" /e /njh /njs /ndl /nc /ns /np /r:1 /w:1 /xf ParaLib.json *.cmd *.txt >nul
if errorlevel 8 goto :copy_fail
if not exist "%DST%\DocFormatTool.exe" goto :copy_fail

if "%~1"=="" (
    start "" "%DST%\DocFormatTool.exe"
    exit /b 0
)
"%DST%\DocFormatTool.exe" %*
exit /b %errorlevel%

:no_exe
echo.
echo   [错误] 没有找到 DocFormatTool.exe
echo   请确认本文件和 DocFormatTool.exe、_internal 文件夹在同一个目录里。
echo.
pause
exit /b 1

:no_internal
echo.
echo   [错误] 没有找到 _internal 文件夹
echo   DocFormatTool.exe 和 _internal 是一个整体，必须一起拷贝、放在同一目录下。
echo   只拷 exe 是跑不起来的。
echo.
pause
exit /b 1

:copy_fail
echo.
echo   [错误] 把程序复制到本机目录时失败。
echo   请检查磁盘剩余空间，或右键本文件 -^> 以管理员身份运行。
echo.
pause
exit /b 1
