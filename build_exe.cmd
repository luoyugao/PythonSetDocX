@echo off
chcp 936 >nul
setlocal
rem ============================================================
rem  打包单文件 EXE（无需 Python 环境即可运行）
rem  产物： dist\DocFormatTool.exe
rem
rem  构建用的虚拟环境刻意放在项目之外，避免被 360 云盘同步上传。
rem  如环境不存在，本脚本会自动创建。
rem
rem  注意：本文件必须保存为 GBK 编码。cmd.exe 按控制台代码页
rem  解析批处理，若存成 UTF-8 会出现行边界错位、把注释当命令执行。
rem ============================================================

set "VENV=%USERPROFILE%\.venvs\docformattool-build"
set "PY=%VENV%\Scripts\python.exe"
cd /d "%~dp0"

if not exist "%PY%" (
    echo [1/3] 未找到构建环境，正在创建 ...
    uv venv --python 3.12 "%VENV%" || goto :fail
    uv pip install --python "%PY%" pyinstaller pywin32 || goto :fail
) else (
    echo [1/3] 构建环境已存在：%VENV%
)

echo [2/3] 正在打包 ...
"%PY%" -m PyInstaller build_onefile.spec --noconfirm --clean || goto :fail

echo [3/3] 完成！产物：
echo        dist\DocFormatTool.exe
echo.
echo 注意：目标电脑需已安装 Microsoft Word 或 WPS 文字（程序通过 COM 调用）。
exit /b 0

:fail
echo.
echo *** 打包失败，请检查上方错误信息 ***
exit /b 1
