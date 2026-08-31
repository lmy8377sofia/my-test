@echo off
chcp 65001 >nul
rem ============================================================
rem  PDF 电子盖章工具 —— 一键启动脚本（Windows）
rem  自动检查依赖，缺失则安装，然后启动程序（无控制台黑框）
rem ============================================================
setlocal

cd /d "%~dp0"

rem ---- 1. 找 Python（优先本目录便携版，其次系统 python）----
set "PY=python"
if exist "portable_python\python.exe" set "PY=%CD%\portable_python\python.exe"
if exist "python\python.exe" set "PY=%CD%\python\python.exe"

"%PY%" --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python。请先安装 Python 3.10 或更高版本：
    echo         https://www.python.org/downloads/windows/
    echo         安装时请勾选 "Add python.exe to PATH"。
    echo.
    pause
    exit /b 1
)

rem ---- 2. 检查并安装依赖 ----
"%PY%" -c "import pymupdf, PIL, numpy" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖（首次运行需要联网，约 1-2 分钟）...
    "%PY%" -m pip install --upgrade pip -q
    "%PY%" -m pip install -r "%~dp0requirements.txt" -q
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络后重试。
        pause
        exit /b 1
    )
)

rem ---- 3. 启动（pythonw 无黑框）----
if exist "%PY%\..\pythonw.exe" (
    start "" "%PY%\..\pythonw.exe" "%~dp0pdf_stamp.py"
) else (
    start "" pythonw "%~dp0pdf_stamp.py" 2>nul
    if errorlevel 1 start "" "%PY%" "%~dp0pdf_stamp.py"
)
endlocal
