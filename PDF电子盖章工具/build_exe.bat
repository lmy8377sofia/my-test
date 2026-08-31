@echo off
chcp 65001 >nul
rem ============================================================
rem  PDF 电子盖章工具 —— 一键打包 exe（Windows）
rem  需要：已安装 Python 3.10+ 并勾选 Add to PATH
rem  产出：dist\PDF电子盖章工具.exe
rem ============================================================
setlocal
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

echo [1/3] 安装打包依赖（PyInstaller）...
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt -q
python -m pip install pyinstaller -q
if errorlevel 1 (
    echo [错误] 依赖安装失败。
    pause
    exit /b 1
)

echo [2/3] 运行自检（预处理 / 自动盖章 / 导出）...
python pdf_stamp.py --selftest
if errorlevel 1 (
    echo [警告] 自检未全部通过，仍将继续打包，请检查上方输出。
)

echo [3/3] PyInstaller 打包...
rem  --onefile 单文件  --windowed 无控制台黑框  --name 输出名
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "PDF电子盖章工具" ^
    --icon "assets\app.ico" ^
    pdf_stamp.py
if errorlevel 1 (
    echo [错误] 打包失败。
    pause
    exit /b 1
)

echo.
echo 打包完成：dist\PDF电子盖章工具.exe
echo 双击即可运行（免安装、绿色便携）。也可以把 exe 单独拷走使用。
pause
endlocal
