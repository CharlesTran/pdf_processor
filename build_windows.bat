@echo off
rem ============================================================
rem  PDF 处理工具 - Windows 手动打包脚本
rem  用法：双击运行，或在命令行执行 build_windows.bat
rem  产物：dist\PDF处理工具.exe
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] 安装依赖...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error

echo [2/3] 使用 PyInstaller 打包...
pyinstaller --noconfirm --clean pdf_tool.spec
if errorlevel 1 goto :error

echo [3/3] 打包完成！
echo.
echo 产物路径: %~dp0dist\PDF处理工具.exe
pause
exit /b 0

:error
echo.
echo 打包失败，请检查上方错误信息。
pause
exit /b 1
