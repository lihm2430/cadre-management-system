@echo off
title 打包干部管理系统为独立 EXE
cd /d "%~dp0"

echo =======================================================
echo          科级干部全周期管理与综合研判系统 - PyInstaller 打包程序
echo =======================================================
echo.
echo 正在安装/检测 PyInstaller 打包工具...
python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo 正在开始单文件打包 (预计耗时 1-3 分钟，请稍候)...
pyinstaller --noconfirm --clean --onefile --windowed --name "干部任免检索系统" --add-data "XXX干部任免审批表(模板).docx;." --hidden-import "cadre_core" --hidden-import "rapidocr_onnxruntime" --hidden-import "openpyxl" --hidden-import "docx" --hidden-import "pypdf" --hidden-import "PySide6.QtPrintSupport" "干部任免检索系统.py"

echo.
if %errorlevel% equ 0 (
    echo =======================================================
    echo   [打包成功] 生成的新版单文件 exe 位于: dist\干部任免检索系统.exe
    echo.
    echo   【无缝替换升级步骤】：
    echo   1. 将旧目录下的【干部任免数据索引.db】先复制一份作为备份。
    echo   2. 将本项目 dist 目录下的【干部任免检索系统.exe】复制过去，直接覆盖旧的 exe。
    echo   3. 双击运行新 exe，系统将自动对原有数据库进行字段无损平滑升级！
    echo =======================================================
) else (
    echo [打包失败] 请检查上方报错信息。
)

echo.
pause
