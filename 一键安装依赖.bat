@echo off
title 科级干部全周期管理与综合研判系统 - 环境与依赖一键安装
cd /d "%~dp0"

echo =======================================================
echo     科级干部全周期管理与综合研判系统
echo               Windows 依赖环境一键安装
echo =======================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python 环境！
    echo 请先安装 Python (建议 3.10 到 3.12 版本)。
    echo 安装时请务必勾选 "Add python.exe to PATH" 添加到环境变量。
    echo.
    pause
    exit /b
)

echo [1/2] 正在检测并升级 pip...
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo [2/2] 正在通过清华源高速安装系统所需依赖库...
echo 依赖列表: PySide6 openpyxl python-docx pypdf rapidocr_onnxruntime
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

if %errorlevel% equ 0 (
    echo.
    echo =======================================================
    echo   [成功] 所有核心依赖已安装完毕！
    echo   您现在可以双击运行【启动系统.bat】直接开启系统。
    echo =======================================================
) else (
    echo.
    echo =======================================================
    echo   [提示] 安装遇到网络波动，可重新双击本脚本重试。
    echo =======================================================
)

echo.
pause
