@echo off
title 科级干部全周期管理与综合研判系统 - 在线更新
cd /d "%~dp0"

echo =======================================================
echo     科级干部全周期管理与综合研判系统 - 在线更新
echo =======================================================
echo.
echo 正在检查 Git 环境并从 GitHub 拉取最新更新...
echo.

git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] 当前电脑未检测到 Git 命令行工具。
    echo 如果您之前是通过文件夹拷贝或下载压缩包的方式部署的：
    echo 1. 可直接将最新修改的 cadre_core.py 复制替换到当前目录；
    echo 2. 或访问 GitHub 仓库下载最新源码包覆盖：
    echo    https://github.com/lihm2430/cadre-management-system
    echo.
    pause
    exit /b
)

git pull origin main

if %errorlevel% equ 0 (
    echo.
    echo =======================================================
    echo   [更新成功] 系统已成功同步 GitHub 最新修复！
    echo   直接双击「启动系统.bat」即可开始使用。
    echo =======================================================
) else (
    echo.
    echo =======================================================
    echo   [更新提示] Git 拉取未完成，请检查网络连接或稍后重试。
    echo =======================================================
)
echo.
pause
