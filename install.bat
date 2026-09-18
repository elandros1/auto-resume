@echo off
REM Auto-Resume installer for Windows
REM Usage: curl -fsSL https://raw.githubusercontent.com/elandros1/auto-resume/main/install.bat | cmd

set REPO=https://github.com/elandros1/auto-resume

echo ================================
echo   Auto-Resume 安装程序
echo ================================
echo.

REM Check Python
where python >nul 2>nul
if errorlevel 1 (
    echo 错误: 未找到 python，请先安装 Python 3.8+
    exit /b 1
)

echo → 正在从 GitHub 安装...
pip install "git+%REPO%.git@main" -q

echo.
echo ================================
echo   安装完成!
echo ================================
echo.
echo 快速开始:
echo   auto-resume new my_resume.json
echo   auto-resume generate -o templates\
echo   auto-resume fill -r my_resume.json -t 学校表格.docx -o output\
echo.
