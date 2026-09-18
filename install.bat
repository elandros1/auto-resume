@echo off
REM Auto-Resume installer for Windows
REM Usage: curl -fsSL https://raw.githubusercontent.com/elandros1/auto-resume/main/install.bat -o install.bat && cmd /c install.bat

setlocal
set REPO=https://github.com/elandros1/auto-resume

echo ================================
echo   Auto-Resume 安装程序
echo ================================
echo.

REM Check Python (single-line if to avoid pipe issues)
where python >nul 2>nul
if errorlevel 1 goto :no_python

echo [1/3] 正在从 GitHub 安装...
pip install "git+%REPO%.git@main" -q
if errorlevel 1 goto :pip_error

echo.
echo [2/3] 验证安装...
python -m auto_resume --help >nul 2>nul
if errorlevel 1 goto :verify_error

echo [3/3] 安装完成!
echo.
echo ================================
echo   安装成功!
echo ================================
echo.
echo 快速开始:
echo   auto-resume new my_resume.json
echo   auto-resume generate -o templates\
echo   auto-resume fill -r my_resume.json -t 学校表格.docx -o output\
echo.
echo AI模式 (适配任何高校模板):
echo   set AI_API_KEY=sk-your-key
echo   auto-resume fill -r my_resume.json -t 模板.docx --ai
echo.
goto :eof

:no_python
echo 错误: 未找到 python，请先安装 Python 3.8+
echo 下载地址: https://www.python.org/downloads/
goto :eof

:pip_error
echo 错误: pip 安装失败，请检查网络连接
echo 手动安装: pip install git+https://github.com/elandros1/auto-resume.git
goto :eof

:verify_error
echo 警告: 安装可能未完全成功，请手动验证
echo 运行: python -m auto_resume --help
goto :eof
