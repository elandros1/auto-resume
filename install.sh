#!/usr/bin/env bash
# Auto-Resume installer — one command, no PyPI needed
# Usage: curl -fsSL https://raw.githubusercontent.com/elandros1/auto-resume/main/install.sh | bash

set -e

REPO="https://github.com/elandros1/auto-resume"
INSTALL_DIR="${1:-$HOME/.local/bin}"
TEMPLATE_DIR="$HOME/.auto-resume/templates"

echo "================================"
echo "  Auto-Resume 安装程序"
echo "================================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 python3，请先安装 Python 3.8+"
    exit 1
fi

PY_VERSION=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "✓ Python $PY_VERSION"

# Install from GitHub directly
echo "→ 正在从 GitHub 安装..."
pip install --user "git+$REPO.git@main" -q 2>&1 || pip install "git+$REPO.git@main" -q 2>&1

# Copy templates to user directory
mkdir -p "$TEMPLATE_DIR"
PACKAGE_DIR=$(python3 -c "import auto_resume; import os; print(os.path.dirname(auto_resume.__file__))" 2>/dev/null)
if [ -n "$PACKAGE_DIR" ] && [ -d "$PACKAGE_DIR/templates" ]; then
    cp -r "$PACKAGE_DIR/templates/"* "$TEMPLATE_DIR/" 2>/dev/null || true
    echo "✓ 模板已复制到 $TEMPLATE_DIR"
fi

# Check if auto-resume is on PATH
if ! command -v auto-resume &> /dev/null; then
    echo ""
    echo "⚠ auto-resume 已安装但不在 PATH 中"
    echo "  请将以下路径添加到 PATH:"
    echo "  export PATH=\"\$HOME/.local/bin:\$PATH\""
    echo "  添加到 ~/.bashrc 或 ~/.zshrc 后重启终端"
fi

echo ""
echo "================================"
echo "  安装完成!"
echo "================================"
echo ""
echo "快速开始:"
echo "  auto-resume new my_resume.json      # 创建简历模板"
echo "  auto-resume generate -o templates/  # 生成模板"
echo "  auto-resume fill -r my_resume.json -t 学校表格.docx -o output/"
echo ""
