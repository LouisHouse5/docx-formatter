#!/bin/bash
set -e

REPO_URL="https://github.com/LouisHouse5/docx-formatter.git"
SKILL_NAME="docx-formatter"
SKILL_DIR="$HOME/.claude/skills/$SKILL_NAME"
TMP_DIR=$(mktemp -d)

echo "============================================"
echo "  Installing docx-formatter skill"
echo "============================================"

# 检查 python-docx
echo ""
echo "[1/4] Checking python-docx..."
if python3 -c "import docx" 2>/dev/null; then
    echo "  python-docx is already installed."
else
    echo "  Installing python-docx..."
    pip install python-docx
fi

# 克隆仓库
echo ""
echo "[2/4] Downloading from GitHub..."
git clone --depth 1 "$REPO_URL" "$TMP_DIR" 2>/dev/null || {
    echo "  Error: Failed to clone repository."
    echo "  Please check your internet connection or the repository URL."
    rm -rf "$TMP_DIR"
    exit 1
}

# 安装技能
echo ""
echo "[3/4] Installing skill to $SKILL_DIR..."
mkdir -p "$SKILL_DIR"
# 清理旧目录（升级场景）
rm -rf "$SKILL_DIR/scripts"
rm -rf "$SKILL_DIR/templates"
# 复制新文件
cp -r "$TMP_DIR/scripts" "$SKILL_DIR/"
cp "$TMP_DIR/SKILL.md" "$SKILL_DIR/"
# 可选：复制示例文件
if [ -d "$TMP_DIR/examples" ]; then
    cp -r "$TMP_DIR/examples" "$SKILL_DIR/"
fi

# 清理临时文件
rm -rf "$TMP_DIR"

# 验证安装
echo ""
echo "[4/4] Verifying installation..."
if [ -f "$SKILL_DIR/SKILL.md" ] && [ -d "$SKILL_DIR/scripts" ]; then
    echo "  Installation successful!"
    echo ""
    echo "  Skill location: $SKILL_DIR"
    echo "  Scripts:        $(ls "$SKILL_DIR/scripts"/*.py | wc -l) Python scripts installed"
    echo ""
    echo "  Usage:"
    echo "    1. In Claude Code: type /docx-format to trigger the skill"
    echo "    2. Directly: python3 $SKILL_DIR/scripts/analyze_template.py <template.docx>"
    echo ""
else
    echo "  Error: Installation verification failed."
    exit 1
fi

echo "============================================"
echo "  Done!"
echo "============================================"
