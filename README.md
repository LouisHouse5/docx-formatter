# docx-formatter

![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)
![python-docx](https://img.shields.io/badge/library-python--docx-4B8BBE?labelColor=2a2a2a)

[中文](#中文) | [English](#english)

---

<a id="中文"></a>

以标准模板 docx 为基准，批量修复目标 docx 文件的所有格式（显式格式 + 隐藏格式），确保格式**逐节、逐行、逐表**完全一致。

## 功能特性

- **全维度格式对齐**：段落、表格、页面、分节、页眉页脚、目录域、样式定义
- **批量处理**：支持文件列表批量处理，支持 JSON 配置文件
- **智能引号修复**：半角引号自动转全角，智能匹配开闭引号
- **东亚字体完整支持**：自动设置 `w:eastAsia` 属性，确保中文排版正确
- **自动化验证**：修复后自动多维度验证，确保与模板完全一致

## 前置依赖

```bash
pip install python-docx
```

## 快速开始

```bash
# 1. 分析模板格式
python3 scripts/analyze_template.py examples/template.docx

# 2. 审核目标文件与模板的差异
python3 scripts/audit_docx.py examples/target.docx examples/template.docx

# 3. 修复目标文件格式
python3 scripts/fix_docx_template.py examples/target.docx --template examples/template.docx

# 4. 验证修复结果
python3 scripts/verify_docx.py examples/target.docx examples/template.docx
```

## 批量处理

```bash
# 从文件列表批量处理
python3 scripts/fix_docx_template.py \
  --batch-file files.txt \
  --template template.docx

# 使用 JSON 配置文件
python3 scripts/fix_docx_template.py \
  --config config.json \
  --template template.docx
```

## 在 AI Agent 中安装

### Claude Code

```bash
# 克隆到用户级 skills 目录（全局可用）
git clone https://github.com/LouisHouse5/docx-formatter.git ~/.claude/skills/docx-formatter

# 或克隆到项目级目录（仅当前项目可用）
git clone https://github.com/LouisHouse5/docx-formatter.git .claude/skills/docx-formatter
```

### Cursor

将 [SKILL.md](SKILL.md) 内容复制到项目规则文件：

```bash
mkdir -p .cursor/rules
cp SKILL.md .cursor/rules/docx-formatter.mdc
```

### Cline (VS Code)

```bash
# 添加到项目根目录的自定义指令文件
cat SKILL.md >> cline-instructions.md
```

### GitHub Copilot

```bash
mkdir -p .github
cp SKILL.md .github/copilot-instructions.md
```

### Gemini CLI

```bash
# 用户级（全局可用）
mkdir -p ~/.gemini
cat SKILL.md >> ~/.gemini/GEMINI.md

# 或项目级
cat SKILL.md >> GEMINI.md
```

### OpenAI Codex CLI

```bash
# 用户级
mkdir -p ~/.codex
cat SKILL.md >> ~/.codex/AGENTS.md

# 或项目级
cat SKILL.md >> AGENTS.md
```

### Windsurf (Codeium)

将 SKILL.md 内容添加到项目 `.windsurfrules` 文件：

```bash
cat SKILL.md >> .windsurfrules
```

## 项目结构

```
docx-formatter/
├── scripts/
│   ├── analyze_template.py      # 深度扫描模板格式
│   ├── audit_docx.py            # 全面对比差异
│   ├── fix_docx_template.py     # 精确修复格式
│   ├── verify_docx.py           # 多维度验证
│   ├── copy_styles.py           # 复制样式定义
│   ├── copy_headers_footers.py  # 复制页眉页脚
│   └── utils.py                 # 公共工具模块
├── examples/
│   ├── template.docx            # 示例模板
│   ├── target.docx              # 示例目标文件
│   ├── batch_config.json        # 批量配置示例
│   └── README.md                # 示例说明
├── tests/
│   ├── test_utils.py            # 工具函数测试
│   ├── test_fix_quotes.py       # 引号修复测试
│   ├── test_table_borders.py    # 表格边框测试
│   ├── test_integration.py      # 集成测试
│   └── run_tests.sh             # 测试运行脚本
├── SKILL.md                     # Skill 定义文件（Agent 指令）
└── README.md
```

## Agent 使用示例

在 Claude Code / Cursor / Copilot 等 AI Agent 中，直接用自然语言描述格式修复需求：

**示例 1 — 单文件格式修复**
```
> 把 report.docx 的格式修复成和 template.docx 一样

Agent 会运行：
  python3 scripts/fix_docx_template.py report.docx --template template.docx
  → 段落、表格、页眉页脚、样式全部对齐模板
```

**示例 2 — 先审核再修复**
```
> 先对比 draft.docx 和 template.docx 的格式差异，确认后再修复

Agent 会分步执行：
  1. python3 scripts/audit_docx.py draft.docx template.docx  → 输出差异报告
  2. 用户确认后：python3 scripts/fix_docx_template.py draft.docx --template template.docx
```

**示例 3 — 批量处理 + 验证**
```
> 把 ./docs 目录里所有 docx 文件批量修复格式，用 template.docx 做模板，修完自动验证

Agent 会生成批量处理脚本：
  1. 遍历目录收集所有 .docx 文件
  2. 逐个调用 fix_docx_template.py 修复
  3. 逐个调用 verify_docx.py 验证结果
  4. 输出修复报告（成功/失败/跳过）
```

**示例 4 — 分析模板格式**
```
> 帮我分析一下 template.docx 里用了哪些格式设置，我要了解模板的详细规格

Agent 会运行：
  python3 scripts/analyze_template.py template.docx
  → 输出字体、字号、行距、页面边距、表格样式等完整格式规格
```

## 运行测试

```bash
cd tests
./run_tests.sh
```

## 详细文档

完整的使用说明、EMU 换算速查、注意事项等，请参阅 [SKILL.md](SKILL.md)。

---

<a id="english"></a>

Batch-fix all formatting (explicit + hidden) in target docx files to match a standard template — section by section, line by line, table by table.

## Features

- **Full-dimension alignment**: Paragraphs, tables, pages, sections, headers/footers, TOC fields, style definitions
- **Batch processing**: File list batch processing, JSON config support
- **Smart quote fix**: Half-width quotes auto-converted to full-width, intelligent open/close matching
- **East Asian font support**: Automatically sets `w:eastAsia` attribute for correct CJK typography
- **Automated verification**: Post-fix multi-dimensional verification against template

## Prerequisites

```bash
pip install python-docx
```

## Quick Start

```bash
# 1. Analyze template formatting
python3 scripts/analyze_template.py examples/template.docx

# 2. Audit differences between target and template
python3 scripts/audit_docx.py examples/target.docx examples/template.docx

# 3. Fix target file formatting
python3 scripts/fix_docx_template.py examples/target.docx --template examples/template.docx

# 4. Verify fix results
python3 scripts/verify_docx.py examples/target.docx examples/template.docx
```

## Batch Processing

```bash
# Batch process from file list
python3 scripts/fix_docx_template.py \
  --batch-file files.txt \
  --template template.docx

# Use JSON config file
python3 scripts/fix_docx_template.py \
  --config config.json \
  --template template.docx
```

## Install in AI Agents

### Claude Code

```bash
# Global (available in all projects)
git clone https://github.com/LouisHouse5/docx-formatter.git ~/.claude/skills/docx-formatter

# Project-level (current project only)
git clone https://github.com/LouisHouse5/docx-formatter.git .claude/skills/docx-formatter
```

### Cursor

Copy [SKILL.md](SKILL.md) into project rules:

```bash
mkdir -p .cursor/rules
cp SKILL.md .cursor/rules/docx-formatter.mdc
```

### Cline (VS Code)

```bash
cat SKILL.md >> cline-instructions.md
```

### GitHub Copilot

```bash
mkdir -p .github
cp SKILL.md .github/copilot-instructions.md
```

### Gemini CLI

```bash
# User-level (global)
mkdir -p ~/.gemini
cat SKILL.md >> ~/.gemini/GEMINI.md

# Project-level
cat SKILL.md >> GEMINI.md
```

### OpenAI Codex CLI

```bash
# User-level
mkdir -p ~/.codex
cat SKILL.md >> ~/.codex/AGENTS.md

# Project-level
cat SKILL.md >> AGENTS.md
```

### Windsurf (Codeium)

Add SKILL.md content to `.windsurfrules`:

```bash
cat SKILL.md >> .windsurfrules
```

## Project Structure

```
docx-formatter/
├── scripts/
│   ├── analyze_template.py      # Deep-scan template formatting
│   ├── audit_docx.py            # Full diff comparison
│   ├── fix_docx_template.py     # Precise format fix
│   ├── verify_docx.py           # Multi-dimensional verification
│   ├── copy_styles.py           # Copy style definitions
│   ├── copy_headers_footers.py  # Copy headers and footers
│   └── utils.py                 # Shared utilities
├── examples/
│   ├── template.docx            # Sample template
│   ├── target.docx              # Sample target file
│   ├── batch_config.json        # Batch config example
│   └── README.md                # Example docs
├── tests/
│   ├── test_utils.py            # Utility tests
│   ├── test_fix_quotes.py       # Quote fix tests
│   ├── test_table_borders.py    # Table border tests
│   ├── test_integration.py      # Integration tests
│   └── run_tests.sh             # Test runner script
├── SKILL.md                     # Skill definition (Agent instructions)
└── README.md
```

## Run Tests

```bash
cd tests
./run_tests.sh
```

## Agent Usage Examples

In AI agents like Claude Code / Cursor / Copilot, describe your formatting fix in natural language:

**Example 1 — Single File Format Fix**
```
> Fix the formatting in report.docx to match template.docx

Agent runs:
  python3 scripts/fix_docx_template.py report.docx --template template.docx
  → Paragraphs, tables, headers/footers, styles all aligned to template
```

**Example 2 — Audit Before Fixing**
```
> First compare draft.docx against template.docx formatting, then fix after I confirm

Agent executes in two steps:
  1. python3 scripts/audit_docx.py draft.docx template.docx  → outputs diff report
  2. After user confirms: python3 scripts/fix_docx_template.py draft.docx --template template.docx
```

**Example 3 — Batch Process + Verify**
```
> Batch-fix all docx files in ./docs using template.docx, then auto-verify each one

Agent generates a batch processing script:
  1. Scan directory for all .docx files
  2. Run fix_docx_template.py on each
  3. Run verify_docx.py on each result
  4. Output summary (success/fail/skipped)
```

**Example 4 — Analyze Template Formatting**
```
> Analyze template.docx and show me all the formatting settings used

Agent runs:
  python3 scripts/analyze_template.py template.docx
  → Outputs fonts, sizes, line spacing, margins, table styles, and more
```

## Full Documentation

For complete usage guide, EMU conversion reference, and caveats, see [SKILL.md](SKILL.md).
