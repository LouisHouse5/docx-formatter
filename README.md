# docx-formatter

以标准模板 docx 为基准，批量修复目标 docx 的所有格式（显式格式 + 隐藏格式），确保格式逐节、逐行、逐表完全一致。

[![Trigger](https://img.shields.io/badge/trigger-%2Fdocx--format-blue)](https://github.com/LouisHouse5/docx-formatter)

## 适用场景

- 批量制作标准化文档（教案、课程标准、报告等）
- 将已有文档的格式统一到指定模板
- 自动化修复 Word 格式差异

## 安装

### 方式一：一键安装脚本

```bash
curl -sL https://raw.githubusercontent.com/LouisHouse5/docx-formatter/main/install.sh | bash
```

### 方式二：手动安装

```bash
# 1. 克隆仓库
git clone https://github.com/LouisHouse5/docx-formatter.git /tmp/docx-formatter

# 2. 复制到 Claude Code 技能目录
mkdir -p ~/.claude/skills/docx-formatter
cp -r /tmp/docx-formatter/scripts /tmp/docx-formatter/SKILL.md ~/.claude/skills/docx-formatter/

# 3. 安装 Python 依赖
pip install python-docx
```

## 依赖

```bash
pip install python-docx
```

## 快速开始

在 Claude Code 中输入 `/docx-format` 触发技能，或直接使用脚本：

```bash
# 1. 分析模板
python3 scripts/analyze_template.py 模板.docx > template_report.txt

# 2. 审核目标文件差异
python3 scripts/audit_docx.py 目标.docx 模板.docx

# 3. 修复格式（编辑 CONFIG 后运行）
cp scripts/fix_docx_template.py fix_target.py
# 修改 fix_target.py 中的 TARGET 和 TEMPLATE 文件名
python3 fix_target.py

# 4. 验证
python3 scripts/verify_docx.py 目标.docx 模板.docx
```

详见 [SKILL.md](SKILL.md) 获取完整文档。

## 文件结构

```
docx-formatter/
├── SKILL.md              # 技能主文档
├── README.md             # 项目说明
├── install.sh            # 一键安装脚本
├── .gitignore
├── examples/             # 示例文件
│   ├── template.docx
│   └── target.docx
└── scripts/              # Python 工具脚本
    ├── analyze_template.py
    ├── audit_docx.py
    ├── fix_docx_template.py
    ├── verify_docx.py
    ├── copy_styles.py
    └── copy_headers_footers.py
```

## 覆盖的格式维度

| 维度 | 显式格式 | 隐藏格式 |
|------|---------|---------|
| **段落** | 字体、字号、加粗、对齐、行距、段前段后、首行缩进 | 样式继承、编号列表格式、大纲级别 |
| **表格** | 单元格字体 | 边框样式、底纹、列宽、合并单元格、对齐方式 |
| **页面** | — | 纸张大小、方向、页边距、页眉页脚距边界 |
| **分节** | — | 分节符类型、首页不同、奇偶页不同 |
| **页眉页脚** | 字体内容 | 段落格式、页码域、页数域 |
| **目录** | — | TOC 域代码、目录级别映射 |
| **样式** | — | 自定义样式完整定义 |
| **其他** | 半角/全角标点 | 制表位、边框、底纹 |

## 许可证

MIT License
