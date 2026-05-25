---
name: docx-formatter
description: "以标准模板 docx 为基准，批量修复目标 docx 文件的所有格式（显式格式 + 隐藏格式），确保格式逐节、逐行、逐表完全一致。支持段落、表格、页面、分节、页眉页脚、目录域、样式定义等全维度对齐"
trigger: /docx-format
---

# Docx 格式对齐工具 (docx-formatter)

**Trigger**: `/docx-format`

**用途**：以标准模板 docx 文件为基准，批量修复目标 docx 文件的所有格式（显式格式 + 隐藏格式），确保格式**逐节、逐行、逐表**完全一致。

## 前置依赖

```bash
pip install python-docx
```

## 两种工作模式

| 模式 | 脚本 | 适用场景 | 原理 |
|------|------|---------|------|
| **深度拷贝** | `copy_format_deep.py` | 两份文档**结构一致**（相同段落数/表格数/section数） | 按索引逐段落 XML 级 deepcopy |
| **规则匹配** | `fix_docx_template.py` | 两份文档**结构不同**或批量处理 | 基于内容规则匹配段落类型 |

### 如何选择模式

```
两份文档结构是否一致？
├── 是 → copy_format_deep.py（推荐，零规则偏差）
└── 否 → fix_docx_template.py（需要调整 CONFIG 和 classify_and_format 规则）
```

**深度拷贝模式的优势**：
- 无需编写/调整内容匹配规则
- 直接复制模板的格式定义（pPr/rPr），保证 100% 对齐
- 通过 verify_docx.py 验证可达 0 段落差异

**规则匹配模式的优势**：
- 支持结构不同的文档（段落数不一致）
- 支持 `--batch-file` 批量处理
- 可通过 CONFIG 和 JSON 配置灵活调整

## 覆盖的格式维度

| 维度 | 显式格式 | 隐藏格式 |
|------|---------|---------|
| **段落** | 字体、字号、加粗、对齐、行距、段前段后、首行缩进 | 样式继承、编号列表格式、大纲级别 |
| **表格** | 单元格字体 | 边框样式、底纹、列宽、合并单元格、对齐方式 |
| **页面** | — | 纸张大小、方向、页边距、装订线、页眉页脚距边界距离 |
| **分节** | — | 分节符类型、页眉页脚链接到前一节、首页不同、奇偶页不同 |
| **页眉页脚** | 字体内容 | 页眉页脚段落格式、页码域、页数域 |
| **目录** | — | TOC 域代码、目录级别映射、页码显示、引导符 |
| **样式** | — | 文档中所有自定义样式的完整定义 |
| **其他** | 半角/全角标点 | 制表位、边框、底纹、保护设置 |

## 工作流程

### 方式 A：深度拷贝模式（结构一致时推荐）

```bash
# 0. 备份
cp 目标.docx 目标_backup.docx

# 1. 复制样式定义
python3 scripts/copy_styles.py 模板.docx 目标.docx

# 2. 复制页眉页脚
python3 scripts/copy_headers_footers.py 模板.docx 目标.docx

# 3. 深度格式拷贝（段落+表格+section 一体化）
python3 scripts/copy_format_deep.py 模板.docx 目标.docx

# 4. 验证
python3 scripts/verify_docx.py 目标.docx 模板.docx
```

### 方式 B：规则匹配模式（结构不同时使用）

```bash
# 1. 深度扫描模板
python3 scripts/analyze_template.py 模板文件.docx > template_report.txt

# 2. 全面审核目标文件
python3 scripts/audit_docx.py 目标文件.docx 模板文件.docx

# 3. 复制样式 + 页眉页脚
python3 scripts/copy_styles.py 模板.docx 目标.docx
python3 scripts/copy_headers_footers.py 模板.docx 目标.docx

# 4. 精确修复
python3 scripts/fix_docx_template.py 目标文件.docx --template 模板文件.docx

# 5. 最终验证
python3 scripts/verify_docx.py 目标文件.docx 模板文件.docx
```

### 批量处理（仅规则匹配模式）

```bash
python3 scripts/fix_docx_template.py \
  --batch-file files.txt \
  --template 模板文件.docx
```

## 关键脚本说明

| 脚本 | 作用 | 是否需要修改 |
|------|------|-------------|
| `copy_format_deep.py` | **深度格式拷贝**（按索引 XML 级 deepcopy） | 否 |
| `analyze_template.py` | **深度扫描**模板所有格式（显式+隐藏） | 否 |
| `audit_docx.py` | 全面对比目标与模板差异 | 否 |
| `fix_docx_template.py` | 规则匹配修复（含隐藏格式） | **是**（CONFIG 和 classify_and_format） |
| `verify_docx.py` | 多维度最终验证 | 否 |
| `copy_styles.py` | 将模板样式复制到目标文件 | 否 |
| `copy_headers_footers.py` | 将模板页眉页脚复制到目标文件 | 否 |
| `utils.py` | 公共工具模块（EMU换算、字体设置、XML操作等） | 否 |

## EMU 换算速查

- `1 pt = 12700 EMU`
- `1 英寸 = 914400 EMU`
- 字号：小四=12pt=`152400`, 三号=16pt=`203200`, 小二=18pt=`228600`, 小初=24pt=`304800`
- 缩进：两字符≈`304800`~`306070`
- 页边距：1英寸=`914400`, 1.25英寸=`1143000`

## 注意事项

1. **务必先备份目标文件**：脚本会直接覆盖保存
2. **bold=None vs False**：`None` 表示继承样式（模板常用），`False` 表示显式不加粗
3. **样式优先级**：直接格式 > 样式定义 > 默认样式。修复时两者都要对齐
4. **页眉页脚复制**：`copy_headers_footers.py` 会覆盖目标文件的所有页眉页脚，谨慎使用
5. **目录域**：自动复制的 TOC 域需在 Word 中右键目录 → "更新域" 才能刷新页码
6. **section break 保护**：`remove_empty_paragraphs()` 已修复 pPr 嵌套 sectPr 的检测，不会再误删含分节符的空段落
7. **封面标题检测**：`classify_and_format()` 的封面标题规则已支持非首行（如"附件6"开头的文档）

## 文件结构

```
~/.claude/skills/docx-formatter/
├── SKILL.md
├── .gitignore
├── examples/
│   ├── template.docx          # 示例模板文件
│   ├── target.docx            # 示例目标文件（修复前）
│   ├── batch_config.json      # 批量处理配置示例
│   └── README.md              # 示例使用说明
├── scripts/
│   ├── copy_format_deep.py    # 深度格式拷贝（新增）
│   ├── analyze_template.py    # 深度扫描模板
│   ├── audit_docx.py          # 全面对比差异
│   ├── fix_docx_template.py   # 规则匹配修复
│   ├── verify_docx.py         # 多维度验证
│   ├── copy_styles.py         # 复制样式定义
│   ├── copy_headers_footers.py # 复制页眉页脚
│   └── utils.py               # 公共工具模块
└── tests/
    ├── test_utils.py          # 工具函数测试
    ├── test_fix_quotes.py     # 引号修复测试
    ├── test_table_borders.py  # 表格边框测试
    ├── test_integration.py    # 集成测试
    └── run_tests.sh           # 测试运行脚本
```

## 进阶用法

### 复制模板样式到目标文件

```bash
python3 scripts/copy_styles.py 模板文件.docx 目标文件.docx
```

### 复制模板页眉页脚到目标文件

```bash
python3 scripts/copy_headers_footers.py 模板文件.docx 目标文件.docx
```

### 使用 JSON 配置文件

```bash
python3 scripts/fix_docx_template.py \
  --config config.json \
  --template 模板文件.docx
```
