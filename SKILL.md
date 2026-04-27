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

### 阶段 1：深度扫描模板

```bash
python3 scripts/analyze_template.py 模板文件.docx > template_report.txt
```

输出包含：
- 每种段落类型的精确格式参数
- **每个 Section** 的页面设置（纸张、页边距、方向）
- **每个页眉/页脚** 的内容和格式
- **所有样式定义**（Normal、Heading1、TOC1 等）
- **每个表格** 的边框、列宽、单元格对齐
- **目录域代码** 分析

### 阶段 2：全面审核目标文件

```bash
python3 scripts/audit_docx.py 目标文件.docx 模板文件.docx
```

输出：
- 段落格式差异（逐行对比）
- 页面/分节差异
- 页眉页脚差异
- 样式缺失/不一致
- 表格结构差异
- 目录格式差异

### 阶段 3：精确修复

```bash
# 单文件修复
python3 scripts/fix_docx_template.py 目标文件.docx 模板文件.docx

# 或指定模板路径（推荐）
python3 scripts/fix_docx_template.py 目标文件.docx --template 模板文件.docx
```

修复内容包括：
1. 所有段落的字体统一（含东亚字体 `w:eastAsia`）
2. 段落格式匹配（对齐、行距、段前段后、首行缩进、样式应用）
3. **页面设置按分节自动同步**（纸张、页边距、方向、页眉页脚距边界 —— 自动从模板读取）
4. **页眉页脚**内容同步（可选：复制模板页眉页脚）
5. **样式定义**同步（复制模板样式到目标文件）
6. **表格格式自动同步**（边框从模板逐表复制、字体、对齐）
7. **目录域自动同步**（从模板复制 TOC 域代码，目标已有则跳过）
8. 半角引号转全角（含单双引号智能开闭匹配）
9. 删除多余空行/段落

### 批量处理

```bash
# 从文件列表批量处理
python3 scripts/fix_docx_template.py \
  --batch-file files.txt \
  --template 模板文件.docx

# 使用 JSON 配置文件
python3 scripts/fix_docx_template.py \
  --config config.json \
  --template 模板文件.docx
```

`files.txt` 每行一个目标文件路径；`config.json` 可覆盖 CONFIG 中的参数（如字体、字号、缩进值等）。

### 阶段 4：最终验证

```bash
python3 scripts/verify_docx.py 目标文件.docx 模板文件.docx
```

理想输出：
```
[段落验证] 共发现 0 处差异
[分节验证] 共发现 0 处差异
[页眉页脚验证] 共发现 0 处差异
[样式验证] 共发现 0 处差异
[表格验证] 共发现 0 处差异
========================================
全部验证通过！目标文件与模板格式完全一致。
```

## 关键脚本说明

| 脚本 | 作用 | 是否需要修改 |
|------|------|-------------|
| `analyze_template.py` | **深度扫描**模板所有格式（显式+隐藏） | 否 |
| `audit_docx.py` | 全面对比目标与模板差异 | 否 |
| `fix_docx_template.py` | 精确修复脚本（含隐藏格式） | **是**（CONFIG 和 classify_and_format） |
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
│   ├── analyze_template.py    # 深度扫描模板
│   ├── audit_docx.py          # 全面对比差异
│   ├── fix_docx_template.py   # 精确修复脚本
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

如果目标文件缺少模板中的自定义样式（如 `toc 1`、`toc 2`、`Heading1` 等）：

```bash
python3 scripts/copy_styles.py 模板文件.docx 目标文件.docx
```

### 复制模板页眉页脚到目标文件

```bash
python3 scripts/copy_headers_footers.py 模板文件.docx 目标文件.docx
```

### 完整流程示例

```bash
# 1. 分析模板（一次性，了解模板格式参数）
python3 scripts/analyze_template.py 模板.docx > template_report.txt

# 2. 审核目标文件
python3 scripts/audit_docx.py 目标.docx 模板.docx

# 3. 复制样式（如目标文件缺少模板样式）
python3 scripts/copy_styles.py 模板.docx 目标.docx

# 4. 复制页眉页脚（如需同步）
python3 scripts/copy_headers_footers.py 模板.docx 目标.docx

# 5. 修复格式（支持命令行参数直接传入）
python3 scripts/fix_docx_template.py 目标.docx --template 模板.docx
# 脚本会自动：
#   - 按段落内容分类并应用对应格式
#   - 从模板自动同步表格边框、页面设置、目录域
#   - 修复引号、删除空行

# 6. 最终验证
python3 scripts/verify_docx.py 目标.docx 模板.docx
```
