# docx-formatter

以标准模板 docx 为基准，批量修复目标 docx 文件的所有格式（显式格式 + 隐藏格式），确保格式**逐节、逐行、逐表**完全一致。

## 功能特性

- **全维度格式对齐**：段落、表格、页面、分节、页眉页脚、目录域、样式定义
- **批量处理**：支持文件列表批量处理，支持 JSON 配置文件
- **智能引号修复**：半角引号自动转全角，智能匹配开闭引号
- **东亚字体完整支持**：自动设置 `w:eastAsia` 属性，确保中文排版正确
- **自动化验证**：修复后自动多维度验证，确保与模板完全一致

## 安装

```bash
pip install python-docx
```

## 快速开始

```bash
cd ~/.claude/skills/docx-formatter

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
├── SKILL.md                     # Claude Code Skill 文档
└── README.md                    # 本文件
```

## 运行测试

```bash
cd tests
./run_tests.sh
```

## 详细文档

完整的使用说明、EMU 换算速查、注意事项等，请参阅 [SKILL.md](SKILL.md)。
