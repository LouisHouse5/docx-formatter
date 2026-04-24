# 示例文件

本目录包含 docx-formatter 的示例文件，用于演示完整的工作流程。

## 文件说明

| 文件 | 说明 |
|------|------|
| `template.docx` | 格式规范的模板文件（标准格式） |
| `target.docx` | 格式混乱的目标文件（待修复） |

## 快速体验

```bash
cd ~/.claude/skills/docx-formatter

# 1. 分析模板
python3 scripts/analyze_template.py examples/template.docx

# 2. 审核差异
python3 scripts/audit_docx.py examples/target.docx examples/template.docx

# 3. 修复格式
# 编辑 scripts/fix_docx_template.py 中的 CONFIG 部分，将 TARGET 改为 'examples/target.docx'
# 然后运行：
python3 scripts/fix_docx_template.py examples/target.docx examples/template.docx

# 4. 验证结果
python3 scripts/verify_docx.py examples/target.docx examples/template.docx
```

## 预期效果

修复前 `target.docx` 存在以下问题：
- 字体不为宋体
- 标题字号不正确
- 段落缺少首行缩进
- 表格边框样式不一致
- 引号为半角符号

修复后，以上问题全部被纠正，格式与 `template.docx` 一致。
