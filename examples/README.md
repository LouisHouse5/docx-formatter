# 示例文件

本目录包含 docx-formatter 的示例文件，用于演示完整的工作流程。

## 文件说明

| 文件 | 说明 |
|------|------|
| `template.docx` | 格式规范的模板文件（标准格式） |
| `target.docx` | 格式混乱的目标文件（待修复） |
| `batch_config.json` | 批量处理的 JSON 配置文件示例 |

## 快速体验

```bash
cd ~/.claude/skills/docx-formatter

# 1. 分析模板
python3 scripts/analyze_template.py examples/template.docx

# 2. 审核差异
python3 scripts/audit_docx.py examples/target.docx examples/template.docx

# 3. 修复格式（直接传命令行参数）
python3 scripts/fix_docx_template.py examples/target.docx examples/template.docx

# 4. 验证结果
python3 scripts/verify_docx.py examples/target.docx examples/template.docx
```

## 批量处理示例

创建 `files.txt`（每行一个目标文件路径）：

```
examples/target.docx
/path/to/another_target.docx
```

执行批量修复：

```bash
python3 scripts/fix_docx_template.py \
  --batch-file files.txt \
  --template examples/template.docx
```

或使用 JSON 配置文件：

```bash
python3 scripts/fix_docx_template.py \
  --config examples/batch_config.json \
  --template examples/template.docx
```

## 预期效果

修复前 `target.docx` 存在以下问题：
- 字体不为宋体
- 标题字号不正确
- 段落缺少首行缩进
- 表格边框样式不一致
- 引号为半角符号

修复后，以上问题全部被纠正，格式与 `template.docx` 一致。
