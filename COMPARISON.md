# docx-formatter vs docx-official 差异对比

本文档详细对比 `docx-formatter`（本技能）与 Claude Code 内置 `docx-official` 技能的区别，帮助用户选择合适的工具。

---

## 一句话总结

| 技能 | 定位 | 适用场景 |
|------|------|---------|
| **docx-official** | 通用文档处理工具箱 | 创建新文档、内容编辑、审阅修订 |
| **docx-formatter** | 专用格式对齐工具 | 批量标准化已有文档的格式 |

两者是**互补关系**，非替代关系。

---

## 核心定位

### docx-official（官方内置）
- 覆盖文档全生命周期：创建、编辑、审阅、分析
- 技术栈多元：JS 创建（docx-js）+ Python XML 编辑 + pandoc 转换
- 强调**内容层面**的精细控制（修订模式、批注、结构化编辑）

### docx-formatter（本技能）
- 专注单一任务：以模板为基准批量修复格式
- 技术栈单一：纯 python-docx + OOXML 底层操作
- 强调**格式层面**的自动化对齐（8 个维度、显式+隐藏格式）

---

## 功能覆盖详细对比

| 功能 | docx-official | docx-formatter |
|------|:-------------:|:--------------:|
| **从零创建文档** | ✅（docx-js） | ❌ |
| **内容编辑（改文字）** | ✅（逐条脚本） | ❌ |
| **修订模式/红字审阅** | ✅（核心功能） | ❌ |
| **添加批注** | ✅ | ❌ |
| **文本提取/转 Markdown** | ✅（pandoc） | ❌ |
| **文档转图像** | ✅（LibreOffice + pdftoppm） | ❌ |
| **格式批量对齐** | ❌ | ✅（核心功能） |
| **隐藏格式修复** | ❌ | ✅（分节、TOC域、样式定义、表格边框等） |
| **段落格式统一** | ❌ | ✅（字体、字号、行距、缩进、对齐等） |
| **页面设置同步** | ❌ | ✅（纸张、页边距、方向、页眉页脚距边界） |
| **页眉页脚复制** | ❌ | ✅ |
| **目录域同步** | ❌ | ✅ |
| **表格边框/列宽同步** | ❌ | ✅ |
| **半角/全角引号统一** | ❌ | ✅ |
| **空行/空段落清理** | ❌ | ✅ |
| **自动化验证闭环** | ❌ | ✅（分析→审核→修复→验证） |

---

## 使用场景决策树

```
用户请求涉及 docx 文件？
  ├── 需要创建全新文档？
  │     └── 使用 docx-official（docx-js 生成）
  │
  ├── 需要修改文档内容（改文字、改结构）？
  │     ├── 涉及修订模式/批注/审阅？
  │     │     └── 使用 docx-official（红字流程）
  │     └── 简单内容修改？
  │           └── 使用 docx-official（Python XML 编辑）
  │
  ├── 需要统一已有文档的格式到模板？
  │     └── 使用 docx-formatter（批量修复）
  │
  ├── 需要提取文档内容分析？
  │     └── 使用 docx-official（pandoc 转 markdown）
  │
  └── 其他文档处理？
        └── 使用 docx-official（通用工具箱）
```

### 典型场景举例

| 场景 | 推荐技能 | 原因 |
|------|---------|------|
| "帮我生成一份新的课程标准文档" | docx-official | 从零创建，内容编排 |
| "这份教案格式和标准模板不一致" | **docx-formatter** | 批量格式对齐 |
| "帮我审阅这份合同并标记修改" | docx-official | 修订模式/红字流程 |
| "把这份报告的字体统一成宋体" | **docx-formatter** | 批量字体修复 |
| "分析这份文档的内容结构" | docx-official | pandoc 转 markdown |
| "修复页眉页脚和目录域格式" | **docx-formatter** | 隐藏格式同步 |

---

## 技术深度对比

| 技术点 | docx-official | docx-formatter |
|--------|---------------|----------------|
| **底层操作** | OOXML 原始 XML 直接编辑 | python-docx 高级 API + 部分 XML |
| **学习成本** | 高（需理解 pack/unpack、RSID、tracked changes） | 低（配置文件 + 运行脚本） |
| **代码复杂度** | 逐字符级 XML 操作 | 段落/表格/分节级批量同步 |
| **自动化程度** | 手工编写每批修改脚本 | 一键分析/审核/修复/验证 |
| **精度控制** | 单字符级 | 段落级（满足格式标准化需求） |
| **依赖环境** | pandoc、LibreOffice、Node.js、Python | 仅 python-docx |
| **技能触发** | 无显式触发（内置自动加载） | `/docx-format` |

---

## 配合使用建议

最佳实践是**先用 docx-formatter 统一格式，再用 docx-official 做内容修改**：

```
┌─────────────────────────────────────────────────┐
│  Step 1: docx-formatter                         │
│  - 统一字体、字号、行距、缩进                    │
│  - 同步页眉页脚、目录域、页面设置                │
│  - 修复表格边框、半角引号                        │
└──────────────────────┬──────────────────────────┘
                       ▼
┌─────────────────────────────────────────────────┐
│  Step 2: docx-official                          │
│  - 内容层面的精细修改                            │
│  - 添加批注、修订标记                            │
│  - 结构调整、新增章节                            │
└─────────────────────────────────────────────────┘
```

---

## 文件结构对比

### docx-official（内置）
```
~/.claude/skills/docx-official/
├── SKILL.md
├── docx-js.md          # JS 创建文档指南
├── ooxml.md            # Python XML 编辑指南
└── ooxml/
    └── scripts/
        ├── unpack.py   # 解包 docx
        └── pack.py     # 打包 docx
```

### docx-formatter（本技能）
```
~/.claude/skills/docx-formatter/
├── SKILL.md
├── README.md
├── COMPARISON.md       # 本文件
├── install.sh
├── .gitignore
├── examples/
│   ├── template.docx
│   └── target.docx
└── scripts/
    ├── analyze_template.py      # 深度扫描模板
    ├── audit_docx.py            # 全面对比差异
    ├── fix_docx_template.py     # 精确修复脚本
    ├── verify_docx.py           # 多维度验证
    ├── copy_styles.py           # 复制样式定义
    └── copy_headers_footers.py  # 复制页眉页脚
```

---

## 总结

| | docx-official | docx-formatter |
|--|---------------|----------------|
| **比喻** | 瑞士军刀 | 专用机床 |
| **擅长** | 什么都能做，内容精细控制 | 只做一件事，格式对齐做到极致 |
| **适用** | 通用文档处理 | 批量标准化格式 |
| **关系** | 互补 | 互补 |

在 Claude Code 中，两个技能会根据你的请求自动协作。当系统检测到格式相关任务时，会优先触发 `docx-formatter`（通过 `/docx-format`）；当涉及通用文档处理时，会调用 `docx-official`。
