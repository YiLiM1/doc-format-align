# doc-format-align

把 Word 文档的**格式**对齐到指定格式模板（.docx）的 ZCode Skill：字体字号、标题层级与编号、段落行距缩进、页边距、页眉页脚页码、表格、目录样式，全部以模板为准。

**内容保护第一原则：只改格式与版式，正文文字一字不改**（段落文字、顺序、结构原样保留）。唯一的文本改动是"标题编号"——当目标标题编号与模板同级编号不一致时，按模板体系重写编号（编号属于格式范围）。

## 功能特性

- **以模板为准**：模板示范了几级标题（支持到 8 级），就对齐几级；模板未示范的层级只对齐字体、不动编号
- **格式全覆盖**：封面大标题、正文、各级标题、行距缩进、页边距、页眉页脚页码、表格样式、目录
- **自动角色识别**：按"封面大标题 > 大纲级别 > 标题样式 > 编号样式"识别目标段落角色，正文列表不会被误判成标题
- **智能豁免**：图片段落保留自适应行距、自动编号列表不冲突缩进、空段落不撑爆封面版式
- **内容保护校验**：对齐前后"去编号文本"逐段比对，正文零改动才允许输出
- **三重验证**：对齐脚本 → 验证报告（verify_format.py）→ 8 项确定性检查（check_result.py）

## 安装

把 `doc-format-align` 目录放到 ZCode 的任一 skill 发现目录：

- `~/.zcode/skills/doc-format-align/`（ZCode 专用）
- `~/.agents/skills/doc-format-align/`（跨工具通用）

skill 会出现在可用技能列表中，模型可自动触发（自然语言触发）或手动调用 `/doc-format-align`。

## 使用

```bash
# 1. 对齐格式（默认直接覆盖目标文件，并生成 <目标名>_备份.docx）
python scripts/align_format.py --target <目标.docx> --template <模板.docx> [--number-mode continuous|nested]

# 2. 生成对齐报告
python scripts/verify_format.py --target <对齐后.docx> --template <模板.docx> --report <报告.md> [--number-mode nested]

# 3. 8 项确定性检查（内容保护 / 图片 / 编号 / 页眉页脚 / 页面设置 / 正文格式 / XML / 目录）
python scripts/check_result.py --target <对齐版.docx> --template <模板.docx> --original <原文件.docx> [--number-mode nested]
```

- `--number-mode`：`continuous`（各级标题全文连续编号，如 一、二、三… / 1、2、3…）或 `nested`（子级在上级下重新从 1 计，适合 1.1 / 1.1.1 技术文档）
- 依赖：Python 3.8+，仅标准库（zipfile / xml.etree），无需安装第三方包

## 目录结构

```
doc-format-align/
├── SKILL.md              # skill 定义：触发、对齐步骤、规则、交付与收尾
├── scripts/
│   ├── align_format.py   # 格式对齐主脚本
│   ├── verify_format.py  # 对齐报告生成
│   ├── check_result.py   # 8 项确定性检查
│   └── ooxml_util.py     # OOXML 解析/注入公共库（"只动格式不动文本"）
└── references/
    └── word-com-notes.md # 技术细节、模板形态与已知限制
```

## 工作原理

1. **提取模板规范**：从模板 docx 读取页面设置、页脚、各级标题与正文的字体/字号/行距/缩进、编号体系（含 Word 自动编号）
2. **识别目标角色**：封面大标题 > 大纲级别 > 标题样式 > 编号样式，正文列表保持正文角色
3. **注入格式 + 重写编号**：按角色注入格式；标题编号按模板体系重写；正文列表编号统一为（1）（2）（3）…每章内重编
4. **套用页面/页脚/表格**：页边距纸张、页脚 PAGE 域、表格样式
5. **内容保护校验**：对齐前后去编号文本逐段比对，零改动才输出

## 注意事项

- 只处理 `.docx`；`.doc` 老格式请先用 Word 另存为 `.docx`
- 脚本是纯 XML 级操作（不依赖 Word 自动化），秒级完成
- 封面分页、图片实际渲染等"所见即所得"项，需用 Word 打开对齐版目检
- 交付默认只保留对齐结果文件，备份/报告/临时文件按需清理（详见 SKILL.md「交付与收尾」）
