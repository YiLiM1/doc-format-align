# Word 格式对齐：技术细节与已知限制

## 实现原理

纯 OOXML 直接注入（不依赖 Word/LibreOffice 运行时）：脚本把 docx 当 zip 读入内存，用标准库
`zipfile` + `xml.etree.ElementTree` 修改格式节点后写回。优点：

- **内容保护天然安全**——只改 `pPr`/`rPr`/`sectPr`/`footer` 等格式节点，正文 `w:t` 文本零触碰
  （唯一例外：标题编号文本，由 `renumber_para` 显式处理）。
- 无 COM 的进程/版本坑；Word 2007（COM 12.0）也能正常打开。

## 内容保护校验

`align_format.py` 在写回前对"去编号后的全部段落文本"做前后快照对比：
- 相同 → 输出"内容保护：通过"，才保存。
- 不同 → 打印差异段落索引并以退出码 2 中止（不写坏文件）。

新增任何会动文本的逻辑时，必须同步考虑快照函数（`ooxml_util.document_text_snapshot`）。

## 模板形态自适应

模板分两种，脚本用同一套"角色识别"逻辑自动适配：

1. **正规文档模板**（自带 Heading 样式 / 大纲级别）：按 `pStyle` 或 `outlineLvl` 直接定位各级示范段落。
2. **说明性范本**（如公司格式范本：每段"示例即规范"）：按段首编号样式识别示范段落——
   `一、` → H1、`（一）` → H2、`1、` → H3；正文取段首为"正文"的段落。

**标题层级数不写死**：模板示范了几级标题就映射几级（支持到 H8）。模板有 5 级 6 级示范时，
示范段落编号 `1.`/`1.1`/`1.1.1`/… 或 Heading 样式自然映射到 H1..H5/H6，无需改代码或 skill 文案。
内部 `H1..H8` 只是占位名，映射多少级由识别结果决定。

**不要**用"关键词包含"匹配模板标题（如"一级标题"字样）——说明性范本的长说明段
（"正文格式要求为…一级标题使用黑体…"）会误命中。

模板的示范段落若有黄色高亮等标注性格式（`w:highlight`），提取规范时忽略它——脚本只在
`_merge_run_fmt` 里合并 `rFonts/sz/b/color`，不含 highlight。

## 编号体系

| 前缀样例 | 体系 | 生成 |
|---|---|---|
| `一、` / `（一）` / `1、` | 中文/阿拉伯 + 前后缀 | `一、二、三…`；`（一）（二）…`；`1、2、3…` |
| `1.1` / `1.1.1` | 点分（depth≥2） | 以最近上级编号为前缀：上级 `2` → `2.1, 2.2`；上级 `2.3` → `2.3.1` |

- `continuous`：各级全文各自连续编号。`nested`：本级出现时其后代级别计数器清零。
- 目标标题段首的**任意形式旧编号**都会被剥离（正则 `_ANY_NUM_RE`），再写入新编号；
  编号后的标题文字不变。正文段落的"1."/"2."（如附件列表）不会被当作标题编号动。
- 模板某级没有编号示范（如模板只有 3 级、目标有 4 级）→ 目标该级标题**不动编号**，只对齐字体格式，报告里提示。

## 目标段落角色识别

优先级：封面大标题（`detect_cover_titles`）→ `pPr/outlineLvl` 大纲级别 → `pStyle` 标题样式
→ 编号样式映射（`build_role_style_map`）→ 正文。

**样式标题体系优先（关键防御）**：`detect_document_heading_scheme` 检测文档是否使用标题样式/大纲级别。
一旦存在，标题识别只按样式走（层级按出现顺序压缩到 H1 起的连续层级，最多 H8），正文里的带编号列表
（"1、xxx""一、xxx"）一律按正文处理——这是保护正文列表不被误判成标题、防止被改写编号的开关。
仅当文档完全没有样式标题时（纯文本排版），才用编号样式映射兜底（公文 `一、`→H1，`（一）`→H2，
`1、`→H3；技术文档 `1`→H1，`1.1`→H2，`1.1.1`→H3，更深点分样式依次映射）。

**toc 样式排除（目录更新后的坑）**：目录（TOC）更新后，条目段落会带 toc 样式（styleId 通常是数字，
靠样式 name 匹配 `toc1`/`toc 1`/`目录 1` 等）。这些**不是标题**，必须排除：
- `toc_style_ids(entries)` 从 styles.xml 收集 toc styleId；
- `detect_document_heading_scheme(paras, toc_ids)` 排除它们，否则目录条目会被误当成深层
  标题层级（实测：模板A 的 toc1/toc3 曾被当成 H5/H4，污染模板规范提取）；
- `detect_para_role` 在 heading_scheme 模式下，体系外样式（含 toc）一律归 BODY，不 fallback
  到固定映射——否则目录条目被识别为标题、正文规范提取也会命中 toc 段。

防御：`1.` 单独出现（如附件"1.XXX"）不视为标题层级，仅当存在 `1.1` 或更深点分样式时才纳入。

**编号样式推断的额外防御（混编文档）**：
- `num_paren`（"（1）"）是正文列表样式：除非它是文档唯一的编号样式（可能是标题），
  否则与顿号/括号体系并存时一律排除（归正文列表处理，编号重写为（1）（2））。
- `cjk_comma` 与 `num_comma` 混编（"一、"与"1、"并存）是同一章节层级的两种写法，
  合并为同一角色（别名同层）——但**仅当两者并存且中间没有 `cjk_paren`（（一））桥接时**
  才合并：有（一）说明是标准公文三级体系（一、/（一）/1、），此时 `1、` 是三级标题。
- 模板正文示范段 rPr 仅含占位属性（hint/cs，无 eastAsia）时，用 Normal 样式的
  rPr 补全 BODY 字体/字号（说明性范本的正文继承 Normal 字体，严格对齐实际显示）。
- 空页眉（header 部件存在但无 w:t 文本，Word 自动创建）不算真页眉——模板视为无页眉，
  目标遗留页眉会被移除。

**封面大标题**：`detect_cover_titles` 扫描文档开头 12 段内居中、无编号、字号≥40（二号）的段落
→ TITLE 角色，套模板大标题格式（方正小标宋二号居中）。模板侧从含"标题"关键词或大号字体的居中
示范段提取 TITLE 规范。

## 编号重编与验证的一致

- `align_format.py --number-mode continuous|nested`（默认 continuous）：nested 下本级出现时
  其后代级别计数器清零（每章内子级重新从 1 计）。
- `verify_format.py` 的 `--number-mode` **必须与 align 一致**，否则报告会把正确编号误报为 ⚠️。

## 格式注入豁免（内容/可读性保护）

`ooxml_util.para_format_skip(p)` 判定段落应跳过的段落级注入项，`apply_role_to_para` 与
`verify_format.role_mismatch_reasons`（`skip` 参数）一致遵循，避免误报：

1. **图片/对象段落**（`w:drawing` / `w:pict` / `w:object`）→ **跳过行距与缩进**。
   模板正文是固定行距（如 560 twips = 28 磅），固定行距会把大图片**裁剪到几乎不可见**
   （图片段实际行高远超固定行距时被压扁/截断）。图片段必须保留自适应行距才能完整显示。
   实测案例：建设方案文档的拓扑图/架构图段落（含 `w:drawing`）曾被注入固定行距后不可见，
   修复后与原文件一致（3 张图全部正常渲染）。
2. **自动编号段落**（`pPr/numPr`，列表项）→ **跳过缩进**。列表的悬挂缩进由 numbering 定义，
   额外写入首行缩进（firstLineChars=200）会与列表缩进冲突、显示错乱。列表段的
   `firstLineChars='0'` 等原文件属性原样保留，不注入也不清空。
   **判定用"有效自动编号"（`para_has_auto_number`，numId≠0）**：自动编号已被关闭
   （numId=0）的段落不豁免——此时编号是手动文本（如"（1）"），段落就是普通正文，
   必须按模板正文注入首行缩进。正文列表项另用 `force_ind=True` 强制注入缩进
   （`apply_role_to_para` 参数，跳过 numPr 豁免），保证（1）段与模板正文缩进一致。
   **注意时序**：格式注入发生在"关闭自动编号"之前，所以列表项若靠 numId=0 判定会被
   numPr 豁免拦下——必须用 force_ind 显式强制。
3. **空段落**（无文本，仅排版占位）→ **跳过行距**。封面/落款用紧凑行距（如 200 twips）
   的空段撑版式，统一注入模板固定行距（560 twips）会撑爆版面——实测封面从 1 页溢出为
   2 页（第 2 页带出页码"2"）。空段只跳过 `spacing`，字体等不受影响。

豁免只针对段落级 `spacing`/`ind` 注入；run 级字体/字号仍照常对齐。

## 样式级自动编号（重复编号的根因）

目标文档的标题样式常自带**样式级自动编号**：`word/styles.xml` 里 heading 样式的 `pPr` 含
`<w:numPr><w:numId w:val="1"/></w:numPr>`，该样式所有段落都会渲染出 "1"、"1.1" 等编号。
脚本为标题插入模板的手动编号（"一、"等）后，两者会**重复显示**（"1 一、引言"）。

修复：`ooxml_util.disable_auto_number(p)` 在标题段落 pPr 写入
`<w:numPr><w:numId w:val="0"/></w:numPr>` 覆盖样式级编号（`numId=0` 表示取消编号，
仅按 pPr 插入顺序放在 `spacing` 之前，OOXML 对 pPr 子元素顺序有要求）。标题段（全部 H1..H8 与
TITLE）在重编编号后统一关闭自动编号。正文段落的自动编号（如项目符号列表）不在此列，保持不动。
`verify_format.py` 与 `check_result.py` 都会检查是否有残留（`para_has_auto_number`）。

## 页眉移除

模板无页眉时，`align_format.remove_headers` 移除目标所有节的页眉引用：每个 `sectPr` 可能带
多个 `headerReference`（default/even/first），**必须 `findall` 全部删除**——只删第一个会残留
even 类型；且后一节无 headerReference 时按 OOXML 规则**继承前一节页眉**，所以每个节都要清干净
（实测：封面节带 default+even 两个引用，漏删 even 导致所有页仍显示页眉）。header 部件与关系
保留为孤儿（无害、可逆）。

## 正文列表编号（保持正文格式）

样式标题体系下，正文中的编号列表条目**保持 BODY 角色、套正文格式（不加粗）**，编号统一重写为
（1）（2）（3）…每章内重编（`renumber_body_lists`）——刻意与标题编号（一、/（一）/1、）不重复：
- 自动编号且非项目符号（`build_num_fmt_map` 查 numbering.xml 的 numFmt != bullet，如（1）（2））；
- 手动"1、""1."或"（1）"开头（num_comma/num_dot1/num_paren）。
- 日期开头（"2024年…"）、"一、"开头的政策条款正文**不**改（是正文引用）。
正文列表段落的编号文本重写为（N），自动编号段落写 `numId=0` 关闭（见样式自动编号节）；
加粗/西文字体残留由 `apply_role_to_para(force_unbold=True)` 清除（之前可能是标题样式）。

## 表格样式注入（跨文档 styleId 冲突）

**不能只复制模板 tblStyle 的 styleId 引用**——目标 styles.xml 里同号 styleId 可能是完全
不同的样式（实测：模板 Table Grid=5，目标文档 styleId=5 是 heading 4 段落样式，导致表格
样式完全没生效）。正确做法（`align_format.apply_table_style`）：
1. 读模板首表 tblStyle 的 styleId，在模板 styles.xml 里查该样式的 **name**（如 Table Grid）；
2. 在目标 styles.xml 按 name 找同名表格样式，用**目标自己的 styleId** 写引用；
3. 目标没有同名表格样式时，把模板样式定义 deepcopy 进目标 styles.xml（styleId 冲突则用
   `DfaTbl<N>` 新 id），再写引用。

**表格字体处理（用户确认的偏好）**：表格内**中文字体**对齐模板（模板表格单元格的 eastAsia，
无显式声明时用模板 Normal 样式的 eastAsia——表格字体继承 Normal），**字号保持目标原样**
（不同表格列数/字号不同，只对齐字体与样式，不强制字号）。注意：目标表格单元格若无显式
rFonts/rPr 需补建，仅写 eastAsia，保留 sz 等其他属性。

## 目录（TOC）处理（需用户确认）

模板没有目录但目标有目录（TOC 域）时，默认规则是**保留目标目录 + 提示在 Word 中右键更新域**。
这是"模板未提及项"之一，**对齐前应向用户确认**处理方式（保留 / 删除 / 保留并自动更新），
不要默默按默认规则处理。

## 页脚与页码格式

- 页脚内容（PAGE 域）套用到所有节；**各节原有的页码格式保留**——脚本不动 `pgNumType`
  （起始值、罗马/阿拉伯数字），封面无页码、目录用罗马数字、正文从 1 起的分节设置不会被破坏。
- 页脚页码验证以 **PDF 底部区域文本**为准（Word 渲染后 PAGE 域生效），COM `Footers(1).Range.Text`
  在这台机器上读不到字段文本属读取怪癖，不可作为判定依据。

## 页脚注入（踩过的坑）

1. **关系 Target 必须相对 `word/` 目录**：新建 footer 部件 `word/footer2.xml` 时，rels 里
   `Target="footer2.xml"`，**不能写 `word/footer2.xml`**（Word 找不到部件会静默丢弃页脚内容）。
   读模板页脚时反过来：rels 里 `Target="footer1.xml"` 要补成 `word/footer1.xml` 才能读字节。
2. **模板页脚要先规范化再复制**：WPS/新版 Word 生成的 footer 原始字节带 `mc:Ignorable="w14 w15 wp14"`
   引用的未声明前缀、跨文档无效的 `pStyle`（引用模板 styles.xml 的 styleId）、`w14:paraId` 追踪属性，
   直接复制会让 Word 解析失败而**清空页脚内容**。必须用 `normalize_footer_xml()` 解析后重序列化。
3. **命名空间由 ET 自动声明**：往 document.xml 加 `r:id` 属性时，ET 的 `register_namespace` 已注册
   `r` 前缀，序列化会自动输出 `xmlns:r`。**不要手动 set `xmlns:` 属性**——ET 会把
   `http://www.w3.org/2000/xmlns/` 序列化成非法的 `xmlns:ns2=".../xmlns/"`，整个 XML 解析失败。
4. 新 footer 部件要同时登记 `[Content_Types].xml` 的 Override。

## 格式注入细节

- 段落级：清掉角色覆盖项（`jc/spacing/ind/snapToGrid`）再写模板值，保证模板优先；`spacing/ind` 保留目标已有、模板未声明的键（如 `before/after`）。
- 段落标记 `pPr/rPr` 与每个 run 的 `rPr` 都注入，保证光标落在段尾时格式正确。
- 正文（BODY）用 `full=False`：只统一中文正文字体（eastAsia）和字号，**保留目标原有西文字体与加粗**；
  标题用 `full=True`：字体/字号/加粗/颜色全部对齐模板。
- 页面设置写入 document.xml 的每个 `sectPr`；页脚引用写入每个节的 `footerReference`。

## 内容保护边界

- 不动：正文文字、段落增删、顺序、表格内文本、批注、图片、域（除页脚 PAGE 域替换）。
- 允许改：标题段首的编号文本（按模板重编）。
- 若目标大量用"加粗+大字号"冒充标题（无编号、无大纲级别、无标题样式），会被识别为正文而只套正文格式——
  这类文档请先请用户确认哪些段算标题（手动编号或设为标题样式），或说明"按正文格式处理"。

## 兼容性

- 脚本：Python 3.9+，仅标准库，无需 pywin32/python-docx/LibreOffice。
- 文档：Word 2007+ 可打开；模板为 WPS 生成时页脚规范化已处理。
- 模板若为 `.doc` 老格式：先请用户在 Word 中另存为 `.docx` 再用。
