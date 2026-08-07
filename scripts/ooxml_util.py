# -*- coding: utf-8 -*-
"""
doc-format-align 公共模块：docx(OOXML) 解析、格式规范提取、段落格式注入。

设计要点（内容保护第一）：
- 全程只操作格式节点（pPr / rPr / sectPr / footer），绝不改动正文 w:t 文本，
  唯一的文本改动是"标题编号文本"，由 align_format 显式调用 renumber_* 完成。
- 不依赖 Word / LibreOffice，纯标准库（zipfile + ElementTree）。
"""
import re
import zipfile
import io
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
CT_NS = '{http://schemas.openxmlformats.org/package/2006/content-types}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'

# 写回 docx 时保持标准命名空间前缀，Word 才认识
_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
}
for _k, _v in _NS.items():
    ET.register_namespace(_k, _v)

# ---------------- 角色定义 ----------------
# TITLE 为封面大标题，H1..H8 对应模板的各级标题（层级数由模板示范决定），BODY 为正文。
# HROLES 在下方 _HEADING_ROLES 处定义（支持模板任意层级的标题体系）。
ROLES = ('TITLE', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'BODY')

# 编号样式推导：样本前缀 -> (数字体系, 前缀/包裹)
_NUM_PATTERNS = [
    (re.compile(r'^（([一二三四五六七八九十]+)）'), 'cjk', '（', '）'),
    (re.compile(r'^\(([一二三四五六七八九十]+)\)'), 'cjk', '(', ')'),
    (re.compile(r'^([一二三四五六七八九十]+)、'), 'cjk', '', '、'),
    (re.compile(r'^([0-9]+)、'), 'arabic', '', '、'),
    (re.compile(r'^([0-9]+)\.'), 'arabic', '', '.'),
    (re.compile(r'^\(([0-9]+)\)'), 'arabic', '(', ')'),
    (re.compile(r'^([0-9]+)[)）]'), 'arabic', '', ')'),
]

_CJK = '一二三四五六七八九十'
_CJK_MAP = {c: i + 1 for i, c in enumerate(_CJK)}


def cjk_number(n):
    """1 -> 一, 10 -> 十, 12 -> 十二, 21 -> 二十一, 105 -> 一百零五（支持到 9999）"""
    if n <= 0:
        return str(n)
    if 1 <= n <= 10:
        return _CJK[n - 1]
    units = ('', '十', '百', '千')
    s = str(n)
    nd = len(s)
    out = ''
    prev_zero = False
    for i, ch in enumerate(s):
        d = int(ch)
        u = units[nd - 1 - i]
        if d == 0:
            prev_zero = True
            continue
        if prev_zero:
            out += '零'
            prev_zero = False
        if d == 1 and u == '十' and i == 0:
            out += '十'
        else:
            out += _CJK[d - 1] + u
    return out or '零'


def derive_number_kind(sample_prefix):
    """从模板示范段落的编号前缀推导编号体系；无编号返回 None"""
    if not sample_prefix:
        return None
    # 点分编号（1.1 / 1.1.1）优先，记录深度
    m = re.match(r'^[0-9]+(?:\.[0-9]+)+', sample_prefix)
    if m:
        return {'num': 'arabic', 'pre': '', 'post': '', 'depth': sample_prefix.count('.') + 1}
    for pat, num_kind, pre, post in _NUM_PATTERNS:
        m = pat.match(sample_prefix)
        if m:
            return {'num': num_kind, 'pre': pre, 'post': post}
    return None


def make_number(kind, n, parent_num=None):
    """按编号体系生成第 n 个编号（n 从 1 开始）。

    点分体系（depth≥2）以 parent_num（上级编号）为前缀，保持层级结构；
    若上级尚未出现（parent_num 缺失），用 "1" 补足到 depth-1 级，
    保证层级深度不变（如 depth=3 → "1.1.N"，而非退化成 "1.N"）。
    """
    if not kind:
        return ''
    if kind['num'] == 'arabic' and kind.get('depth', 1) > 1:
        m = re.match(r'^[0-9]+(?:\.[0-9]+)*', parent_num or '')
        if m:
            base = m.group(0)
        else:
            # 无上级编号：用 1 补足到 depth-1 个点段
            base = '.'.join(['1'] * (kind['depth'] - 1))
        return f"{base}.{n}"
    num = cjk_number(n) if kind['num'] == 'cjk' else str(n)
    return kind['pre'] + num + kind['post']


# ---------------- 基础读取 ----------------

def read_docx_entries(path):
    """读取 docx 为 {文件名: bytes}"""
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def write_docx(entries, path):
    """写回 docx（保持全部条目，仅替换修改过的）"""
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)


def parse_xml(data):
    return ET.fromstring(data)


def ensure_ns_declared(root, prefixes):
    """确保根元素声明了指定的命名空间前缀（r 等）。"""
    declared = {k for k, v in root.attrib.items() if k.startswith('{http://www.w3.org/2000/xmlns/}')}
    for p in prefixes:
        if p not in declared:
            root.set('{http://www.w3.org/2000/xmlns/}' + p, _NS[p])


def normalize_footer_xml(footer_xml):
    """解析模板页脚 XML，规范化命名空间并移除跨文档无效的引用。

    模板页脚常由 WPS/新版 Word 生成，原始字节含：
    - mc:Ignorable 引用的未声明前缀（w14/w15/wp14），旧版 Word 解析会丢弃内容
    - pStyle 引用模板 styles.xml 的 styleId，目标文档中无效
    - paraId 等 wp14 追踪属性
    这里用 ET 解析后重新序列化，命名空间全部正规化，并剥离上述无效引用。
    """
    root = ET.fromstring(footer_xml)
    # 移除 mc:Ignorable 属性（其列出的 w14/w15/wp14 前缀未声明，旧版 Word 可能拒绝）
    for k in list(root.attrib):
        if k.startswith('{http://schemas.openxmlformats.org/markup-compatibility/2006}') and 'Ignorable' in k:
            del root.attrib[k]
    # 移除所有段落的 pStyle（跨文档无效）
    for pPr in root.iter(W + 'pPr'):
        ps = pPr.find(W + 'pStyle')
        if ps is not None:
            pPr.remove(ps)
    # 移除 paraId 等追踪属性
    for el in root.iter():
        for k in list(el.attrib):
            if k.startswith('{http://schemas.microsoft.com/office/word/2010/') and k.endswith('}paraId'):
                del el.attrib[k]
    return to_bytes(root)


def to_bytes(elem):
    return ET.tostring(elem, encoding='utf-8', xml_declaration=True)


# ---------------- 段落工具 ----------------

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def para_runs(p):
    return p.findall(W + 'r')


def para_first_text_run(p):
    for r in p.findall(W + 'r'):
        if r.find(W + 't') is not None:
            return r
    return None


def para_pPr(p):
    pPr = p.find(W + 'pPr')
    if pPr is None:
        pPr = ET.SubElement(p, W + 'pPr')
        p.insert(0, pPr)
    return pPr


def run_rPr(r):
    rPr = r.find(W + 'rPr')
    if rPr is None:
        rPr = ET.Element(W + 'rPr')
        r.insert(0, rPr)
    return rPr


def get_attr(elem, tag):
    if elem is None:
        return None
    e = elem.find(W + tag)
    return e


# ---------------- 段落角色判定（目标文档） ----------------

# 标题层级上限：支持模板 1~8 级标题（H1..H8）；模板有几级示范就映射几级。
# H1..H8 只是内部占位名，具体映射多少级由模板识别结果决定，不写死"三级/四级"。
_HEADING_ROLES = ('H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8')
HROLES = _HEADING_ROLES
MAX_HEADING_LEVEL = len(_HEADING_ROLES)

# 大纲级别 -> 角色（outline 0-7 -> H1-H8；更深归最深层 H8）
_OUTLINE_TO_ROLE = {str(i): _HEADING_ROLES[i] for i in range(MAX_HEADING_LEVEL)}
for _i in range(MAX_HEADING_LEVEL, 10):
    _OUTLINE_TO_ROLE[str(_i)] = _HEADING_ROLES[-1]

# pStyle styleId -> 角色（兼容中文 Word 的数字 styleId 与英文 styleId）。
# 数字 styleId 从 1 开始，可到 10（Word 2007 默认最多 9 级标题样式，styleId 1-9；
# 用户自定义可更多）。超出 H8 的归最深层 H8。
_STYLE_TO_ROLE = {
    **{str(i): _HEADING_ROLES[min(i - 1, MAX_HEADING_LEVEL - 1)]
       for i in range(1, 11)},
    **{f'Heading{i}': _HEADING_ROLES[i - 1] for i in range(1, MAX_HEADING_LEVEL + 1)},
    **{f'heading {i}': _HEADING_ROLES[i - 1] for i in range(1, MAX_HEADING_LEVEL + 1)},
}

# pStyle styleId -> 标题级别数字（用于层级压缩排序）
_STYLE_LEVEL = {
    **{str(i): i for i in range(1, 11)},
    **{f'Heading{i}': i for i in range(1, MAX_HEADING_LEVEL + 1)},
    **{f'heading {i}': i for i in range(1, MAX_HEADING_LEVEL + 1)},
}


def toc_style_ids(entries):
    """返回 styles.xml 中 name 匹配 toc 模式（toc1/toc 1/目录 1/...）的 styleId 集合。

    目录更新后条目段落带 toc 样式，这些**不是标题**，层级推断/角色识别必须排除。
    注意 styleId 通常是数字（如 '5'），toc 判断必须靠样式 name。
    """
    if 'word/styles.xml' not in entries:
        return set()
    root = parse_xml(entries['word/styles.xml'])
    pat = re.compile(r'(?:^|\s)(?:toc|目录)\s*\d*$', re.IGNORECASE)
    ids = set()
    for st in root.findall(W + 'style'):
        name = st.find(W + 'name')
        if name is not None and name.get(W + 'val') and pat.search(name.get(W + 'val')):
            ids.add(st.get(W + 'styleId'))
    return ids


def detect_document_heading_scheme(paras, toc_ids=None):
    """检测文档是否使用样式/大纲标题体系。

    返回 {信号: 角色} 压缩映射（文档的样式标题层级按出现顺序压缩到 H1..H8），
    信号为 'outline:<级别>' 或 'style:<styleId>'；文档无任何样式标题时返回 None
    （此时才启用编号样式兜底识别）。

    关键：文档一旦使用标题样式，正文里的带编号列表（"1、xxx""一、xxx"）一律按
    正文处理，绝不因编号长得像标题而被误判——这是保护正文列表的关键开关。
    **toc 样式不是标题**：目录更新后条目段落带 toc 样式（name 含 toc/目录），
    必须排除，否则会被误当成深层标题层级（实测：模板A 的 toc1/toc3 曾被当成 H5/H4）。
    """
    toc_re = re.compile(r'^toc\s*\d+$', re.IGNORECASE)
    toc_ids = toc_ids or set()
    sigs = set()
    for p in paras:
        pPr = p.find(W + 'pPr')
        if pPr is None:
            continue
        ol = pPr.find(W + 'outlineLvl')
        ps = pPr.find(W + 'pStyle')
        if ol is not None:
            # 同段既有 outlineLvl 又有标题样式：只记 outlineLvl 信号，避免同一段落
            # 被拆成两个层级（实测：招标文件"第一章"段 outline:0+style:2 重复计数，
            # 层级压缩错乱，模板规范提取失真）
            sigs.add(('outline:' + ol.get(W + 'val'), int(ol.get(W + 'val'))))
            continue
        if ps is not None and ps.get(W + 'val') in _STYLE_TO_ROLE:
            sid = ps.get(W + 'val')
            if sid in toc_ids or toc_re.match(sid):
                continue  # toc 样式不是标题
            sigs.add(('style:' + sid, _STYLE_LEVEL[sid]))
    if not sigs:
        return None
    # 按层级升序，压缩映射到 H1..H8（层级数由文档实际用到的样式级别决定）
    ordered = sorted(sigs, key=lambda s: s[1])
    scheme = {}
    n_roles = min(len(ordered), MAX_HEADING_LEVEL)
    for i, (sig, _lvl) in enumerate(ordered[:n_roles]):
        scheme[sig] = _HEADING_ROLES[i]
    return scheme


def classify_style(text):
    """把段首编号归类为样式标签（用于文档级层级推断）"""
    if not text:
        return None
    if re.match(r'^[一二三四五六七八九十]+、', text):
        return 'cjk_comma'
    if re.match(r'^[（(][一二三四五六七八九十]+[）)]', text):
        return 'cjk_paren'
    if re.match(r'^[0-9]+、', text):
        return 'num_comma'
    m = re.match(r'^([0-9]+\.)+[0-9]+', text)
    if m:
        n = m.group(0).count('.') + 1
        return 'num_dot%d' % n
    if re.match(r'^[0-9]+\.', text):
        return 'num_dot1'
    if re.match(r'^[（(][0-9]+[）)]', text) or re.match(r'^[0-9]+[）]', text):
        return 'num_paren'
    return None


# 编号样式偏好链：按公文/技术文档惯例从一级到最深层（层级数由实际出现的样式决定）
_STYLE_ORDER = (['cjk_comma', 'cjk_paren', 'num_comma']
                + [f'num_dot{i}' for i in range(1, MAX_HEADING_LEVEL + 1)]
                + ['num_paren'])


def build_role_style_map(paras):
    """文档级层级推断：根据全文出现的编号样式推断各编号样式对应角色。

    出现样式按偏好链压缩分配到 H1..H8——
    公文（一、/（一）/1、）→ H1/H2/H3；技术文档（1/1.1/1.1.1）→ H1/H2/H3。
    防御：
    - num_dot1（"1."）单独出现时更像附件/列表项而非标题，仅在 num_dot2 或
      更深的点分样式并存时才视为标题层级。
    - num_paren（"（1）"）是正文列表样式：除非它是文档唯一的编号样式（可能是
      标题），否则与顿号/括号等标题体系并存时一律排除（归正文列表处理）。
    - cjk_comma 与 num_comma 混编（"一、"与"1、"并存）是同一章节层级的两种
      写法，合并为一个槽位并映射到同一角色（别名同层），后续层级顺延。
    """
    styles = set()
    for p in paras:
        t = para_text(p)
        if not t:
            continue
        c = classify_style(t)
        if c:
            styles.add(c)
    if not styles:
        return {}
    # num_dot1 防御：需 num_dot2 或更深的点分样式并存才视为标题层级
    deeper_dots = [f'num_dot{i}' for i in range(2, MAX_HEADING_LEVEL + 1)]
    if 'num_dot1' in styles and not any(s in styles for s in deeper_dots):
        styles.discard('num_dot1')
    # num_paren 防御：与其他标题体系并存时视为正文列表，排除
    if 'num_paren' in styles and len(styles) > 1:
        styles.discard('num_paren')
    # cjk_comma 与 num_comma 同层合并（章节级两种写法）——仅当两者并存且
    # 中间没有 cjk_paren 桥接时才合并：有 （一） 说明是标准公文三级体系
    # （一、/（一）/1、），此时 1、 是三级标题而非一级别名。
    merged = ('cjk_comma' in styles and 'num_comma' in styles
              and 'cjk_paren' not in styles)
    available = [c for c in _STYLE_ORDER if c in styles]
    if merged:
        available = [c for c in available if c != 'num_comma']
    roles = _HEADING_ROLES[:len(available)]
    mapping = dict(zip(available, roles))
    if merged:
        mapping['num_comma'] = mapping['cjk_comma']  # 别名同层
    return mapping


def detect_cover_titles(paras, max_check=12, min_sz=40):
    """封面大标题识别：文档开头 max_check 段内，无编号、字号 >= min_sz 的段落。

    公文封面的大标题通常居中使用二号(44)或更大字号；部分模板的大标题未写 jc=center
    （对齐属性缺失，Word 按默认左对齐渲染），因此"未写居中"也接受——只要字号 >= min_sz
    且非样式标题即可。正文里居中的引用标题一般字号较小，不会被误判。
    """
    cover = set()
    for i, p in enumerate(paras):
        if i >= max_check:
            break
        text = para_text(p)
        if not text.strip():
            continue
        old, _ = match_leading_number(text)
        if old:
            continue  # 带编号的不是封面标题
        pPr = p.find(W + 'pPr')
        if pPr is not None:
            ol = pPr.find(W + 'outlineLvl')
            ps = pPr.find(W + 'pStyle')
            if ol is not None or (ps is not None and ps.get(W + 'val') in _STYLE_TO_ROLE):
                continue  # 样式标题不当作封面标题
            jc = pPr.find(W + 'jc')
            if jc is not None and jc.get(W + 'val') != 'center':
                continue  # 显式非居中（left/right/justify）的大字号段不视为封面标题
        else:
            continue
        run = para_first_text_run(p)
        if run is None:
            continue
        rPr = run.find(W + 'rPr')
        sz = rPr.find(W + 'sz') if rPr is not None else None
        szval = int(sz.get(W + 'val')) if sz is not None else 0
        if szval >= min_sz:
            cover.add(id(p))
    return cover


def detect_para_role(p, text, style_map=None, heading_scheme=None, cover_set=None):
    """判定一个段落的角色。

    优先级：
    0. 封面大标题（cover_set 命中）
    1. 大纲级别 outlineLvl（0-3 → H1-H4，更深的归 H4）
    2. 标题样式 pStyle（有 heading_scheme 时按压缩映射，否则按固定映射）
    3. 编号样式映射 —— 仅当文档无样式标题体系（heading_scheme is None）时启用，
       避免把正文里的带编号列表误判成标题
    4. 正文
    """
    if cover_set is not None and id(p) in cover_set:
        return 'TITLE'
    pPr = p.find(W + 'pPr')
    if pPr is not None:
        ol = pPr.find(W + 'outlineLvl')
        if ol is not None:
            return _OUTLINE_TO_ROLE.get(ol.get(W + 'val'), _HEADING_ROLES[-1])
        ps = pPr.find(W + 'pStyle')
        if ps is not None and ps.get(W + 'val') in _STYLE_TO_ROLE:
            if heading_scheme:
                # 有样式标题体系时只认体系内的样式；体系外样式（如 toc）一律按正文，
                # 不 fallback 到固定映射（避免把目录条目误判成深层标题）
                return heading_scheme.get('style:' + ps.get(W + 'val'), 'BODY')
            return _STYLE_TO_ROLE[ps.get(W + 'val')]
    if heading_scheme is None and style_map:
        c = classify_style(text)
        if c and c in style_map:
            return style_map[c]
    return 'BODY'


# ---------------- 模板规范提取 ----------------

def _merge_run_fmt(pPr_rPr, run_rPr):
    """取 pPr/rPr 与 run rPr 中更具体的格式（run 优先）"""
    base = {}
    for src in (pPr_rPr, run_rPr):
        if src is None:
            continue
        rf = src.find(W + 'rFonts')
        if rf is not None:
            d = {}
            for a in ('hint', 'ascii', 'hAnsi', 'eastAsia', 'cs'):
                v = rf.get(W + a)
                if v:
                    d[a] = v
            base['rFonts'] = d
        for tag in ('b', 'bCs', 'sz', 'szCs', 'color', 'highlight'):
            e = src.find(W + tag)
            if e is not None:
                val = e.get(W + 'val')
                # b 元素存在即视为加粗（val 显式 0/false 除外）
                if tag in ('b', 'bCs'):
                    base[tag] = not (val in ('0', 'false'))
                else:
                    base[tag] = val
    return base


def extract_role_spec_from_para(p, style_rpr_map=None, num_lvl_map=None):
    """从示范段落提取 RoleSpec：rPr 格式 + 段落级格式 + 编号样本。

    style_rpr_map：{styleId: rPr 摘要}（来自 styles.xml，见 build_style_rpr_map）。
    正规模板的示范段落常用 Heading 样式且自身无 run 级 rPr（字体定义在样式里），
    此时用样式的 rPr 补全字体/字号（段落级/run 级显式 rPr 优先）。
    num_lvl_map：{numId: lvl0 编号样本}（见 build_num_lvl_map）。示范段落的编号若是
    Word 自动编号（无段首手动编号文本），从 numbering.xml 推导编号样本与体系。
    """
    text = para_text(p)
    pPr = p.find(W + 'pPr')
    pPr_rPr = pPr.find(W + 'rPr') if pPr is not None else None
    run = para_first_text_run(p)
    run_rPr = run.find(W + 'rPr') if run is not None else None
    fmt = _merge_run_fmt(pPr_rPr, run_rPr)
    # 正规模板补全：段落用了样式但自身无 rPr 格式时，取样式 rPr
    if style_rpr_map and pPr is not None:
        ps = pPr.find(W + 'pStyle')
        if ps is not None:
            sfmt = style_rpr_map.get(ps.get(W + 'val'))
            if sfmt:
                merged = dict(sfmt)
                if not fmt.get('rFonts'):
                    merged['rFonts'] = sfmt.get('rFonts')
                else:
                    merged['rFonts'] = {**sfmt.get('rFonts', {}), **fmt['rFonts']}
                for k in ('sz', 'szCs', 'b', 'bCs', 'color'):
                    if fmt.get(k) is None and sfmt.get(k) is not None:
                        merged[k] = sfmt[k]
                    elif fmt.get(k) is not None:
                        merged[k] = fmt[k]
                fmt = merged

    spec = {
        'rFonts': fmt.get('rFonts'),
        'sz': fmt.get('sz'),
        'szCs': fmt.get('szCs'),
        'b': fmt.get('b'),          # True/False/None
        'bCs': fmt.get('bCs'),
        'color': fmt.get('color'),
        'jc': None,
        'spacing': {},
        'ind': {},
        'snapToGrid': None,
        'number': None,
    }
    if pPr is not None:
        jc = pPr.find(W + 'jc')
        if jc is not None:
            spec['jc'] = jc.get(W + 'val')
        sp = pPr.find(W + 'spacing')
        if sp is not None:
            for a in ('before', 'after', 'line', 'lineRule'):
                v = sp.get(W + a)
                if v:
                    spec['spacing'][a] = v
        ind = pPr.find(W + 'ind')
        if ind is not None:
            for a in ('left', 'right', 'firstLine', 'firstLineChars', 'hanging', 'hangingChars', 'leftChars'):
                v = ind.get(W + a)
                if v:
                    spec['ind'][a] = v
        stg = pPr.find(W + 'snapToGrid')
        if stg is not None:
            spec['snapToGrid'] = stg.get(W + 'val') or '0'
    # 编号样本：从段首文本截取编号前缀
    sample = None
    if text:
        m, _ = match_leading_number(text)
        if m:
            sample = m
    # 示范段落的编号常为 Word 自动编号（numPr，段首无手动编号文本）：
    # 从 numbering.xml 的 lvl0 lvlText 推导编号体系（如 '%1、'+chineseCounting → 一、）
    if sample is None and num_lvl_map is not None and pPr is not None:
        numPr = pPr.find(W + 'numPr')
        if numPr is not None:
            numId_el = numPr.find(W + 'numId')
            if numId_el is not None and numId_el.get(W + 'val') not in ('0', None):
                lvl = num_lvl_map.get(numId_el.get(W + 'val'))
                if lvl:
                    sample = lvl['sample']
    if sample:
        spec['number'] = {'sample': sample, 'kind': derive_number_kind(sample)}
    return spec


# 模板示范段落角色识别（两种模板形态）：
# 1) 说明性范本：段首编号（一、/（一）/1、…）作为示范段落（"示例即规范"）
# 2) 正规文档：Heading 样式 / 大纲级别
# 两种都走 detect_para_role，不依赖"几级标题"关键词——模板有 5 级 6 级标题同样适用。


def detect_template_para_role(p, text, style_map=None, heading_scheme=None):
    """模板示范段落的角色识别。

    优先用强信号（大纲级别/标题样式/编号样式），不采用"关键词包含"匹配——
    说明性范本的长说明段（如"正文格式要求为…一级标题使用黑体…"）会误命中关键词。
    """
    return detect_para_role(p, text, style_map, heading_scheme)


def build_style_rpr_map(entries):
    """从 styles.xml 构建 {styleId: rPr 摘要}，供正规模板规范补全。

    摘要取样式 pPr/rPr 与 rPr 的 rFonts(eastAsia/ascii)/sz/b。
    """
    if 'word/styles.xml' not in entries:
        return {}
    root = parse_xml(entries['word/styles.xml'])
    m = {}
    for st in root.findall(W + 'style'):
        sid = st.get(W + 'styleId')
        if not sid:
            continue
        rpr = None
        pPr = st.find(W + 'pPr')
        if pPr is not None:
            rpr = pPr.find(W + 'rPr')
        if rpr is None:
            rpr = st.find(W + 'rPr')
        summary = {}
        if rpr is not None:
            rf = rpr.find(W + 'rFonts')
            if rf is not None:
                d = {}
                for a in ('eastAsia', 'ascii', 'hAnsi'):
                    v = rf.get(W + a)
                    if v:
                        d[a] = v
                summary['rFonts'] = d
            for tag in ('sz', 'szCs', 'b', 'color'):
                e = rpr.find(W + tag)
                if e is not None:
                    if tag == 'b':
                        val = e.get(W + 'val')
                        summary[tag] = not (val in ('0', 'false'))
                    else:
                        summary[tag] = e.get(W + 'val')
        if summary:
            m[sid] = summary
    return m


def extract_template_spec(entries):
    """从模板 docx 提取完整格式规范。

    返回 dict:
      page_setup: pgSz/pgMar/cols/docGrid 属性
      footer_xml: bytes（模板 footer 完整 XML，无则 None）
      has_header: bool
      roles: {role: RoleSpec}
      body_role: RoleSpec（正文）
    """
    doc = parse_xml(entries['word/document.xml'])
    body = doc.find(W + 'body')
    paras = body.findall(W + 'p')
    tpl_style_map = build_role_style_map(paras)
    tpl_heading_scheme = detect_document_heading_scheme(paras, toc_style_ids(entries))
    style_rpr_map = build_style_rpr_map(entries)
    num_lvl_map = build_num_lvl_map(entries)

    roles = {}
    for role in HROLES:
        roles[role] = None

    # 模板大标题（TITLE）：无编号、居中、大字号，或含"标题"关键词的居中段。
    # 兜底：公文封面大标题常不写 jc=center（XML 无对齐属性），取首个无编号的大字号
    # （≥二号）段作为 TITLE，保证这类模板的大标题格式也能对齐。
    title_spec = None
    fallback_title = None
    for p in paras:
        text = para_text(p)
        if not text.strip():
            continue
        pPr = p.find(W + 'pPr')
        if pPr is None:
            continue
        run = para_first_text_run(p)
        rPr = run.find(W + 'rPr') if run is not None else None
        sz = rPr.find(W + 'sz') if rPr is not None else None
        szval = int(sz.get(W + 'val')) if sz is not None else 0
        jc = pPr.find(W + 'jc')
        centered = jc is not None and jc.get(W + 'val') == 'center'
        if szval >= 40:
            if centered:
                title_spec = extract_role_spec_from_para(p, style_rpr_map, num_lvl_map)
                break
            if fallback_title is None and pPr.find(W + 'numPr') is None:
                fallback_title = p
        if '标题' in text and centered:
            title_spec = extract_role_spec_from_para(p, style_rpr_map, num_lvl_map)
            break
    if title_spec is None and fallback_title is not None:
        title_spec = extract_role_spec_from_para(fallback_title, style_rpr_map, num_lvl_map)
    roles['TITLE'] = title_spec

    # 第一遍：识别示范段落（每角色取第一个命中）
    for p in paras:
        text = para_text(p)
        if not text.strip():
            continue
        role = detect_template_para_role(p, text, tpl_style_map, tpl_heading_scheme)
        if role in HROLES and roles.get(role) is None:
            roles[role] = extract_role_spec_from_para(p, style_rpr_map, num_lvl_map)

    # 正文：优先取段首以"正文"开头的示范段落，否则第一个无编号普通段
    toc_ids = toc_style_ids(entries)
    body_spec = None
    for p in paras:
        text = para_text(p)
        if not text.strip():
            continue
        if text.startswith('正文') and detect_para_role(p, text, tpl_style_map, tpl_heading_scheme) == 'BODY':
            body_spec = extract_role_spec_from_para(p, style_rpr_map, num_lvl_map)
            break
    # 正文兜底：第一个普通正文段（排除大标题/目录/toc 样式等段）
    if body_spec is None:
        for p in paras:
            text = para_text(p)
            if not text.strip():
                continue
            if '标题' in text:
                continue
            if detect_para_role(p, text, tpl_style_map, tpl_heading_scheme) != 'BODY':
                continue
            pPr = p.find(W + 'pPr')
            ps = pPr.find(W + 'pStyle') if pPr is not None else None
            if ps is not None and ps.get(W + 'val') in toc_ids:
                continue  # 排除 toc 样式段（目录条目）
            jc = pPr.find(W + 'jc') if pPr is not None else None
            run = para_first_text_run(p)
            rPr = run.find(W + 'rPr') if run is not None else None
            sz = rPr.find(W + 'sz') if rPr is not None else None
            szval = int(sz.get(W + 'val')) if sz is not None else 0
            if jc is not None and jc.get(W + 'val') == 'center':
                continue  # 排除居中段（大标题/目录标题）
            if szval >= 40:
                continue  # 排除大字号段（大标题）
            if '目录' in text:
                continue
            body_spec = extract_role_spec_from_para(p, style_rpr_map, num_lvl_map)
            break
    roles['BODY'] = body_spec
    # BODY 规范补全：说明性范本的正文继承 Normal 样式字体（示范段无显式 rPr）。
    # 若提取的 BODY rFonts/sz 缺失，用 Normal 样式的 rPr 补全（严格对齐模板实际显示字体）。
    if body_spec is not None and style_rpr_map:
        normal = None
        for sid, sfmt in style_rpr_map.items():
            if sid in ('1', 'Normal', 'a'):
                normal = sfmt
                break
        if normal:
            # 补全条件：rFonts 缺失或没有 eastAsia 值（示范段 rPr 仅有 hint/cs 等占位属性）
            cur_ea = (body_spec.get('rFonts') or {}).get('eastAsia')
            if not cur_ea and normal.get('rFonts'):
                merged = dict(body_spec.get('rFonts') or {})
                merged.update({k: v for k, v in normal['rFonts'].items() if v})
                body_spec['rFonts'] = merged
            if not body_spec.get('sz') and normal.get('sz'):
                body_spec['sz'] = normal['sz']
                body_spec['szCs'] = normal.get('szCs') or normal['sz']

    # 页面设置
    sect = body.find(W + 'sectPr')
    page_setup = None
    if sect is not None:
        page_setup = {}
        pg = sect.find(W + 'pgSz')
        if pg is not None:
            page_setup['pgSz'] = {a: pg.get(W + a) for a in ('w', 'h', 'orient') if pg.get(W + a)}
        pm = sect.find(W + 'pgMar')
        if pm is not None:
            page_setup['pgMar'] = {a: pm.get(W + a) for a in
                                   ('top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter') if pm.get(W + a)}
        cols = sect.find(W + 'cols')
        if cols is not None:
            page_setup['cols'] = {a: cols.get(W + a) for a in ('num', 'space') if cols.get(W + a)}
        dg = sect.find(W + 'docGrid')
        if dg is not None:
            page_setup['docGrid'] = {a: dg.get(W + a) for a in ('type', 'linePitch', 'charSpace') if dg.get(W + a)}

    # 页眉页脚（注意：Target 通常相对 word/ 目录，如 'footer1.xml'）
    footer_xml = None
    has_header = False
    rels = parse_xml(entries['word/_rels/document.xml.rels'])
    for rel in rels:
        t = rel.get('Type', '')
        if t.endswith('/footer'):
            fname = rel.get('Target')
            if fname:
                if fname.startswith('/'):
                    fname = fname[1:]
                elif not fname.startswith('word/'):
                    fname = 'word/' + fname
            if fname in entries:
                footer_xml = entries[fname]
        elif t.endswith('/header'):
            # header 部件存在但无 w:t 文本（Word 自动创建的空页眉）视为无页眉
            hname = rel.get('Target')
            if hname:
                if hname.startswith('/'):
                    hname = hname[1:]
                elif not hname.startswith('word/'):
                    hname = 'word/' + hname
                if hname in entries:
                    hxml = entries[hname].decode('utf-8', 'ignore')
                    # 无 w:t 文本的空页眉不算真页眉，保持 has_header=False
                    if '<w:t>' in hxml or '<w:t ' in hxml:
                        has_header = True

    return {
        'page_setup': page_setup,
        'footer_xml': footer_xml,
        'has_header': has_header,
        'roles': roles,
    }


# ---------------- 格式注入 ----------------

def _set_or_create(pPr, tag, attrs):
    """在 pPr 下设置子元素属性（元素不存在则新建，存在则保留未指定的属性）"""
    e = pPr.find(W + tag)
    if e is None:
        e = ET.SubElement(pPr, W + tag)
    for k, v in attrs.items():
        if v is None:
            continue
        e.set(W + k, v)
    return e


def _clear_children(pPr, tags):
    for tag in tags:
        e = pPr.find(W + tag)
        if e is not None:
            pPr.remove(e)


def _apply_run_fmt(rPr, fmt, full):
    """把 RoleSpec 的 run 级格式写入 rPr。

    full=True 表示连加粗/颜色也统一（标题用）；
    full=False 只覆盖字体(eastAsia/sz)，保留目标原有加粗等局部强调（正文用）。
    """
    rf = rPr.find(W + 'rFonts')
    if rf is None:
        rf = ET.SubElement(rPr, W + 'rFonts')
    if fmt.get('rFonts'):
        if full:
            for a in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                v = fmt['rFonts'].get(a)
                if v:
                    rf.set(W + a, v)
        else:
            # 正文：只统一中文正文字体，不强行改西文字体（保留数字/英文原字体）
            if 'eastAsia' in fmt['rFonts']:
                rf.set(W + 'eastAsia', fmt['rFonts']['eastAsia'])
    else:
        if full:
            for a in ('ascii', 'hAnsi', 'eastAsia'):
                if rf.get(W + a):
                    del rf.attrib[W + a]
    # 字号
    if fmt.get('sz'):
        for tag, val in (('sz', fmt['sz']), ('szCs', fmt.get('szCs') or fmt['sz'])):
            e = rPr.find(W + tag)
            if e is None:
                e = ET.SubElement(rPr, W + tag)
            e.set(W + 'val', val)
    # 加粗
    if full and fmt.get('b') is not None:
        e = rPr.find(W + 'b')
        if fmt['b']:
            if e is None:
                e = ET.SubElement(rPr, W + 'b')
            elif e.get(W + 'val') in ('0', 'false'):
                del e.attrib[W + 'val']
        else:
            if e is not None:
                rPr.remove(e)
            ec = rPr.find(W + 'bCs')
            if ec is not None:
                rPr.remove(ec)
    # 颜色
    if full and fmt.get('color'):
        e = rPr.find(W + 'color')
        if e is None:
            e = ET.SubElement(rPr, W + 'color')
        e.set(W + 'val', fmt['color'])


def para_format_skip(p):
    """返回该段落应跳过的段落级格式注入项（set，含 'spacing'/'ind' 等）。

    三类豁免（都是内容/可读性/版式保护）：
    - 含图片/对象（w:drawing / w:pict / w:object）的段落：**跳过行距与缩进注入**。
      模板正文是固定行距（如 28 磅），固定行距会把大图片裁剪到几乎不可见——
      图片段落必须保留自适应行距才能完整显示。
    - **仍有有效自动编号**（numPr 且 numId≠0）的列表段落：**跳过缩进注入**。列表的
      缩进由 numbering 定义（悬挂缩进），额外写入首行缩进会与列表缩进冲突、显示错乱。
      注意：自动编号已被关闭（numId=0）的段落不豁免——此时编号是手动文本（如"（1）"），
      段落就是普通正文，必须按模板正文注入首行缩进（与模板正文格式一致）。
    - 空段落（无文本，仅排版占位）：**跳过行距注入**。封面/落款常用紧凑行距
      （如 200 twips）的空段撑版式，统一注入模板固定行距会撑爆版面（封面溢出成两页）。
    """
    skip = set()
    if not para_text(p).strip():
        skip.add('spacing')
    for el in p.iter():
        if el.tag.endswith(('}drawing', '}pict', '}object')):
            skip |= {'spacing', 'ind'}
            break
    if para_has_auto_number(p):
        skip.add('ind')
    return skip


def apply_role_to_para(p, spec, role, full_run=True, force_unbold=False, force_ind=False):
    """把角色格式注入段落。不修改任何 w:t 文本。

    full_run=True（标题）全量对齐；full_run=False（正文）保留局部强调。
    force_unbold=True 额外移除 run 的加粗与西文字体（用于"正文列表改回正文格式"——
    之前被标题格式注入过加粗，需清掉；同时清掉 ascii/hAnsi 让西文回默认主题字体）。
    force_ind=True 强制注入缩进（即使 para_format_skip 因自动编号而豁免）——用于
    "正文列表项"：这些段落最终是手动编号正文（自动编号随后被关闭），必须像模板正文
    一样有首行缩进。
    图片段落跳过行距/缩进、仍有有效自动编号的段落跳过缩进（见 para_format_skip）。
    """
    if spec is None:
        return
    pPr = para_pPr(p)
    skip = para_format_skip(p)
    if force_ind:
        skip.discard('ind')
    elif role != 'BODY' and para_has_auto_number(p):
        # 标题段落：缩进以模板示范段落 pPr 为准（其自动编号随后改写为手动编号文本并
        # 关闭，编号缩进不再生效），不因"有自动编号"而跳过缩进注入；
        # 正文列表段仍保持 para_format_skip 的豁免（缩进由编号定义，强加会错乱）。
        skip.discard('ind')

    # 段落级：先清掉该角色要覆盖的项，再写入（保证模板值优先）
    _clear_children(pPr, tuple(t for t in ('jc', 'spacing', 'ind', 'snapToGrid') if t not in skip))
    if spec.get('jc') and 'jc' not in skip:
        _set_or_create(pPr, 'jc', {'val': spec['jc']})
    if spec.get('spacing') and 'spacing' not in skip:
        _set_or_create(pPr, 'spacing', spec['spacing'])
    if spec.get('ind') and 'ind' not in skip:
        _set_or_create(pPr, 'ind', spec['ind'])
    if spec.get('snapToGrid') is not None and 'snapToGrid' not in skip:
        _set_or_create(pPr, 'snapToGrid', {'val': spec['snapToGrid']})

    # 段落标记格式
    pPr_rPr = pPr.find(W + 'rPr')
    if pPr_rPr is None:
        pPr_rPr = ET.SubElement(pPr, W + 'rPr')
    _apply_run_fmt(pPr_rPr, spec, full_run)
    if force_unbold:
        _strip_run_bold(pPr_rPr)

    # 各 run
    for r in p.findall(W + 'r'):
        rPr = run_rPr(r)
        _apply_run_fmt(rPr, spec, full_run)
        if force_unbold:
            _strip_run_bold(rPr)


def _strip_run_bold(rPr):
    """移除 run 的加粗（b/bCs）与西文字体（rFonts ascii/hAnsi/cs），保留中文字体 eastAsia。

    用于"正文列表改回正文格式"：之前被标题格式注入的加粗和西文字体残留需清除，
    让西文回主题默认字体、正文不加粗。
    """
    for tag in ('b', 'bCs'):
        e = rPr.find(W + tag)
        if e is not None:
            rPr.remove(e)
    rf = rPr.find(W + 'rFonts')
    if rf is not None:
        for a in ('ascii', 'hAnsi', 'cs'):
            if rf.get(W + a):
                del rf.attrib[W + a]


# ---------------- 标题编号重写 ----------------

_ANY_NUM_RE = re.compile(
    r'^(（[一二三四五六七八九十]+）|\([一二三四五六七八九十]+\)|[一二三四五六七八九十]+、'
    r'|[0-9]+、|[0-9]+(?:\.[0-9]+)+|[0-9]+[.）)]?'
    r'|（[0-9]+）|\([0-9]+\))')


def match_leading_number(text):
    """返回段首编号 (匹配文本, 长度)，无则 (None, 0)"""
    if not text:
        return None, 0
    m = _ANY_NUM_RE.match(text)
    if m:
        return m.group(0), m.end()
    return None, 0


def _strip_prefix_across_runs(para, length):
    """从段落前部 runs 移除 length 个字符（编号文本），不产生新文本。"""
    if length <= 0:
        return
    remaining = length
    for r in para.findall(W + 'r'):
        t = r.find(W + 't')
        if t is None:
            continue
        s = t.text or ''
        if remaining <= 0:
            break
        if len(s) <= remaining:
            remaining -= len(s)
            t.text = ''
        else:
            t.text = s[remaining:]
            remaining = 0


def _prepend_number(para, prefix):
    """把编号前缀写入段落第一个文本 run（其文本已去掉旧编号）。"""
    r = para_first_text_run(para)
    if r is None:
        # 无文本 run（空标题）：追加一个 run 承载编号
        pPr = para.find(W + 'pPr')
        r = ET.Element(W + 'r')
        t = ET.SubElement(r, W + 't')
        t.text = prefix
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        para.insert(para.index(pPr) + 1 if pPr is not None else 0, r)
        return
    t = r.find(W + 't')
    t.text = prefix + (t.text or '')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def renumber_para(para, new_number):
    """用 new_number 替换段落段首任意形式的旧编号；无编号则前缀插入。"""
    text = para_text(para)
    old, old_len = match_leading_number(text)
    if old is not None:
        _strip_prefix_across_runs(para, old_len)
    if new_number:
        _prepend_number(para, new_number)


def renumber_document(paras_roles, roles_spec, mode='continuous'):
    """按模板编号体系重写标题编号。

    paras_roles: [(para_element, role), ...] 按文档顺序
    roles_spec:  {'H1': RoleSpec, ...}
    mode: 'continuous' 各级全文各自连续编号；'nested' 子级在上级下重新从 1 开始
    返回: {role: {'renumbered': n, 'skipped': n, 'reason': str}}
    """
    counters = {r: 0 for r in HROLES}
    last_num = {}
    stats = {r: {'renumbered': 0, 'skipped': 0, 'reason': ''} for r in HROLES}

    # 确定各角色编号体系
    kinds = {}
    for role in HROLES:
        spec = roles_spec.get(role) or {}
        num = spec.get('number')
        kinds[role] = num['kind'] if num else None

    for p, role in paras_roles:
        if role not in HROLES:
            continue
        kind = kinds[role]
        if kind is None:
            stats[role]['skipped'] += 1
            stats[role]['reason'] = '模板该级别无编号示范'
            continue
        # nested 模式：本级出现时，其所有后代级别的计数器重置（从 1 重新计）
        if mode == 'nested':
            role_idx = _HEADING_ROLES.index(role)
            for d in _HEADING_ROLES[role_idx + 1:]:
                counters[d] = 0
        counters[role] += 1
        parent = _HEADING_ROLES[_HEADING_ROLES.index(role) - 1] if role != _HEADING_ROLES[0] else None
        new_num = make_number(kind, counters[role], last_num.get(parent))
        renumber_para(p, new_num)
        last_num[role] = new_num
        stats[role]['renumbered'] += 1
    return stats


def renumber_body_lists(paras_roles, num_fmt_map, list_style='paren'):
    """把正文列表条目的编号重写为指定样式（默认（1）（2）…），每章内重编。

    规则（nested 语义，与标题编号一致）：
    - 按文档顺序遍历，遇到任意标题角色（H1..H8/TITLE）时正文列表计数重置为 1，
      即每个标题章节下的列表从（1）开始；
    - 只处理角色为 BODY 的编号列表条目（para_is_list_item 判定，需 num_fmt_map 区分
      bullet）；手动编号（"1、""1."）替换为（N），Word 自动编号（（1））关闭后写入手动（N）。
    用于"正文列表用正文格式、编号不与标题编号重复"的场景——
    标题编号为一、/（一）/1、，正文列表用（1）（2）（3），互不重复。
    list_style: 'paren' -> （N）；'plain' -> 1、2、3（备用）
    返回重写的段落数。
    """
    if list_style not in ('paren', 'plain'):
        list_style = 'paren'
    pre, post = ('（', '）') if list_style == 'paren' else ('', '、')
    n = 0
    counter = 0
    for p, role in paras_roles:
        if role in _HEADING_ROLES or role == 'TITLE':
            counter = 0
            continue
        if role != 'BODY':
            continue
        text = para_text(p)
        if not para_is_list_item(p, text, num_fmt_map):
            continue
        counter += 1
        n += 1
        renumber_para(p, pre + str(counter) + post)
        pPr = p.find(W + 'pPr')
        if pPr is not None and pPr.find(W + 'numPr') is not None:
            disable_auto_number(p)
    return n


# ---------------- 自动编号处理 ----------------

# pPr 中应位于 numPr 之后的元素（OOXML 规定的 pPr 子元素顺序）
_AFTER_NUMPR = (W + 'spacing', W + 'ind', W + 'jc', W + 'snapToGrid',
                W + 'outlineLvl', W + 'divId', W + 'rPr', W + 'sectPr')


def disable_auto_number(p):
    """关闭段落的自动编号（numPr numId=0）。

    目标文档的标题样式常自带样式级自动编号（styles.xml 中 numPr），会渲染出
    "1"、"1.1" 等编号；脚本已为标题插入模板的手动编号文本（"一、"等），两者
    会重复显示。本函数在段落 pPr 写入 <w:numPr><w:numId w:val="0"/></w:numPr>
    覆盖样式级编号，仅保留手动编号文本。
    """
    pPr = para_pPr(p)
    numPr = pPr.find(W + 'numPr')
    if numPr is None:
        numPr = ET.Element(W + 'numPr')
        idx = None
        for i, child in enumerate(pPr):
            if child.tag in _AFTER_NUMPR:
                idx = i
                break
        if idx is not None:
            pPr.insert(idx, numPr)
        else:
            pPr.append(numPr)
    ilvl = numPr.find(W + 'ilvl')
    if ilvl is None:
        ilvl = ET.SubElement(numPr, W + 'ilvl')
    ilvl.set(W + 'val', '0')
    numId = numPr.find(W + 'numId')
    if numId is None:
        numId = ET.SubElement(numPr, W + 'numId')
    numId.set(W + 'val', '0')


def para_has_auto_number(p):
    """段落是否仍有有效自动编号（numPr 且 numId 非 0）"""
    pPr = p.find(W + 'pPr')
    if pPr is None:
        return False
    numPr = pPr.find(W + 'numPr')
    if numPr is None:
        return False
    numId = numPr.find(W + 'numId')
    if numId is None:
        return True  # 有 numPr 但无 numId（继承样式）视为有编号
    return numId.get(W + 'val') not in ('0', None)


def build_num_fmt_map(entries):
    """从 numbering.xml 建立 {numId: lvl0 numFmt} 映射；无 numbering.xml 返回 {}。

    用于区分"编号列表"（decimal：（1）（2）…）与"项目符号列表"（bullet）。
    """
    if 'word/numbering.xml' not in entries:
        return {}
    root = parse_xml(entries['word/numbering.xml'])
    abs_map = {}
    for a in root.findall(W + 'abstractNum'):
        abs_map[a.get(W + 'abstractNumId')] = a
    m = {}
    for num in root.findall(W + 'num'):
        nid = num.get(W + 'numId')
        abid_el = num.find(W + 'abstractNumId')
        if abid_el is None or abid_el.get(W + 'val') not in abs_map:
            continue
        a = abs_map[abid_el.get(W + 'val')]
        lvl = a.find(W + 'lvl')
        fmt = lvl.find(W + 'numFmt') if lvl is not None else None
        m[nid] = fmt.get(W + 'val') if fmt is not None else 'decimal'
    return m


_NUM_FMT_FIRST = {
    'chineseCounting': '一', 'chineseLower': '一', 'chineseCountThousand': '一',
    'ideographTraditional': '一',
    'decimal': '1', 'decimalZero': '1',
    'upperRoman': 'I', 'lowerRoman': 'i',
}


def build_num_lvl_map(entries):
    """从 numbering.xml 建立 {numId: {'sample': <lvl0 第 1 项编号文本样本>}}。

    模板示范段落的标题编号常是 Word 自动编号（段 pPr numPr，编号文本定义在
    numbering.xml 的 abstractNum lvl0 lvlText 中，如 '%1、' + chineseCounting
    渲染为 '一、'）。derive_number_kind 需要"文本样本"（如 '一、'），这里把
    lvl0 定义换算成样本：%1 替换为第一项编号文本（chineseCounting→一、decimal→1）。
    """
    if 'word/numbering.xml' not in entries:
        return {}
    root = parse_xml(entries['word/numbering.xml'])
    abs_map = {}
    for a in root.findall(W + 'abstractNum'):
        abs_map[a.get(W + 'abstractNumId')] = a
    m = {}
    for num in root.findall(W + 'num'):
        nid = num.get(W + 'numId')
        abid_el = num.find(W + 'abstractNumId')
        if abid_el is None or abid_el.get(W + 'val') not in abs_map:
            continue
        a = abs_map[abid_el.get(W + 'val')]
        lvl = a.find(W + 'lvl')
        if lvl is None:
            continue
        fmt = lvl.find(W + 'numFmt')
        fv = fmt.get(W + 'val') if fmt is not None else 'decimal'
        txt = lvl.find(W + 'lvlText')
        lv = txt.get(W + 'val') if txt is not None else '%1'
        first = _NUM_FMT_FIRST.get(fv)
        if not first or '%1' not in lv:
            continue
        m[nid] = {'sample': lv.replace('%1', first)}
    return m


def heading_number_kinds(roles_spec):
    """{role: 编号体系 kind}——模板各标题级别的编号体系；模板未示范的层级为 None。

    供 align/verify/check 三方共用：有 kind 的级别"编号重写为手动文本并关闭自动编号"，
    无 kind 的级别（模板未示范）"只对齐字体、不动编号"（保留其原有自动编号）。
    """
    kinds = {}
    for role in HROLES:
        num = (roles_spec.get(role) or {}).get('number')
        kinds[role] = num['kind'] if num else None
    return kinds


def shifted_heading_spec(roles_spec, shift=0):
    """按层级偏移构建 shifted 角色规范：目标 H(n) 用模板 H(n+shift) 的规范/编号。

    用于模板最高级是封面/章节特有层、需忽略的场景（如招标文件"第一章"章节层，
    目标的一级标题应对齐模板二级标题：shift=1）。TITLE/BODY 不变。
    """
    if not shift:
        return roles_spec
    out = {}
    for role in HROLES:
        idx = _HEADING_ROLES.index(role) + shift
        src = _HEADING_ROLES[min(idx, len(_HEADING_ROLES) - 1)]
        out[role] = roles_spec.get(src)
    out['TITLE'] = roles_spec.get('TITLE')
    out['BODY'] = roles_spec.get('BODY')
    return out


def para_is_list_item(p, text, num_fmt_map):
    """段落是否为"编号列表条目"（非项目符号、非正文引用）。

    两类：
    - 带自动编号 numPr 且 numId 非 0、编号格式非 bullet（如（1）（2））
    - 手动编号开头："1、"/"1."（num_comma/num_dot1）或"（1）"/"（N）"（num_paren）
      都视为列表条目。
    日期开头（"2024年…"）、"一、"开头的政策条款正文不属于列表条目。
    """
    pPr = p.find(W + 'pPr')
    if pPr is not None:
        numPr = pPr.find(W + 'numPr')
        if numPr is not None:
            numId = numPr.find(W + 'numId')
            if numId is not None:
                nid = numId.get(W + 'val')
                if nid not in (None, '0'):
                    return num_fmt_map.get(nid, 'decimal') != 'bullet'
    c = classify_style(text)
    return c in ('num_comma', 'num_dot1', 'num_paren')


# ---------------- 文本快照（内容保护验证用） ----------------

def para_text_without_number(p):
    text = para_text(p)
    old, old_len = match_leading_number(text)
    if old:
        return text[old_len:]
    return text


def document_text_snapshot(entries):
    """所有段落去编号后的文本序列（内容保护验证：对齐前后应完全一致）。"""
    doc = parse_xml(entries['word/document.xml'])
    body = doc.find(W + 'body')
    return [para_text_without_number(p) for p in body.findall(W + 'p')]
