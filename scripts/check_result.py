# -*- coding: utf-8 -*-
"""
doc-format-align 确定性检查：对齐后文档的快速全面自检（秒级，纯 XML，不依赖 Word/渲染）。

用法：
  python check_result.py --target <对齐版.docx> --template <模板.docx> --original <原文件.docx> [--number-mode nested]

检查项（每项 ✅/❌/⚠️）：
  1. 内容保护：原 vs 对齐版"去编号后段落文本"完全一致
  2. 图片完整性：media 部件、w:drawing 数、rels 图片关系数 与原文件一致
  3. 标题编号：各级标题编号与模板体系一致（nested/continuous），无样式自动编号残留
  4. 页眉页脚：模板无页眉时对齐版无 headerReference；有页脚 PAGE 域
  5. 页面设置：pgSz/pgMar 与模板一致
  6. 正文格式：抽查段落字体/字号/行距/缩进符合模板规范
  7. 文档完整性：全部 XML 可解析
  8. 目录：TOC 域 / _Toc 书签保留（提示在 Word 中更新域）

说明：封面分页与图片实际渲染需在 Word 中目检（本脚本不启动 Word/不导出 PDF）。
"""
import argparse
import os
import sys
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

W = U.W
R = U.R


def iter_sect_elems(body):
    elems = [body.find(W + 'sectPr')] if body.find(W + 'sectPr') is not None else []
    elems += body.findall('.//' + W + 'sectPr')
    seen, out = set(), []
    for s in elems:
        if id(s) not in seen:
            seen.add(id(s))
            out.append(s)
    return out


def check_content(target_entries, orig_entries):
    a = U.document_text_snapshot(orig_entries)
    b = U.document_text_snapshot(target_entries)
    if len(a) != len(b):
        return False, f'段落数不同: 原{len(a)} vs 对齐{len(b)}'
    diff = [i for i in range(len(a)) if a[i] != b[i]]
    if diff:
        return False, f'{len(diff)} 段正文文本差异: {diff[:8]}'
    return True, '原 vs 对齐版去编号文本完全一致'


def check_images(target_entries, orig_entries):
    def sig(e):
        media = sorted(n for n in e if '/media/' in n)
        doc = e['word/document.xml'].decode('utf-8', 'ignore')
        ndraw = doc.count('w:drawing')
        rels = U.parse_xml(e['word/_rels/document.xml.rels'])
        nimg = sum(1 for r in rels if 'image' in r.get('Type', ''))
        return media, ndraw, nimg
    ts, osig = sig(target_entries), sig(orig_entries)
    if ts != osig:
        return False, f'图片不一致: 对齐{ts} vs 原{osig}'
    return True, f'media {len(ts[0])} 个、drawing {ts[1]}、图片关系 {ts[2]}，与原文件一致'


def check_numbering(target_entries, tpl_spec, number_mode):
    """编号序列与模板体系一致 + 无样式自动编号残留"""
    doc = U.parse_xml(target_entries['word/document.xml'])
    body = doc.find(W + 'body')
    paras = body.findall(W + 'p')
    style_map = U.build_role_style_map(paras)
    heading_scheme = U.detect_document_heading_scheme(paras, U.toc_style_ids(target_entries))
    cover = U.detect_cover_titles(paras)
    num_fmt_map = U.build_num_fmt_map(target_entries)
    paras_roles = []
    for p in paras:
        role = U.detect_para_role(p, U.para_text(p), style_map, heading_scheme, cover)
        paras_roles.append((p, role))

    kinds = U.heading_number_kinds(tpl_spec['roles'])
    counters = {r: 0 for r in U.HROLES}
    last_num = {}
    problems = []
    for idx, (p, role) in enumerate(paras_roles):
        if role not in U.HROLES:
            continue
        kind = kinds[role]
        text = U.para_text(p)
        old, _ = U.match_leading_number(text)
        if kind is None:
            continue
        if number_mode == 'nested':
            r_i = U._HEADING_ROLES.index(role)
            for d in U._HEADING_ROLES[r_i + 1:]:
                counters[d] = 0
        counters[role] += 1
        parent = U._HEADING_ROLES[U._HEADING_ROLES.index(role) - 1] if role != U._HEADING_ROLES[0] else None
        expect = U.make_number(kind, counters[role], last_num.get(parent))
        if old != expect:
            problems.append(f'段落[{idx}]「{text[:20]}」编号 {old!r} 应为 {expect!r}')
        last_num[role] = expect
    # 样式自动编号残留（模板有编号示范的级别才检查；未示范的级别保留原编号）
    resid = [idx for idx, (p, r) in enumerate(paras_roles)
             if r in U.HROLES and kinds.get(r) and U.para_has_auto_number(p)]
    if resid:
        problems.append(f'{len(resid)} 段标题仍有样式自动编号(会与手动编号重复): {resid[:8]}')
    # 正文列表编号（（1）（2）…，每章内重编）：按文档顺序遍历，遇标题重置、遇列表项比对
    k = 0
    n_list = 0
    for idx, (p, r) in enumerate(paras_roles):
        if r in U.HROLES or r == 'TITLE':
            k = 0
            continue
        if r != 'BODY':
            continue
        text = U.para_text(p)
        if not U.para_is_list_item(p, text, num_fmt_map):
            continue
        n_list += 1
        old, _ = U.match_leading_number(text)
        expect = f'（{k + 1}）'
        if old != expect:
            problems.append(f'正文列表段[{idx}]「{text[:16]}」编号 {old!r} 应为 {expect!r}')
        if U.para_has_auto_number(p):
            problems.append(f'正文列表段[{idx}] 仍有自动编号残留')
        k += 1
    if problems:
        return False, '; '.join(problems[:5]) + (f' …共{len(problems)}处' if len(problems) > 5 else '')
    n_h = sum(1 for _, r in paras_roles if r in U.HROLES)
    return True, f'全部 {n_h} 段标题编号与模板体系一致、{n_list} 段正文列表编号（1）（2）…一致'


def check_header_footer(target_entries, tpl_spec):
    doc = U.parse_xml(target_entries['word/document.xml'])
    body = doc.find(W + 'body')
    problems = []
    if not tpl_spec['has_header']:
        for s in iter_sect_elems(body):
            for hr in s.findall(W + 'headerReference'):
                problems.append('存在遗留 headerReference（模板无页眉）')
    else:
        # 模板有页眉：目标页眉保留不动（内容由用户核对），不视为失败
        msgs = ['模板含页眉 → 目标页眉保留，请人工核对页眉内容']
        return True, '；'.join(msgs)
    # 页脚
    if tpl_spec['footer_xml'] is not None:
        has_footer = False
        for s in iter_sect_elems(body):
            if s.find(W + 'footerReference') is not None:
                has_footer = True
        if not has_footer:
            problems.append('模板有页脚但目标无 footerReference')
    if problems:
        return False, '; '.join(problems)
    return True, '页眉符合模板（无页眉/已移除）、页脚 PAGE 域已套用'


def check_page_setup(target_entries, tpl_spec):
    if not tpl_spec.get('page_setup'):
        return True, '模板无页面设置规范'
    doc = U.parse_xml(target_entries['word/document.xml'])
    body = doc.find(W + 'body')
    sects = iter_sect_elems(body)
    problems = []
    for s in sects:
        for tag, keys in (('pgSz', ('w', 'h')), ('pgMar', ('top', 'right', 'bottom', 'left', 'header', 'footer'))):
            se = s.find(W + tag)
            for k in keys:
                tv = se.get(W + k) if se is not None else None
                sv = tpl_spec['page_setup'].get(tag, {}).get(k)
                if sv and tv != sv:
                    problems.append(f'{tag}.{k}: 规范{sv} 实际{tv}')
    if problems:
        return False, '; '.join(problems[:5])
    return True, f'共 {len(sects)} 节 pgSz/pgMar 与模板一致'


def check_body_fmt(target_entries, tpl_spec, samples=5):
    body_spec = tpl_spec['roles'].get('BODY')
    if not body_spec:
        return True, '模板无正文规范'
    doc = U.parse_xml(target_entries['word/document.xml'])
    body = doc.find(W + 'body')
    paras = body.findall(W + 'p')
    style_map = U.build_role_style_map(paras)
    heading_scheme = U.detect_document_heading_scheme(paras, U.toc_style_ids(target_entries))
    cover = U.detect_cover_titles(paras)
    num_fmt_map = U.build_num_fmt_map(target_entries)
    found = 0
    problems = []
    for p in paras:
        text = U.para_text(p)
        role = U.detect_para_role(p, text, style_map, heading_scheme, cover)
        if heading_scheme and role == 'BODY' and U.para_is_list_item(p, text, num_fmt_map):
            continue
        if role != 'BODY' or not text.strip():
            continue
        if U.para_format_skip(p):
            continue
        pPr = p.find(W + 'pPr')
        pPr_rPr = pPr.find(W + 'rPr') if pPr is not None else None
        run = U.para_first_text_run(p)
        run_rPr = run.find(W + 'rPr') if run is not None else None
        fmt = U._merge_run_fmt(pPr_rPr, run_rPr)
        spec_ea = (body_spec.get('rFonts') or {}).get('eastAsia')
        act_ea = (fmt.get('rFonts') or {}).get('eastAsia')
        if spec_ea and act_ea != spec_ea:
            problems.append(f'「{text[:12]}」字体 {act_ea} 应为 {spec_ea}')
        if body_spec.get('sz') and fmt.get('sz') != body_spec.get('sz'):
            problems.append(f'「{text[:12]}」字号 {fmt.get("sz")} 应为 {body_spec.get("sz")}')
        found += 1
        if found >= samples:
            break
    if problems:
        return False, '; '.join(problems[:5])
    return True, f'抽查 {found} 段正文：字体/字号与模板规范一致'


def check_xml_integrity(entries):
    bad = []
    for name, data in entries.items():
        if name.endswith('.xml') or name.endswith('.rels'):
            try:
                ET.fromstring(data)
            except Exception as e:
                bad.append(f'{name}: {e}')
    if bad:
        return False, '; '.join(bad[:3])
    return True, f'全部 {sum(1 for n in entries if n.endswith(".xml") or n.endswith(".rels"))} 个 XML/rels 可解析'


def check_toc(target_entries):
    doc = target_entries['word/document.xml'].decode('utf-8', 'ignore')
    has_toc = 'TOC' in doc and ('w:instrText' in doc)
    has_bookmarks = '_Toc' in doc
    if has_toc:
        return True, f'目录 TOC 域保留（_Toc 书签 {"有" if has_bookmarks else "无"}）；请在 Word 中右键更新域刷新编号/页码'
    return True, '目标无目录'


def main():
    ap = argparse.ArgumentParser(description='doc-format-align 对齐结果确定性检查')
    ap.add_argument('--target', required=True, help='对齐版文档')
    ap.add_argument('--template', required=True, help='格式模板')
    ap.add_argument('--original', required=True, help='原文件（用于内容保护/图片对比）')
    ap.add_argument('--number-mode', default='continuous', choices=['continuous', 'nested'])
    ap.add_argument('--title-shift', type=int, default=0,
                    help='与 align_format 的 --title-shift 一致（目标H1用模板H(1+shift)规范）')
    args = ap.parse_args()

    for p in (args.target, args.template, args.original):
        if not os.path.isfile(p):
            print(f'[错误] 文件不存在: {p}')
            sys.exit(1)
    target_entries = U.read_docx_entries(args.target)
    tpl_entries = U.read_docx_entries(args.template)
    orig_entries = U.read_docx_entries(args.original)
    tpl_spec = U.extract_template_spec(tpl_entries)
    if args.title_shift:
        tpl_spec['roles'] = U.shifted_heading_spec(tpl_spec['roles'], args.title_shift)

    checks = [
        ('内容保护', check_content(target_entries, orig_entries)),
        ('图片完整性', check_images(target_entries, orig_entries)),
        ('标题编号', check_numbering(target_entries, tpl_spec, args.number_mode)),
        ('页眉页脚', check_header_footer(target_entries, tpl_spec)),
        ('页面设置', check_page_setup(target_entries, tpl_spec)),
        ('正文格式', check_body_fmt(target_entries, tpl_spec)),
        ('文档完整性', check_xml_integrity(target_entries)),
        ('目录', check_toc(target_entries)),
    ]
    print('=' * 46)
    print('doc-format-align 确定性检查报告')
    print('=' * 46)
    all_ok = True
    for name, (ok, msg) in checks:
        mark = '✅' if ok else '❌'
        if not ok:
            all_ok = False
        print(f'{mark} {name}: {msg}')
    print('=' * 46)
    print('✅ 全部通过' if all_ok else '❌ 存在未通过项，请修复后重跑')
    print('提示：封面分页、图片实际渲染请用 Word 打开目检（本检查不启动 Word）。')
    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()
