# -*- coding: utf-8 -*-
"""
doc-format-align：把目标 docx 的格式对齐到模板 docx。

用法：
  python align_format.py --target <目标.docx> --template <模板.docx> [--output <输出.docx>] [--number-mode continuous|nested]

行为：
  1. 备份目标 -> <目标名>_备份.docx
  2. 提取模板规范（页面设置、页脚、各角色格式、编号体系）
  3. 注入：页面设置、页脚页码、段落角色格式、表格主样式、标题编号重写
  4. 内容保护：正文文本一字不改，仅标题编号文本按模板体系重写
"""
import argparse
import os
import re
import shutil
import sys
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

W = U.W
R = U.R
CT_NS = U.CT_NS


def apply_page_setup(entries, page_setup):
    """把模板页面设置写入目标 document.xml 的每个 sectPr。返回改动段落数。"""
    if not page_setup:
        return 0
    doc = U.parse_xml(entries['word/document.xml'])
    body = doc.find(W + 'body')
    sect_elems = [body.find(W + 'sectPr')] if body.find(W + 'sectPr') is not None else []
    sect_elems += body.findall('.//' + W + 'sectPr')
    # 去重（body.find('.//') 会包含 body 直属的）
    seen = set()
    unique = []
    for s in sect_elems:
        if id(s) not in seen:
            seen.add(id(s))
            unique.append(s)
    for s in unique:
        for tag in ('pgSz', 'pgMar', 'cols', 'docGrid'):
            tgt = s.find(W + tag)
            src = page_setup.get(tag)
            if src is None:
                continue
            if tgt is None:
                tgt = ET.SubElement(s, W + tag)
            else:
                # 清掉目标残留属性再写模板值
                for k in list(tgt.attrib):
                    if k.startswith(W):
                        del tgt.attrib[k]
            for k, v in src.items():
                tgt.set(W + k, v)
    entries['word/document.xml'] = U.to_bytes(doc)
    return len(unique)


def apply_footer(entries, footer_xml, doc_root_rels):
    """把模板页脚写入目标：新 footer 部件 + 各节 footerReference + 关系/Content-Type。"""
    if footer_xml is None:
        return
    # 规范化：命名空间正规化 + 剥离跨文档无效引用（否则 Word 会丢弃页脚内容）
    footer_xml = U.normalize_footer_xml(footer_xml)
    doc = U.parse_xml(entries['word/document.xml'])
    body = doc.find(W + 'body')

    # 找到已有的 footer 部件（若有，直接替换其 XML）
    existing_footer = None
    rels = U.parse_xml(entries['word/_rels/document.xml.rels'])
    for rel in rels:
        t = rel.get('Type', '')
        if t.endswith('/footer'):
            existing_footer = rel.get('Target')
            if existing_footer and existing_footer.startswith('/'):
                existing_footer = existing_footer[1:]
            elif existing_footer and not existing_footer.startswith('word/'):
                existing_footer = 'word/' + existing_footer
            break
    if existing_footer and existing_footer in entries:
        entries[existing_footer] = footer_xml
        rel_id = rel.get('Id')
    else:
        # 新建 footer 部件
        fname = 'word/footer2.xml'
        if fname in entries:
            i = 2
            while f'word/footer{i}.xml' in entries:
                i += 1
            fname = f'word/footer{i}.xml'
        entries[fname] = footer_xml
        # 关系（Target 相对 word/ 目录，不能带 word/ 前缀！）
        rel_id = 'rIdDfaFooter'
        rel = ET.SubElement(rels, U.PKG + 'Relationship')
        rel.set('Id', rel_id)
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
        rel.set('Target', fname.split('/')[-1])
        # Content-Type
        ct = U.parse_xml(entries['[Content_Types].xml'])
        already = False
        for ov in ct.findall(CT_NS + 'Override'):
            if ov.get('PartName') == '/' + fname:
                already = True
                break
        if not already:
            ov = ET.SubElement(ct, CT_NS + 'Override')
            ov.set('PartName', '/' + fname)
            ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
        entries['[Content_Types].xml'] = U.to_bytes(ct)
    entries['word/_rels/document.xml.rels'] = U.to_bytes(rels)

    # 各节 footerReference -> 新 footer
    sect_elems = [body.find(W + 'sectPr')] if body.find(W + 'sectPr') is not None else []
    sect_elems += body.findall('.//' + W + 'sectPr')
    seen = set()
    for s in sect_elems:
        if id(s) in seen:
            continue
        seen.add(id(s))
        fr = s.find(W + 'footerReference')
        if fr is None:
            fr = ET.SubElement(s, W + 'footerReference')
            # footerReference 需放在指定顺序（sectPr 子元素序）——插到 pgSz 前
            pg = s.find(W + 'pgSz')
            if pg is not None:
                s.remove(fr)
                s.insert(list(s).index(pg), fr)
        fr.set(R + 'id', rel_id)
        fr.set(W + 'type', 'default')
    # 说明：不需要手动加 xmlns:r 声明——ET.register_namespace 已注册 r 前缀，
    # 序列化时会自动在根元素输出 xmlns:r="...relationships"。手动 set xmlns 属性
    # 反而会被 ET 序列化成非法的 xmlns:ns2="...xmlns/"，必须避免。
    entries['word/document.xml'] = U.to_bytes(doc)


def remove_headers(doc, entries):
    """移除目标文档所有节的页眉（模板无页眉时调用，严格对齐模板）。

    每个 sectPr 可能带多个 headerReference（default/even/first 三种类型），
    必须用 findall 全部删除——只删第一个会残留 even 等类型；同时后一节无
    headerReference 时按 OOXML 规则会继承前一节页眉，所以必须每个节都清干净。
    header 部件与关系保留为孤儿（无害），保证操作可逆——原文件未动，随时可恢复。
    返回移除的页眉引用数。
    """
    body = doc.find(W + 'body')
    sect_elems = [body.find(W + 'sectPr')] if body.find(W + 'sectPr') is not None else []
    sect_elems += body.findall('.//' + W + 'sectPr')
    seen, n = set(), 0
    for s in sect_elems:
        if id(s) in seen:
            continue
        seen.add(id(s))
        for hr in s.findall(W + 'headerReference'):
            s.remove(hr)
            n += 1
    return n


def main():
    ap = argparse.ArgumentParser(description='doc-format-align 格式对齐')
    ap.add_argument('--target', required=True, help='目标文档路径')
    ap.add_argument('--template', required=True, help='格式模板路径')
    ap.add_argument('--output', default=None, help='输出路径（默认覆盖目标，先备份）')
    ap.add_argument('--number-mode', default='continuous', choices=['continuous', 'nested'],
                    help='编号重编方式：continuous=各级全文连续编号；nested=子级在上级下重新编号')
    ap.add_argument('--title-shift', type=int, default=0,
                    help='目标标题层级相对模板的偏移：1=目标H1用模板H2规范、目标H2用模板H3…'
                         '（模板最高级是封面/章节特有层、需忽略时用）。编号体系跟随偏移。')
    args = ap.parse_args()

    target = os.path.abspath(args.target)
    template = os.path.abspath(args.template)
    for p, what in ((target, '目标'), (template, '模板')):
        if not os.path.isfile(p):
            print(f'[错误] {what}文档不存在: {p}')
            sys.exit(1)

    # 1. 备份
    if args.output:
        out = os.path.abspath(args.output)
    else:
        base, ext = os.path.splitext(target)
        backup = f'{base}_备份{ext}'
        shutil.copy2(target, backup)
        print(f'[备份] {target} -> {backup}')
        out = target

    # 2. 读取 & 提取模板规范
    tpl_entries = U.read_docx_entries(template)
    spec = U.extract_template_spec(tpl_entries)
    print('[模板] 页面设置:', '有' if spec['page_setup'] else '无')
    print('[模板] 页脚:', '有' if spec['footer_xml'] else '无')
    for role in U.ROLES:
        s = spec['roles'].get(role)
        if s:
            fonts = (s.get('rFonts') or {}).get('eastAsia')
            num = s.get('number')
            numinfo = f" 编号={num['kind']['num']} {num['sample']!r}" if num else ''
            print(f'[模板] {role}: 字体={fonts} 字号={s.get("sz")} 加粗={s.get("b")}{numinfo}')
        else:
            print(f'[模板] {role}: 未识别')

    # 3. 读取目标
    entries = U.read_docx_entries(target)
    snapshot_before = U.document_text_snapshot(entries)

    # 4. 页面设置
    n = apply_page_setup(entries, spec['page_setup'])
    print(f'[对齐] 页面设置：更新 {n} 节')

    # 5. 页脚
    if spec['footer_xml'] is not None:
        apply_footer(entries, spec['footer_xml'], None)
        print('[对齐] 页脚页码：已套用模板页脚')
    else:
        print('[对齐] 页脚页码：模板无页脚，跳过')

    # 5b. 页眉：模板无页眉时移除目标遗留页眉（严格对齐模板；可逆——只删引用不删部件）
    if not spec['has_header']:
        doc0 = U.parse_xml(entries['word/document.xml'])
        n_hr = remove_headers(doc0, entries)
        if n_hr:
            entries['word/document.xml'] = U.to_bytes(doc0)
            print(f'[对齐] 页眉：模板无页眉，已移除目标 {n_hr} 处遗留页眉引用')
        else:
            print('[对齐] 页眉：模板无页眉，目标无页眉')

    # 6. 段落角色格式注入 + 编号重写
    doc = U.parse_xml(entries['word/document.xml'])
    body = doc.find(W + 'body')
    paras = body.findall(W + 'p')
    style_map = U.build_role_style_map(paras)
    heading_scheme = U.detect_document_heading_scheme(paras, U.toc_style_ids(entries))
    cover_set = U.detect_cover_titles(paras)
    if heading_scheme:
        print('[目标] 样式标题体系:', {k: v for k, v in heading_scheme.items()})
    if cover_set:
        print(f'[目标] 封面大标题: {len(cover_set)} 段（套模板大标题格式）')
    paras_roles = []
    role_counts = {}
    for p in paras:
        text = U.para_text(p)
        role = U.detect_para_role(p, text, style_map, heading_scheme, cover_set)
        role_counts[role] = role_counts.get(role, 0) + 1
        paras_roles.append((p, role))
    # 编号列表条目保持为正文（不提升为标题），记录供正文注入与编号重写使用：
    # 正文列表 = 正文格式 + 编号（1）（2）（3）（与标题编号 一、/（一）/1、 不重复）。
    num_fmt_map = U.build_num_fmt_map(entries)
    list_items = [i for i, (p, role) in enumerate(paras_roles)
                  if role == 'BODY' and U.para_is_list_item(p, U.para_text(p), num_fmt_map)]
    if list_items:
        print(f'[目标] 正文列表条目: {len(list_items)} 段（保持正文格式，编号用（1）（2）…）')
    print('[目标] 角色统计:', {k: role_counts.get(k, 0) for k in U.ROLES})

    # 层级偏移：目标 H(n) 用模板 H(n+shift) 规范/编号（如忽略模板"第一章"章节层 shift=1）
    spec_roles = spec['roles']
    if args.title_shift:
        spec_roles = U.shifted_heading_spec(spec_roles, args.title_shift)
        print(f'[目标] 标题层级偏移: +{args.title_shift}（目标H1→模板H{1 + args.title_shift}规范）')

    for i, (p, role) in enumerate(paras_roles):
        # TITLE 与标题一样全量对齐；正文保留局部强调
        U.apply_role_to_para(p, spec_roles.get(role), role, full_run=(role != 'BODY'))
        # 正文列表条目：正文格式且去除加粗/西文字体残留（之前可能是标题样式），
        # 并强制注入首行缩进（这些段最终是手动编号正文，须与模板正文格式一致）
        if i in list_items:
            U.apply_role_to_para(p, spec_roles.get('BODY'), 'BODY',
                                 full_run=False, force_unbold=True, force_ind=True)

    # 7. 编号重写（按模板各角色编号体系，含层级偏移）
    stats = U.renumber_document(paras_roles, spec_roles, mode=args.number_mode)
    for role in U.HROLES:
        s = stats[role]
        flag = '✓' if s['reason'] or s['renumbered'] else '–'
        print(f"[编号] {role}: 重编 {s['renumbered']} 段" +
              (f"，跳过 {s['skipped']} 段（{s['reason']}）" if s['skipped'] else ''))
    # 标题段落关闭样式自带自动编号（否则"1"与手动"一、"重复显示）。
    # 只对"模板有编号示范、已改写为手动编号"的级别关闭；模板未示范的层级
    # （如模板只有 H1、目标还有 H2）"只对齐字体、不动编号"——保留其原有自动编号。
    h_kinds = U.heading_number_kinds(spec_roles)
    n_closed = 0
    for p, role in paras_roles:
        if role == 'TITLE' or (role in U.HROLES and h_kinds.get(role)):
            if U.para_has_auto_number(p) or p.find(U.W + 'pPr') is not None and \
                    p.find(U.W + 'pPr').find(U.W + 'pStyle') is not None:
                U.disable_auto_number(p)
                n_closed += 1
    if n_closed:
        print(f'[编号] 已关闭 {n_closed} 段标题的样式自动编号（避免与手动编号重复）')

    # 7b. 正文列表编号：正文列表条目重写为（1）（2）（3）…（不与标题编号重复）
    if list_items:
        n_list = U.renumber_body_lists(paras_roles, num_fmt_map, list_style='paren')
        print(f'[编号] 正文列表: {n_list} 段编号重写为（1）（2）…')
    entries['word/document.xml'] = U.to_bytes(doc)

    # 8. 表格主样式
    apply_table_style(entries, tpl_entries)

    # 9. 内容保护验证
    snapshot_after = U.document_text_snapshot(entries)
    mismatches = [i for i, (a, b) in enumerate(zip(snapshot_before, snapshot_after)) if a != b]
    if mismatches:
        print(f'[警告] 内容保护校验发现 {len(mismatches)} 处正文文本差异（段落索引: {mismatches[:10]}），已中止保存！')
        sys.exit(2)
    print('[校验] 内容保护：通过（正文文本零改动）')

    # 10. 写回
    U.write_docx(entries, out)
    print(f'[完成] 已输出: {out}')
    print('[提示] 建议运行 verify_format.py 生成对齐报告并目检')


def _table_cell_font(tbl, sroot=None):
    """提取表格的中文字体（eastAsia）。

    遍历表格所有单元格 run 的 rPr，取第一个显式 eastAsia；单元格均无显式声明时，
    用模板 Normal 样式 rPr 的 eastAsia（表格字体通常继承 Normal）。
    返回字体名或 None。
    """
    for tc in tbl.findall('.//' + W + 'tc'):
        for p in tc.findall(W + 'p'):
            for r in p.findall(W + 'r'):
                rPr = r.find(W + 'rPr')
                if rPr is None:
                    continue
                rf = rPr.find(W + 'rFonts')
                if rf is not None and rf.get(W + 'eastAsia'):
                    return rf.get(W + 'eastAsia')
    # 无显式声明：用 Normal 样式字体（表格继承 Normal）
    if sroot is not None:
        for st in sroot.findall(W + 'style'):
            nm = st.find(W + 'name')
            if nm is not None and nm.get(W + 'val') == 'Normal':
                rpr = st.find(W + 'rPr')
                if rpr is not None:
                    rf = rpr.find(W + 'rFonts')
                    if rf is not None and rf.get(W + 'eastAsia'):
                        return rf.get(W + 'eastAsia')
    return None


def _apply_table_font(tbl, font):
    """把指定中文字体写入表格所有单元格 run（只改 eastAsia，保留字号等原样）。"""
    for tc in tbl.findall('.//' + W + 'tc'):
        for p in tc.findall(W + 'p'):
            for r in p.findall(W + 'r'):
                rPr = r.find(W + 'rPr')
                if rPr is None:
                    rPr = ET.SubElement(r, W + 'rPr')
                    r.insert(0, rPr)
                rf = rPr.find(W + 'rFonts')
                if rf is None:
                    rf = ET.SubElement(rPr, W + 'rFonts')
                rf.set(W + 'eastAsia', font)


def apply_table_style(target_entries, tpl_entries):
    """把模板表格样式套用到目标所有表格，并按用户偏好对齐字体。

    处理规则（用户确认）：
    - **表格样式**：按样式 name（如 Table Grid）匹配/复制到目标 styles.xml（不能只复制
      styleId 引用——跨文档 styleId 可能冲突，实测模板 Table Grid=5、目标 5 是 heading 4）；
    - **中文字体**：模板表格单元格的中文字体（继承 Normal 时取 Normal 的 eastAsia）写入
      目标所有表格单元格，**字号保持目标原样**（不同表格列数/字号不同，只对齐字体与样式）。
    """
    tpl_doc = U.parse_xml(tpl_entries['word/document.xml'])
    tpl_body = tpl_doc.find(W + 'body')
    tpl_tbl = tpl_body.find('.//' + W + 'tbl')
    if tpl_tbl is None:
        return
    tpl_tblPr = tpl_tbl.find(W + 'tblPr')
    if tpl_tblPr is None:
        return
    tpl_style_id = None
    ps = tpl_tblPr.find(W + 'tblStyle')
    if ps is not None:
        tpl_style_id = ps.get(W + 'val')
    # 模板表格中文字体（继承 Normal 时用 Normal 的 eastAsia）
    tpl_sroot = None
    if 'word/styles.xml' in tpl_entries:
        tpl_sroot = U.parse_xml(tpl_entries['word/styles.xml'])
    tpl_font = _table_cell_font(tpl_tbl, tpl_sroot)

    tpl_style_name = None
    tpl_style_def = None
    if tpl_style_id and tpl_sroot is not None:
        for st in tpl_sroot.findall(W + 'style'):
            if st.get(W + 'styleId') == tpl_style_id:
                nm = st.find(W + 'name')
                tpl_style_name = nm.get(W + 'val') if nm is not None else None
                tpl_style_def = st
                break
    # 目标 styles.xml：按 name 找同名表格样式，或复制定义
    target_style_id = None
    if 'word/styles.xml' in target_entries:
        sroot2 = U.parse_xml(target_entries['word/styles.xml'])
        if tpl_style_name:
            for st in sroot2.findall(W + 'style'):
                nm = st.find(W + 'name')
                if nm is not None and nm.get(W + 'val') == tpl_style_name:
                    target_style_id = st.get(W + 'styleId')
                    break
        if target_style_id is None and tpl_style_def is not None:
            import copy as _copy
            new_st = _copy.deepcopy(tpl_style_def)
            new_id = tpl_style_id
            if sroot2.find(f'.//{W}style[@{{{W}}}styleId="{tpl_style_id}"]') is not None:
                i = 1
                while sroot2.find(f'.//{W}style[@{{{W}}}styleId="DfaTbl{i}"]') is not None:
                    i += 1
                new_id = f'DfaTbl{i}'
                new_st.set(W + 'styleId', new_id)
            sroot2.append(new_st)
            target_entries['word/styles.xml'] = U.to_bytes(sroot2)
            target_style_id = new_id
    # 目标表格：写 tblStyle 引用 + 注入中文字体（字号保持原样）
    doc = U.parse_xml(target_entries['word/document.xml'])
    body = doc.find(W + 'body')
    n = 0
    for tbl in body.findall('.//' + W + 'tbl'):
        tblPr = tbl.find(W + 'tblPr')
        if tblPr is None:
            tblPr = ET.SubElement(tbl, W + 'tblPr')
            tbl.insert(0, tblPr)
        if target_style_id:
            ts = tblPr.find(W + 'tblStyle')
            if ts is None:
                ts = ET.SubElement(tblPr, W + 'tblStyle')
            ts.set(W + 'val', target_style_id)
        if tpl_font:
            _apply_table_font(tbl, tpl_font)
        n += 1
    if n:
        target_entries['word/document.xml'] = U.to_bytes(doc)
        if tpl_font:
            print(f'[对齐] 表格：{n} 个表格套用模板样式 "{tpl_style_name or tpl_style_id}"'
                  f'（目标 styleId={target_style_id}），中文字体→{tpl_font}（字号保持原样）')
        else:
            print(f'[对齐] 表格：{n} 个表格套用模板样式 "{tpl_style_name or tpl_style_id}"'
                  f'（目标 styleId={target_style_id}）')
    else:
        print('[对齐] 表格：目标无表格')


if __name__ == '__main__':
    main()
