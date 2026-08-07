# -*- coding: utf-8 -*-
"""
doc-format-align 验证：对比对齐结果与模板规范，生成对齐报告。

用法：
  python verify_format.py --target <对齐后.docx> --template <模板.docx> [--report <报告.md>]
"""
import argparse
import os
import sys
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

W = U.W
R = U.R
CT_NS = U.CT_NS


def para_actual_format(p):
    """读取段落实际格式关键项（与 RoleSpec 同构）"""
    pPr = p.find(W + 'pPr')
    pPr_rPr = pPr.find(W + 'rPr') if pPr is not None else None
    run = U.para_first_text_run(p)
    run_rPr = run.find(W + 'rPr') if run is not None else None
    base = U._merge_run_fmt(pPr_rPr, run_rPr)
    fmt = {
        'rFonts': base.get('rFonts'),
        'sz': base.get('sz'),
        'b': base.get('b'),
        'jc': None,
        'spacing': {},
        'ind': {},
    }
    if pPr is not None:
        jc = pPr.find(W + 'jc')
        if jc is not None:
            fmt['jc'] = jc.get(W + 'val')
        sp = pPr.find(W + 'spacing')
        if sp is not None:
            for a in ('line', 'lineRule', 'before', 'after'):
                v = sp.get(W + a)
                if v:
                    fmt['spacing'][a] = v
        ind = pPr.find(W + 'ind')
        if ind is not None:
            for a in ('firstLineChars', 'firstLine', 'leftChars', 'left'):
                v = ind.get(W + a)
                if v:
                    fmt['ind'][a] = v
    return fmt


def role_mismatch_reasons(actual, spec, full=True, skip=None):
    """返回实际格式 vs 规范的差异列表；空列表 = 一致。

    skip：该段落应跳过的检查项（'spacing'/'ind'，见 ooxml_util.para_format_skip），
    图片段落与列表段落不检查对应项，避免误报。
    """
    skip = skip or set()
    if spec is None:
        return ['模板未定义该角色规范']
    reasons = []
    # 字体：比较 eastAsia
    spec_ea = (spec.get('rFonts') or {}).get('eastAsia')
    act_ea = (actual.get('rFonts') or {}).get('eastAsia')
    if spec_ea and act_ea != spec_ea:
        reasons.append(f'字体: 规范{spec_ea} 实际{act_ea}')
    # 字号
    if spec.get('sz') and actual.get('sz') != spec.get('sz'):
        reasons.append(f'字号: 规范{spec.get("sz")} 实际{actual.get("sz")}')
    # 加粗：模板不加粗(b=False)时，实际"无 b 元素"(None)与显式 val=0(False)都算不加粗
    if full and spec.get('b') is not None:
        act_b = actual.get('b')
        if spec['b'] and act_b is not True:
            reasons.append(f'加粗: 规范{spec.get("b")} 实际{act_b}')
        elif not spec['b'] and act_b not in (None, False):
            reasons.append(f'加粗: 规范{spec.get("b")} 实际{act_b}')
    # 行距
    if 'spacing' not in skip:
        for a in ('line', 'lineRule'):
            if spec.get('spacing', {}).get(a) and actual.get('spacing', {}).get(a) != spec['spacing'][a]:
                reasons.append(f'行距.{a}: 规范{spec["spacing"][a]} 实际{actual.get("spacing", {}).get(a)}')
    # 缩进
    if 'ind' not in skip:
        for a in ('firstLineChars', 'firstLine'):
            if spec.get('ind', {}).get(a) and actual.get('ind', {}).get(a) != spec['ind'][a]:
                reasons.append(f'缩进.{a}: 规范{spec["ind"][a]} 实际{actual.get("ind", {}).get(a)}')
    # 对齐
    if spec.get('jc') and actual.get('jc') != spec.get('jc'):
        reasons.append(f'对齐: 规范{spec.get("jc")} 实际{actual.get("jc")}')
    return reasons


def find_footer_in(entries):
    """检测 docx 是否有页脚部件"""
    rels_name = 'word/_rels/document.xml.rels'
    if rels_name not in entries:
        return None
    rels = U.parse_xml(entries[rels_name])
    for rel in rels:
        t = rel.get('Type', '')
        if t.endswith('/footer'):
            target = rel.get('Target', '').lstrip('/')
            if not target.startswith('word/'):
                target = 'word/' + target
            if target in entries:
                return target
    return None


def build_report(target_entries, tpl_entries, spec, target_path, number_mode='continuous'):
    """构建报告文本"""
    L = []
    L.append('# 文档格式对齐报告')
    L.append('')
    L.append(f'- 目标文档: `{target_path}`')

    doc = U.parse_xml(target_entries['word/document.xml'])
    body = doc.find(W + 'body')
    paras = body.findall(W + 'p')
    style_map = U.build_role_style_map(paras)
    heading_scheme = U.detect_document_heading_scheme(paras, U.toc_style_ids(target_entries))
    cover_set = U.detect_cover_titles(paras)
    num_fmt_map = U.build_num_fmt_map(target_entries)
    paras_roles = []
    for p in paras:
        role = U.detect_para_role(p, U.para_text(p), style_map, heading_scheme, cover_set)
        # 正文列表条目保持 BODY 角色（正文格式 + （1）（2）编号），不提升为标题
        paras_roles.append((p, role))

    # 1. 页面设置
    L.append('')
    L.append('## 1. 页面设置')
    tgt_sect = body.find(W + 'sectPr')
    if tgt_sect is not None and spec.get('page_setup'):
        ok = True
        for tag, keys in (('pgSz', ('w', 'h')), ('pgMar', ('top', 'right', 'bottom', 'left', 'header', 'footer'))):
            se = tgt_sect.find(W + tag)
            for k in keys:
                tv = se.get(W + k) if se is not None else None
                sv = spec['page_setup'].get(tag, {}).get(k)
                if sv and tv != sv:
                    ok = False
                    L.append(f'- ⚠️ {tag}.{k}: 规范={sv} 实际={tv}')
        if ok:
            L.append('- ✅ 页边距/纸张与模板一致')
    else:
        L.append('- ⚠️ 无法对比页面设置')

    # 2. 页脚
    L.append('')
    L.append('## 2. 页眉页脚与页码')
    tpl_footer = spec.get('footer_xml') is not None
    tgt_footer = find_footer_in(target_entries) is not None
    if tpl_footer and tgt_footer:
        L.append('- ✅ 页脚（页码）已套用模板')
    elif tpl_footer and not tgt_footer:
        L.append('- ⚠️ 模板有页脚，目标未检出页脚')
    elif not tpl_footer:
        L.append('- ✅ 模板无页脚（无需对齐）')
    if spec.get('has_header'):
        L.append('- 🖐 模板含页眉，请人工核对')

    # 3. 角色格式
    L.append('')
    L.append('## 3. 段落格式（按角色）')
    for role in U.ROLES:
        spec_role = spec['roles'].get(role)
        role_paras = [idx for idx, (p, r) in enumerate(paras_roles) if r == role]
        if not role_paras:
            L.append(f'- {role}: 目标无该角色段落' + ('（模板规范: 见上）' if spec_role else ''))
            continue
        mism = []
        for idx in role_paras:
            skip = U.para_format_skip(paras[idx])
            reasons = role_mismatch_reasons(para_actual_format(paras[idx]), spec_role,
                                            full=(role != 'BODY'), skip=skip)
            if reasons:
                mism.append((idx, reasons))
        if not spec_role:
            L.append(f'- {role}: ⚠️ 模板未识别该角色规范（{len(role_paras)} 段未应用）')
            continue
        if mism:
            L.append(f'- {role}: ⚠️ {len(mism)}/{len(role_paras)} 段与规范不完全一致')
            for idx, reasons in mism[:8]:
                txt = U.para_text(paras[idx])[:18]
                L.append(f'  - 段落[{idx}]「{txt}」: {"; ".join(reasons)}')
            if len(mism) > 8:
                L.append(f'  - …共 {len(mism)} 段')
        else:
            L.append(f'- {role}: ✅ {len(role_paras)} 段全部与模板一致')

    # 4. 标题编号
    L.append('')
    L.append('## 4. 标题编号')
    counters = {r: 0 for r in U.HROLES}
    last_num = {}
    kinds = U.heading_number_kinds(spec['roles'])
    n_changed = 0
    n_ok = 0
    for idx, (p, role) in enumerate(paras_roles):
        if role not in U.HROLES:
            continue
        kind = kinds[role]
        text = U.para_text(p)
        old, old_len = U.match_leading_number(text)
        if kind is None:
            continue
        # nested 模式：本级出现时后代计数器清零（与 align 保持一致）
        if number_mode == 'nested':
            role_idx = U._HEADING_ROLES.index(role)
            for d in U._HEADING_ROLES[role_idx + 1:]:
                counters[d] = 0
        counters[role] += 1
        parent = U._HEADING_ROLES[U._HEADING_ROLES.index(role) - 1] if role != U._HEADING_ROLES[0] else None
        expect = U.make_number(kind, counters[role], last_num.get(parent))
        if old == expect:
            n_ok += 1
        else:
            n_changed += 1
            L.append(f'- ⚠️ 段落[{idx}]「{text[:24]}」: 编号 {old!r} → 应为 {expect!r}')
        last_num[role] = expect
    if n_changed == 0:
        L.append(f'- ✅ 全部标题编号与模板体系一致（{n_ok} 段）')
    else:
        L.append(f'- 共 {n_changed} 段编号待确认，{n_ok} 段已一致')
    # 自动编号残留检查：模板有编号示范的级别应已关闭自动编号（避免与手动编号重复）。
    # 模板未示范的级别（如模板只有 H1、目标还有 H2）保留原有自动编号，不视为残留。
    resid = [idx for idx, (p, r) in enumerate(paras_roles)
             if r in U.HROLES and kinds.get(r) and U.para_has_auto_number(p)]
    if resid:
        L.append(f'- ⚠️ {len(resid)} 段标题仍有样式自动编号（会与手动编号重复）: {resid[:8]}')

    # 4b. 正文列表编号（（1）（2）…，每章内重编，不与标题编号重复）
    # 按文档顺序遍历：遇标题重置计数、遇 BODY 列表项比对编号
    k = 0
    n_list = 0
    list_bad = 0
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
            list_bad += 1
            L.append(f'- ⚠️ 段落[{idx}]「{text[:22]}」: 编号 {old!r} 应为 {expect!r}')
        if U.para_has_auto_number(p):
            list_bad += 1
            L.append(f'- ⚠️ 段落[{idx}] 仍有自动编号残留')
        k += 1
    if n_list:
        if list_bad == 0:
            L.append(f'- ✅ {n_list} 段正文列表编号为（1）（2）…（每章内重编），与标题编号不重复')
        else:
            L.append(f'- ⚠️ 共 {list_bad} 处正文列表编号问题（{n_list} 段）')

    # 5. 表格
    L.append('')
    L.append('## 5. 表格样式')
    tbls = body.findall('.//' + W + 'tbl')
    if not tbls:
        L.append('- ✅ 目标无表格')
    else:
        L.append(f'- 🖐 目标有 {len(tbls)} 个表格，请人工核对表样式是否与模板一致')

    # 6. 目录
    L.append('')
    L.append('## 6. 目录 (TOC)')
    L.append('- 🖐 若目标含目录，请在 Word 中右键更新目录以套用新样式')

    # 7. 汇总
    L.append('')
    L.append('## 汇总')
    L.append('- ✅ 表示与模板一致；⚠️ 表示需注意；🖐 表示需人工核对')
    return '\n'.join(L)


def main():
    ap = argparse.ArgumentParser(description='doc-format-align 对齐验证')
    ap.add_argument('--target', required=True, help='对齐后的目标文档')
    ap.add_argument('--template', required=True, help='格式模板')
    ap.add_argument('--report', default=None, help='报告输出路径（md）')
    ap.add_argument('--number-mode', default='continuous', choices=['continuous', 'nested'],
                    help='编号预期方式，须与 align_format 一致')
    ap.add_argument('--title-shift', type=int, default=0,
                    help='与 align_format 的 --title-shift 一致（目标H1用模板H(1+shift)规范）')
    args = ap.parse_args()

    target = os.path.abspath(args.target)
    template = os.path.abspath(args.template)
    for p, what in ((target, '目标'), (template, '模板')):
        if not os.path.isfile(p):
            print(f'[错误] {what}文档不存在: {p}')
            sys.exit(1)

    tpl_entries = U.read_docx_entries(template)
    spec = U.extract_template_spec(tpl_entries)
    if args.title_shift:
        spec['roles'] = U.shifted_heading_spec(spec['roles'], args.title_shift)
    tgt_entries = U.read_docx_entries(target)
    report = build_report(tgt_entries, tpl_entries, spec, target, args.number_mode)

    print(report)
    if args.report:
        with open(args.report, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f'\n[报告已写入] {args.report}')


if __name__ == '__main__':
    main()
