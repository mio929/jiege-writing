#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""按中文课程论文/综述的格式规范，把 JSON 稿件渲染成 .docx。

用法:
    python build_docx.py spec.json -o 论文.docx
    python build_docx.py spec.json --check        # 只校验，不写文件

spec.json 结构见 assets/example_spec.json。要点：
  - blocks 里用 {{key}} 标参考文献，构建时按首次出现顺序自动编号并插入 REF 域；
  - refs 的值可以是已排好的字符串，也可以是 8 元组
    [作者, 题名, 刊名, 年, 卷, 期, 页码, DOI]，由脚本拼成 GB/T 7714-2025 格式；
  - 输出的编号是 Word 自动编号（[1]），正文引用是「段落编号」交叉引用，可 Ctrl+单击跳转。
"""
import argparse
import json
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt

CN_FONT = "宋体"
EN_FONT = "Times New Roman"
# 输出语种：zh 中文稿（宋体+Times New Roman），en 英文稿（全文 Times New Roman）
LANG = "zh"

TOKEN = re.compile(r"\{\{([^}]+)\}\}")
HEAD_NUM = re.compile(r"^(\d+(?:\.\d+)*)\s+(.*)$")
CJK = re.compile(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef\u2018\u2019\u201c\u201d]")
GREEK = re.compile(r"[\u0370-\u03ff\u1f00-\u1fff]")
GREEK_SPLIT = re.compile(r"([\u0370-\u03ff\u1f00-\u1fff]+)")

BM_REF = "_RefW"
BM_SEC = "_Sec"
ID_REF_BASE = 500
ID_SEC_BASE = 700
TAB_TWIPS = 420          # 参考文献悬挂缩进与制表位（21 pt）

PPR_SUCCESSORS = (
    "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
    "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing",
    "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)


# ------------------------------------------------------------------ 字体
def normalize_rfonts(rfonts, cn=CN_FONT, en=EN_FONT):
    for t in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(t)) is not None:
            del rfonts.attrib[qn(t)]
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:cs"), en)


def set_rfonts(rpr, cn=CN_FONT, en=EN_FONT):
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rstyle = rpr.find(qn("w:rStyle"))
        (rstyle.addnext(rfonts) if rstyle is not None else rpr.insert(0, rfonts))
    normalize_rfonts(rfonts, cn, en)


def patch_doc_defaults(doc, body_pt, line):
    dd = doc.styles.element.find(qn("w:docDefaults"))
    rpr = dd.find(qn("w:rPrDefault")).find(qn("w:rPr"))
    set_rfonts(rpr)
    half = str(int(body_pt * 2))
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rpr.append(el)
        el.set(qn("w:val"), half)

    ppr = dd.find(qn("w:pPrDefault")).find(qn("w:pPr"))
    if ppr.find(qn("w:widowControl")) is None:
        wc = OxmlElement("w:widowControl")
        wc.set(qn("w:val"), "0")
        ppr.insert(0, wc)
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        ppr.append(sp)
    sp.set(qn("w:line"), str(int(240 * line)))
    sp.set(qn("w:lineRule"), "auto")


def patch_styles(doc, body_pt, line):
    """styles.xml 内每个 rFonts 都规范化，含段落标记属性里的主题字体引用。"""
    for rf in doc.styles.element.iter(qn("w:rFonts")):
        normalize_rfonts(rf)
    normal = doc.styles["Normal"]
    set_rfonts(normal.element.get_or_add_rPr())
    normal.font.size = Pt(body_pt)
    pf = normal.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.widow_control = False


def patch_settings(doc):
    """兼容模式提到 Word 2013+，否则 Word 标题栏显示「兼容性模式」。"""
    compat = doc.settings.element.find(qn("w:compat"))
    if compat is None:
        return
    fe = compat.find(qn("w:useFELayout"))
    if fe is not None:
        compat.remove(fe)
    for cs in compat.findall(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")


# ------------------------------------------------------------------ 段落
SUP_CHARS = "⁰¹²³⁴⁵⁶⁷⁸⁹"
SUP_MAP = str.maketrans(SUP_CHARS, "0123456789")
RICH = re.compile(r"(\^\{[^}]*\}|_\{[^}]*\})")
SUP_RUN = re.compile(r"([%s]+)" % SUP_CHARS)


def add_run(p, text, size, bold=False, sup=False, sub=False):
    """带行内语法的文本：^{上标}、_{下标}，以及 Unicode 上标字符（²³¹）自动转真上标。"""
    if not text:
        return _emit(p, text, size, bold, sup, False, sub)
    last = None
    for tok in RICH.split(text):
        if not tok:
            continue
        if tok.startswith("^{") and tok.endswith("}"):
            last = add_run(p, tok[2:-1], size, bold, sup=True)
            continue
        if tok.startswith("_{") and tok.endswith("}"):
            last = add_run(p, tok[2:-1], size, bold, sub=True)
            continue
        for seg in SUP_RUN.split(tok):
            if not seg:
                continue
            if all(c in SUP_CHARS for c in seg):
                last = add_run(p, seg.translate(SUP_MAP), size, bold, sup=True)
            else:
                last = _split_greek(p, seg, size, bold, sup, sub)
    return last


def _split_greek(p, text, size, bold, sup, sub):
    if text and GREEK.search(text):
        last = None
        for seg in GREEK_SPLIT.split(text):
            if seg:
                last = _emit(p, seg, size, bold, sup, bool(GREEK.fullmatch(seg)), sub)
        return last
    return _emit(p, text, size, bold, sup, False, sub)


def _emit(p, text, size, bold, sup, greek=False, sub=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    rpr = r._element.get_or_add_rPr()
    set_rfonts(rpr)
    rfonts = rpr.find(qn("w:rFonts"))
    if LANG == "en":
        # 英文稿：四类字体位一律 Times New Roman，不设中文提示
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(attr), EN_FONT)
        rfonts.set(qn("w:hint"), "default")
    elif greek:
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(attr), EN_FONT)
        rfonts.set(qn("w:hint"), "default")
    elif text and CJK.search(text):
        rfonts.set(qn("w:hint"), "eastAsia")
    if sup:
        r.font.superscript = True
    if sub:
        r.font.subscript = True
    return r


def new_para(doc, indent_chars=0, align=None, before=0, after=0, line=1.5,
             hanging_pt=0, body_pt=12.0, indent_pt=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.widow_control = False
    if hanging_pt:
        pf.left_indent = Pt(hanging_pt)
        pf.first_line_indent = Pt(-hanging_pt)
        pf.tab_stops.add_tab_stop(Pt(hanging_pt))
    elif indent_pt:
        pf.first_line_indent = Pt(indent_pt)
    elif indent_chars:
        pf.first_line_indent = Pt(body_pt * indent_chars)
        p._p.get_or_add_pPr().find(qn("w:ind")).set(
            qn("w:firstLineChars"), str(int(indent_chars * 100)))
    return p


# ------------------------------------------------------------------ 编号
def add_numbering(doc, ref_pt):
    numbering = doc.part.numbering_part.element
    aids = [int(a.get(qn("w:abstractNumId")))
            for a in numbering.findall(qn("w:abstractNum"))]
    nids = [int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))]
    aid, nid = max(aids + [0]) + 1, max(nids + [0]) + 1

    abstract = parse_xml(
        '<w:abstractNum %s w:abstractNumId="%d">'
        '<w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0">'
        '<w:start w:val="1"/>'
        '<w:numFmt w:val="decimal"/>'
        '<w:suff w:val="tab"/>'
        '<w:lvlText w:val="[%%1]"/>'
        '<w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
        '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
        '</w:lvl></w:abstractNum>'
        % (nsdecls("w"), aid, TAB_TWIPS, TAB_TWIPS,
           EN_FONT, EN_FONT, CN_FONT, EN_FONT,
           int(ref_pt * 2), int(ref_pt * 2)))

    first_num = numbering.find(qn("w:num"))
    (first_num.addprevious(abstract) if first_num is not None else numbering.append(abstract))
    numbering.append(parse_xml(
        '<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
        % (nsdecls("w"), nid, aid)))
    return nid


def apply_numbering(p, num_id):
    ppr = p._p.get_or_add_pPr()
    numpr = OxmlElement("w:numPr")
    for tag, val in (("w:ilvl", "0"), ("w:numId", str(num_id))):
        el = OxmlElement(tag)
        el.set(qn("w:val"), val)
        numpr.append(el)
    ppr.insert_element_before(numpr, *PPR_SUCCESSORS)


# ------------------------------------------------------------------ 域
def _field(p, instr, cached, size, sup=False):
    r = add_run(p, None, size, sup=sup)
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "begin")
    r._element.append(fc)

    r = add_run(p, None, size, sup=sup)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    r._element.append(it)

    r = add_run(p, None, size, sup=sup)
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "separate")
    r._element.append(fc)

    add_run(p, cached, size, sup=sup)

    r = add_run(p, None, size, sup=sup)
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), "end")
    r._element.append(fc)


def open_bookmark(p, name, bid):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    p._p.get_or_add_pPr().addnext(start)


def close_bookmark(p, bid):
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    p._p.append(end)


# ------------------------------------------------------------------ 参考文献格式化
CN_P = {"sep": "，", "dot": "．", "colon": "：", "lp": "（", "rp": "）"}
EN_P = {"sep": ", ", "dot": ". ", "colon": ": ", "lp": "(", "rp": ")"}
CJK_RX = re.compile(r"[\u4e00-\u9fff]")


def _is_cn(*texts):
    """任一处出现汉字即按中文条目处理，标点随之切换。"""
    return any(CJK_RX.search(t or "") for t in texts)


def _names(authors, cn):
    """责任者超过 3 人取前 3 位，中文补「等」，西文补「et al」。"""
    if isinstance(authors, str):
        if cn:
            return authors.replace(", ", CN_P["sep"]).replace(",", CN_P["sep"])
        return authors
    names = [n for n in authors if n]
    if not names:
        return ""
    if len(names) > 3:
        head = (CN_P["sep"] if cn else EN_P["sep"]).join(names[:3])
        return head + ("，等" if cn else ", et al")
    return (CN_P["sep"] if cn else EN_P["sep"]).join(names)


def _publication(ref, cn, kind):
    """期刊给「年，卷（期）：页」，图书与学位论文给「出版地：出版者，年」。"""
    P = CN_P if cn else EN_P
    end = CN_P["dot"] if cn else "."
    if kind in ("J", "C", "A"):
        s = P["sep"] + str(ref.get("year", ""))
        if ref.get("vol"):
            s += P["sep"] + str(ref["vol"])
        if ref.get("issue"):
            s += P["lp"] + str(ref["issue"]) + P["rp"]
        if ref.get("pages"):
            s += P["colon"] + str(ref["pages"])
        return s + end
    s = "%s%s%s%s%s" % (ref.get("place", ""), P["colon"], ref.get("publisher", ""),
                        P["sep"], ref.get("year", ""))
    if ref.get("pages"):
        s += P["colon"] + str(ref["pages"])
    return s + end


def format_ref(ref):
    """把一条文献著录成 GB/T 7714 字符串。

    两种输入都认：
      - 字典 {"authors": [...], "title", "type", "journal" 或 "publisher", "year", ...}
      - 8/9 元组 [作者, 题名, 刊名, 年, 卷, 期, 页码, DOI, 类型]

    中文条目标点用全角（，．：（）），西文条目用半角；DOI 不列入条目。
    类型取 J（期刊）、M（图书）、D（学位论文）、C/A（会议）、R（报告）、S（标准）。
    """
    if isinstance(ref, dict):
        kind = (ref.get("type") or "J").upper()
        cn = ref.get("lang") == "cn" or _is_cn(ref.get("title"), ref.get("journal"))
        P = CN_P if cn else EN_P
        s = "%s%s%s[%s]%s" % (_names(ref.get("authors", ""), cn), P["dot"],
                               ref.get("title", ""), kind, P["dot"])
        if kind in ("J", "C", "A"):
            s += str(ref.get("journal", ""))
        return s + _publication(ref, cn, kind)

    vals = list(ref) + [""] * 9
    authors, title, journal, year, vol, issue, pages, _doi, kind = vals[:9]
    kind = (kind or "J").upper()
    cn = _is_cn(title, journal)
    P = CN_P if cn else EN_P
    s = "%s%s%s[%s]%s" % (_names(authors, cn), P["dot"], title, kind, P["dot"])
    if kind in ("J", "C", "A"):
        s += str(journal)
    return s + _publication(
        dict(year=year, vol=vol, issue=issue, pages=pages, place="", publisher=""),
        cn, kind)


def resolve_refs(raw):
    out = {}
    for k, v in raw.items():
        out[k] = format_ref(v) if isinstance(v, (list, tuple, dict)) else str(v)
    return out


# ------------------------------------------------------------------ 字数
HAN = re.compile(r"[\u4e00-\u9fff]")
WORD = re.compile(r"[A-Za-z][A-Za-z0-9'\-]*")


def _count(text):
    """中文数汉字，英文数词。"""
    clean = TOKEN.sub("", text or "")
    if LANG == "en":
        return len(WORD.findall(clean))
    return len(HAN.findall(clean))


def word_stats(spec, blocks):
    """按口径分列统计：正文、标题、摘要关键词。

    正文只算 p 段，不含 h1 标题与 __TABLE_n__ 占位段。
    """
    body = heads = 0
    for kind, text in blocks:
        n = _count(text)
        if kind == "h1":
            heads += n
        elif not text.startswith("__TABLE"):
            body += n
    front = sum(_count(x) for x in (spec.get("abstract"), spec.get("keywords")))
    return {"body": body, "headings": heads, "front": front}


def parse_target(t):
    """word_target 接受 [lo, hi] 或单值；单值按 ±5% 展开。"""
    if isinstance(t, (list, tuple)) and len(t) == 2:
        return int(t[0]), int(t[1])
    v = int(t)
    return int(round(v * 0.95)), int(round(v * 1.05))


def report_words(spec, stats):
    tgt = spec.get("word_target")
    unit = "词" if LANG == "en" else "汉字"
    msg = "字数：正文 %d %s（另标题 %d、摘要关键词 %d）" % (
        stats["body"], unit, stats["headings"], stats["front"])
    if not tgt:
        print(msg + "；未设 word_target", file=sys.stderr)
        return
    lo, hi = parse_target(tgt)
    if lo <= stats["body"] <= hi:
        tail = "在目标区间 %d–%d 内" % (lo, hi)
    elif stats["body"] < lo:
        tail = "低于目标 %d–%d，差 %d 字，按论证缺口补（见 writing-guide）" % (
            lo, hi, lo - stats["body"])
    else:
        tail = "超出目标 %d–%d，多 %d 字，删整段而不是等比例砍" % (
            lo, hi, stats["body"] - hi)
    print(msg + "；" + tail, file=sys.stderr)


# ------------------------------------------------------------------ 构建
def build(spec, out_path, check_only=False):
    refs = resolve_refs(spec.get("refs", {}))
    blocks = [tuple(b) for b in spec["blocks"]]
    st = spec.get("styles", {})
    title_pt = float(st.get("title_pt", 18))
    author_pt = float(st.get("author_pt", 14))
    body_pt = float(st.get("body_pt", 12))
    ref_pt = float(st.get("ref_pt", 10.5))
    line = float(st.get("line", 1.5))
    global LANG
    LANG = (spec.get("lang") or "zh").lower()
    en = LANG == "en"
    refs_title = spec.get("refs_title", "References" if en else "参考文献")
    lbl_abs, lbl_kw = ("Abstract: ", "Keywords: ") if en else ("摘要：", "关键词：")
    ind = dict(indent_pt=36) if en else dict(indent_chars=2)

    order, seq, secs = {}, [], []
    for kind, text in blocks:
        if kind == "h1":
            m = HEAD_NUM.match(text)
            if m:
                secs.append(m.group(1))
        for m in TOKEN.finditer(text):
            body = m.group(1)
            if body.startswith("@"):
                continue
            for k in body.split("|"):
                k = k.strip()
                if k not in order:
                    order[k] = len(order) + 1
                    seq.append(k)

    missing = [k for k in order if k not in refs]
    if missing:
        raise SystemExit("引用了未提供的文献键：%s" % missing)
    bad = sorted({s for s in re.findall(r"\{\{@([\d.]+)\}\}", "".join(t for _, t in blocks))
                  if s not in secs})
    if bad:
        raise SystemExit("引用了不存在的章节号：%s" % bad)
    unused = [k for k in refs if k not in order]
    if unused:
        print("提示：这些文献未被正文引用，不会进入列表：%s" % unused, file=sys.stderr)

    stats = word_stats(spec, blocks)
    report_words(spec, stats)

    if check_only:
        return {"refs": len(seq), "headings": len(secs), "words": stats}

    doc = Document()
    patch_doc_defaults(doc, body_pt, line)
    patch_settings(doc)
    patch_styles(doc, body_pt, line)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.17)

    p = new_para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, after=12, line=line, body_pt=body_pt)
    add_run(p, spec["title"], title_pt, bold=True)
    p = new_para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, after=18, line=line, body_pt=body_pt)
    add_run(p, spec.get("author", ""), author_pt, bold=True)

    p = new_para(doc, after=6, line=line, body_pt=body_pt, **ind)
    add_run(p, lbl_abs, body_pt, bold=True)
    fill(p, spec.get("abstract", ""), body_pt, False, order, ref_pt)
    p = new_para(doc, after=12, line=line, body_pt=body_pt, **ind)
    add_run(p, lbl_kw, body_pt, bold=True)
    add_run(p, spec.get("keywords", ""), body_pt)

    sec_ids = {s: ID_SEC_BASE + i for i, s in enumerate(secs, start=1)}
    for kind, text in blocks:
        if kind == "h1":
            m = HEAD_NUM.match(text)
            if m is None:
                raise SystemExit(
                    "标题需以「编号 + 半角空格」开头，如「1 引言」「2.1 方法」「2.1.1 样地设置」：%s" % text)
            num, rest = m.group(1), m.group(2)
            depth = num.count(".")
            before, after = {0: (18, 6), 1: (12, 4)}.get(depth, (6, 3))
            p = new_para(doc, before=before, after=after, line=line, body_pt=body_pt)
            p.paragraph_format.keep_with_next = True
            bid = sec_ids[num]
            open_bookmark(p, BM_SEC + num.replace(".", "_"), bid)
            add_run(p, num, body_pt, bold=True)
            close_bookmark(p, bid)
            add_run(p, " " + rest, body_pt, bold=True)
        else:
            p = new_para(doc, line=line, body_pt=body_pt, **ind)
            fill(p, text, body_pt, False, order, ref_pt)

    p = new_para(doc, before=18, after=6, line=line, body_pt=body_pt)
    p.paragraph_format.keep_with_next = True
    add_run(p, refs_title, body_pt, bold=True)

    num_id = add_numbering(doc, ref_pt)
    for i, key in enumerate(seq, start=1):
        p = new_para(doc, line=line, hanging_pt=TAB_TWIPS / 20.0, body_pt=body_pt)
        apply_numbering(p, num_id)
        open_bookmark(p, "%s%03d" % (BM_REF, i), ID_REF_BASE + i)
        add_run(p, refs[key], ref_pt)
        close_bookmark(p, ID_REF_BASE + i)

    doc.save(out_path)
    return {"refs": len(seq), "headings": len(secs), "words": stats, "out": out_path}


def compress_nums(nums):
    """把编号压成区间：连续 3 个以上用短横线，2 个仍用逗号分开（GB/T 7714 习惯）。

    [1,2,3,5,7,8] -> [(1,3), (5,None), (7,8)]
    """
    nums = sorted(set(nums))
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:
            out.append((nums[i], nums[j]))
        else:
            out.extend((n, None) for n in nums[i:j + 1])
        i = j + 1
    return out


def format_multi(nums):
    """[1,2,3] -> '[1-3]'；[1,3] -> '[1,3]'；[1,3,4,5,8] -> '[1,3-5,8]'。"""
    parts = []
    for a, b in compress_nums(nums):
        parts.append("%d-%d" % (a, b) if b is not None else str(a))
    return "[%s]" % ",".join(parts)


def fill(p, text, size, bold, order, ref_pt):
    pos = 0
    for m in TOKEN.finditer(text):
        if m.start() > pos:
            add_run(p, text[pos:m.start()], size, bold)
        body = m.group(1)
        if body.startswith("@"):
            _field(p, " REF %s%s \\h " % (BM_SEC, body[1:].replace(".", "_")), body[1:], size)
        else:
            nums = sorted(order[k.strip()] for k in body.split("|"))
            if len(nums) == 1:
                # 单条：用域，可 Ctrl+单击跳转。编号格式已是 [%1]，不再包括号
                _field(p, " REF %s%03d \\n \\h " % (BM_REF, nums[0]),
                       "[%d]" % nums[0], size, sup=True)
            else:
                # 多条：按区间压缩成纯文本。Word 的域无法表示 [1-3] 这种范围，
                # 每次构建都会重算编号，所以写死的文本始终与文献表一致。
                add_run(p, format_multi(nums), size, sup=True)
        pos = m.end()
    if pos < len(text):
        add_run(p, text[pos:], size, bold)


def main():
    ap = argparse.ArgumentParser(description="渲染中文论文 docx")
    ap.add_argument("spec", help="JSON 稿件文件")
    ap.add_argument("-o", "--out", default=None, help="输出 docx 路径")
    ap.add_argument("--check", action="store_true", help="只校验引用关系")
    args = ap.parse_args()

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)
    out = args.out or spec.get("out", "paper.docx")
    info = build(spec, out, check_only=args.check)
    print(json.dumps(info, ensure_ascii=False))


if __name__ == "__main__":
    main()
