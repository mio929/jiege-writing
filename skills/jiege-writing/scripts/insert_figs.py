#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把 build_docx.py 产出的 docx 里的图占位替换成「图 + 图注」。

用法
----
    python scripts/insert_figs.py spec.json draft.docx -o final.docx
    python scripts/insert_figs.py spec.json draft.docx -o final.docx --fig 1=figs/f1.jpg:8.6
    python scripts/insert_figs.py spec.json draft.docx --check

spec 约定
---------
1. 正文里图片占位段落，文本为 ``__FIG_1__``（``__FIG1__`` 也认）。
2. 占位段之后紧跟图注段，文本以 ``__CAP_1__`` 开头。
3. 图注里的文献编号写 ``@key@``，脚本按 spec 正文的首次出现顺序换成 ``[n]``，
   与 build_docx.py 编出的编号一致。用 ``@`` 而不是 ``{{}}``，是为了让图注
   不参与「首次出现」的排序。
4. 图片路径与宽度写在 spec 的 ``figures`` 里，命令行 ``--fig`` 可覆盖::

       "figures": {"1": {"src": "figs/f1.jpg", "width_cm": 8.6}}

为什么要包一层表格
------------------
Word 的 ``keepNext``（与下段同页）对含浮动/内嵌图片的段落并不可靠，实测图片
会留在页底、图注被推到下一页。把图和图注放进同一个单元格、行属性设
``cantSplit``，二者才真正绑在一起。图与图注写进同一段落也不可行——图片会压住
图注文字。
"""
import argparse
import json
import os
import re
import sys

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

FIG_TOKEN = re.compile(r"^__FIG_?(\d+)__$")
CAP_PREFIX = "__CAP_"
DEFAULT_ROW_CM = 14.2      # 版心 14.66 cm，留一点余量
MAX_WIDTH_CM = 14.6


# ------------------------------------------------------------------ spec
def load_spec(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def ref_order(spec):
    """按正文首次出现顺序给文献编号，口径与 build_docx.py 一致。"""
    order = {}
    for block in spec.get("blocks", []):
        text = block[1] if len(block) > 1 else ""
        for m in re.finditer(r"\{\{([^}]+)\}\}", text):
            body = m.group(1)
            if body.startswith("@"):
                continue
            for key in body.split("|"):
                key = key.strip()
                if key not in order:
                    order[key] = len(order) + 1
    return order


def figure_spec(spec, overrides):
    figs = {}
    for key, val in (spec.get("figures") or {}).items():
        figs[str(key)] = {"src": val["src"], "width_cm": float(val.get("width_cm", 12.0))}
    figs.update(overrides)
    return figs


# ------------------------------------------------------------------ docx
def style_run(run, pt):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:ascii"), "Times New Roman")
    rf.set(qn("w:hAnsi"), "Times New Roman")
    rf.set(qn("w:eastAsia"), "宋体")
    rf.set(qn("w:cs"), "Times New Roman")
    rf.set(qn("w:hint"), "eastAsia")
    run.font.size = Pt(pt)


def strip_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        borders.append(el)
    tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def forbid_split(table):
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def replace_caption(text, order):
    for key, num in order.items():
        text = text.replace("@%s@" % key, "[%d]" % num)
    left = re.findall(r"@[a-z0-9_]+@", text)
    if left:
        raise SystemExit("图注里有未定义或未在正文引用的键：%s" % left)
    return text


def build_figure_block(doc, src, width_cm, caption, order, cap_pt):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(DEFAULT_ROW_CM)
    table.rows[0].cells[0].width = Cm(DEFAULT_ROW_CM)
    strip_borders(table)
    forbid_split(table)

    cell = table.rows[0].cells[0]
    ip = cell.paragraphs[0]
    ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ip.paragraph_format.space_before = Pt(8)
    ip.paragraph_format.space_after = Pt(3)
    ip.paragraph_format.first_line_indent = Pt(0)
    ip.paragraph_format.line_spacing = 1.0
    run = ip.add_run()
    style_run(run, 12)
    run.add_picture(src, width=Cm(width_cm))

    cp = cell.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(3)
    cp.paragraph_format.space_after = Pt(8)
    cp.paragraph_format.first_line_indent = Pt(0)
    cp.paragraph_format.line_spacing = 1.0
    style_run(cp.add_run(replace_caption(caption, order)), cap_pt)
    return table


# ------------------------------------------------------------------ main
def parse_fig_args(items):
    out = {}
    for item in items or []:
        m = re.match(r"^(\d+)=(.+?)(?::([\d.]+))?$", item)
        if not m:
            raise SystemExit("--fig 格式应为 编号=路径[:宽度cm]，收到：%s" % item)
        out[m.group(1)] = {"src": m.group(2), "width_cm": float(m.group(3) or 12.0)}
    return out


def main():
    ap = argparse.ArgumentParser(description="把图占位替换为图 + 图注")
    ap.add_argument("spec")
    ap.add_argument("draft")
    ap.add_argument("-o", "--out")
    ap.add_argument("--fig", action="append", help="编号=图片路径[:宽度cm]，可重复")
    ap.add_argument("--caption-pt", type=float, default=9.0)
    ap.add_argument("--check", action="store_true", help="只列出待插入的图，不写文件")
    args = ap.parse_args()

    spec = load_spec(args.spec)
    order = ref_order(spec)
    figs = figure_spec(spec, parse_fig_args(args.fig))

    doc = docx.Document(args.draft)
    paras = list(doc.paragraphs)
    todo, problems = [], []

    for i, p in enumerate(paras):
        m = FIG_TOKEN.match(p.text.strip())
        if not m:
            continue
        num = m.group(1)
        cap = paras[i + 1] if i + 1 < len(paras) else None
        if cap is None or not cap.text.strip().startswith(CAP_PREFIX):
            problems.append("图 %s 后面没有 __CAP_%s__ 图注段" % (num, num))
            continue
        if num not in figs:
            problems.append("spec 的 figures 里没有第 %s 张图的路径" % num)
            continue
        src = figs[num]["src"]
        if not os.path.exists(src):
            problems.append("图片不存在：%s" % src)
            continue
        if figs[num]["width_cm"] > MAX_WIDTH_CM:
            problems.append("图 %s 宽度 %.1f cm 超过版心 %.1f cm"
                            % (num, figs[num]["width_cm"], MAX_WIDTH_CM))
        caption = cap.text.strip()
        caption = replace_caption(caption.replace(CAP_PREFIX + num + "__", ""), order)
        todo.append((num, src, figs[num]["width_cm"], caption, p, cap))

    for msg in problems:
        print("✗ " + msg, file=sys.stderr)
    if problems:
        raise SystemExit(1)

    if args.check:
        for num, src, w, caption, _, _ in todo:
            print("图 %-3s %5.1f cm  %-42s %s" % (num, w, os.path.basename(src), caption[:50]))
        print("共 %d 张，未写文件（--check）" % len(todo))
        return

    if not args.out:
        raise SystemExit("需要 -o 指定输出文件")
    for num, src, width, caption, ph, cap in todo:
        block = build_figure_block(doc, src, width, caption, order, args.caption_pt)
        ph._p.addprevious(block._tbl)      # 表格插到占位段之前
        ph._p.getparent().remove(ph._p)    # 删占位段
        cap._p.getparent().remove(cap._p)  # 删图注段

    doc.save(args.out)
    print("已插入 %d 张图 → %s" % (len(todo), args.out))


if __name__ == "__main__":
    main()
