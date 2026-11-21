#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""交付前自检：读回 .docx，把格式指标打出来，异常项标红提示。

用法:
    python check_docx.py 论文.docx
    python check_docx.py 论文.docx --strict     # 有异常则退出码 1
    python check_docx.py 论文.docx --refs-title 参考文献
"""
import argparse
import re
import sys
import zipfile
from collections import Counter

from docx import Document
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CN_FONT = "宋体"
EN_FONT = "Times New Roman"
CJK = re.compile(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef\u2018\u2019\u201c\u201d]")
GREEK = re.compile(r"[\u0370-\u03ff\u1f00-\u1fff]")


def run_text(r):
    return "".join(t.text or "" for t in r.iter(W + "t"))


def inspect(path, refs_title="参考文献", target=None, by="body"):
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read("word/document.xml"))
    paras = doc.find(W + "body").findall(W + "p")
    runs = list(doc.iter(W + "r"))
    problems = []

    # ---------- 字体 ----------
    combos = Counter((rf.get(W + "ascii"), rf.get(W + "eastAsia"))
                     for rf in doc.iter(W + "rFonts"))
    theme = sum(1 for p in ("word/document.xml", "word/styles.xml")
                for rf in etree.fromstring(z.read(p)).iter(W + "rFonts")
                if any(rf.get(W + t) is not None
                       for t in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")))
    no_rf = [r for r in runs
             if r.find(W + "rPr") is None or r.find(W + "rPr").find(W + "rFonts") is None]
    cjk_runs = [r for r in runs if CJK.search(run_text(r)) and not GREEK.search(run_text(r))]
    no_hint = [r for r in cjk_runs
               if r.find(W + "rPr").find(W + "rFonts").get(W + "hint") != "eastAsia"]
    greek_runs = [r for r in runs if GREEK.search(run_text(r))]
    greek_bad = [r for r in greek_runs
                 if r.find(W + "rPr").find(W + "rFonts").get(W + "eastAsia") != EN_FONT]

    # ---------- 字号 / 上标 ----------
    sizes = Counter(t.get(W + "val") for t in doc.iter(W + "sz"))
    sup = [r for r in runs
           if r.find(W + "rPr") is not None and r.find(W + "rPr").find(W + "vertAlign") is not None]
    sup_sizes = Counter(r.find(W + "rPr").find(W + "sz").get(W + "val")
                        for r in sup if r.find(W + "rPr").find(W + "sz") is not None)

    # ---------- 段落 ----------
    spacing = Counter()
    widow_off = 0
    for p in paras:
        pr = p.find(W + "pPr")
        s = pr.find(W + "spacing") if pr is not None else None
        spacing[(s.get(W + "line"), s.get(W + "lineRule")) if s is not None else None] += 1
        wc = pr.find(W + "widowControl") if pr is not None else None
        if wc is not None and wc.get(W + "val") == "0":
            widow_off += 1
    if widow_off != len(paras):
        problems.append("有 %d 个段落未关闭孤行控制" % (len(paras) - widow_off))

    # ---------- 域与书签 ----------
    instr = [e.text or "" for e in doc.iter(W + "instrText")]
    ref_fields = [i for i in instr if "\\n" in i]
    sec_fields = [i for i in instr if "\\n" not in i]
    bookmarks = {b.get(W + "name") for b in doc.iter(W + "bookmarkStart")}
    dangling = [i for i in instr
                if (m := re.match(r"^\s*REF\s+(\S+)", i)) and m.group(1) not in bookmarks]
    no_link = [i for i in instr if "\\h" not in i]
    if dangling:
        problems.append("%d 个域指向不存在的书签" % len(dangling))
    if no_link:
        problems.append("%d 个域缺少 \\h，无法跳转" % len(no_link))

    # ---------- 正文与文献 ----------
    d = Document(path)
    text = [p.text for p in d.paragraphs]
    try:
        cut = text.index(refs_title)
    except ValueError:
        cut = len(text)
        problems.append("未找到「%s」标题行" % refs_title)

    # 分列统计：题目与空段跳过，摘要关键词、表题、各级标题单独计
    front = heads = caps = 0
    body_parts = []
    for i, t in enumerate(text[:cut]):
        s = t.strip()
        if i == 0 or not s:
            continue
        if s.startswith("摘要：") or s.startswith("摘要:"):
            front += len(re.findall(r"[\u4e00-\u9fff]", s))
            continue
        if s.startswith("关键词：") or s.startswith("关键词:"):
            front += len(re.findall(r"[\u4e00-\u9fff]", s))
            continue
        if re.match(r"^表\s*\d", s):
            caps += len(re.findall(r"[\u4e00-\u9fff]", s))
            continue
        if re.match(r"^\d+(\.\d+)*\s\S", s) and len(s) < 40:
            heads += len(re.findall(r"[\u4e00-\u9fff]", s))
            continue
        body_parts.append(s)

    body = "".join(body_parts)
    han = len(re.findall(r"[\u4e00-\u9fff]", body))
    words = len(re.findall(r"[A-Za-z][A-Za-z0-9\-.']*", body))
    tbl_han = sum(len(re.findall(r"[\u4e00-\u9fff]", c.text))
                  for t in d.tables for r in t.rows for c in r.cells)
    approx = han + words + front + heads + caps + tbl_han
    entries = len(text) - cut - 1 if cut < len(text) else 0
    cited = set()
    for i in ref_fields:
        m = re.search(r"REF\s+(\S+)", i)
        if m:
            cited.add(int(m.group(1)[-3:]))
    if entries and len(cited) != entries:
        problems.append("正文引用 %d 条，文末条目 %d 条，不一一对应" % (len(cited), entries))

    # ---------- 输出 ----------
    print("文件：%s" % path)
    print("段落数：%d" % len(paras))
    print("字体组合：%s" % dict(combos))
    if no_rf:
        problems.append("%d 个 run 未设置 rFonts" % len(no_rf))
    if theme:
        problems.append("仍有 %d 处主题字体引用（asciiTheme 等）" % theme)
    if no_hint:
        problems.append("%d 个含中文的 run 缺 hint=\"eastAsia\"" % len(no_hint))
    if greek_bad:
        problems.append("%d 个希腊字母 run 的 eastAsia 不是 Times New Roman" % len(greek_bad))
    print("中文 run：%d（缺 hint %d）｜希腊字母 run：%d" % (len(cjk_runs), len(no_hint), len(greek_runs)))
    print("字号分布（半磅）：%s" % dict(sizes))
    print("上标：%d 处，字号 %s" % (len(sup), dict(sup_sizes)))
    if len(sup_sizes) > 1:
        problems.append("上标字号不统一：%s" % dict(sup_sizes))
    print("行距：%s" % dict(spacing))
    if len(spacing) > 1:
        problems.append("行距不统一：%s" % dict(spacing))
    print("孤行控制关闭：%d / %d" % (widow_off, len(paras)))
    print("域：文献 %d ｜章节 %d ｜书签 %d" % (len(ref_fields), len(sec_fields), len(bookmarks)))
    print("文献：%d 条，正文引用涉及 %d 条" % (entries, len(cited)))
    print("字数：正文 %d 汉字 + %d 英文词 ｜一览 %d ｜摘要关键词 %d ｜表题 %d ｜表格内 %d"
          % (han, words, heads, front, caps, tbl_han))
    print("      Word 口径估算 %d（学校多按此查，含表格）" % approx)
    if target:
        lo, hi = target
        val = han if by == "body" else approx
        label = "正文汉字" if by == "body" else "Word 估算"
        if lo <= val <= hi:
            print("      目标 %d–%d：达标（按%s %d）" % (lo, hi, label, val))
            if by == "body" and approx > hi:
                print("      注意：Word 口径为 %d，已超 %d 字。学校若按 Word 查，"
                      "改用 --by word 复核" % (approx, approx - hi))
        else:
            gap = val - hi if val > hi else lo - val
            problems.append("%s %d 不在目标区间 %d–%d（差 %d 字）"
                            % (label, val, lo, hi, gap))
    heads = [t for t in text[:cut]
             if re.match(r"^\d+(\.\d+)*\s\S", t) and len(t) < 40]
    print("标题 %d 个：%s" % (len(heads), " / ".join(h.split()[0] for h in heads)))

    print()
    if problems:
        print("发现 %d 项问题：" % len(problems))
        for p in problems:
            print("  ✗ %s" % p)
    else:
        print("✓ 未发现格式异常")
    return problems


def main():
    ap = argparse.ArgumentParser(description="docx 交付前自检")
    ap.add_argument("docx")
    ap.add_argument("--refs-title", default="参考文献")
    ap.add_argument("--target", help="目标字数区间，如 4500-5000；单值按 ±5%% 展开")
    ap.add_argument("--by", choices=["body", "word"], default="body",
                    help="--target 按哪个口径核：body=正文汉字（默认），word=Word 估算含表格")
    ap.add_argument("--strict", action="store_true", help="有问题时退出码为 1")
    args = ap.parse_args()

    target = None
    if args.target:
        nums = [int(x) for x in re.split(r"[-~,，\s]+", args.target) if x.strip()]
        if not nums:
            sys.exit("--target 格式不对，写成 4500-5000 或 5000")
        target = (nums[0], nums[1]) if len(nums) > 1 else (
            int(round(nums[0] * 0.95)), int(round(nums[0] * 1.05)))

    problems = inspect(args.docx, args.refs_title, target, args.by)
    sys.exit(1 if (problems and args.strict) else 0)


if __name__ == "__main__":
    main()
