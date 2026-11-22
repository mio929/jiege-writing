#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从 PMC 开放获取论文里取原图与图题，供课程论文引用配图。

用法
----
    python scripts/fetch_pmc_figs.py PMC7460318 --list
    python scripts/fetch_pmc_figs.py PMC7460318 -o figs/
    python scripts/fetch_pmc_figs.py PMC7460318 -o figs/ --only g004,g006

只处理 PMC（PubMed Central）收录的开放获取文献。订阅期刊的图取不到，也不要
去破解；换一篇 OA 文献，或者改用自己的实测数据。

两步走的原因：eutils 的 XML 只给文件名（如 ``ijerph-17-05680-g004.jpg``），
真实地址挂在文章页里的 ``cdn.ncbi.nlm.nih.gov/pmc/blobs/...`` 上，必须再抓
一次页面才能拿到可下载的 URL。直接拼 ``/bin/文件名`` 现在会 403。
"""
import argparse
import os
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
HEADERS = {"User-Agent": UA,
           "Accept": "text/html,application/xhtml+xml,image/avif,image/webp,image/*,*/*;q=0.8"}
XLINK = "{http://www.w3.org/1999/xlink}href"


def get(url, timeout=90):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def pmcid_of(arg):
    arg = arg.strip()
    if arg.upper().startswith("PMC"):
        return arg.upper()
    return "PMC" + arg


def read_meta(pmcid):
    """返回 (标题, [(图号, 文件名, 图题)])。"""
    xml = get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
              "?db=pmc&id=%s&retmode=xml" % pmcid[3:])
    root = ET.fromstring(xml)
    title = "".join(root.find(".//article-title").itertext()) if root.find(".//article-title") is not None else ""
    figs = []
    for fig in root.iter("fig"):
        label = fig.find("label")
        cap = fig.find("caption")
        href = None
        for g in fig.iter("graphic"):
            href = g.get(XLINK)
            break
        text = re.sub(r"\s+", " ", "".join(cap.itertext())).strip() if cap is not None else ""
        figs.append((label.text.strip() if label is not None else "?",
                     href or "", text[:300]))
    return title, figs


def blob_urls(pmcid):
    html = get("https://pmc.ncbi.nlm.nih.gov/articles/%s/" % pmcid).decode("utf-8", "ignore")
    urls = re.findall(r"https://cdn\.ncbi\.nlm\.nih\.gov/pmc/blobs/[^\s\"']+?\.(?:jpg|png|gif)", html)
    seen, out = set(), []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def tag_of(url):
    name = os.path.basename(urllib.parse.urlparse(url).path)
    m = re.search(r"(g\d{3}|fig\d+|[a-z]+-g\d{3})", name)
    return m.group(1) if m else name


def main():
    ap = argparse.ArgumentParser(description="下载 PMC 论文配图")
    ap.add_argument("pmcid")
    ap.add_argument("-o", "--out", default="figs")
    ap.add_argument("--list", action="store_true", help="只打印图号与图题")
    ap.add_argument("--only", default="", help="只下指定图号，逗号分隔，如 g004,g006")
    args = ap.parse_args()

    pmcid = pmcid_of(args.pmcid)
    title, figs = read_meta(pmcid)
    print("%s  %s" % (pmcid, title[:90]))
    for label, href, text in figs:
        print("  %-8s %-34s %s" % (label, os.path.basename(href), text[:110]))
    if args.list:
        return

    only = {x.strip().lower() for x in args.only.split(",") if x.strip()}
    os.makedirs(args.out, exist_ok=True)
    got = 0
    for url in blob_urls(pmcid):
        tag = tag_of(url)
        if only and tag.lower() not in only:
            continue
        dest = os.path.join(args.out, "%s_%s" % (pmcid, urllib.parse.unquote(
            os.path.basename(urllib.parse.urlparse(url).path))))
        data = get(url)
        if len(data) < 6000:
            print("  跳过（体积过小，可能是占位图）%s" % dest)
            continue
        open(dest, "wb").write(data)
        print("  下载 %-46s %7d 字节" % (os.path.basename(dest), len(data)))
        got += 1
    if not got:
        print("没下到图。检查 PMC 编号，或该文不在 OA 子集里。", file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
