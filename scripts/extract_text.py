# -*- coding: utf-8 -*-
"""A4/B2 按分栏阅读顺序导出英文全文（供翻译与要点提取）

用法:
    python extract_text.py <paper.pdf> <fulltext.txt> [--pages 1-12] [--cols auto|1|2]
                           [--cut 297.6]

顺序：每页先左栏自上而下，再右栏自上而下；行内按 x 排序，行距 > 3pt 断行，
段落间按块插入空行。页眉页脚（y < 40 或 y > 页高-40）自动丢弃。

输出含 "=== PAGE n (col k) ===" 分节标记，便于翻译时定位图/公式归属页。
"""
import sys
import os
import re
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fitz


def parse_pages(spec, total):
    if not spec:
        return list(range(total))
    out = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            out += list(range(int(a) - 1, min(int(b), total)))
        else:
            out.append(int(part) - 1)
    return out


def lines_of(page, x0, x1, line_gap=15.0):
    """返回该栏内按 y 排序的行文本列表"""
    rows = []
    for blk in page.get_text("dict")["blocks"]:
        if blk["type"] != 0:
            continue
        for ln in blk["lines"]:
            bx0, by0, bx1, by1 = ln["bbox"]
            if bx1 <= x0 + 1 or bx0 >= x1 - 1:
                continue
            cy = (by0 + by1) / 2
            if cy < 40 or cy > page.rect.height - 40:
                continue
            txt = "".join(s["text"] for s in ln["spans"]).strip()
            if txt:
                rows.append((round(cy, 1), bx0, txt))
    rows.sort(key=lambda r: (r[0], r[1]))
    # 合并成段：相邻行 y 中心差 > line_gap 视为换段
    paras, buf, prev = [], [], None
    for cy, bx, txt in rows:
        if prev is not None and cy - prev > line_gap:
            paras.append(" ".join(buf))
            buf = []
        buf.append(txt)
        prev = cy
    if buf:
        paras.append(" ".join(buf))
    return [re.sub(r"\s+", " ", p).strip() for p in paras if p.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out")
    ap.add_argument("--pages", default=None)
    ap.add_argument("--cols", default="auto", choices=["auto", "1", "2"])
    ap.add_argument("--cut", type=float, default=None)
    ap.add_argument("--line-gap", type=float, default=15.0,
                    help="相邻行 y 中心超过该值判为换段（单栏密排可调至 13）")
    a = ap.parse_args()

    doc = fitz.open(a.pdf)
    L = ["# %s  (%d pages)" % (a.pdf, doc.page_count)]
    n = 0
    for pno in parse_pages(a.pages, doc.page_count):
        page = doc[pno]
        W = page.rect.width
        if a.cols == "1":
            bounds = [(0, W)]
        elif a.cols == "2":
            cut = a.cut or W / 2
            bounds = [(0, cut), (cut, W)]
        else:
            from probe_pdf import column_bounds
            bounds = column_bounds(page)
        L.append("\n=== PAGE %d (%d-col) ===" % (pno + 1, len(bounds)))
        for k, (x0, x1) in enumerate(bounds):
            L.append("--- col %d ---" % (k + 1))
            for p in lines_of(page, x0, x1, a.line_gap):
                L.append(p)
                n += 1
    open(a.out, "w", encoding="utf-8").write("\n".join(L) + "\n")
    print("paragraphs=%d -> %s" % (n, a.out))


if __name__ == "__main__":
    main()
