# -*- coding: utf-8 -*-
"""A2 插图紧贴裁剪：图注锚定 + 内容块自下向上聚类 + 去白边

用法:
    python crop_figures.py <paper.pdf> <fig_dir>
        [--fig-regex REGEX] [--z 4] [--max-gap 22] [--top-margin 46] [--pad 3]
        [--override 3=105,46,500,175 --override 101=150,374,430,497]
        [--names names.json]

原理:
    图注正上方即该图。取 [上一图注底部, 本图注顶部] 为候选带，
    带内收集 图像矩形 + 面积阈以上的矢量绘制矩形，
    以贴近图注的块为种子向上聚类（纵向间隙 > max-gap 则截断），
    得紧贴框后按 z 倍渲染，再按灰度阈值去白边。

矢量图 / 附录小图常因绘制块检测不全而漏边 -> 用 --override 指定该图的宽松带，
交由去白边自动收紧。--names 提供 {图号: 输出文件名} 映射（图号用 int，附录图可用 101 起）。

输出: <fig_dir>/crop_report.txt
"""
import sys
import os
import re
import json
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import fitz
import numpy as np
from PIL import Image

DEF_FIG = r"^(?:Fig\.\s*(\d+[a-z]?|A\d+|S\d+)\.|Figure\s+\d+\.|图\s*\d+)"
TAB_RE = re.compile(r"^(?:Table\s*\d+\s*|表\s*\d+)")


def norm_id(s):
    """'A1' -> 101, 'S2' -> 202, '5' -> 5, '5b' -> 5 (主图)"""
    s = s.strip()
    if s[0] in "Aa" and s[1:].isdigit():
        return 100 + int(s[1:])
    if s[0] in "Ss" and s[1:].isdigit():
        return 200 + int(s[1:])
    m = re.match(r"^(\d+)", s)
    return int(m.group(1)) if m else 0


def content_rects(page):
    """页面内容块：图像矩形 + 面积阈以上的矢量绘制矩形。

    2026-09 演进：原宽高分别 >6 会漏掉矢量图的细长轴线
    （如 Fig.A1 竖轴 w=2.9 h=39、横轴 w=108 h=2.9），导致裁剪框缺轴/刻度；
    改面积阈值 w*h > 36 后细长轴线可入选。"""
    r = []
    for im in page.get_images(full=True):
        for q in page.get_image_rects(im[0]):
            if q.width > 2 and q.height > 2:
                r.append((q.x0, q.y0, q.x1, q.y1))
    for d in page.get_drawings():
        q = d["rect"]
        if q.width * q.height > 36:
            r.append((q.x0, q.y0, q.x1, q.y1))
    return r


def cluster(cand, deepest, max_gap):
    """自下向上聚类，返回并入选中的块"""
    used, cur = set(), []
    for i, r in enumerate(cand):
        if r[3] >= deepest - max_gap:
            cur.append(r)
            used.add(i)
    top = min(r[1] for r in cur)
    while True:
        grew = False
        for i, r in enumerate(cand):
            if i not in used and r[1] < top and r[3] >= top - max_gap:
                cur.append(r)
                used.add(i)
                grew = True
        if not grew:
            return cur
        top = min(r[1] for r in cur)


def trim_white(img, z):
    a = np.asarray(img.convert("L"))
    m = a < 248
    rws = np.where(m.any(axis=1))[0]
    cls = np.where(m.any(axis=0))[0]
    if len(rws) == 0 or len(cls) == 0:
        return None
    pad = int(2 * z)
    return img.crop((max(0, cls[0] - pad), max(0, rws[0] - pad),
                     min(img.width, cls[-1] + pad + 1), min(img.height, rws[-1] + pad + 1)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("outdir")
    ap.add_argument("--fig-regex", default=DEF_FIG)
    ap.add_argument("--z", type=float, default=4.0, help="渲染倍率，4≈288DPI")
    ap.add_argument("--max-gap", type=float, default=22.0)
    ap.add_argument("--top-margin", type=float, default=46.0)
    ap.add_argument("--pad", type=float, default=3.0)
    ap.add_argument("--override", action="append", default=[], help="图号=x0,y0,x1,y1")
    ap.add_argument("--names", default=None, help="JSON: {图号: 文件名}")
    a = ap.parse_args()

    os.makedirs(a.outdir, exist_ok=True)
    fig_re = re.compile(a.fig_regex)
    names = json.load(open(a.names, encoding="utf-8")) if a.names else {}
    ov = {}
    for item in a.override:
        k, v = item.split("=", 1)
        ov[int(k)] = tuple(float(x) for x in v.split(","))

    doc = fitz.open(a.pdf)
    rep, saved, manual = [], 0, 0
    for pno in range(doc.page_count):
        page = doc[pno]
        caps = []
        for b in page.get_text("blocks"):
            t = b[4].strip().replace("\n", " ")
            if TAB_RE.match(t):
                continue
            m = fig_re.match(t)
            if m:
                caps.append((round(b[1], 1), round(b[3], 1), round(b[0], 1),
                             round(b[2], 1), norm_id(m.group(1)), t[:45]))
        if not caps:
            continue
        caps.sort(key=lambda c: c[0])
        rects = content_rects(page)
        if not rects:
            continue

        for idx, (cy0, cy1, cx0, cx1, num, head) in enumerate(caps):
            if num == 0:
                continue
            band_top = caps[idx - 1][1] + 2 if idx > 0 else a.top_margin
            band_bot = cy0 - 2
            if num in ov:
                x0, y0, x1, y1 = ov[num]
                y0, y1 = max(band_top, y0), min(band_bot, y1)
            else:
                cand = [r for r in rects
                        if r[1] >= band_top - 2 and r[3] <= band_bot + 2
                        and not (r[2] < cx0 - 30 or r[0] > cx1 + 30)]
                if not cand:
                    rep.append("MANUAL p%d Fig.%s [%s] 带内无内容块 band[%.0f,%.0f]"
                               % (pno + 1, head[:14], head, band_top, band_bot))
                    manual += 1
                    continue
                cur = cluster(cand, max(r[3] for r in cand), a.max_gap)
                x0 = min(r[0] for r in cur) - a.pad
                x1 = max(r[2] for r in cur) + a.pad
                y0 = max(band_top, min(r[1] for r in cur) - a.pad)
                y1 = min(band_bot, max(r[3] for r in cur) + a.pad)
                x0 = max(0, x0)
                x1 = min(page.rect.width, x1)

            pix = page.get_pixmap(matrix=fitz.Matrix(a.z, a.z), clip=fitz.Rect(x0, y0, x1, y1))
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            out = trim_white(img, a.z)
            if out is None:
                rep.append("EMPTY p%d Fig.%s clip=(%.0f,%.0f,%.0f,%.0f)" % (pno + 1, head[:14], x0, y0, x1, y1))
                continue
            fn = names.get(str(num), "fig%03d" % num) + ".png"
            out.save(os.path.join(a.outdir, fn), dpi=(int(a.z * 72), int(a.z * 72)))
            saved += 1
            rep.append("OK p%-3d Fig.%-4s %-34s %4dx%-4dpx clip=(%.0f,%.0f,%.0f,%.0f)"
                       % (pno + 1, num, fn, out.size[0], out.size[1], x0, y0, x1, y1))

    rep.append("SAVED %d   MANUAL %d   EMPTY %d" % (saved, manual, sum(1 for r in rep if r.startswith("EMPTY"))))
    open(os.path.join(a.outdir, "crop_report.txt"), "w", encoding="utf-8").write("\n".join(rep) + "\n")
    print("SAVED %d  MANUAL %d  -> %s" % (saved, manual, os.path.join(a.outdir, "crop_report.txt")))


if __name__ == "__main__":
    main()
