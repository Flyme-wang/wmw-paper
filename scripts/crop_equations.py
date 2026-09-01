# -*- coding: utf-8 -*-
"""A3 编号公式裁剪（含正文残留剔除闭环）

用法:
    python crop_equations.py <paper.pdf> <windows.json> <eq_dir>
        [--prose-font CharisSIL ...] [--z 5] [--dpi 360]

windows.json 格式（page 为 1-based）:
    [{"page": 2, "num": "2.1", "y0": 274, "y1": 302, "x0": 36, "x1": 293}, ...]

五步闭环（顺序不可省）:
    1) 定位编号 word，把 "(2.1)" 从裁剪框右侧排除
    2) 以编号 y 中心为界，把窗口切成上/下两半
    3) 用正体字体 + prose_tokens() 判据识别正文 span，据此上下收缩窗口
    4) 按 z 倍渲染，再按暗像素去白边
    5) 报告 TRIM / SKIPTRIM / EMPTY 与最终 bbox 命中的文字，逐条人工过

--prose-font 可重复给多个（不同刊不同：Elsevier 多 CharisSIL，IEEE 多 TimesBT）。
不传则自动取该 PDF 中字符数最多的"非斜体非粗体"字体。

输出: <eq_dir>/eq_report.txt + <eq_dir>/eq_words.txt
"""
import sys
import os
import re
import json
import argparse
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8")   # 幂等；不可重复包 TextIOWrapper（会关掉底层 buffer）
except Exception:
    pass
import fitz
import numpy as np
from PIL import Image

ENG = re.compile(r"^[A-Za-z][A-Za-z'\u2019\-]{1,}[.,;:!?\]\)]*$")
YRT = re.compile(r"^(19|20)\d\d[):.,;]*$")
FUNC = set("""the and of to in into as is are be been by on or for with that this these those
can could we will were was from at an all any each both when where which then than has have had
not but their its it they them some more most such only also et al i.e. e.g. etc via while
whereas between below above after before because used use gives give here there these
eq eqs eq. fig figs sec ref refs""".split())
MATHNAMES = {"sigma", "delta", "alpha", "beta", "lambda", "gamma", "times", "prime",
             "approx", "infty", "partial", "nabla", "therefore", "because", "sign"}


def prose_tokens(text):
    """一个正体 span 是否构成正文：>=2 个功能词 / >=1 个长词 / 任一四位数年份"""
    n_func = n_long = n_year = 0
    for t in text.split():
        core = t.strip(".,;:!?)][(")
        if YRT.match(t) or YRT.match(core):
            n_year += 1
        elif core.lower() in FUNC:
            n_func += 1
        elif ENG.match(t) and len(core) >= 5 and core.lower() not in MATHNAMES:
            n_long += 1
    return (n_func >= 2) or n_long >= 1 or n_year >= 1


def prose_spans(page, y0, y1, x0, x1, prose_fonts):
    """窗口内正文性行的 (y0, y1, text) —— line 级判据，不依赖字体。

    2026-09 演进：原 span 级 + 字体白名单（prose_fonts）会漏掉 Bold 标题
    （如 "4. Model verification"）与页眉行，导致裁剪图混入残留文字。
    改为 line 级：窗口内完整 line 且 prose_tokens() 为真即视为正文。
    数学公式的碎片 line 是单符号短文本，prose_tokens 判 False，不会误剔。
    """
    out = []
    for blk in page.get_text("dict")["blocks"]:
        if blk["type"] != 0:
            continue
        for ln in blk["lines"]:
            sx0, sy0, sx1, sy1 = ln["bbox"]
            txt = " ".join(s["text"] for s in ln["spans"]).strip()
            if not txt or not prose_tokens(txt):
                continue
            if sx1 > x0 and sx0 < x1 and sy1 > y0 + 1.0 and sy0 < y1 - 1.0:
                out.append((sy0, sy1, txt))
    return out


def bands_of(page, x0, y0, x1, y1, z):
    """栅格白通道分带 -> [(pdf_y0, pdf_y1), ...]，用于确认窗口内实际内容行"""
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z), clip=fitz.Rect(x0, y0, x1, y1))
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3]
    dark = (a.mean(axis=2) < 245).any(axis=1)
    out, s = [], None
    for i, d in enumerate(dark):
        if d and s is None:
            s = i
        elif not d and s is not None:
            if i - s >= 2:
                out.append((y0 + s / z - 0.4, y0 + i / z + 0.4))
            s = None
    if s is not None and len(dark) - s >= 2:
        out.append((y0 + s / z - 0.4, y0 + len(dark) / z + 0.4))
    return out


def detect_prose_fonts(doc):
    c = Counter()
    for pg in doc:
        for b in pg.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            for ln in b["lines"]:
                for s in ln["spans"]:
                    f = s["font"]
                    if "Italic" in f or "Bold" in f:
                        continue
                    if prose_tokens(s["text"]):
                        c[f] += len(s["text"])
    return {k for k, _ in c.most_common(3)}


def find_tags(doc, eq_re):
    """返回 [(page1, num, x0, y_center, x1, y0, y1), ...]"""
    out = []
    for pno in range(doc.page_count):
        for w in doc[pno].get_text("words"):
            m = eq_re.match(w[4].strip())
            if m:
                out.append((pno + 1, m.group(1), w[0], (w[1] + w[3]) / 2, w[2], w[1], w[3]))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("windows")
    ap.add_argument("outdir")
    ap.add_argument("--prose-font", action="append", default=[])
    ap.add_argument("--z", type=float, default=5.0)
    ap.add_argument("--dpi", type=int, default=360)
    ap.add_argument("--eq-regex", default=r"^\((\d{1,2}\.[A-Za-z]?\d{1,2}|[A-D]\.\d{1,2}|[A-C]\d)\)$")
    a = ap.parse_args()

    os.makedirs(a.outdir, exist_ok=True)
    doc = fitz.open(a.pdf)
    pf = set(a.prose_font) or detect_prose_fonts(doc)
    wins = json.load(open(a.windows, encoding="utf-8"))
    tag_re = lambda num: re.compile(r"^\(%s\)$" % re.escape(num))

    rep, words_out = ["prose_fonts=%s" % sorted(pf)], []
    for w in wins:
        pno, num = w["page"] - 1, w["num"]
        x0, y0, x1, y1 = float(w["x0"]), float(w["y0"]), float(w["x1"]), float(w["y1"])
        page = doc[pno]
        # 1) 排除编号：窗口内可能有多个同号 word（真编号 + 正文行末交叉引用），
        #    取 y 中心最接近窗口中心的那个（交叉引用在正文行内，y 偏移大）
        tag_y = None
        cands = []
        for t in page.get_text("words"):
            if tag_re(num).match(t[4].strip()) and t[1] >= y0 - 8 and t[3] <= y1 + 8 \
                    and t[0] >= x0 and t[2] <= x1:
                cands.append(t)
        if cands:
            mid_y = (y0 + y1) / 2.0
            t = min(cands, key=lambda t2: abs((t2[1] + t2[3]) / 2.0 - mid_y))
            x1 = min(x1, t[0] - 5)
            tag_y = (t[1], t[3])
        # 2-3) 正文残留剔除
        pr = prose_spans(page, y0, y1, x0, x1, pf)
        ref = (tag_y[0] + tag_y[1]) / 2.0 if tag_y else (y0 + y1) / 2.0
        above = [s for s in pr if (s[0] + s[1]) / 2 < ref]
        below = [s for s in pr if (s[0] + s[1]) / 2 >= ref]
        ny0, ny1 = y0, y1
        if above:
            ny0 = max([s[1] - 0.8 for s in above] + [y0])   # 正文行下沿上方 0.8pt（descent 空隙）
        if below:
            ny1 = min([s[0] - 0.8 for s in below] + [y1])   # 正文行上沿上方 0.8pt（ascender 空隙）
        # 注：below 曾写 s[0]+0.8，会把正文行顶部 0.8pt 留在窗口内，正文行横贯全宽时
        # 去白边会把裁剪图横向撑满（如 C.4）；-0.8 对称于 above，正文文字完全排除。
        if ny1 - ny0 < 8:
            ny0, ny1 = y0, y1
            rep.append("SKIPTRIM %-6s p%d would over-cut" % (num, pno + 1))
        elif (ny0, ny1) != (y0, y1):
            rep.append("TRIM %-6s p%-2d y(%.1f,%.1f)->(%.1f,%.1f) drop %s" % (
                num, pno + 1, y0, y1, ny0, ny1,
                " / ".join(s[2] for s in above + below)[:110]))
        # 4) 渲染 + 去白边
        pix = page.get_pixmap(matrix=fitz.Matrix(a.z, a.z), clip=fitz.Rect(x0, ny0, x1, ny1))
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        m = np.asarray(img.convert("L")) < 245
        rws = np.where(m.any(axis=1))[0]
        cls = np.where(m.any(axis=0))[0]
        if len(rws) == 0 or len(cls) == 0:
            rep.append("EMPTY %-6s p%d win(%.0f,%.1f,%.0f,%.1f)" % (num, pno + 1, x0, ny0, x1, ny1))
            continue
        pad = int(2 * a.z)
        box = (max(0, cls[0] - pad), max(0, rws[0] - pad),
               min(img.width, cls[-1] + pad + 1), min(img.height, rws[-1] + pad + 1))
        img = img.crop(box)
        rx0, ry0 = x0 + box[0] / a.z, ny0 + box[1] / a.z
        rx1, ry1 = x0 + box[2] / a.z, ny0 + box[3] / a.z
        hit = [t[4] for t in page.get_text("words")
               if t[0] < rx1 and t[2] > rx0 and t[1] < ry1 and t[3] > ry0]
        words_out.append("[%s] %s" % (num, " | ".join(hit) if hit else "(no text)"))
        fn = "eq" + num.replace(".", "_") + ".png"
        img.save(os.path.join(a.outdir, fn), dpi=(a.dpi, a.dpi))
        rep.append("eq%-6s p%-2d %-12s %4dx%-4dpx crop=(%.0f,%.1f,%.0f,%.1f)" % (
            num, pno + 1, fn, img.size[0], img.size[1], rx0, ry0, rx1, ry1))

    open(os.path.join(a.outdir, "eq_report.txt"), "w", encoding="utf-8").write("\n".join(rep) + "\n")
    open(os.path.join(a.outdir, "eq_words.txt"), "w", encoding="utf-8").write("\n".join(words_out) + "\n")
    print("SAVED %d  REPORT %d  EMPTY %d  prose_fonts=%s"
          % (len(words_out), len(rep), sum(1 for r in rep if r.startswith("EMPTY")), sorted(pf)))


if __name__ == "__main__":
    main()
