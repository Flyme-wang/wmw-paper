# -*- coding: utf-8 -*-
"""A3 前置：为每个编号公式生成"宽松候选窗口"，供人工核对后再交给 crop_equations.py

用法:
    python eq_windows.py <paper.pdf> --out eq_windows.json [--prose-font CharisSIL ...]
                         [--out-report eq_windows_report.txt]

窗口确定规则（与版面栏数无关，不依赖分栏检测 —— 跨栏图会让分栏检测失稳）:
    y 范围 = 编号所在栏内，上方最近正文行的中心 到 下方最近正文行的中心（宽松窗口，
             靠 crop_equations 的正文剔除闭环收窄；下限 46pt 避开页眉带）
    x 范围 = 编号在左半区 -> 左栏 [左缘, 0.5W]
             编号在右半区 -> 先查该页是否单栏（跨书脊正文行 >=5）或窗口 y 带内
                             有跨书脊内容（非正文 span / 图像）:
                             是 -> 全幅 [左缘, 右缘]（附录单栏页、跨栏公式）
                             否 -> 右栏 [0.5W, 右缘]

视觉行聚类为 line 级（y 聚类 + x 重叠检查）：左右栏相邻行 y 差可 <2.5pt，
word 级聚类会把它们合并导致判侧错误；line 级不会。

结果写入 UTF-8 文件。**不要用 PowerShell `>` 重定向**（默认 UTF-16LE，json.load 会失败）。
输出条目: {"page": 1-based, "num": "2.1", "y0":.., "y1":.., "x0":.., "x1":.., "wide":bool}
多公式共号、跨页公式、相邻公式对（无正文行分隔，如 3.10/3.11）仍需人工调整 y0/y1。
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
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fitz

from crop_equations import prose_tokens, detect_prose_fonts
from probe_pdf import eq_tags

DEF_EQ = r"^\((\d{1,2}\.[A-Za-z]?\d{1,2}|[A-D]\.\d{1,2}|[A-C]\d)\)$"
YBAND = 120.0          # 向上/下搜索正文边界的最大距离
YMIN = 46.0            # 页面内容区 y 下限（页眉带通常 y<46，避免页眉进入窗口）
SINGLE_TH = 5          # 跨书脊正文行数 >= 此值 -> 单栏页


def word_rows(page, ytol=2.5, xtol=6.0):
    """把 line 聚成视觉行（y 聚类 + x 重叠检查，左右栏相邻行不会合并）；
    返回 [(y_center, x0, x1, is_prose, text), ...]

    2026-09 演进：word 级聚类会把 y 差 <2.5pt 的左右栏相邻行合并（如
    p6 y187.5 右栏行 vs y189.9 左栏行），合并后行中心判侧翻转，导致
    3.12/3.13 共享窗口。line 级 + x 重叠检查（xtol=6）避免跨栏合并。"""
    H = page.rect.height
    items = []
    for blk in page.get_text("dict")["blocks"]:
        if blk["type"] != 0:
            continue
        for ln in blk["lines"]:
            x0, y0, x1, y1 = ln["bbox"]
            cy = (y0 + y1) / 2
            if cy < YMIN or cy > H - YMIN:
                continue
            txt = " ".join(s["text"] for s in ln["spans"]).strip()
            if not txt:
                continue
            items.append((cy, x0, x1, txt))
    items.sort(key=lambda t: (t[0], t[1]))
    rows, cur = [], []
    for it in items:
        cy, x0, x1, txt = it
        if cur:
            cur_cy = sum(q[0] for q in cur) / len(cur)
            cur_x0 = min(q[1] for q in cur)
            cur_x1 = max(q[2] for q in cur)
            x_ok = not (x0 > cur_x1 + xtol or x1 < cur_x0 - xtol)
            if cy - cur_cy > ytol or not x_ok:
                rows.append(cur)
                cur = []
        cur.append(it)
    if cur:
        rows.append(cur)
    out = []
    for r in rows:
        cy = sum(q[0] for q in r) / len(r)
        x0 = min(q[1] for q in r)
        x1 = max(q[2] for q in r)
        txt = " ".join(q[3] for q in r)
        out.append((cy, x0, x1, prose_tokens(txt), txt))
    return out


def cross_band(page, y0, y1):
    """窗口 y 带内是否存在横跨书脊（x0<mid<x1）的非正文 span 或图像矩形
    （跨栏公式本体为碎片 span/小图像，正文行不跨书脊，图注最多 1-3 行且为正文）"""
    mid = page.rect.width / 2
    for blk in page.get_text("dict")["blocks"]:
        if blk["type"] != 0:
            continue
        for ln in blk["lines"]:
            for sp in ln["spans"]:
                if not sp["text"].strip():
                    continue
                sx0, sy0, sx1, sy1 = sp["bbox"]
                if sy1 > y0 and sy0 < y1 and sx0 < mid < sx1 \
                        and not prose_tokens(sp["text"]):
                    return True
    for im in page.get_images(full=True):
        for r in page.get_image_rects(im[0]):
            if r.y1 > y0 and r.y0 < y1 and r.x0 < mid < r.x1:
                return True
    return False


def is_single_column(page):
    """单栏页判据：跨书脊正文 line 数 >= SINGLE_TH。
    双栏页正文行不跨书脊（只有图注跨栏，1-3 条）；附录推导页（如本样本 p18-22）
    每行从页左缘排到页右缘，跨书脊行数 >=5，此时右端编号的公式本体必在左半区，
    窗口须全幅（x 到左缘），否则裁剪 EMPTY。"""
    mid = page.rect.width / 2
    H = page.rect.height
    cross = 0
    for blk in page.get_text("dict")["blocks"]:
        if blk["type"] != 0:
            continue
        for ln in blk["lines"]:
            x0, y0, x1, y1 = ln["bbox"]
            cy = (y0 + y1) / 2
            if cy < YMIN or cy > H - YMIN:
                continue
            txt = " ".join(s["text"] for s in ln["spans"]).strip()
            if not txt or not prose_tokens(txt):
                continue
            if x0 < mid - 20 and x1 > mid + 20:
                cross += 1
                if cross >= SINGLE_TH:
                    return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--out", default="eq_windows.json")
    ap.add_argument("--out-report", default=None)
    ap.add_argument("--prose-font", action="append", default=[])
    ap.add_argument("--eq-regex", default=DEF_EQ)
    ap.add_argument("--fallback", type=float, default=35.0)
    a = ap.parse_args()

    doc = fitz.open(a.pdf)
    pf = set(a.prose_font) or detect_prose_fonts(doc)
    eq_re = re.compile(a.eq_regex)
    tags = []
    for pno in range(doc.page_count):
        for num, tx0, tyc, tx1 in eq_tags(doc[pno], eq_re):
            tags.append((pno + 1, num, tx0, tyc, tx1))
    if not tags:
        print("未找到任何编号；请检查 --eq-regex")
        return

    rows, rep, seen = [], [], set()
    for p1, num, tx0, tyc, tx1 in tags:
        if (p1, num) in seen:
            rep.append("DUP  p%-3d (%s) 同页重复出现，可能是行末交叉引用 -> 人工确认后删除" % (p1, num))
            continue
        seen.add((p1, num))
        page = doc[p1 - 1]
        W = page.rect.width
        H = page.rect.height
        rl = word_rows(page)
        left = tx1 < 0.55 * W
        mid = 0.5 * W
        mL = min(r[1] for r in rl)
        mR = max(r[2] for r in rl)

        def side_of(r):
            # 以行右缘判侧：正文行右缘 <= 栏右缘（左栏 <mid，右栏 >mid）
            return "R" if r[2] > mid else "L"

        me = "L" if left else "R"
        up = dn = None
        for cy, x0, x1, is_p, txt in rl:
            if not is_p or side_of((cy, x0, x1, is_p, txt)) != me:
                continue
            if cy < tyc and (up is None or cy > up[0]):
                up = (cy, x1, txt)
            elif cy > tyc and (dn is None or cy < dn[0]):
                dn = (cy, x0, txt)
        y0 = (up[0] + 6) if up else tyc - a.fallback
        y1 = (dn[0] - 6) if dn else tyc + a.fallback
        y0 = max(YMIN, y0 - 6)
        y1 = min(H - YMIN, y1 + 6)

        # x 范围：右半区编号需判断是否全幅（单栏页 / 跨栏公式）
        wide = False
        if not left:
            wide = is_single_column(page) or cross_band(page, y0, y1)
        if wide or left:
            x0, x1 = (mL, mid + 6) if left else (mL, mR + 4)
        else:
            x0, x1 = mid - 6, mR + 4

        rows.append({"page": p1, "num": num, "y0": round(y0, 1), "y1": round(y1, 1),
                     "x0": round(x0, 1), "x1": round(x1, 1), "wide": wide})
        rep.append("WIN  p%-3d (%-6s) y(%6.1f,%6.1f) x(%5.1f,%5.1f) %-5s %s" % (
            p1, num, y0, y1, x0, x1, "FULL" if wide else ("LEFT" if left else "RIGHT"),
            ("cut above:%s" % (up[2][:26] if up else "-")) +
            (" | below:%s" % (dn[2][:26] if dn else "-"))))

    rows.sort(key=lambda r: (r["page"], r["y0"]))
    json.dump(rows, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if a.out_report:
        open(a.out_report, "w", encoding="utf-8").write("\n".join(rep) + "\n")
    print("windows=%d  (full-width=%d) -> %s" % (len(rows), sum(1 for r in rows if r["wide"]), a.out))
    print("prose_fonts=%s" % sorted(pf))
    dups = [r for r in rep if r.startswith("DUP")]
    if dups:
        print("DUP %d 条（行末交叉引用误判），见报告" % len(dups))
    print("下一步：核对 %s，再跑 crop_equations.py" % (a.out_report or a.out))


if __name__ == "__main__":
    main()
