# -*- coding: utf-8 -*-
"""A1 结构探测：页面尺寸 / 分栏边界 / span 字体统计 / 图注表注坐标 / 编号公式坐标 / 图像块并集

用法:
    python probe_pdf.py <paper.pdf> [probe_out.txt] [--fig-regex REGEX]

结果写入 UTF-8 文件（不要经 PowerShell 管道，中文会乱码）。
"""
import sys
import io
import re
import argparse
from collections import Counter

try:
    sys.stdout.reconfigure(encoding="utf-8")   # 幂等；PowerShell 管道默认 GBK
except Exception:
    pass
import fitz

DEF_FIG = r"^(?:Fig\.\s*(\d+[a-z]?|A\d+|S\d+)\.|Figure\s+\d+\.|图\s*\d+)"
DEF_TAB = r"^(?:Table\s*(\d+)\s*|表\s*\d+)"
DEF_EQ = r"^\((\d{1,2}\.[A-Za-z]?\d{1,2}|[A-D]\.\d{1,2}|[A-C]\d)\)$"


def text_lines(page):
    """正文区内的行 bbox [(x0,x1), ...]（已除页眉页脚）"""
    H = page.rect.height
    out = []
    for blk in page.get_text("dict")["blocks"]:
        if blk["type"] != 0:
            continue
        for ln in blk["lines"]:
            x0, y0, x1, y1 = ln["bbox"]
            if 40 < (y0 + y1) / 2 < H - 40 and x1 - x0 > 1:
                out.append((x0, x1))
    return out


def column_bounds(page):
    """双栏判定：取中部 [0.38W, 0.62W] 内“无任何文字行跨越”的最宽竖直空隙带

    两种常见错法不可靠：
      * word 中心直方图 —— 单栏页有大图/大公式时中部同样会空；
      * 固定 0.42/0.50 等比例试探 —— 同一文档不同页会给出不同切点，
        直接搞乱 column_edges 的右边界，导致公式编号漏检。
    文字行不会跨栏，故“零跨越带”是稳定信号。
    """
    W = page.rect.width
    lfs = text_lines(page)
    if len(lfs) < 12:
        return [(0, W)]
    cov = [0] * (int(W) + 2)
    for x0, x1 in lfs:
        for i in range(max(0, int(x0)), min(int(W) + 1, int(x1))):
            cov[i] += 1
    lo, hi = int(0.38 * W), int(0.62 * W)
    best_w, best_c, run = 0, None, 0
    for i in range(lo, hi + 1):
        if cov[i] == 0:
            run += 1
        else:
            if run >= 8 and run > best_w:
                best_w, best_c = run, i - run // 2
            run = 0
    if run >= 8 and run > best_w:
        best_c = hi - run // 2
    if best_c is None:
        return [(0, W)]
    g = float(best_c)
    n = len(lfs)
    left = sum(1 for x0, x1 in lfs if x1 <= g)
    right = sum(1 for x0, x1 in lfs if x0 >= g)
    if left + right < 0.85 * n or min(left, right) < 0.15 * n:
        return [(0, W)]
    return [(0, g), (g, W)]


def column_edges(page):
    """每栏文字的实测左右边界 [(x0,x1), ...]（已除页眉页脚）"""
    cols = column_bounds(page)
    H = page.rect.height
    ws = [w for w in page.get_text("words") if 40 < (w[1] + w[3]) / 2 < H - 40]
    out = []
    for x0, x1 in cols:
        sub = [w for w in ws if x0 <= (w[0] + w[2]) / 2 <= x1]
        out.append((min(w[0] for w in sub), max(w[2] for w in sub)) if sub else (x0, x1))
    return out


def right_edges(page, tol=10.0, min_count=3, min_frac=0.4):
    """本页列右缘候选：word x1 聚类（差 <=tol 并簇），取簇内众数

    不能用簇最大值：个别超长 word（如 "(b)" 跨栏图注、数学式残片）
    会把簇拉到 300+，真实右缘 288.6 反而被漏掉。
    正文行右对齐形成的右缘 word 数量最多，取众数最稳；
    仅当簇成员 >=min_count 且众数在右半区才视为列右缘。
    注意：附录推导页行尾参差（数学式行尾散布多个 x）时候选偏多
    （实测 2025-LiSihai-CG p22 有 26 个候选），故本函数在 eq_tags 中
    只作"贴缘"弱约束（容差 8pt），主过滤由 eq_tags 的
    "段内 word 总数==1"判据承担（见 eq_tags docstring）。
    """
    W = page.rect.width
    xs = sorted(w[2] for w in page.get_text("words"))
    clus = []
    for x in xs:
        if clus and x - clus[-1][0] <= tol:
            clus[-1].append(x)
        else:
            clus.append([x])
    out = []
    for c in clus:
        if len(c) < min_count:
            continue
        mode = Counter(round(x, 1) for x in c).most_common(1)[0][0]
        if mode >= min_frac * W:
            out.append(mode)
    return out


def eq_tags(page, eq_re, gap=15.0, min_frac=0.42, ytol=2.5):
    """真实公式编号 = gap 切分段中唯一 word（段内 word 总数==1），且贴列右缘

    为何不用分栏边界/栏右边：跨栏大图或全幅图注会横跨书脊，
    导致 column_bounds 在正文页上判为单栏，左栏编号全部漏检。
    按视觉行切段则与版面栏数无关。

    判据演变（v1 -> v2，2026-09）：
    v1 曾要求"该视觉行右半区内只有这一个 word"，目的是过滤行末交叉引用
    "…Eq. (2.2)"。但正文行只要其余 word 的 x1 全部 <0.42W，编号仍是
    右半区唯一 word -> 误检。实测 2025-LiSihai-CG p22 附录推导行
    "…in Eq. (3.15)"（与 Eq. 空隙约 2.4pt < gap）即因此被当成编号。
    真编号行的编号与公式本体（图像/文本片段）水平空隙 >150pt，必被
    gap=15 切成独立段；行末引用与 "Eq." 空隙 <5pt 必同段。
    故 v2 收紧为"段内 word 总数 == 1"：编号必须是其段唯一元素。
    此外编号 x1 须贴近本页某一列右缘（weak 约束，见 right_edges）。
    """
    W, H = page.rect.width, page.rect.height
    ws = sorted([w for w in page.get_text("words") if 40 < (w[1] + w[3]) / 2 < H - 40],
                key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    groups, cur, cy = [], [], None
    for w in ws:
        y = (w[1] + w[3]) / 2
        if cy is None or abs(y - cy) <= ytol:
            cur.append(w)
            cy = y if cy is None else (cy * (len(cur) - 1) + y) / len(cur)
        else:
            groups.append(cur)
            cur, cy = [w], y
    if cur:
        groups.append(cur)

    out = []
    for g in groups:
        g.sort(key=lambda w: w[0])
        segs, seg = [], [g[0]]
        for w in g[1:]:
            if w[0] - seg[-1][2] > gap:
                segs.append(seg)
                seg = []
            seg.append(w)
        segs.append(seg)
        for seg in segs:
            last = seg[-1]
            m = eq_re.match(last[4].strip())
            if m and last[2] >= min_frac * W:
                # v2 判据：编号必须是其段唯一 word。
                # 真编号与公式本体空隙 >150pt 必单独成段；
                # 行末引用 "Eq. (3.15)" 与 Eq. 同段（段长>=2）一律过滤。
                if len(seg) != 1:
                    continue
                # 编号 x1 须贴近本页某一列右缘（弱约束；附录推导页候选偏多
                # 也不影响，主过滤已由 len(seg)==1 完成）
                edges = right_edges(page)
                if edges and not any(abs(last[2] - e) <= 8 for e in edges):
                    continue
                out.append((m.group(1), last[0], (last[1] + last[3]) / 2, last[2]))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out", nargs="?", default="probe_out.txt")
    ap.add_argument("--fig-regex", default=DEF_FIG)
    ap.add_argument("--tab-regex", default=DEF_TAB)
    ap.add_argument("--eq-regex", default=DEF_EQ)
    a = ap.parse_args()

    doc = fitz.open(a.pdf)
    L = ["FILE %s" % a.pdf,
         "PAGES %d  SIZE %.1f x %.1f pt" % (doc.page_count, doc[0].rect.width, doc[0].rect.height)]

    fonts = Counter()
    for pg in doc:
        for b in pg.get_text("dict")["blocks"]:
            if b["type"] == 0:
                for ln in b["lines"]:
                    for s in ln["spans"]:
                        if s["text"].strip():
                            fonts[s["font"]] += len(s["text"])
    L += ["", "## SPAN FONTS (name: chars) —— 正体字体名用于公式正文剔除判据"]
    L += ["   %-28s %d" % (k, v) for k, v in fonts.most_common(15)]

    L += ["", "## COLUMNS per page"]
    for pno in range(doc.page_count):
        cols = column_bounds(doc[pno])
        if len(cols) > 1:
            L.append("   p%-3d two-col cut_x=%.1f" % (pno + 1, cols[0][1]))

    fig_re, tab_re, eq_re = re.compile(a.fig_regex), re.compile(a.tab_regex), re.compile(a.eq_regex)
    L += ["", "## CAPTIONS (page-1-based y0-y1 x0-x1)"]
    ncap = 0
    for pno in range(doc.page_count):
        for b in doc[pno].get_text("blocks"):
            t = b[4].strip().replace("\n", " ")
            kind = "FIG" if fig_re.match(t) else ("TAB" if tab_re.match(t) else None)
            if kind:
                ncap += 1
                L.append("   p%-3d %s y%6.1f-%-6.1f x%5.1f-%-5.1f :: %s"
                         % (pno + 1, kind, b[1], b[3], b[0], b[2], t[:60]))
    L.append("   TOTAL CAPTIONS = %d" % ncap)

    L += ["", "## EQUATION TAGS (numbered, segment-rightmost)"]
    neq = Counter()
    for pno in range(doc.page_count):
        for num, tx0, tyc, tx1 in eq_tags(doc[pno], eq_re):
            neq[num.split(".")[0]] += 1
            L.append("   p%-3d (%-6s) y%6.1f tag_x %5.1f-%-5.1f" % (pno + 1, num, tyc, tx0, tx1))
    L.append("   TOTAL EQ = %d  by-series %s" % (sum(neq.values()), dict(neq)))
    L.append("   提示：by-series 应与实际公式数一致（如 2:15 3:18 A:13 B:4 C:10）")

    L += ["", "## IMAGE BLOCKS (union of rects > 40pt)"]
    for pno in range(doc.page_count):
        pg = doc[pno]
        big = []
        for im in pg.get_images(full=True):
            for r in pg.get_image_rects(im[0]):
                if r.width > 40 and r.height > 40:
                    big.append(r)
        if big:
            u = big[0]
            for r in big[1:]:
                u |= r
            L.append("   p%-3d union y%6.1f-%-6.1f x%5.1f-%-5.1f  (%d big rects)"
                     % (pno + 1, u.y0, u.y1, u.x0, u.x1, len(big)))

    open(a.out, "w", encoding="utf-8").write("\n".join(L) + "\n")
    print("written", a.out)
    print("pages=%d captions=%d equations=%d" % (doc.page_count, ncap, sum(neq.values())))


if __name__ == "__main__":
    main()
