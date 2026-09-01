# -*- coding: utf-8 -*-
"""A6 交付前验证：media 计数 -> LibreOffice 转 PDF -> 逐页统计 -> 表格完整性 -> 拼图

用法:
    python verify_render.py <out.docx> <workdir>
        [--pdf 已渲染的PDF] [--soffice PATH] [--expect-media N] [--tables 表1 表2 表3]
        [--sheet-pages 6] [--dpi 1.6] [--only-pdf]

步骤:
    1) zipfile 数 word/media/*        —— 与预期 图数+公式数 比对
    2) soffice headless 转 PDF        —— 独立 UserInstallation 目录，避免与已开 GUI 冲突；
                                      已有渲染结果时可 --pdf 直接复用跳过转换
    3) 逐页 块数/字符数/尾部留白      —— 写 <workdir>/pages.txt
    4) 表格完整性                     —— 每个表名只应出现在 1 个页面，否则跨页断裂
    5) 拼图                           —— 每 sheet-pages 页一张，写 <workdir>/sheet_XX.png

**拼图后必须逐张 Read**。注意：小字号密排文本在缩略图中可能显示为空白，
判断页面有无内容一律以 pages.txt 的 chars 为准。
"""
import sys
import os
import glob
import time
import zipfile
import argparse
import subprocess

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import fitz
from PIL import Image, ImageDraw

SOFFICE_CAND = [
    os.environ.get("SOFFICE", ""),
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "tools/LibreOffice/program/soffice.exe",
    "tools/LibreOffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "soffice",
    "libreoffice",
]


def find_soffice(explicit):
    cands = [explicit] if explicit else SOFFICE_CAND
    for c in cands:
        if c and os.path.exists(c):
            return c
        if c and os.path.basename(c) == c:
            from shutil import which
            w = which(c)
            if w:
                return w
    raise SystemExit("未找到 soffice，请用 --soffice 指定 LibreOffice 可执行文件路径")


def count_media(docx):
    with zipfile.ZipFile(docx) as z:
        return len([n for n in z.namelist() if n.startswith("word/media/")])


def to_pdf(soffice, docx, workdir):
    os.makedirs(workdir, exist_ok=True)
    profile = os.path.join(workdir, "lo_profile")
    url = "file:///" + profile.replace("\\", "/")
    cmd = [soffice, "--headless", "--norestore", "-env:UserInstallation=" + url,
           "--convert-to", "pdf", "--outdir", workdir, docx]
    try:
        subprocess.run(cmd, timeout=300, capture_output=True)
    except subprocess.TimeoutExpired:
        pass
    stem = os.path.splitext(os.path.basename(docx))[0]
    target = os.path.join(workdir, stem + ".pdf")
    for _ in range(20):                      # soffice 可能先返回再落盘
        if os.path.exists(target) and os.path.getsize(target) > 1024:
            return target
        time.sleep(0.5)
    got = glob.glob(os.path.join(workdir, "*.pdf"))
    if got:
        return max(got, key=os.path.getmtime)
    raise SystemExit("转换失败：未在 %s 找到 PDF" % workdir)


def page_stats(pdf, out_txt):
    doc = fitz.open(pdf)
    L = ["PDF %s  pages=%d" % (pdf, doc.page_count)]
    for i, pg in enumerate(doc):
        txt = pg.get_text().strip()
        blocks = pg.get_text("dict")["blocks"]
        ymax = 0
        for b in blocks:
            if b["bbox"][1] > pg.rect.height - 30:      # 跳过页脚页码
                continue
            if b["type"] == 1:
                ymax = max(ymax, b["bbox"][3])
            elif b["type"] == 0 and "".join(
                    s["text"] for l in b["lines"] for s in l["spans"]).strip():
                ymax = max(ymax, b["bbox"][3])
        nimg = sum(1 for b in blocks if b["type"] == 1)
        L.append("p%-3d blocks=%-3d imgs=%-2d chars=%-5d bottom_gap=%.0fpt"
                 % (i + 1, len(blocks), nimg, len(txt), pg.rect.height - 30 - ymax))
    open(out_txt, "w", encoding="utf-8").write("\n".join(L) + "\n")
    return doc


def check_tables(doc, names):
    """表名只出现在 1 页 -> 未跨页；出现在 >=2 页 -> 可能断裂（正文引用需人工排除）"""
    rep = []
    for nm in names:
        pages = [i + 1 for i, pg in enumerate(doc) if nm in pg.get_text()]
        rep.append("%-8s pages=%s %s" % (nm, pages, "OK" if len(pages) <= 1 else "<< 检查是否跨页断裂"))
    return rep


def make_sheets(doc, workdir, per=6, zoom=1.6):
    CW, CH = 700, 990
    cols = 3
    paths, n = [], doc.page_count
    for start in range(0, n, per):
        cnt = min(per, n - start)
        rows = (cnt + cols - 1) // cols
        sheet = Image.new("RGB", (CW * cols + 40, CH * rows + 40), "white")
        dr = ImageDraw.Draw(sheet)
        for k in range(cnt):
            idx = start + k
            pm = doc[idx].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
            im.thumbnail((CW - 20, CH - 50))
            c, r = k % cols, k // cols
            x, y = 10 + c * (CW + 10), 30 + r * (CH + 10)
            sheet.paste(im, (x, y))
            dr.text((x, y - 16), "PDF p%d" % (idx + 1), fill="red")
        p = os.path.join(workdir, "sheet_%02d.png" % (start // per + 1))
        sheet.save(p)
        paths.append(p)
    return paths


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("workdir")
    ap.add_argument("--soffice", default=None)
    ap.add_argument("--pdf", default=None, help="直接使用已有渲染 PDF（跳过 soffice 转换）")
    ap.add_argument("--expect-media", type=int, default=None)
    ap.add_argument("--tables", nargs="*", default=[])
    ap.add_argument("--sheet-pages", type=int, default=6)
    ap.add_argument("--dpi", type=float, default=1.6)
    ap.add_argument("--only-pdf", action="store_true")
    a = ap.parse_args()

    media = count_media(a.docx)
    kb = os.path.getsize(a.docx) // 1024
    print("media=%d  docx=%dKB" % (media, kb))
    if a.expect_media is not None:
        print("MEDIA %s (expect %d)" % ("OK" if media == a.expect_media else "MISMATCH!", a.expect_media))

    pdf = a.pdf or to_pdf(find_soffice(a.soffice), a.docx, a.workdir)
    if not a.pdf:
        print("pdf=%s" % pdf)
    if a.only_pdf:
        return

    doc = page_stats(pdf, os.path.join(a.workdir, "pages.txt"))
    print("pages=%d  stats -> %s" % (doc.page_count, os.path.join(a.workdir, "pages.txt")))

    if a.tables:
        for line in check_tables(doc, a.tables):
            print("  " + line)

    blanks = [i + 1 for i, pg in enumerate(doc) if len(pg.get_text().strip()) < 5]
    print("blank/low-text pages: %s" % (blanks or "none"))

    paths = make_sheets(doc, a.workdir, a.sheet_pages, a.dpi)
    print("sheets (%d) -> 逐张 Read:" % len(paths))
    for p in paths:
        print("  " + p)


if __name__ == "__main__":
    main()
