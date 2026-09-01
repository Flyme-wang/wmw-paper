# wmw-paper

Qoder / Codex 个人技能：**学术文献双轨工作流**（与 [wmw-ppt](https://github.com/Flyme-wang/wmw-ppt) 同系列）。

- **轨道 A · 论文翻译**：把英文学术论文 PDF 翻译为中文图文 Word——探测版面、按图注精确定位裁剪插图、编号公式图像化嵌入（360 DPI 原尺寸）、python-docx 排版、LibreOffice 渲染逐页视觉核验，零内容遗漏。
- **轨道 B · 文献整理**：文献下载归档（403/Token 回退）、单篇要点卡片、主题分类、综述/背景调研报告、参考文献排序。

## 目录结构

```
wmw-paper/
├── SKILL.md            # 技能主指令（工作流 + 检查清单 + 已知限制）
├── reference.md        # 排版参数、尺寸保真链路、编号剥离等细节
└── scripts/
    ├── probe_pdf.py        # A1 结构探测：分栏/图注表注/编号公式/字体统计
    ├── extract_text.py     # A2 按版面顺序提取全文文本
    ├── crop_figures.py     # A3 插图精确裁剪（按图注定位 + 内容包围盒）
    ├── eq_windows.py       # A3 公式裁剪前置：生成宽松候选窗口（含行末引用过滤）
    ├── crop_equations.py   # A3 编号公式裁剪（正文残留自动剔除闭环）
    ├── docx_kit.py         # A5 python-docx 排版库（图/公式/表格/伪代码块）
    └── verify_render.py    # A6 LibreOffice 渲染 + 逐页统计 + 拼图视觉核验
```

## 快速开始（轨道 A）

```bash
# 1. 结构探测（输出 captions=30 / equations=60 等与人工基准核对）
python scripts/probe_pdf.py paper.pdf probe_out.txt

# 2. 公式候选窗口（人工核对后交裁剪）
python scripts/eq_windows.py paper.pdf --out eq_windows.json --out-report eq_windows_report.txt

# 3. 编号公式裁剪（正文残留剔除闭环，eq_report.txt 逐条人工过）
python scripts/crop_equations.py paper.pdf eq_windows.json eq_out/ --prose-font CharisSIL

# 4. 插图裁剪
python scripts/crop_figures.py paper.pdf fig_out/

# 5. 渲染核验（docx → PDF → 逐页拼图）
python scripts/verify_render.py out.docx verify_out/ --expect-media 87 --pdf rendered.pdf
```

## 依赖

- Python 3.10+：PyMuPDF (fitz)、numpy、Pillow、python-docx
- LibreOffice（headless 渲染验证；若本机 soffice 崩溃可用 `--pdf` 复用已渲染 PDF）
- 推荐：`sys.stdout.reconfigure(encoding='utf-8')` 处理 PowerShell 管道中文

## 已知限制

见 `SKILL.md` 末尾 "Known Limitations" 小节（编号检测假设公式本体为图像、相邻公式对共享窗口需人工调整、soffice 崩溃回退等）。

## 许可

个人技能库，仅本人工作流使用。
