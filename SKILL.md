---
name: wmw-paper
description: Two-track academic literature workflow. Track A - translate an English paper PDF into a fully illustrated Chinese Word document with zero content loss (probe PDF structure, crop every figure/table/numbered equation at exact position, embed equations as native-size images, build with python-docx, then LibreOffice-render and verify page by page). Track B - organize literature sets (download ledger with 403/token fallback, per-paper key-point cards, thematic classification, 综述/背景调研报告, reference-list sorting). Use when the user asks to 翻译论文 / 中文翻译 / 图文并茂 / 不要有遗落 / 图片截图位置要准确, to 整理文献 / 文献要点 / 综述 / 文献下载归档 / 参考文献排序, or to turn a PDF article or a folder of PDFs into a Chinese report or Word document.
---

# WMW Paper — 文献翻译与整理

两条工作线，可单独使用也可串联（先 B 摸清文献池 → 再 A 精译关键篇）：

- **Track A**：单篇论文 PDF → 图文并茂中文 Word（重位置准确 + 完整性可证明）
- **Track B**：文献集整理 → 台账 / 要点卡片 / 分类综述 / 调研报告

---

## Non-Negotiable Rules (最高优先级)

| Rule | Requirement | Why |
|------|-------------|-----|
| **公式一律图像化** | 禁止用 Unicode/LaTeX/纯文本重排编号公式，必须从 PDF 截取为图像嵌入 | 手写转写必然引入符号错误，上下标、矩阵、积分号无法保真 |
| **公式按原尺寸嵌入** | 截取固定 DPI（360），嵌入宽度 = `px / DPI * 2.54` cm | Word 中视觉大小与论文一致，长公式不超版心 |
| **逐项计数闭环** | 生成后核对 `word/media/` 数量 == 图数 + 公式数，并打印块类型统计 | 唯一能"证明无遗漏"的手段，肉眼读 docx 不可靠 |
| **渲染后逐页视觉核验** | LibreOffice 转 PDF → 拼图 → **Read 每张拼图** | 统计脚本会漏报排版问题（表格跨页、图被裁切、公式溢出） |
| **图片裁到紧贴内容** | 图注正上方为图区，去白边后不得混入正文行或编号 `(2.1)` | "截图位置准确"是硬要求，多裁一行正文即视为失败 |
| **表格整表不跨页** | 每行加 `w:cantSplit`，除末行外全部段落 `keep_with_next=True` | 参数表被截断后上下半段无法对照 |
| **引用不译不收录** | 文内 `(Author, Year)` 保留英文原样；文末 References 不放入 Word | 用户固定要求，无需再问 |
| **下载失败必须留痕** | 403 / token 失效 / 超时的文献转存开放 HTML 全文并写入"待补清单" | 文献整理最怕静默丢失 |
| **要点卡片必须可回溯** | 每张卡片标注对应 PDF 文件名，结论只写文献里有的内容 | 防止把推断当引用 |

---

# Track A · 单篇论文 → 图文中文 Word

复制清单跟踪进度：

```
- [ ] A1 探测：页数/页面尺寸/分栏边界/全部图注表注坐标/全部编号公式坐标
- [ ] A2 插图裁剪：逐图紧贴裁剪 → 每张 Read 验证
- [ ] A3 公式裁剪：编号定位 → 剔除正文残留 → 去白边 → 逐张 Read 验证
- [ ] A4 全文提取：按分栏阅读顺序导出英文全文，通读后再动笔翻译
- [ ] A5 生成：content.py（纯数据）+ build.py（调 docx_kit 渲染）
- [ ] A6 验证：media 计数 → LibreOffice 转 PDF → 逐页拼图 → Read 全部拼图
```

## A1 Probe — 先摸清结构，不要边猜边裁

```bash
python scripts/probe_pdf.py "<paper.pdf>" probe_out.txt
```

输出：页面尺寸、左右栏 x 边界、span 字体统计、每条 `Fig. n` / `Table n` 的页码与 bbox、每个 `(n.m)` 编号的页码与 y 中心、页面内大图像块并集。

**必须人工确认**：图在注上方还是下方、有无跨栏大图、公式编号是否右对齐、附录编号体系（A1/A.1 vs Fig. A1）。判断错则后面全部返工。

## A2 Crop figures

```bash
python scripts/crop_figures.py "<paper.pdf>" <fig_dir> [--names names.json] \
       [--override 3=105,46,500,175] [--override 101=150,374,430,497]
```

以图注 y 顶部为下边界、上一图注 y 底部为上边界，在带内对**图像块 + 矢量绘制矩形**做自下向上聚类（纵向间隙 > 22pt 截断），`--z 4` 渲染后去白边。矢量图/附录小图易漏边 → 用 `--override 图号=x0,y0,x1,y1` 指定宽松带，靠去白边自动收紧；`--names names.json` 提供 `{图号: 文件名}`（附录图号记作 101=A1）。报告写 `<fig_dir>/crop_report.txt`，`MANUAL`/`EMPTY` 条目必须清零。**每张图 `Read` 一遍**。

## A3 Crop numbered equations

```bash
python scripts/eq_windows.py "<paper.pdf>" --out eq_windows.json   # 生成宽松候选窗口
# 人工核对/修正 eq_windows.json 中 (page,y0,y1,x0,x1,num)
python scripts/crop_equations.py "<paper.pdf>" eq_windows.json <eq_dir> --prose-font CharisSIL
```

窗口规则（eq_windows.py，输出条目含 `wide` 布尔）：
y = 编号所在栏内上下最近正文行中心之间（下限 46pt 避开页眉带）；
x = 编号在左半区 → 左栏 `[左缘, 0.5W]`；编号在右半区 → 先查单栏页（跨书脊正文行 ≥5，如附录推导页）或窗口 y 带内跨书脊内容（跨栏公式/图像），命中则全幅 `[左缘, 右缘]`，否则右栏 `[0.5W, 右缘]`。

**禁用 PowerShell `>` 重定向**（默认 UTF-16LE，`json.load` 会失败），一律用 `--out`。

**正文残留剔除闭环**（公式裁剪最容易失败处）：
1. 用编号文字 bbox 把 `(2.1)` 从裁剪框右侧排除（窗口内同号 word 取 y 中心最接近窗口中心者，避免行末交叉引用干扰）
2. 以编号 y 中心为界切上/下两半
3. 用 line 级判据识别正文行：窗口内完整 line 且 `prose_tokens()`（≥2 个功能词 / ≥1 个 ≥5 字母长词 / 任一四位数年份）为真即收缩窗口 —— 不依赖字体，Bold 标题与页眉也能剔除；公式碎片 line 是单符号短文本，判 False 不受影响
4. 渲染后按暗像素去白边，记录最终 bbox 命中文字
5. 逐条看报告里的 `TRIM / SKIPTRIM / EMPTY`，`EMPTY` 必须为 0

字体名因刊而异（Elsevier 多 `CharisSIL`，Springer 多 `MM*`，IEEE 多 `TimesBT`）→ 不传 `--prose-font` 时脚本自动取字符数最多的正体字体，并打印实际采用值。

## A4 Extract full text

```bash
python scripts/extract_text.py "<paper.pdf>" fulltext.txt
```

按分栏阅读顺序（先左栏自上而下，再右栏）。**必须完整通读**再翻译——跳读会导致术语错译与漏段。

## A5 Build the docx

拆两个文件，内容与排版分离：

- `<paper>_content.py` — 纯数据：`META` / `TRANSLATOR_NOTE` / `BLOCKS` / `TABLES`
- `<paper>_build.py` — 只 `import docx_kit` 渲染

`BLOCKS` 块类型（顺序即文档顺序）：

| kind | payload | 渲染 |
|------|---------|------|
| `h1` / `h2` | str | 深蓝标题，`keep_with_next` |
| `p` / `li` | str | 正文段（`li` 无首行缩进：编号假设、结论条） |
| `abs` | str | 灰底 + 左侧深蓝竖线 |
| `kw` | str | 全粗体关键词 |
| `note` | str | 9.5pt 灰小注（表注、译注） |
| `fig` | `"文件\|图注\|宽度cm"` | 居中图 + 9pt 灰图注，`keep_with_next` |
| `eqimg` | `"2.1"` | 公式图像居中制表位 + 右制表位 `(2.1)` |
| `table` | `("表 1　题名", "T1")` | Table Grid，深蓝表头白字，整表不跨页 |
| `algo` | `[行, ...]` | 16pt 精确行距伪代码 |

行内标记：`*斜体*`、`^上标^`、`~下标~`（`add_marked_runs` 解析）。

```bash
python <paper>_build.py
```

版面参数与标题区/摘要/译注/附录模板见 [reference.md](reference.md)。

## A6 Verify — 不可跳过

```bash
python scripts/verify_render.py "<out.docx>" tmp_render --expect-media 87 --tables 表1 表2 表3
# 已有渲染 PDF 时加 --pdf "<渲染.pdf>" 跳过 soffice 转换
```

1. `zipfile` 数 `word/media/` → 与预期（如图 27 + 公式 60 = 87）比对，不符直接报 `MISMATCH!`
2. LibreOffice headless 转 PDF（自动探测 `soffice` 路径，用独立 `lo_profile` 避免与已开 GUI 冲突，轮询等落盘）
3. 逐页统计块数/图像数/字符数/底部留白 → `tmp_render/pages.txt`，并列出空白页
4. 表名出现页数检查 → 一张表出现在 ≥2 页即跨页断裂
5. 3×2 拼图 `tmp_render/sheet_NN.png` → **逐张 `Read`**

```
- [ ] 图注紧跟对应图，无孤立图注
- [ ] 公式编号右对齐，长公式未超版心、未与正文重叠
- [ ] 每张表完整落在一页（查表名是否只出现在单一页）
- [ ] 附录各节标题、伪代码行号连续
- [ ] 无空白页、无重复页
```

**拼图对小字号密排文本可能显示为空白** → 判断页面有无内容一律以 `page.get_text()` 为准，不得据拼图误删内容。

---

# Track B · 文献集整理

## B1 下载与归档台账

目录命名 `<主题>文献下载/`，内含 `下载记录与待补清单.md`：

```markdown
# 下载记录与待补清单
## 文件夹说明
统一保存目录：`<主题>文献下载`
本次已下载并校验为真实 PDF 的文献共 N 篇。另有 M 篇因 403、token 失效、
下载接口超时或数据库权限限制，已保存开放 HTML 全文或列入待手动下载清单。
## 已成功下载 PDF
| 序号 | 文件名 | 主要用途 |
## 已保存 HTML 全文或链接文件
## 待手动下载清单
| 序号 | 题目 | 出处 | 失败原因 | 建议获取途径 |
```

要点：文件名统一 `年份-第一作者-中文短题.pdf`；**下载后必须校验是真实 PDF**（首字节 `%PDF`），HTML 假 PDF 是最常见陷阱；失败项进"待补清单"并注明原因与替代途径，不留空。

## B2 逐篇要点卡片

一篇一张，固定五字段（字段名可按学科调整，但同一批必须一致）：

```markdown
### N. Author, Year：中文题目
文件：`<pdf 文件名>`
- **研究对象**：
- **方法**：
- **核心结论**：
- **方法启示**：      # 或"三维意义""对本项目的直接可用参数"
- **可用于综述的点**：
```

批量抽取文本用 `python scripts/extract_text.py`（多 PDF 时逐个导出到 `pdf_texts/`），再逐篇填卡。**禁止凭标题臆造结论。**

## B3 分类与综述

先写"总体认识"：把文献池按方法学或地质角色分 3–5 类，每类一句话界定边界，并点名对当前问题**最直接支撑**的文献与仅作**理论铺垫**的文献。再按分类展开逐篇卡片。

## B4 背景调研报告

长文报告用中文数字大章 + 阿拉伯小节号，章序固定：

```
一、研究目的与范围（1.1 目的 1.2 范围）
二、数据来源与处理方法（2.1 检索策略与关键词 2.2 文献构成 2.3 审核筛选）
三、区域地质背景
四、构造演化与地层
五、地层系统与岩相古地理
六、沉积体系与储层特征
七、热演化与生排烃
八、岩石力学与地应力参数
九、建模参数取值建议
十、已有成果与不足 → 研究空白 → 下一步
```

**数值模拟类项目必须落到"建模参数取值范围"表**（参数 / 推荐值 / 区间 / 来源文献），这是调研报告对建模最有用的产出。

## B5 参考文献排序

按正文首次出现顺序编号；同一处多篇按年份升序；核对卷期、DOI、页码范围与作者拼写。输出独立 `参考文献-排序.docx`，不与译文混排。

---

## Failure Modes & Fixes

| 症状 | 根因 | 处理 |
|------|------|------|
| 公式图里混入半行正文 | 未做正体 span 判据收缩 | 必须走 A3 五步闭环，不能只靠去白边 |
| 公式上下标被切掉 | 窗口按文字 bbox 而非栅格暗行划分 | 用 `bands_of()` 栅格白通道分带，行高阈值 ≥2px |
| 表格跨页断裂 | 缺 `cantSplit` + `keep_with_next` | 见 A5 `table` |
| 图与图注分属两页 | 图段未设 `keep_with_next` | `figure()` 中图段设 True |
| 表格后凭空页 | 表后未加空段 | `build_table()` 末尾追加 4pt 空段 |
| PowerShell 输出中文乱码 | 管道 GBK 编码 | 脚本内 `sys.stdout.reconfigure(encoding='utf-8')`，结果写 UTF-8 文件而非管道 |
| MuPDF `No common ancestor in structure tree` 刷屏 | 结构树警告 | 命令尾加 `2>$null` |
| 写文件报 `save file failed, reason: unknown` | 工具假报错 | 用 `python -c` 读回确认已写入，**不要重复写** |
| 下载得到几 KB 的"PDF" | 站点返回 HTML 错误页 | 校验首字节 `%PDF`，失败转 HTML 存档 + 待补清单 |

---

## Known Limitations（2026-09 在 Elsevier 双栏样本上实测）

- **编号公式检测假设公式本体为图像/碎片 line（不产生正文 word）**：编号与本体水平空隙 >150pt 才会被 gap=15 切成独立段，故判据为"段内 word 总数==1"；行末交叉引用 `…Eq. (3.15)`（与 Eq. 空隙 <5pt）必同段被过滤。若公式本体是完整正文行（极少数刊排版），需人工补编号。
- **附录推导页行尾参差不齐**时 right_edges 候选过松（每簇众数 >=0.4W 即入候选）；它只是弱约束，主过滤由 len(seg)==1 承担，不影响检出。
- **相邻公式对**（如 3.10/3.11、B.1/B.2、C.1/C.2 等无正文行分隔）共享一个窗口，裁剪图会含两个公式，需人工调整 y0/y1。
- **正文行末残词独立成 line**（如 "zero:"）prose_tokens 判 False 不剔除，可能混入公式图底部，需人工核查 eq_report.txt 的 TRIM drop 文本。
- **公式含 "where …" 定义行**时自动窗口保留完整（比人工"只裁公式"更宽，内容完整，属可接受差异）。
- **crop_figures 矢量图内容块按面积阈 w*h>36 收集**，极细装饰线（面积<36）可能漏检导致略缺边；附录矢量小图建议 --override 指定宽松带。
- **verify_render 依赖本机 LibreOffice 转 PDF**：soffice 异常（如本机 0xC0000409 fail-fast）时用 `--pdf` 复用已有渲染 PDF。

---

## Translation & Writing Guardrails

- 术语首现给中英对照，此后统一缩写（BP/HF/DD/XFEM 等保持原文）
- 原文笔误（单位、符号不一致）→ 就地加"原文如此"，不改数据
- 原文自相矛盾处 → 译注标明"据文义更正"
- 基金号、DOI、收稿/修回日期、专有名词原样保留不译
- 交付时给出**内容清单表**（图/公式/表/附录计数），让用户可自行核对
- 译文不嵌入工作区，直接交付到源 PDF 同目录；命名 `<原论文标识>_<中文题名>_中文翻译.docx`；若进入多版本迭代，按全局规范改为 `v{n}_...` 前缀

---

## Resources

- [reference.md](reference.md) — 版面参数、标题区/摘要/译注模板全文、`META` 字段清单、整理类输出模板与范例
- `scripts/probe_pdf.py` — A1 结构探测
- `scripts/crop_figures.py` — A2 插图紧贴裁剪
- `scripts/eq_windows.py` / `scripts/crop_equations.py` — A3 公式窗口生成与裁剪
- `scripts/extract_text.py` — A4/B2 分栏阅读顺序导出
- `scripts/verify_render.py` — A6 计数 + 转 PDF + 逐页统计 + 表格检查 + 拼图
- `scripts/docx_kit.py` — python-docx 排版组件库（A5 直接 import）

依赖：`pip install pymupdf pillow numpy python-docx`（+ 本机 LibreOffice，仅 A6 需要）。
