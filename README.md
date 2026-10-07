# PDF → Word 高保真重建

> 不是把 PDF「转码」成 docx，而是把 PDF「重建」成 docx。
> 一套给 AI Agent 用的技能，也是一条能独立跑的 CLI 流水线。

PDF 是**排版结果的快照**——一堆画在坐标上的字。任何「读出文字 → 塞进 Word」的转换器
都会丢掉段落、列表、表格、页边距、行距这些结构信息，结果就是「字对了，版全乱」。

本项目走另一条路：**先用眼睛看懂版式，再按语义重建 OOXML，最后渲染回图逐页比对**。

- 面向 Agent：`SKILL.md` 定义完整方法论与流程
- 面向人：`python scripts/pdf2docx.py in.pdf out.docx` 一条命令跑完

[English README](README.en.md)

---

## 实测效果

一份 6 页文字型简历（612×792pt），重建后渲染回 PDF 与原件的逐页比对：

| 页 | 文本相似度 | 像素差异 | 判定 |
|---|---|---|---|
| 1 | **1.000** | 10.8% | pass |
| 2 | **1.000** | 9.8% | pass |
| 3 | **1.000** | 7.1% | pass |
| 4 | 0.998 | 7.2% | pass |
| 5 | 0.956 | 8.5% | warn |
| 6 | 0.992 | 3.8% | pass |

**页数 6 → 6，文本相似度 0.956–1.000，fail = 0。**

像素差异永远不可能为 0——换行位置、字体渲染、微调间距都会造成差异。
所以**内容相似度是硬指标，像素差异只是参考**。

---

## 核心思路

大多数 PDF→Word 工具死在同一件事上：把 PDF 当数据源，而不是当**画面**。

| # | 机制 | 为什么关键 |
|---|---|---|
| 1 | **用眼睛读 PDF** | 版式是视觉信息。「这是浅灰底三列表格」「这是悬挂缩进的列表」——只有看图才判得准 |
| 2 | **当重建，不当转码** | 必须按语义重建出段落、标题、列表、真表格，而不是把字倒进文档 |
| 3 | **直写 OOXML** | 精确控制页面尺寸、页边距、字号、行距、段后间距、字体 eastAsia |
| 4 | **渲染-比对-迭代** | 生成后渲染回图，与原 PDF 逐页比对，测出偏差再改 |

第 4 条最常被忽略，却是「保得住格式」的分水岭。

---

## 特性

- **类型自适应**：自动分诊 `text` / `mixed` / `image`(扫描件) / `vector`，选对应路线
- **图片型 PDF 双通道**：内置 OCR（RapidOCR，纯 Python 无外挂）给坐标，视觉判读给语义
- **内置高度预算**：检测内容顶到页面天花板的「满页」，自动收缩段后间距，防止分页漂移
- **量化验收**：逐页输出文本相似度 + 像素差异 + 并排对照图（原图 / 结果 / 差异热图）
- **零硬编码**：不写死任何绝对路径；渲染器自动降级 Word → LibreOffice → 跳过视觉验收
- **不依赖外部程序**：不调用 pdf2zh / pandoc 之类的第三方转换器

---

## 安装

需要 Python 3.9+。把仓库复制到任意位置后：

```bash
python scripts/bootstrap.py                                  # 只自检，不改环境
python scripts/bootstrap.py --install                        # 缺什么装什么
python scripts/bootstrap.py --install --with-ocr --mirror    # 连 OCR 一起装（国内镜像）
```

依赖分两档，缺可选依赖不会阻断文字型流程：

| 档位 | 组件 | 缺失后果 |
|---|---|---|
| 必需 | `pymupdf` `Pillow` `python-docx` | 无法运行 |
| 可选 | `pdfplumber` `rapidocr-onnxruntime` `opencv-python-headless` | 复杂表格抽取变弱；图片型/扫描件无法 OCR |

**渲染器**（用于把 docx 渲染回 PDF 做验收）会自动探测：

```
Microsoft Word (COM，Windows，保真最高)
   ↓ 没有
LibreOffice (headless，跨平台)
   ↓ 都没有
跳过视觉验收 —— 内容抽取与 docx 生成仍然可用
```

需要指定时用环境变量：

| 变量 | 作用 |
|---|---|
| `PDF2DOCX_SOFFICE` | 指定 `soffice` / `libreoffice` 路径 |
| `PDF2DOCX_WORD` | 指定 `WINWORD.EXE` 路径 |
| `PDF2DOCX_RENDERER` | 强制渲染器：`word` / `libreoffice` / `none` |

---

## 使用

### 一条命令

```bash
python scripts/pdf2docx.py 输入.pdf 输出.docx
```

串起 分诊 → 抽取 → 重建 → 渲染 → 验收，中间产物放在 `_pdf2docx/`，最后打印比对表。

```bash
--pages 1-3        # 只处理部分页（先小样试跑，省时间）
--workdir W        # 指定中间产物目录
--no-verify        # 跳过渲染验收
```

### 分步使用（要精细控制时）

```bash
# 0. 分诊：这是什么类型的 PDF，该走哪条路
python scripts/probe.py in.pdf --json probe.json

# 1. 文字型：直接抽结构（含表格与图片）
python scripts/extract.py in.pdf --out blocks.json --images imgs/

# 2. 图片型：出图给人/Agent 看 + OCR 拿坐标
python scripts/render_pages.py in.pdf --outdir pages/ --dpi 160 --tile
python scripts/ocr.py in.pdf --out ocr.json --dpi 200
python scripts/ocr_to_spec.py ocr.json in.pdf --out blocks.json   # OCR 初稿

# 3. 生成 docx（也可直接吃 Markdown）
python scripts/build_docx.py blocks.json out.docx
python scripts/build_docx.py rebuild.md  out.docx --preset page.json

# 4. 验收：渲染回 PDF 并逐页比对
python scripts/render_docx.py out.docx out.pdf
python scripts/compare.py in.pdf out.pdf --outdir diff/
```

### 作为 Agent 技能

把 `SKILL.md` 与 `scripts/` 放进你的 Agent 技能目录即可。
`SKILL.md` 里含完整的方法论、分诊判据、验收标准，以及一组**实测踩坑记录**。

---

## 图片型 / 扫描件 PDF

这是最难的一类，也是本项目花了最多力气的地方。

**关键事实：OCR 和视觉各有硬伤，必须融合。**

| 通道 | 强在 | 弱在 |
|---|---|---|
| 视觉判读（Agent 看图） | 判层级、分栏、表格归属、阅读顺序；修正专业术语与中英混排 | 拿不到精确坐标、字号、颜色 |
| OCR / 几何抽取 | 坐标准、字号准 | 会丢字、吞空格、写错专业词 |

实测例证：OCR 把 `John Ka-Kit Leung (梁嘉傑)` 读成 `JohnKa-Kit Leung(梁嘉)`——
**丢了「傑」、吞了空格**；而看图能读对，却量不出它在页面的 `y=75.8`。

> **融合规则：视觉决定「是什么」，OCR/几何决定「在哪、多大」。**

`pdf2docx.py` 对图片型会自动产出 **OCR 初稿**，并在输出里明确标注质量边界。
要达到接近原件的质量，需要 Agent 按 `SKILL.md` 的
「视觉判读 → Markdown 重建」路径介入复核。

---

## 验收标准

| 指标 | 回答什么 | 期望 |
|---|---|---|
| `text_similarity` | **内容**有没有丢、有没有错 | ≥ 0.995 理想；≥ 0.98 合格 |
| `pixel_hot_ratio` | **版面**接近程度 | ≤ 0.20 正常 |
| 页数一致 | 分页有没有漂 | 必须相等 |

`compare.py` 会为每页生成并排对照图：**左 = 原图，中 = 重建结果，右 = 差异热图**。

---

## 实测踩坑（都在代码里修掉了）

写这份工具时踩过的坑，全部是**量化定位**、不是猜的：

- **页边距不能按「最后一页最后一行的位置」推**。内容少的页天然排不满，那反映的是内容长度，不是页边距。实测：按单页推得底边距 136pt，真实 72pt——可用高度白丢 56pt，内容直接溢出。**必须跨页取极值聚合**。
- **段内行距基准要取分布的低分位（p25），不能取中位数**。段落密集的页里段间距可能占多数，中位数会落在段间距上：实测某页行距被算成 24.72pt（真实 14.88pt），重建高度虚高 100pt+。
- **行距 ≠ 行高 × 1.32**。实测 17.03 vs 真实 14.88，误差足以把满页内容挤到下一页。行距应直接量**相邻行 y0 的差值**。
- **段后间距属于「前一段」，不是后一段**。挂反了，标题后面就会贴住正文。
- **PDF 常把项目符号「•」单独放一行**。只按行首字符判列表会得到一堆孤立的 `•`；要按几何关系合并（符号行 x0 更小，差值即悬挂缩进宽度）。
- **分页要用「分节」而不是分页符**。一旦某页溢出，分页符会让后面所有页跟着错位；分节能切断级联误差。
- **PowerShell 参数别用 `$Input`**——它是自动变量，会取到空值。
- **`pwsh -File` 下别指望 `ExportAsFixedFormat` 的参数绑定**，12 个可选参数不稳，用 `SaveAs2($path, 17)`。
- **PyMuPDF ≥1.28 的 `get_pixmap(dpi=)` 只接受 int**，传 float 抛 TypeError。
- **Windows 控制台中文乱码**：脚本开头要 `sys.stdout.reconfigure(encoding="utf-8")`。

---

## 目录结构

```
.
├── SKILL.md                  # Agent 技能定义（方法论 / 流程 / 验收 / 踩坑）
├── README.md                 # 中文说明（本文件）
├── README.en.md              # English README
├── LICENSE                   # MIT
└── scripts/
    ├── _kit.py               # 底座：单位换算、字体映射、OOXML 助手
    ├── _runtime.py           # 运行时探测：依赖、渲染器、路径（分发关键）
    ├── bootstrap.py          # 一键安装与自检
    ├── pdf2docx.py           # 对外入口：一条命令跑完全流程
    ├── probe.py              # Phase 0  分诊
    ├── extract.py            # Phase 1  文字型结构抽取
    ├── render_pages.py       # Phase 1  图片型：出图供视觉判读
    ├── ocr.py                # Phase 1  图片型：OCR 取文字与坐标
    ├── ocr_to_spec.py        # Phase 2  OCR 结果 → 结构化规格
    ├── build_docx.py         # Phase 3  生成 docx（含高度预算）
    ├── render_docx.py        # Phase 4  渲染回 PDF（自动选渲染器）
    ├── render_docx.ps1       # Phase 4  Word COM 渲染分支
    ├── compare.py            # Phase 4  逐页比对与差异热图
    ├── diag_lines.py         # 开发自查：行几何与行距分布
    ├── diag_fit.py           # 开发自查：每页内容高度对比
    └── make_scan_fixture.py  # 开发自查：把 PDF 光栅化成扫描件样本
```

---

## 已知限制

- 复杂多栏 + 跨栏表格的阅读顺序仍可能判错，图片型需逐页复核。
- 数学公式会被还原为线性文本，不做公式对象重建。
- 表格跨页的拆分依赖 Word 自身行为，可能与原 PDF 不一致。
- 图纸、海报这类**以图形为主**的页面，「可编辑」与「保观感」需要取舍。

---

## License

[MIT](LICENSE)
