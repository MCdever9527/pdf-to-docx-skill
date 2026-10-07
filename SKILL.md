---
name: pdf-to-docx
description: 把 PDF 高保真重建为可编辑的 .docx，覆盖文字型、图片型（扫描件/图形页）与混合型。当用户要求「PDF 转 Word」「PDF 转 docx」「保持原格式」「扫描件转 Word」「照着 PDF 重建文档」或格式被转码工具搞乱需要重做时使用。
version: 1.0.0
---

# PDF → Word 高保真重建

把 PDF 变成**可编辑、且版式贴近原件**的 .docx。整套工具链内置在 DSH 运行时里，
不调用任何外部独立程序（不依赖 pdf2zh / pandoc / 第三方转换器）。

## 一、先理解：为什么 Claude 在这件事上表现好

不是因为它用了某个更强的库，而是因为它在这四个机制上都不偷懒。任何一条缺失，效果都会塌：

| # | 机制 | 具体含义 |
|---|---|---|
| 1 | **用眼睛读 PDF** | 把页面渲染成图直接看。版式是视觉信息——「这是浅灰底三列表格」「这是悬挂缩进的列表」，只有看图才判得准。 |
| 2 | **当重建，不当转码** | PDF 是**排版结果的快照**，不是结构源。它是「一堆画在坐标上的字」。「转码」思维必然丢结构；必须按语义**重建**出段落、标题、列表、真表格。 |
| 3 | **直写 OOXML** | 用 python-docx 精确控制页面尺寸、页边距、字号、行距、段后间距、字体 eastAsia，而不是靠中间格式碰运气。 |
| 4 | **渲染-比对-迭代** | 生成后**渲染回图**，与原 PDF 逐页比对，测出偏差再改。这一步最常被忽略，却是「保得住格式」的分水岭。 |

本技能把这四条固化成可执行的脚本流水线。

## 二、安装与分发（换一台机器也能跑）

技能目录是**自包含**的，整个 `pdf-to-docx/` 复制到任意机器后先自检：

```bash
python scripts/bootstrap.py                                  # 只自检，不改环境
python scripts/bootstrap.py --install                        # 缺什么装什么
python scripts/bootstrap.py --install --with-ocr --mirror    # 连 OCR 一起装，走国内镜像
```

**零硬编码**：脚本不写死任何本机绝对路径，按
「环境变量 → 常见安装位置 → PATH」逐级探测。需要覆盖时用环境变量：

| 变量 | 作用 |
|---|---|
| `PDF2DOCX_SOFFICE` | 指定 `soffice` / `libreoffice` 路径 |
| `PDF2DOCX_WORD` | 指定 `WINWORD.EXE` 路径 |
| `PDF2DOCX_RENDERER` | 强制渲染器：`word` / `libreoffice` / `none` |

**渲染器自动降级链**：

```
Microsoft Word (COM，Windows，保真最高)
   ↓ 没有
LibreOffice (headless，跨平台)
   ↓ 都没有
跳过视觉验收 —— 内容抽取与 docx 生成仍然可用
```

依赖分两档，缺可选依赖不会阻断文字型流程：

| 档位 | 组件 | 缺失后果 |
|---|---|---|
| 必需 | pymupdf、Pillow、python-docx | 无法运行 |
| 可选 | pdfplumber、rapidocr-onnxruntime、opencv | 复杂表格抽取变弱；图片型/扫描件无法 OCR |

## 三、快速开始（一条命令）

```bash
python scripts/pdf2docx.py in.pdf out.docx
```

它会串起 分诊 → 抽取 → 重建 → 渲染 → 验收，并把中间产物放进 `_pdf2docx/`，
最后打印逐页比对表和结论。常用开关：

```bash
--pages 1-3        # 只处理部分页（先小样试跑，省时间）
--workdir W        # 指定中间产物目录
--no-verify        # 跳过渲染验收
```

图片型 PDF 走同一条命令会给出**OCR 初稿**，并在输出里明确提示质量边界。

## 四、工具链（已内置）

Python 侧（由 bootstrap 安装到当前运行时）：

| 能力 | 组件 |
|---|---|
| PDF 解析 / 渲染 | PyMuPDF（`pymupdf`） |
| 表格抽取 | PyMuPDF `find_tables` + pdfplumber |
| OCR（图片型） | rapidocr-onnxruntime（onnxruntime，纯 Python，模型内置） |
| 图像分析 | OpenCV headless + Pillow |
| docx 生成 | python-docx |

渲染回图侧：**本机 Word COM**（`render_docx.ps1`）。
用 Word 本体渲染，是因为验收要回答「这份 docx 在 Word 里打开长什么样」——
换任何第三方渲染器，分页与字体回退的差异都会掩盖真实问题。

## 五、流程

```
Phase 0  分诊      probe.py          → 判定类型，选路线
Phase 1  取素材    extract.py / render_pages.py
Phase 2  定结构    几何事实 + 视觉判读（+ OCR）
Phase 3  生成      build_docx.py     → .docx
Phase 4  验收      render_docx.ps1 → compare.py → 迭代
```

### Phase 0 · 分诊

```bash
python <skill>/scripts/probe.py in.pdf --json probe.json
```

按「文字层字符数 / 图片面积占比 / 矢量绘制数」把每页分成：

| 类型 | 特征 | 路线 |
|---|---|---|
| `text` | 有完整文字层 | **A**：几何直接可读，不需要 OCR |
| `mixed` | 有文字层 + 大图 | **B**：文字精确抽取 + 图片按位还原 |
| `image` | 每页几乎 0 字符、整页位图 | **C**：OCR + 视觉判读双通道 |
| `scanned` | 整页位图、有倾斜噪声 | **C**：先校正去噪再走 C |
| `vector` | 文字转曲线，无文字层 | **D**：从 `get_drawings()` 反推色块，OCR 补文字 |

判断依据是**实测数字**，不是文件名或猜测。

### Phase 1 · 取素材

**路线 A/B（有文字层）**——直接抽几何：

```bash
python <skill>/scripts/extract.py in.pdf --out blocks.json [--images imgs/]
```

抽出 span 级 `(text, font, size, color, bbox, flags)`，再聚成 行 → 段落 → 文档块。

**路线 C/D（无文字层）**——先把页面变成我能看的东西：

```bash
python <skill>/scripts/render_pages.py in.pdf --outdir pages/ --dpi 160 --tile
```

然后**逐页 read_image 看图**。同时跑 OCR 拿坐标：

```bash
python <skill>/scripts/ocr.py in.pdf --out ocr.json --dpi 200
```

### Phase 2 · 定结构（关键分工）

两条通道各有短板，必须**融合**，这也是整个方案的核心：

| 通道 | 强在 | 弱在 |
|---|---|---|
| **视觉判读（我自己看图）** | 判层级、分栏、表格归属、阅读顺序；修正专业术语与中英混排 | 拿不到精确坐标、字号、颜色值 |
| **OCR / 几何抽取** | 坐标准、字号准、能算 | 会丢字、吞空格、写错专业词 |

实测例证：OCR 把 `John Ka-Kit Leung (梁嘉傑)` 读成 `JohnKa-Kit Leung(梁嘉)`——
**丢了「傑」、吞了空格**；而看图能读对，却量不出它在页面的 y=75.8。

**融合规则：视觉决定「是什么」，OCR/几何决定「在哪、多大」。**

图片型 PDF 的推荐产出物不是 blocks.json，而是**我直接写的 Markdown 重建稿**，
然后把页面几何交给 `build_docx.py` 的 `--preset`：

```bash
python <skill>/scripts/build_docx.py rebuild.md out.docx --preset page.json
```

（Markdown 里的 `##` → 色块标题，`-` → 列表，`|` → 真表格。）

### Phase 3 · 生成

`build_docx.py` 负责把结构写成真 OOXML。三个决定成败的细节：

1. **页面尺寸与页边距按实测值写入**，不套 Word 默认值。
2. **字体同时写 `w:rFonts` 的 `ascii/hAnsi`（拉丁）与 `eastAsia`（东亚）**——
   只设 `run.font.name` 中文会回退乱字体。
3. **高度预算**：原 PDF 的某些页是「满页」，重建只要多出几磅就会把末尾几行挤到
   下一页，而一旦有一页溢出，**后面所有页的对应关系全部错乱**。
   引擎因此按页估算高度并对段后间距做等比收缩；分页用**分节**（section break）
   而非分页符，切断级联误差。

### Phase 4 · 验收（不能省）

```bash
pwsh -NoProfile -File <skill>/scripts/render_docx.ps1 -InFile out.docx -OutFile out.pdf
python <skill>/scripts/compare.py in.pdf out.pdf --outdir diff/
```

产出三种证据：**数值表** + **逐页并排对照图**（左原 / 中结果 / 右差异热图）。

## 六、验收标准

两条指标回答的是**不同问题**，不要混为一谈：

| 指标 | 回答什么 | 期望 |
|---|---|---|
| `text_similarity` | **内容**有没有丢、有没有错 | ≥ 0.995 理想；≥ 0.98 合格 |
| `pixel_hot_ratio` | **版面**接近程度 | ≤ 0.20 正常 |
| 页数一致 | 分页有没有漂 | 必须相等 |

像素差异**永远不可能为 0**：换行位置、字体渲染、微调间距都会产生差异。
拿像素差当唯一标准会陷入无意义的追打；**内容相似度才是硬指标**。

实测基准（一份 6 页文字型简历，612×792pt）：

```
结论: pass=5 warn=1 fail=0
页 1  相似度 1.000  像素差 10.7%
页 2  相似度 1.000  像素差  9.7%
页 3  相似度 1.000  像素差  7.2%
页 4  相似度 0.998  像素差  7.2%
页 5  相似度 0.956  像素差  8.4%
页 6  相似度 0.992  像素差  3.7%
```

## 七、踩过的坑（都是实测出来的，别重犯）

- **页边距不能按「第一页最后一行的位置」推**。内容少的页天然排不满，
  那反映的是内容长度，不是页边距。实测：按单页推得底边距 136pt，
  真实值 72pt——可用高度被压掉 56pt，内容直接溢出。**必须跨页取极值聚合**。
- **段内行距基准要取分布的小分位（p25），不能取中位数**。
  段落密集的页里段间距可能占多数，中位数会落在段间距上：
  实测某页行距被算成 24.72pt（真实 14.88pt），重建后高度虚高 100pt+。
- **行距≠行高**。用 `行高 × 1.32` 估行距会偏大（实测 17.03 vs 真实 14.88），
  足已把满页内容挤到下一页。行距应当直接量**相邻行 y0 的差值**。
- **段后间距属于「前一段」，不是后一段**。挂反了标题后面就会贴住正文。
- **PDF 常把项目符号「•」单独放一行**，与文字行分成两个对象。
  只按行首字符判列表会得到一堆孤立的 `•`。要按几何关系合并：
  符号行 x0 更小、文字行更深，差值就是悬挂缩进宽度。
- **PowerShell 参数别用 `$Input`**——它是自动变量，会导致取到空值。
- **`pwsh -File` 里不要指望 `ExportAsFixedFormat` 的参数绑定**，
  12 个可选参数在 PS 下不稳，用 `SaveAs2($path, 17)`。
- **PyMuPDF ≥1.28 的 `get_pixmap(dpi=)` 只接受 int**，传 float 抛 TypeError。
- **控制台中文乱码**：脚本开头要 `sys.stdout.reconfigure(encoding="utf-8")`。

## 八、图片型 PDF 的额外注意

- **渲染 dpi 选 160**：9pt 小字能看清，又不至于让单页图太大。
- **扫描件先做倾斜校正**再 OCR，否则行会被切碎。
- **装饰性背景**（色块、弧线、纹理）有两条策略，按需求选：
  - 要**可编辑**：用色块标题的段落底纹 + 边框重绘（`build_docx.py` 已支持 `shade` / `border_bottom`）。
  - 要**高保真**：整页底图 + 文字层，代价是背景不可编辑。
- **正文照片要抽原图**（`extract.py --images`），不要截屏，否则清晰度会掉。
- **专业术语务必用视觉复核 OCR 结果**：`ISO55000`、`FSOE`、`SMIEEE`、`CEnv`
  这类缩写，OCR 极易写错，而错一个字符整份文档的专业度就打了折扣。

## 九、局限（诚实说明）

- 复杂多栏 + 跨栏表格的版面，段落阅读顺序仍可能判错，需要人工/视觉逐页复核。
- 数学公式会被还原为线性文本，不做公式对象重建。
- 表格跨页时的拆分逻辑依赖 Word 自身行为，与原 PDF 可能不一致。
- 图纸、海报这类**以图形为主**的页面，重建的可编辑版本在观感上必然与原图有差距，
  此时应主动与用户确认：要「可编辑」还是「保观感」。
