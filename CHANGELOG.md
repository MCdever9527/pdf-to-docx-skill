# Changelog

All notable changes to this project are documented here.
This project adheres to [Semantic Versioning](https://semver.org/).

## [1.1.0] — 2026-10-07

以一次 5 页扫描件公文的重建复盘为准，把**流程效率**固化成两个新脚本，
并修掉一个会让自检直接崩掉的缺陷。

### Added

- **`verify_sheet.py`** — 抽检拼版：把需要人眼确认的行带拼成 1~2 张长图，一次看完。
  支持 `--at` 手工给点 / `--spec` 批量给点 / `--auto` 自动切行带；产出带 `#序号`
  标尺的拼版图与 `verify_map.json` 对照表。实测把「逐行裁剪 + 逐行读图」的
  35 次调用压到 3 次。只用 Pillow 实现——**不引入 numpy**（它是本技能的可选依赖）。
- **`prespec.py`** — 渲染前的高度 / 折行预检，纯算术、不启动 Word。按 `build_docx.py`
  同一口径估算每页 `Σ(折行数 × 行距) + (段数 − 1) × 段间距`，与「原页实测内容高度」
  比对，给出 `ok` / `tight` / `overflow` 与建议收缩系数；退出码 1 = 有页溢出。
  实测把 Word 渲染迭代从 7 次压到 2 次。
- `SKILL.md` / `README.md` 增补效率纪律章节，以及扫描件（无文字层）的实测基准。
- 新增踩坑记录：扫描件逐页平移错位的反解方法；Word 默认不做 kerning 导致贴边行
  提前断行；表格单元格残留空段落撑高边框盒；跨页段落需按原件断点切分；
  用表格排双栏落款会整块跳页；标签与取值的制表位对齐规律。
- `SKILL.md` 增加脚本清单表；README 目录结构补全 `measure_ink.py`。

### Fixed

- **`bootstrap.py` 自检直接崩溃**：`_runtime.env_report()` 已把
  `missing_required` / `missing_optional` 归约成 pip 名称**字符串列表**，
  `bootstrap.py` 却按 dict 取值（`[d["pip"] for d in …]`），必然抛
  `TypeError: string indices must be integers`。现已改为直接使用字符串列表，
  并在两处加了契约说明。
- **`compare.py` 对无文字层的原件报假指标**：图片型 / 扫描件原文没有文字层，
  文本相似度取不到基准，旧版会恒返回 `0.000`，把「无法评估」误报成「内容全错」。
  现在该指标为 `None` 并显示 `N/A`，判定改为只看版面接近度
  （≤10% pass，≤25% warn），并在结论下方打印一行说明。

### Changed

- **仓库与已部署的技能副本对齐**（此前双向漂移）：
  `build_docx.py` 取新版（含页眉页脚 PAGE 域、折行估算、制表符支持，
  比库内版本多约 6.7 KB）；`compare.py` 取新版（扫描件感知）；
  `measure_ink.py` 补入库。

## [1.0.0] — 2026-10-07

首个可用版本。核心是「重建而非转码」的完整流水线，以及一套可复现的量化验收方法。

### Added

**分诊与抽取**
- `probe.py` — 按文字层字符数 / 图片面积占比 / 矢量绘制数把页面分为
  `text` / `mixed` / `image` / `scanned` / `vector`，并给出建议路线
- `extract.py` — span → 行 → 段落 → 文档块的结构抽取，含真表格（矢量线检测）
  与按绘制位置裁切的图片导出
- `render_pages.py` — 页面渲染为图 + 全局拼版总览，供视觉判读
- `ocr.py` — 内置 RapidOCR（onnxruntime，纯 Python，无外部程序）行级文字与坐标
- `ocr_to_spec.py` — OCR 结果 → 结构化规格（图片型自动初稿）

**重建**
- `build_docx.py` — 结构规格 / Markdown → 真 OOXML：
  - 页面尺寸与页边距按实测写入，不套 Word 默认值
  - 字体同时写 `ascii`/`hAnsi` 与 `eastAsia`，避免中文回退乱字体
  - 段落底纹与下边框按 OOXML schema 顺序插入
  - 分页采用**分节**而非分页符，切断溢出后的级联错位
  - **高度预算**：按页估算并收缩段后间距，防止满页内容溢出

**验收**
- `render_docx.py` — docx → pdf 统一入口，自动选择 Word COM 或 LibreOffice
- `render_docx.ps1` — Word COM 渲染分支
- `compare.py` — 逐页文本相似度 + 像素差异 + 并排对照图（原图/结果/差异热图）

**分发**
- `bootstrap.py` — 一键安装与自检，必需/可选依赖分档，支持国内镜像
- `_runtime.py` — 运行时探测：依赖、渲染器、路径，零硬编码
- `pdf2docx.py` — 对外入口，一条命令跑完整个流程

**开发自查**
- `diag_lines.py` — 行几何与行距分布诊断
- `diag_fit.py` — 原页与重建页的内容高度对比
- `make_scan_fixture.py` — 把任意 PDF 光栅化为图片型样本（用于验证 OCR 管线）

### Fixed

以下问题均在开发期通过量化诊断定位并修复：

- 页边距误按单页推断，导致底边距被高估（实测 136pt vs 真实 72pt），
  可用高度白丢 56pt 引发内容溢出 → 改为跨页极值聚合
- 段内行距基准取中位数，在段落密集页落在段间距上
  （实测 24.72pt vs 真实 14.88pt），重建高度虚高 100pt+ → 改用 p25
- 行距按 `行高 × 1.32` 估算偏大（17.03 vs 14.88），足以挤爆满页 → 改为量相邻行 `y0` 差值
- 段后间距挂到后一段，导致标题贴住正文 → 改为归属前一段
- 孤立项目符号行未按几何合并，产生大量空 `•` → 按 x0 关系合并并还原悬挂缩进
- PowerShell 参数名 `$Input` 撞自动变量导致取空值 → 改名 `InFile` / `OutFile`
- `ExportAsFixedFormat` 在 `pwsh -File` 下参数绑定不稳 → 改用 `SaveAs2($path, 17)`
- PyMuPDF ≥1.28 `get_pixmap(dpi=)` 仅接受 int → 显式取整
- Windows 控制台中文乱码 → 脚本入口统一 `reconfigure(encoding="utf-8")`
