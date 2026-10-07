# Changelog

All notable changes to this project are documented here.
This project adheres to [Semantic Versioning](https://semver.org/).

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
