#!/usr/bin/env python3
"""render_pages.py — 把 PDF 页面导出成图，供视觉判读（Phase 1-B）。

图片型 / 扫描型 PDF 没有文字层，只能靠「看」。本步骤不解析任何内容，
只负责把页面变成 Agent 能读的素材：
  - 逐页 PNG（默认 160dpi，够看清 9pt 小字）
  - 全局拼版总览图（一次看清整篇的版式规律）
  - 按区块切小图（局部放大，用于确认表格线、符号、颜色）

用法:
  python render_pages.py in.pdf --outdir pages/ [--dpi 160] [--pages 1-4] [--tile]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import dump_json, load_fitz  # noqa: E402

from PIL import Image  # noqa: E402


def render_page(page, dpi: int, out: Path, png: bool = True) -> tuple[Path, tuple[int, int]]:
    pix = page.get_pixmap(dpi=int(dpi))
    path = out.with_suffix(".png" if png else ".jpg")
    pix.save(str(path))
    return path, (pix.width, pix.height)


def make_overview(images: list[Path], out: Path, cols: int = 4,
                  thumb_w: int = 420) -> Path | None:
    """把多页缩略图拼成一张总览，方便一眼看出整篇的版式规律。"""
    if not images:
        return None
    thumbs = []
    for p in images:
        im = Image.open(p).convert("RGB")
        ratio = thumb_w / im.width
        thumbs.append(im.resize((thumb_w, int(im.height * ratio)), Image.LANCZOS))
    rows = (len(thumbs) + cols - 1) // cols
    cell_h = max(t.height for t in thumbs)
    gap = 10
    canvas = Image.new("RGB", (cols * thumb_w + (cols + 1) * gap,
                               rows * cell_h + (rows + 1) * gap), (235, 235, 235))
    for i, t in enumerate(thumbs):
        r, c = divmod(i, cols)
        canvas.paste(t, (gap + c * (thumb_w + gap), gap + r * (cell_h + gap)))
    canvas.save(str(out))
    return out


def parse_pages(spec: str | None, total: int) -> list[int]:
    if not spec:
        return list(range(total))
    picks: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            picks.extend(range(int(a) - 1, int(b)))
        elif part:
            picks.append(int(part) - 1)
    return sorted(set(i for i in picks if 0 <= i < total))


def main() -> int:
    ap = argparse.ArgumentParser(description="PDF 页面 → 图片素材（供视觉判读）")
    ap.add_argument("pdf")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--dpi", type=int, default=160)
    ap.add_argument("--pages", default=None)
    ap.add_argument("--tile", action="store_true", help="额外生成全局拼版总览")
    ap.add_argument("--json", default=None)
    args = ap.parse_args()

    fits = load_fitz()
    src = Path(args.pdf)
    if not src.exists():
        print(f"[错误] 不存在: {src}", file=sys.stderr)
        return 2

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    doc = fits.open(str(src))
    rendered: list[Path] = []
    meta = []
    try:
        for i in parse_pages(args.pages, doc.page_count):
            page = doc[i]
            path, (w, h) = render_page(page, args.dpi,
                                       outdir / f"p{i + 1:03d}.png")
            rendered.append(path)
            meta.append({
                "page": i + 1,
                "png": str(path),
                "pix_px": [w, h],
                "size_pt": [round(page.rect.width, 2), round(page.rect.height, 2)],
            })
    finally:
        doc.close()

    overview = None
    if args.tile:
        overview = make_overview(rendered, outdir / "overview.png")

    print(f"已渲染 {len(rendered)} 页 → {outdir}")
    for m in meta[:12]:
        print(f"  p{m['page']:<3} {m['pix_px'][0]}x{m['pix_px'][1]}px  {Path(m['png']).name}")
    if overview:
        print(f"  总览图: {overview}")
    if args.json:
        dump_json({"file": str(src), "dpi": args.dpi, "pages": meta,
                   "overview": str(overview) if overview else None}, args.json)
        print(f"[JSON] {args.json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
