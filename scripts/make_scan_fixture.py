#!/usr/bin/env python3
"""make_scan_fixture.py — 自制「图片型 PDF」测试样本（仅用于自测）。

把任意带文字层的 PDF 光栅化成纯图像 PDF，模拟扫描件 / 图片型 PDF：
每页只剩一张位图，没有任何文字层，所有文字都以像素形式存在。

这样做的价值：原始 PDF 的文字层可以当作 ground truth，
用来客观评估 OCR + 重建管线的准确率。

用法:
  python make_scan_fixture.py <src.pdf> <out.pdf> [--dpi 200] [--pages 1-3] [--jpeg]
"""
from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import load_fitz  # noqa: E402


def build(src: str, out: str, dpi: int = 200, pages: str | None = None,
          use_jpeg: bool = False) -> dict:
    fits = load_fitz()
    sdoc = fits.open(src)
    odoc = fits.open()
    meta = {"src": src, "out": out, "dpi": dpi, "pages": [], "format": "jpeg" if use_jpeg else "png"}
    try:
        indices = range(sdoc.page_count)
        if pages:
            picks: list[int] = []
            for part in pages.split(","):
                part = part.strip()
                if "-" in part:
                    a, b = part.split("-", 1)
                    picks.extend(range(int(a) - 1, int(b)))
                elif part:
                    picks.append(int(part) - 1)
            indices = sorted(set(i for i in picks if 0 <= i < sdoc.page_count))

        for i in indices:
            spage = sdoc[i]
            # 用比目标更高的 dpi 渲染，模拟扫描件的采样损失
            pix = spage.get_pixmap(dpi=dpi)
            data = pix.tobytes("jpeg" if use_jpeg else "png")

            npage = odoc.new_page(width=spage.rect.width, height=spage.rect.height)
            rect = npage.rect
            if use_jpeg:
                npage.insert_image(rect, stream=data, keep_proportion=False)
            else:
                # PNG 走无损，避免二次压缩干扰评估
                npage.insert_image(rect, stream=data, keep_proportion=False)
            meta["pages"].append({
                "page": i + 1,
                "size_pt": [round(spage.rect.width, 2), round(spage.rect.height, 2)],
                "pix_px": [pix.width, pix.height],
                "bytes": len(data),
            })

        out_path = Path(out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        odoc.save(str(out_path), garbage=4, deflate=True)
        meta["page_count"] = odoc.page_count
        return meta
    finally:
        odoc.close()
        sdoc.close()


def main() -> int:
    ap = argparse.ArgumentParser(description="把 PDF 光栅化为图片型 PDF（造测试样本）")
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--pages", default=None, help="如 1-3")
    ap.add_argument("--jpeg", action="store_true", help="用 JPEG 而非 PNG")
    args = ap.parse_args()

    meta = build(args.src, args.out, args.dpi, args.pages, args.jpeg)
    print(f"已生成图片型 PDF: {meta['out']}")
    print(f"  页数={meta['page_count']}  dpi={meta['dpi']}  格式={meta['format']}")
    for p in meta["pages"][:8]:
        print(f"  p{p['page']}: {p['pix_px'][0]}x{p['pix_px'][1]}px  {p['bytes']//1024}KB")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
