#!/usr/bin/env python3
"""measure_ink.py — 图片型 PDF 的内容边界实测（开发自查用）。

图片型 PDF 没有文字层，页边距只能从「墨迹边界」量：
对每页做二值化，取非白像素的外接范围，再跨页聚合极值。

用法:
  python measure_ink.py in.pdf [--dpi 150] [--threshold 200]
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
from _kit import load_fitz  # noqa: E402

import numpy as np  # noqa: E402
from PIL import Image  # noqa: E402


def ink_bbox(page, dpi: int, threshold: int, margin_px: int) -> tuple | None:
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
    a = np.asarray(img)
    mask = a < threshold
    # 忽略边缘 margin_px 内的扫描噪声
    if margin_px > 0:
        mask[:margin_px, :] = False
        mask[-margin_px:, :] = False
        mask[:, :margin_px] = False
        mask[:, -margin_px:] = False
    rows = np.where(mask.any(axis=1))[0]
    cols = np.where(mask.any(axis=0))[0]
    if len(rows) == 0 or len(cols) == 0:
        return None
    k = 72.0 / dpi
    return (cols[0] * k, rows[0] * k, (cols[-1] + 1) * k, (rows[-1] + 1) * k)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--threshold", type=int, default=200)
    ap.add_argument("--margin-px", type=int, default=6)
    args = ap.parse_args()

    fits = load_fitz()
    doc = fits.open(args.pdf)
    try:
        boxes = []
        print(f"{'页':>4} {'内容边界(pt) x0,y0,x1,y1':<34} {'上':>7}{'左':>7}{'右':>7}{'下':>7}")
        for i in range(doc.page_count):
            page = doc[i]
            bb = ink_bbox(page, args.dpi, args.threshold, args.margin_px)
            if bb is None:
                print(f"{i+1:>4} {'(空白页)':<34}")
                continue
            x0, y0, x1, y1 = bb
            boxes.append({"page": i + 1, "bbox": bb,
                          "w": page.rect.width, "h": page.rect.height})
            print(f"{i+1:>4} {x0:7.1f},{y0:7.1f},{x1:7.1f},{y1:7.1f}   "
                  f"{y0:7.1f}{x0:7.1f}{page.rect.width - x1:7.1f}{page.rect.height - y1:7.1f}")

        if boxes:
            w = boxes[0]["w"]
            h = max(b["h"] for b in boxes)
            left = min(b["bbox"][0] for b in boxes)
            top = min(b["bbox"][1] for b in boxes)
            right = w - max(b["bbox"][2] for b in boxes)
            bottom = h - max(b["bbox"][3] for b in boxes)
            print(f"\n跨页聚合页边距: L{left:.1f} R{right:.1f} T{top:.1f} B{bottom:.1f} pt")
            print(f"页面 {w:.0f}x{h:.0f}pt  可用高度 {h - top - bottom:.1f}pt")
        return 0
    finally:
        doc.close()


if __name__ == "__main__":
    raise SystemExit(main())