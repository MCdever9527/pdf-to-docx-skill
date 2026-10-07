#!/usr/bin/env python3
"""diag_lines.py — 抽取质量诊断（开发自查用）。

打印页面每一行的几何事实与相邻行距，用来判断
「段落合并阈值该取多少」「bullet 符号是不是独立成行」。
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
from extract import spans_to_lines  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--page", type=int, default=1)
    ap.add_argument("--max", type=int, default=60)
    args = ap.parse_args()

    fits = load_fitz()
    doc = fits.open(args.pdf)
    try:
        page = doc[args.page - 1]
        lines = spans_to_lines(page)
        print(f"页 {args.page}  共 {len(lines)} 行   页面 {page.rect.width:.0f}x{page.rect.height:.0f}")
        print(f"{'#':>3} {'y0':>7} {'y1':>7} {'x0':>7} {'h':>5} {'gap':>6} {'sz':>5} {'B':>2} 文本")
        prev_bottom = None
        gaps = []
        for i, l in enumerate(lines[: args.max]):
            gap = "" if prev_bottom is None else f"{l['bbox'][1] - prev_bottom:6.2f}"
            if prev_bottom is not None:
                gaps.append(round(l["bbox"][1] - prev_bottom, 2))
            b = "B" if l["bold"] else " "
            txt = l["text"][:58]
            print(f"{i:>3} {l['bbox'][1]:>7.1f} {l['bbox'][3]:>7.1f} {l['bbox'][0]:>7.1f} "
                  f"{l['height_pt']:>5.1f} {gap if gap != '' else '     -'} "
                  f"{l['size_pt']:>5.1f} {b:>2} {txt}")
            prev_bottom = l["bbox"][3]
        if gaps:
            s = sorted(gaps)
            print(f"\n行距分布: min={s[0]:.1f} p25={s[len(s)//4]:.1f} 中位={s[len(s)//2]:.1f} "
                  f"p75={s[3*len(s)//4]:.1f} max={s[-1]:.1f}")
        return 0
    finally:
        doc.close()


if __name__ == "__main__":
    raise SystemExit(main())
