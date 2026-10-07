#!/usr/bin/env python3
"""diag_fit.py — 页面容量诊断（开发自查用）。

量出「原文每页内容占多高」与「重建后同一页内容占多高」，
直接定位是哪一页、超了多少 pt —— 而不是靠猜参数。
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


def page_metrics(doc, i: int) -> dict:
    page = doc[i]
    bbox = page.get_text("blocks")
    ys0, ys1 = [], []
    for b in bbox:
        if b[6] != 0:
            continue
        if not (b[4] or "").strip():
            continue
        ys0.append(b[1])
        ys1.append(b[3])
    # 图片也占高度
    for info in page.get_image_info():
        ib = info.get("bbox", (0, 0, 0, 0))
        ys0.append(ib[1]); ys1.append(ib[3])
    top = min(ys0) if ys0 else 0.0
    bottom = max(ys1) if ys1 else 0.0
    return {
        "page": i + 1,
        "page_h": round(page.rect.height, 1),
        "content_top": round(top, 1),
        "content_bottom": round(bottom, 1),
        "content_h": round(bottom - top, 1),
        "text_chars": len(page.get_text("text").strip()),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("original")
    ap.add_argument("result")
    args = ap.parse_args()

    fits = load_fitz()
    da, db = fits.open(args.original), fits.open(args.result)
    try:
        ma = [page_metrics(da, i) for i in range(da.page_count)]
        mb = [page_metrics(db, i) for i in range(db.page_count)]
        print(f"原文 {da.page_count} 页   重建 {db.page_count} 页")
        print(f"{'页':>4} {'原内容高':>9} {'原顶/底':>16} {'重建高':>9} {'重建顶/底':>16} {'字符 原/果':>14}")
        for i in range(max(len(ma), len(mb))):
            a = ma[i] if i < len(ma) else None
            b = mb[i] if i < len(mb) else None
            sa = f"{a['content_h']:>9.1f} {a['content_top']:>7.1f}/{a['content_bottom']:<7.1f}" if a else f"{'-':>9} {'-':>16}"
            sb = f"{b['content_h']:>9.1f} {b['content_top']:>7.1f}/{b['content_bottom']:<7.1f}" if b else f"{'-':>9} {'-':>16}"
            chars = f"{a['text_chars'] if a else 0:>6}/{b['text_chars'] if b else 0:<6}"
            print(f"{i+1:>4} {sa} {sb} {chars}")
        if ma:
            print(f"\n原页可用高度参考: 页面 {ma[0]['page_h']:.0f}pt")
        return 0
    finally:
        da.close(); db.close()


if __name__ == "__main__":
    raise SystemExit(main())
