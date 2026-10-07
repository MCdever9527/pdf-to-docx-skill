#!/usr/bin/env python3
"""ocr.py — 图片型 PDF 的文字层重建（Phase 2）。

用内置 RapidOCR（onnxruntime，纯 Python，随 DSH 运行时走，不依赖外部程序）
把页面位图转成「行级文字 + 精确 bbox」。

输出是纯几何+文本事实，不做语义判断——语义判断交给视觉层（Agent 看图）。

用法:
  python ocr.py <input.pdf> --out ocr.json [--pages 1-3] [--dpi 200]
  python ocr.py <page.png> --out ocr.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import dump_json, load_fitz  # noqa: E402

_ENGINE = None


def get_engine():
    """懒加载 RapidOCR。首次会加载内置 onnx 模型（约数秒）。"""
    global _ENGINE
    if _ENGINE is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as exc:
            raise SystemExit(
                "缺少 OCR 引擎。请执行： python -m pip install rapidocr-onnxruntime"
            ) from exc
        _ENGINE = RapidOCR()
    return _ENGINE


def quad_to_bbox(quad) -> list[float]:
    """RapidOCR 给的是四点多边形 → 轴对齐 bbox。"""
    xs = [float(p[0]) for p in quad]
    ys = [float(p[1]) for p in quad]
    return [min(xs), min(ys), max(xs), max(ys)]


def ocr_image(image_path: str | Path) -> list[dict]:
    """对单张图片做 OCR，返回行级结果（像素坐标）。"""
    engine = get_engine()
    result, _ = engine(str(image_path))
    lines: list[dict] = []
    if not result:
        return lines
    for item in result:
        quad, text, score = item[0], item[1], item[2]
        bbox = quad_to_bbox(quad)
        lines.append({
            "text": text,
            "score": round(float(score), 4),
            "bbox_px": [round(v, 1) for v in bbox],
            "quad_px": [[round(float(p[0]), 1), round(float(p[1]), 1)] for p in quad],
            "height_px": round(bbox[3] - bbox[1], 1),
        })
    lines.sort(key=lambda l: (round(l["bbox_px"][1] / 8), l["bbox_px"][0]))
    return lines


def px_to_pt(bbox_px: list[float], dpi: float) -> list[float]:
    k = 72.0 / dpi
    return [round(v * k, 2) for v in bbox_px]


def ocr_pdf(pdf_path: str, pages: str | None, dpi: int,
            workdir: Path) -> dict:
    fits = load_fitz()
    doc = fits.open(pdf_path)
    workdir.mkdir(parents=True, exist_ok=True)
    out = {"file": pdf_path, "dpi": dpi, "pages": []}
    try:
        indices = list(range(doc.page_count))
        if pages:
            picks: list[int] = []
            for part in pages.split(","):
                part = part.strip()
                if "-" in part:
                    a, b = part.split("-", 1)
                    picks.extend(range(int(a) - 1, int(b)))
                elif part:
                    picks.append(int(part) - 1)
            indices = sorted(set(i for i in picks if 0 <= i < doc.page_count))

        for i in indices:
            page = doc[i]
            pix = page.get_pixmap(dpi=dpi)
            img_path = workdir / f"page_{i + 1:03d}.png"
            pix.save(str(img_path))

            lines = ocr_image(img_path)
            for ln in lines:
                ln["bbox_pt"] = px_to_pt(ln["bbox_px"], dpi)

            # 行高换算成 pt 字号，是后续判读标题/正文的重要线索
            heights = [ln["height_px"] for ln in lines if ln["height_px"] > 0]
            out["pages"].append({
                "index": i + 1,
                "size_pt": [round(page.rect.width, 2), round(page.rect.height, 2)],
                "pix_px": [pix.width, pix.height],
                "line_count": len(lines),
                "median_line_height_px": round(sorted(heights)[len(heights) // 2], 1) if heights else 0,
                "lines": lines,
            })
    finally:
        doc.close()
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="内置 OCR：图片/扫描型 PDF → 行级文字+坐标")
    ap.add_argument("input", help="PDF 或图片")
    ap.add_argument("--out", required=True, help="输出 JSON")
    ap.add_argument("--pages", default=None)
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--workdir", default=None, help="页面 PNG 落盘目录")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists():
        print(f"[错误] 不存在: {src}", file=sys.stderr)
        return 2

    is_pdf = src.suffix.lower() == ".pdf"
    if is_pdf:
        workdir = Path(args.workdir) if args.workdir else src.parent / "_ocr_pages"
        result = ocr_pdf(str(src), args.pages, args.dpi, workdir)
    else:
        lines = ocr_image(src)
        result = {"file": str(src), "dpi": args.dpi,
                  "pages": [{"index": 1, "line_count": len(lines), "lines": lines}]}

    dump_json(result, args.out)
    total = sum(p["line_count"] for p in result["pages"])
    print(f"OCR 完成: {src.name}")
    print(f"  页数={len(result['pages'])}  行数合计={total}  dpi={args.dpi}")
    for p in result["pages"][:6]:
        head = " / ".join(l["text"] for l in p["lines"][:4])
        print(f"  p{p['index']:<3} {p['line_count']:>3}行  {head[:90]}")
    print(f"\n[JSON] {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
