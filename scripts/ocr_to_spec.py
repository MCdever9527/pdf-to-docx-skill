#!/usr/bin/env python3
"""ocr_to_spec.py — 把 OCR 行结果转成结构化规格（图片型路线的自动初稿）。

OCR 只给「文字 + 坐标」，这里补上结构判断：
  - 行 → 段落（用行距分布找段间隔，与 extract.py 同一套判据）
  - 行高 → 字号估算
  - 行首符号 → 列表项
  - 内容边界 → 页边距

产出的 spec 可直接喂给 build_docx.py。
但要清楚：这一步是**机器初稿**，OCR 的丢字/吞空格/错专业词都会带进来，
正式交付前应当由 Agent 看图复核（见 SKILL.md）。
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import clean_text, dump_json, has_cjk, load_fitz  # noqa: E402

BULLET_CHARS = "•·▪‣◦∙-–—*"


def px_to_pt(v: float, dpi: float) -> float:
    return v * 72.0 / dpi


def group_lines(lines: list[dict], page_h_pt: float) -> list[dict]:
    """行 → 段落：按行距分布定阈值，与文字型抽取用同一套判据。"""
    if not lines:
        return []
    ordered = sorted(lines, key=lambda l: (l["bbox_pt"][1], l["bbox_pt"][0]))

    gaps: list[float] = []
    for i in range(1, len(ordered)):
        g = ordered[i]["bbox_pt"][1] - ordered[i - 1]["bbox_pt"][3]
        if g > 0:
            gaps.append(g)
    if gaps:
        gs = sorted(gaps)
        ref = gs[min(int(len(gs) * 0.25), len(gs) - 1)]
        threshold = max(ref * 2.5, ref + 2.5)
    else:
        ref, threshold = 2.0, 5.0

    paras: list[dict] = []
    cur: dict | None = None
    for ln in ordered:
        text = clean_text(ln.get("text", "")).strip()
        if not text:
            continue
        x0, y0, x1, y1 = ln["bbox_pt"]
        height = max(y1 - y0, 1.0)

        if cur is None:
            cur = {"text": text, "x0": x0, "y0": y0, "y1": y1,
                   "median_h": height, "lines": 1}
            continue

        gap = y0 - cur["y1"]
        same_scale = abs(height - cur["median_h"]) <= max(cur["median_h"] * 0.35, 2.0)
        tight = gap <= threshold

        if tight and same_scale and not text[:1] in BULLET_CHARS:
            joiner = "" if (has_cjk(cur["text"][-1:]) or has_cjk(text[:1])) else " "
            cur["text"] = (cur["text"] + joiner + text).strip()
            cur["y1"] = y1
            cur["lines"] += 1
            continue

        paras.append(cur)
        cur = {"text": text, "x0": x0, "y0": y0, "y1": y1,
               "median_h": height, "lines": 1}

    if cur is not None:
        paras.append(cur)

    for p in paras:
        p["_gap_after"] = 0.0
    for i in range(len(paras) - 1):
        g = paras[i + 1]["y0"] - paras[i]["y1"]
        paras[i]["_gap_after"] = max(g, 0.0)
        paras[i]["_threshold"] = threshold
    if paras:
        paras[-1]["_threshold"] = threshold
    return paras


def build_spec(ocr: dict, pdf_path: str, gap_split: bool = True) -> dict:
    fits = load_fitz()
    doc = fits.open(pdf_path)
    try:
        page_w = doc[0].rect.width
        page_h = doc[0].rect.height
    finally:
        doc.close()

    dpi = float(ocr.get("dpi") or 200)

    # 全局字号基准：按字符数加权的中位行高（行高约为字号的 1.2~1.3 倍）
    heights: list[tuple[float, int]] = []
    for pg in ocr.get("pages", []):
        for ln in pg.get("lines", []):
            h = px_to_pt(ln.get("height_px", 0) or 0, dpi)
            if h > 0:
                heights.append((h, max(len(ln.get("text", "")), 1)))
    heights.sort(key=lambda x: x[0])
    total = sum(w for _, w in heights) or 1
    acc, body_line_h = 0, 14.0
    for h, w in heights:
        acc += w
        if acc >= total / 2:
            body_line_h = h
            break
    body_size = round(body_line_h * 0.78, 1)      # 行高 → 字号的经验换算
    line_pt = round(body_line_h * 1.18, 2)

    blocks: list[dict] = []
    bounds_x0, bounds_x1, bounds_y0, bounds_y1 = [], [], [], []

    for pi, pg in enumerate(ocr.get("pages", [])):
        if pi > 0:
            blocks.append({"kind": "page_break", "text": ""})
        paras = group_lines(pg.get("lines", []), page_h)

        for p in paras:
            x0, y0, y1 = p["x0"], p["y0"], p["y1"]
            bounds_x0.append(x0); bounds_x1.append(p["x0"] + (p["y1"] - p["y0"]) * 0)
            bounds_y0.append(y0); bounds_y1.append(y1)

            ratio = p["median_h"] / body_line_h if body_line_h else 1.0
            size = round(body_size * max(min(ratio, 2.6), 0.8), 1)
            gap_after = p["_gap_after"]
            threshold = p.get("_threshold", 5.0)
            space_after = round(min(max(gap_after - 1.5, 0.0), 30.0), 1) if gap_after > threshold else 0.0

            is_bullet = p["text"][:1] in BULLET_CHARS
            text = p["text"].lstrip(BULLET_CHARS).strip() if is_bullet else p["text"]
            is_heading = (not is_bullet) and ratio >= 1.12 and len(text) <= 90

            style = {
                "size_pt": size,
                "bold": is_heading,
                "latin": "Arial",
                "east_asian": "微软雅黑",
                "line_pt": round(size * 1.35, 2) if is_heading else line_pt,
                "space_after_pt": space_after,
            }
            entry = {"kind": "bullet" if is_bullet else
                             ("heading" if is_heading else "paragraph"),
                     "level": 1 if is_heading else 0, "text": text,
                     "lines": p["lines"], "style": style}
            if is_bullet:
                entry["indent_pt"] = 18
                entry["hanging_pt"] = -18
            blocks.append(entry)

    margins = {
        "left": round(max(min(bounds_x0), 0.0), 1) if bounds_x0 else 72.0,
        "right": round(max(page_w - max(bounds_x1), 0.0), 1) if bounds_x1 else 72.0,
        "top": round(max(min(bounds_y0), 0.0), 1) if bounds_y0 else 72.0,
        "bottom": round(max(page_h - max(bounds_y1), 0.0), 1) if bounds_y1 else 72.0,
    }

    return {
        "source": pdf_path,
        "origin": "ocr-draft",      # 标记来源，提示需要人工/视觉复核
        "page": {
            "width_pt": round(page_w, 2),
            "height_pt": round(page_h, 2),
            "margins_pt": margins,
            "latin": "Arial",
            "east_asian": "微软雅黑",
            "size_pt": body_size,
        },
        "page_meta": [
            {"index": pi + 1, "line_count": len(pg.get("lines", []))}
            for pi, pg in enumerate(ocr.get("pages", []))
        ],
        "blocks": blocks,
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="OCR 结果 → 结构化规格")
    ap.add_argument("ocr_json")
    ap.add_argument("pdf")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    import json
    ocr = json.loads(Path(args.ocr_json).read_text(encoding="utf-8"))
    spec = build_spec(ocr, args.pdf)
    dump_json(spec, args.out)

    kinds = Counter(b["kind"] for b in spec["blocks"])
    m = spec["page"]["margins_pt"]
    print(f"OCR 初稿: {spec['page']['width_pt']:.0f}x{spec['page']['height_pt']:.0f}pt "
          f"页边距 L{m['left']} R{m['right']} T{m['top']} B{m['bottom']}")
    print(f"  段落={kinds.get('paragraph', 0)} 标题={kinds.get('heading', 0)} "
          f"列表={kinds.get('bullet', 0)} 分页={kinds.get('page_break', 0)}")
    print(f"  正文字号估算={spec['page']['size_pt']}pt")
    print(f"\n[spec] {args.out}")
    print("⚠ 这是 OCR 机器初稿，交付前请看图复核（OCR 会丢字/吞空格/错专业词）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
