#!/usr/bin/env python3
"""probe.py — PDF 分诊（Phase 0）。

只读。回答三个问题：
  1. 这份 PDF 属于哪一类（文字型 / 图片型 / 扫描型 / 混合）
  2. 每页的版式指纹（页面尺寸、字体字号色值分布、图片与矢量块的位置）
  3. 后续该走哪条重建路线

用法:
  python probe.py <input.pdf> [--json out.json] [--pages 1-3]
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

# Windows 控制台默认 GBK，中文报告会乱码；强制 UTF-8 输出
for _stream in ("stdout", "stderr"):
    try:
        getattr(sys, _stream).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import (  # noqa: E402
    dump_json,
    int_to_hex_rgb,
    map_font,
    round_half,
)

fitz = None


def _fitz():
    global fitz
    if fitz is None:
        from _kit import load_fitz
        fitz = load_fitz()
    return fitz


# ------------------------------------------------------------------ 分诊

SCAN_TEXT_CHARS = 80        # 每页低于此字符数 + 有大图 → 疑似图片型
IMAGE_AREA_RATIO = 0.25     # 图片覆盖超过此比例 → 图片主导


def classify_page(chars: int, span_count: int, image_area_ratio: float,
                  drawing_count: int) -> str:
    if chars <= SCAN_TEXT_CHARS and image_area_ratio >= IMAGE_AREA_RATIO:
        return "image"          # 图片型：几乎无文字层，靠图承载内容
    if chars <= SCAN_TEXT_CHARS and drawing_count > 20:
        return "vector"         # 矢量型：文字被转曲线，靠矢量绘制承载
    if chars <= SCAN_TEXT_CHARS:
        return "scanned"        # 扫描型：整页位图
    if image_area_ratio >= IMAGE_AREA_RATIO:
        return "mixed"          # 混合：有文字层也有大图
    return "text"               # 文字型：正常可抽取


def parse_pages(spec: str | None, total: int) -> list[int]:
    """'1-3,7' → [0,1,2,6]（对外 1-based，对内 0-based）"""
    if not spec:
        return list(range(total))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start, end = int(a), int(b)
        else:
            start = end = int(part)
        for n in range(start, end + 1):
            if 1 <= n <= total:
                out.append(n - 1)
    return sorted(set(out))


# ------------------------------------------------------------------ 单页分析

def analyze_page(page, index: int) -> dict:
    fits = _fitz()
    rect = page.rect
    page_area = max(rect.width * rect.height, 1e-6)

    # 文字层
    raw = page.get_text("dict")
    spans: list[dict] = []
    for block in raw.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                txt = span.get("text", "")
                if not txt.strip():
                    continue
                spans.append(span)

    text = page.get_text("text")
    chars = len(text.strip())

    font_counter: Counter[str] = Counter()
    size_counter: Counter[float] = Counter()
    color_counter: Counter[str] = Counter()
    for sp in spans:
        latin, ea = map_font(sp.get("font", ""))
        font_counter[f"{sp.get('font','?')} -> {latin}/{ea}"] += 1
        size_counter[round_half(float(sp.get("size", 0)))] += 1
        color_counter[int_to_hex_rgb(sp.get("color", 0))] += 1

    # 图片：用 get_image_info 拿实际绘制位置
    images = []
    image_area = 0.0
    try:
        for info in page.get_image_info(xrefs=True):
            bbox = [float(v) for v in info.get("bbox", (0, 0, 0, 0))]
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
            area = max(w * h, 0.0)
            # 裁剪到页内再算占比，避免越界图虚高
            cw = min(bbox[2], rect.x1) - max(bbox[0], rect.x0)
            ch = min(bbox[3], rect.y1) - max(bbox[1], rect.y0)
            clipped = max(cw, 0) * max(ch, 0)
            image_area += clipped
            images.append({
                "xref": info.get("xref"),
                "bbox": [round(v, 2) for v in bbox],
                "size_px": [info.get("width"), info.get("height")],
                "area_pt2": round(area, 1),
            })
    except Exception:
        pass
    image_area_ratio = min(image_area / page_area, 1.0)

    # 矢量绘制：色块与线条
    draws = []
    fill_counter: Counter[str] = Counter()
    try:
        for item in page.get_drawings():
            r = item.get("rect")
            if r is None:
                continue
            fill = item.get("fill")
            stroke = item.get("color")
            hexfill = None
            if fill is not None:
                hexfill = "#{:02X}{:02X}{:02X}".format(
                    *[int(round(c * 255)) for c in fill[:3]]
                )
                fill_counter[hexfill] += 1
            draws.append({
                "rect": [round(float(v), 1) for v in (r.x0, r.y0, r.x1, r.y1)],
                "fill": hexfill,
                "stroke": None if stroke is None else "#{:02X}{:02X}{:02X}".format(
                    *[int(round(c * 255)) for c in stroke[:3]]),
                "width": round(float(r.width) * float(r.height), 1),
            })
    except Exception:
        pass

    kind = classify_page(chars, len(spans), image_area_ratio, len(draws))

    return {
        "index": index + 1,
        "size_pt": [round(rect.width, 2), round(rect.height, 2)],
        "kind": kind,
        "text_chars": chars,
        "span_count": len(spans),
        "image_count": len(images),
        "image_area_ratio": round(image_area_ratio, 4),
        "drawing_count": len(draws),
        "fonts": [{"raw": k.split(" -> ")[0], "mapped": k.split(" -> ")[1], "n": v}
                  for k, v in font_counter.most_common(8)],
        "sizes": [{"pt": k, "n": v} for k, v in sorted(size_counter.items(), key=lambda x: -x[1])[:8]],
        "text_colors": [{"hex": k, "n": v} for k, v in color_counter.most_common(6)],
        "fill_colors": [{"hex": k, "n": v} for k, v in fill_counter.most_common(8)],
        "images": images,
        "text_head": text.strip()[:400],
    }


# ------------------------------------------------------------------ 主流程

def suggest_route(page_kinds: Counter) -> dict:
    dominant = page_kinds.most_common(1)[0][0] if page_kinds else "text"
    routes = {
        "text": {
            "route": "A/text",
            "why": "有完整文字层，几何信息可直接从 PDF 精确读取，不需要 OCR。",
            "steps": ["extract 抽 span/表格/图片", "build_docx 重建", "compare 验收"],
        },
        "mixed": {
            "route": "B/mixed",
            "why": "文字层可用，但大图承载了主要视觉，需保留原图并还原版面。",
            "steps": ["extract 抽文字+图片bbox", "vision 判读图文关系", "build_docx 图文混排", "compare"],
        },
        "image": {
            "route": "C/image",
            "why": "几乎无文字层，文字以像素存在，必须走 OCR + 视觉判读双通道。",
            "steps": ["vision 生成页图与区块", "ocr 拿行级文字与坐标", "合并视觉结构 → blocks.json",
                      "build_docx 重建为真文字/真表格", "compare 逐页比对迭代"],
        },
        "scanned": {
            "route": "C/scanned",
            "why": "整页位图，需 OCR；版面噪声大，需先做去噪与倾斜校正。",
            "steps": ["vision 校正+切块", "ocr", "人工/视觉复核", "build_docx", "compare"],
        },
        "vector": {
            "route": "D/vector",
            "why": "文字被转成曲线，无文字层但矢量几何精确，可由 drawings 反推色块与线条。",
            "steps": ["vision 用 get_drawings 反推色块", "ocr 补文字", "build_docx 重绘形状+文字", "compare"],
        },
    }
    return routes.get(dominant, routes["text"])


def main() -> int:
    ap = argparse.ArgumentParser(description="PDF 分诊与版式指纹")
    ap.add_argument("pdf")
    ap.add_argument("--json", default=None, help="输出 JSON 路径")
    ap.add_argument("--pages", default=None, help="限定页范围，如 1-3,7")
    args = ap.parse_args()

    fits = _fitz()
    path = Path(args.pdf)
    if not path.exists():
        print(f"[错误] 文件不存在: {path}", file=sys.stderr)
        return 2

    doc = fits.open(str(path))
    try:
        targets = parse_pages(args.pages, doc.page_count)
        reports = []
        for i in targets:
            reports.append(analyze_page(doc[i], i))

        kinds = Counter(r["kind"] for r in reports)
        route = suggest_route(kinds)

        result = {
            "file": str(path),
            "page_count": doc.page_count,
            "analyzed_pages": [i + 1 for i in targets],
            "kind_summary": dict(kinds),
            "route": route,
            "pages": reports,
        }

        # 控制台简报
        print(f"文件      : {path.name}")
        print(f"总页数    : {doc.page_count}")
        print(f"类型分布  : {dict(kinds)}")
        print(f"建议路线  : {route['route']}")
        print(f"理由      : {route['why']}")
        print("-" * 72)
        print(f"{'页':>4} {'类型':<9}{'字符':>7}{'span':>6}{'图':>4}{'图占比':>9}{'矢量':>6}  尺寸(pt)")
        for r in reports:
            print(f"{r['index']:>4} {r['kind']:<9}{r['text_chars']:>7}{r['span_count']:>6}"
                  f"{r['image_count']:>4}{r['image_area_ratio']:>9.2%}{r['drawing_count']:>6}"
                  f"  {r['size_pt'][0]:.0f}x{r['size_pt'][1]:.0f}")
        print("-" * 72)
        for r in reports[:3]:
            if r["sizes"]:
                print(f"p{r['index']} 字号: " + ", ".join(
                    f"{s['pt']}pt×{s['n']}" for s in r["sizes"][:6]))
            if r["fill_colors"]:
                print(f"p{r['index']} 色块: " + ", ".join(
                    f"{c['hex']}×{c['n']}" for c in r["fill_colors"][:6]))

        if args.json:
            dump_json(result, args.json)
            print(f"\n[JSON] {args.json}")
        return 0
    finally:
        doc.close()


if __name__ == "__main__":
    raise SystemExit(main())
