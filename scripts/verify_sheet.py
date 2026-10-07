#!/usr/bin/env python3
"""verify_sheet.py — 抽检拼版：把需要人眼确认的行带拼成 1~2 张长图。

为什么需要它
------------
图片型 / 扫描件 PDF 没有文字层，专业名词与拼写只能靠看图核对。逐行裁剪、逐行
read_image 是最慢的做法——实测一次 5 页公文因此用了约 35 次看图调用，其中光确认
一个词（Liquidating vs Liqudating）就裁了 6 次。把待确认行带拼成一张长图，一次
即可看完，看图调用降到 3 次以内。

用法
----
    # 1) 手工给点：--at 页码:y0:y1[:标签]   （y 为该 PDF 页内的 pt）
    python verify_sheet.py in.pdf --at "1:93:112:P1 caption" --at 3:679:695

    # 2) 批量给点：JSON（同 --at 的字段）
    python verify_sheet.py in.pdf --spec regions.json --outdir sh/

    # 3) 自动：对指定页做墨迹行带检测，把整页所有行都拼进去
    python verify_sheet.py in.pdf --auto 1,2 --outdir sh/

产出
----
    sh/sheet_001.png …   拼版图（每格左侧有 #序号 标尺）
    sh/verify_map.json   序号 → (页码, y0, y1, 标签)，看图时对照

regions.json 格式（两种都收）::

    {"items": [{"page": 1, "y0": 93, "y1": 112, "label": "P1 caption"}]}
    [[1, 93, 112, "P1 caption"], [3, 679, 695, ""]]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

GAUGE_W = 78          # 左侧标尺宽度（px）
PAD = 6               # 格间距（px）
_INK_LUT = [255 if i < 170 else 0 for i in range(256)]   # 灰度 → 墨迹掩码


# ------------------------------------------------------------------ 字体
def _font(size: int):
    """找一个能画 ASCII 序号的字体；找不到就退回 PIL 位图字体。"""
    for name in ("arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "segoeui.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None


# ------------------------------------------------------------------ 行带检测
def detect_bands(page: fitz.Page, dpi: int = 150, thr: int = 3,
                 min_h_px: int = 4) -> list[tuple[float, float]]:
    """按墨迹行剖面切出行带，返回 [(y0_pt, y1_pt), …]。

    只用 Pillow 实现（二值化 → 缩到宽 1 取行均值），不引入 numpy——
    numpy 在本技能里属可选依赖，核心路径不该被它卡住。
    """
    pm = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    img = Image.frombytes("L", (pm.width, pm.height), pm.samples)
    bw = img.point(_INK_LUT, "L")                       # 255 = 墨迹
    prof = bw.resize((1, img.height), Image.BOX)        # 每行墨迹占比 × 255
    # 还原成「该行的墨迹像素数」，这样 thr 的语义与按像素计数时一致。
    # 用 tobytes() 取行值：prof 宽 1，正好一行一个字节，且不触发 Pillow 的
    # getdata() 弃用告警。
    rows = [b * pm.width / 255.0 for b in prof.tobytes()]
    scale = 72.0 / dpi
    bands: list[tuple[float, float]] = []
    start = None
    for i, v in enumerate(rows):
        if v > thr and start is None:
            start = i
        elif v <= thr and start is not None:
            if i - start >= min_h_px:
                bands.append((start * scale, i * scale))
            start = None
    if start is not None:
        bands.append((start * scale, len(rows) * scale))
    return bands


# ------------------------------------------------------------------ 参数解析
def parse_at(spec: str) -> dict:
    """'页码:y0:y1[:标签]' → dict。标签里允许出现 ':'。"""
    parts = spec.split(":")
    if len(parts) < 3:
        raise ValueError(f"--at 需要 页码:y0:y1[:标签]，收到 {spec!r}")
    return {"page": int(parts[0]), "y0": float(parts[1]), "y1": float(parts[2]),
            "label": ":".join(parts[3:]).strip()}


def load_spec(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    rows = data.get("items", []) if isinstance(data, dict) else data
    out = []
    for r in rows:
        if isinstance(r, dict):
            out.append({"page": int(r["page"]), "y0": float(r["y0"]),
                        "y1": float(r["y1"]), "label": str(r.get("label", ""))})
        else:
            out.append({"page": int(r[0]), "y0": float(r[1]), "y1": float(r[2]),
                        "label": str(r[3]) if len(r) > 3 else ""})
    return out


def parse_pages(text: str, total: int) -> list[int]:
    if text.strip().lower() in ("all", "*"):
        return list(range(1, total + 1))
    out: list[int] = []
    for chunk in text.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            out += list(range(int(a), int(b) + 1))
        else:
            out.append(int(chunk))
    return out


# ------------------------------------------------------------------ 拼版
def render_tile(doc: fitz.Document, item: dict, dpi: int, zoom: float,
                max_w: int) -> Image.Image:
    page = doc[item["page"] - 1]
    y0 = max(0.0, min(item["y0"], page.rect.height))
    y1 = max(y0 + 1.0, min(item["y1"], page.rect.height))
    clip = fitz.Rect(0.0, y0, page.rect.width, y1)
    pm = page.get_pixmap(dpi=int(dpi), clip=clip)
    img = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
    scale = zoom * min(1.0, max_w / img.width)
    if abs(scale - 1.0) > 0.01:
        img = img.resize((max(1, int(img.width * scale)),
                          max(1, int(img.height * scale))), Image.LANCZOS)
    return img


def build_sheets(tiles: list[tuple[int, Image.Image, str]], outdir: Path,
                 max_side: int, gauge_w: int = GAUGE_W) -> list[Path]:
    font = _font(30)
    sheets: list[Path] = []
    idx = 0
    while idx < len(tiles):
        rows: list[tuple[int, Image.Image, str]] = []
        height = PAD
        while idx < len(tiles) and height < max_side:
            _, img, _ = tiles[idx]
            if rows and height + img.height + PAD > max_side:
                break
            rows.append(tiles[idx])
            height += img.height + PAD
            idx += 1
        width = gauge_w + max((im.width for _, im, _ in rows), default=1) + PAD * 2
        canvas = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(canvas)
        y = PAD
        for n, im, _ in rows:
            canvas.paste(im, (gauge_w + PAD, y))
            draw.line([(0, y - PAD // 2), (width, y - PAD // 2)], fill=(200, 200, 200))
            if font is not None:
                draw.text((6, y + 2), f"#{n:02d}", fill="black", font=font)
            y += im.height + PAD
        out = outdir / f"sheet_{len(sheets) + 1:03d}.png"
        canvas.save(out)
        sheets.append(out)
    return sheets


# ------------------------------------------------------------------ 主流程
def main() -> int:
    ap = argparse.ArgumentParser(description="抽检拼版：待确认行带 → 一张长图")
    ap.add_argument("pdf")
    ap.add_argument("--spec", help="regions.json（见文件头格式说明）")
    ap.add_argument("--at", action="append", default=[],
                    help="页码:y0:y1[:标签]，可重复（y 为页内 pt）")
    ap.add_argument("--auto", help="自动检测行带：如 1,2 或 1-5 或 all")
    ap.add_argument("--outdir", default=None, help="默认 <pdf同级>/_verify")
    ap.add_argument("--dpi", type=int, default=700, help="裁剪渲染 dpi（默认 700）")
    ap.add_argument("--zoom", type=float, default=1.0, help="裁剪后再放大倍数")
    ap.add_argument("--max-tile-w", type=int, default=2600,
                    help="单格最大宽度 px，超了等比缩（默认 2600）")
    ap.add_argument("--max-side", type=int, default=3800,
                    help="单张拼版最大高度 px，超了分张（默认 3800）")
    ap.add_argument("--pad", type=float, default=3.0, help="行带上下留白 pt")
    args = ap.parse_args()

    pdf = Path(args.pdf)
    if not pdf.exists():
        print(f"[错误] 不存在: {pdf}", file=sys.stderr)
        return 2
    outdir = Path(args.outdir) if args.outdir else pdf.parent / "_verify"
    outdir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(pdf))
    items: list[dict] = []
    if args.spec:
        items += load_spec(Path(args.spec))
    for s in args.at:
        items.append(parse_at(s))
    if args.auto:
        for pno in parse_pages(args.auto, doc.page_count):
            if not (1 <= pno <= doc.page_count):
                continue
            page = doc[pno - 1]
            for i, (y0, y1) in enumerate(detect_bands(page)):
                items.append({"page": pno, "y0": y0 - args.pad, "y1": y1 + args.pad,
                              "label": f"p{pno} band {i}"})
    if not items:
        print("[错误] 没有给任何区域：用 --at / --spec / --auto 之一", file=sys.stderr)
        doc.close()
        return 2

    tiles: list[tuple[int, Image.Image, str]] = []
    mapping: list[dict] = []
    for n, it in enumerate(items, 1):
        img = render_tile(doc, it, args.dpi, args.zoom, args.max_tile_w)
        tiles.append((n, img, it.get("label", "")))
        mapping.append({"n": n, "page": it["page"], "y0": round(it["y0"], 1),
                        "y1": round(it["y1"], 1), "label": it.get("label", ""),
                        "tile_px": [img.width, img.height]})
    doc.close()

    sheets = build_sheets(tiles, outdir, args.max_side)
    (outdir / "verify_map.json").write_text(
        json.dumps({"pdf": str(pdf), "sheets": [str(s) for s in sheets],
                    "items": mapping}, ensure_ascii=False, indent=2),
        encoding="utf-8")

    print(f"拼版 {len(sheets)} 张，共 {len(items)} 格 → {outdir}")
    for s in sheets:
        print(f"  {s}")
    print("  对照表: verify_map.json")
    for m in mapping:
        print(f"  #{m['n']:02d}  p{m['page']}  y {m['y0']:6.1f}-{m['y1']:6.1f}  "
              f"{m['tile_px'][1]}px  {m['label']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())