#!/usr/bin/env python3
"""compare.py — 验收闭环（Phase 4）。

把「原始 PDF」与「重建并渲染回来的 PDF」逐页比对，输出可复核的证据：
  1. 页面尺寸一致性（决定分页是否会漂）
  2. 文本层相似度（原 PDF 有文字层时，作为内容保真的客观指标）
  3. 像素差异（版面保真的客观指标，含差异热图与并排对照图）

判定阈值保守，宁可报 warn 也不放过真实问题。

用法:
  python compare.py original.pdf result.pdf --outdir diff/
"""
from __future__ import annotations

import argparse
import sys
from difflib import SequenceMatcher
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import dump_json, load_fitz  # noqa: E402

import numpy as np  # noqa: E402
from PIL import Image  # noqa: E402


def normalize_text(s: str) -> str:
    return "".join(ch for ch in (s or "") if not ch.isspace())


def text_similarity(a: str, b: str) -> float:
    na, nb = normalize_text(a), normalize_text(b)
    if not na and not nb:
        return 1.0
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def render_gray(page, dpi: float, size_px: tuple[int, int] | None = None) -> Image.Image:
    # PyMuPDF ≥1.28 的 get_pixmap 要求 dpi 为 int，传 float 会抛 TypeError
    pix = page.get_pixmap(dpi=int(round(dpi)))
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("L")
    if size_px and (img.width, img.height) != size_px:
        img = img.resize(size_px, Image.LANCZOS)
    return img


def diff_images(a: Image.Image, b: Image.Image) -> tuple[float, float, Image.Image]:
    """返回 (平均差异 0-1, 差异像素占比, 热图)。"""
    aa = np.asarray(a, dtype=np.int16)
    bb = np.asarray(b, dtype=np.int16)
    d = np.abs(aa - bb)
    mean = float(d.mean()) / 255.0
    hot = (d > 64)  # 显著差异阈值
    ratio = float(hot.mean())
    heat = np.zeros((*d.shape, 3), dtype=np.uint8)
    heat[..., 0] = np.clip(d * 3, 0, 255)
    heat[..., 1] = np.clip(255 - d * 3, 0, 255)
    # 把「原图」「结果图」的轮廓叠进去，便于定位
    heat[..., 2] = (aa // 3).astype(np.uint8)
    return mean, ratio, Image.fromarray(heat, "RGB")


def side_by_side(a: Image.Image, b: Image.Image, heat: Image.Image,
                 label_a: str, label_b: str) -> Image.Image:
    h = max(a.height, b.height, heat.height)
    gap = 12
    canvas = Image.new("RGB", (a.width + b.width + heat.width + gap * 2, h), (255, 255, 255))
    canvas.paste(a, (0, 0))
    canvas.paste(b, (a.width + gap, 0))
    canvas.paste(heat, (a.width + b.width + gap * 2, 0))
    return canvas


def verdict_for(size_ok: bool, sim: float, pix_ratio: float) -> str:
    """判定语义（重要：两条指标回答的是不同问题）

    text_similarity  —— 内容保真：文字有没有丢、有没有错。这是硬指标。
    pixel_hot_ratio  —— 版面接近度：字体渲染、断行位置、微调间距必然带来像素差，
                        换任何引擎都不可能像素级一致，所以阈值要留出合理余量。
    """
    if not size_ok:
        return "fail"
    if sim >= 0.995 and pix_ratio <= 0.05:
        return "pass"          # 内容无损且版面几乎重合
    if sim >= 0.98 and pix_ratio <= 0.20:
        return "pass"          # 内容无损、版面接近（正常重建的预期区间）
    if sim >= 0.90 and pix_ratio <= 0.32:
        return "warn"
    return "fail"


def main() -> int:
    ap = argparse.ArgumentParser(description="原始 PDF vs 重建 PDF 逐页比对")
    ap.add_argument("original")
    ap.add_argument("result")
    ap.add_argument("--outdir", default="diff")
    ap.add_argument("--dpi", type=float, default=110)
    ap.add_argument("--pages", default=None, help="只比对指定页，如 1-3,7（用于页数不等的场景）")
    ap.add_argument("--json", default=None)
    ap.add_argument("--no-images", action="store_true", help="只出数值不出图")
    args = ap.parse_args()

    fits = load_fitz()
    a_path, b_path = Path(args.original), Path(args.result)
    for p in (a_path, b_path):
        if not p.exists():
            print(f"[错误] 不存在: {p}", file=sys.stderr)
            return 2

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    da, db = fits.open(str(a_path)), fits.open(str(b_path))
    try:
        if args.pages:
            indices: list[int] = []
            for part in args.pages.split(","):
                part = part.strip()
                if "-" in part:
                    x, y = part.split("-", 1)
                    indices.extend(range(int(x) - 1, int(y)))
                elif part:
                    indices.append(int(part) - 1)
            indices = sorted(set(i for i in indices if i >= 0))
        else:
            indices = list(range(max(da.page_count, db.page_count)))

        pages = []
        for i in indices:
            entry: dict = {"page": i + 1}
            if i >= da.page_count:
                entry.update({"verdict": "fail", "note": "结果多出的页"})
                pages.append(entry)
                continue
            if i >= db.page_count:
                entry.update({"verdict": "fail", "note": "结果缺失的页"})
                pages.append(entry)
                continue

            pa, pb = da[i], db[i]
            sa = (round(pa.rect.width, 1), round(pa.rect.height, 1))
            sb = (round(pb.rect.width, 1), round(pb.rect.height, 1))
            size_ok = abs(sa[0] - sb[0]) <= 2.0 and abs(sa[1] - sb[1]) <= 2.0

            ta = pa.get_text("text")
            tb = pb.get_text("text")
            sim = text_similarity(ta, tb)

            # 目标尺寸取原页，用于对齐后比较
            target = (int(round(pa.rect.width * args.dpi / 72)),
                      int(round(pa.rect.height * args.dpi / 72)))
            ia = render_gray(pa, args.dpi, target)
            ib = render_gray(pb, args.dpi, target)
            mean_diff, hot_ratio, heat = diff_images(ia, ib)

            entry.update({
                "size_orig_pt": list(sa),
                "size_result_pt": list(sb),
                "size_ok": size_ok,
                "text_similarity": round(sim, 4),
                "text_chars_orig": len(ta.strip()),
                "text_chars_result": len(tb.strip()),
                "pixel_mean_diff": round(mean_diff, 4),
                "pixel_hot_ratio": round(hot_ratio, 4),
                "verdict": verdict_for(size_ok, sim, hot_ratio),
            })

            if not args.no_images:
                img_path = outdir / f"page_{i + 1:03d}_compare.png"
                side_by_side(ia, ib, heat, "orig", "result").save(img_path)
                entry["compare_image"] = str(img_path)

            pages.append(entry)

        passes = sum(1 for p in pages if p.get("verdict") == "pass")
        fails = sum(1 for p in pages if p.get("verdict") == "fail")
        warns = sum(1 for p in pages if p.get("verdict") == "warn")

        result = {
            "original": str(a_path),
            "result": str(b_path),
            "page_count_original": da.page_count,
            "page_count_result": db.page_count,
            "summary": {"pass": passes, "warn": warns, "fail": fails},
            "pages": pages,
        }

        print(f"原始: {a_path.name} ({da.page_count}页)")
        print(f"结果: {b_path.name} ({db.page_count}页)")
        print(f"结论: pass={passes} warn={warns} fail={fails}")
        print("-" * 88)
        print(f"{'页':>4} {'判定':<6}{'尺寸':<8}{'文本相似':>10}{'像素均差':>10}{'差异像素':>10}  字符 原/果")
        for p in pages:
            if "verdict" not in p:
                continue
            size_flag = "OK" if p.get("size_ok") else "MISMATCH"
            print(f"{p['page']:>4} {p['verdict']:<6}{size_flag:<8}"
                  f"{p.get('text_similarity', 0):>10.3f}{p.get('pixel_mean_diff', 0):>10.3f}"
                  f"{p.get('pixel_hot_ratio', 0):>10.1%}"
                  f"  {p.get('text_chars_orig', 0):>5}/{p.get('text_chars_result', 0):<5}")
        print("-" * 88)
        print(f"对照图目录: {outdir}")

        if args.json:
            dump_json(result, args.json)
            print(f"[JSON] {args.json}")
        elif (outdir / "compare.json").parent.exists():
            dump_json(result, outdir / "compare.json")

        return 0 if fails == 0 else 1
    finally:
        da.close()
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
