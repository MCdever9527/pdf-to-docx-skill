#!/usr/bin/env python3
"""extract.py — 文字型 / 混合型 PDF 的精确结构抽取（Phase 1-A）。

从 PDF 文字层直接读取几何事实，精度高于任何图像分析：
  span(text/font/size/color/bbox/flags) → 行 → 段落 → 文档块
  + 真表格（矢量线检测） + 图片（按实际绘制位置裁切）

本文件里三个决定「格式保不保得住」的关键判断，全部由实测数据驱动，
不使用拍脑袋的固定倍数：
  1. 段落切分阈值 = 段内行距中位数 × 2.2（不是固定倍数）
  2. 行距 line_pt = 段内相邻行 y0 差值的中位数（不是 行高 × 常数）
  3. 项目符号：PDF 常把「•」单独放一行，需与后续文字行合并成悬挂缩进列表项

用法:
  python extract.py in.pdf --out blocks.json [--pages 1-3] [--images ./imgs]
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
from _kit import (  # noqa: E402
    clean_text,
    dump_json,
    has_cjk,
    int_to_hex_rgb,
    load_fitz,
    map_font,
    round_half,
)

BULLET_CHARS = "•·▪‣◦∙*—-–"
FONT_SIZE_EPS = 0.7


# ------------------------------------------------------------------ 行构建

def spans_to_lines(page) -> list[dict]:
    """把 span 聚成行：同一基线的 span 合成一行。"""
    raw = page.get_text("dict")
    lines: list[dict] = []
    for block in raw.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = [s for s in line.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            text = clean_text("".join(s["text"] for s in spans))
            if not text.strip():
                continue
            x0 = min(s["bbox"][0] for s in spans)
            y0 = min(s["bbox"][1] for s in spans)
            x1 = max(s["bbox"][2] for s in spans)
            y1 = max(s["bbox"][3] for s in spans)
            main = max(spans, key=lambda s: len(s["text"]))
            sizes = [round_half(float(s["size"])) for s in spans]
            flags = [int(s.get("flags", 0)) for s in spans]
            bold = any(f & 16 for f in flags) or "bold" in main.get("font", "").lower()
            italic = (any(f & 2 for f in flags)
                      or "italic" in main.get("font", "").lower()
                      or "oblique" in main.get("font", "").lower())
            color = int_to_hex_rgb(main.get("color", 0))
            latin, ea = map_font(main.get("font", ""))
            lines.append({
                "text": text.strip(),
                "bbox": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                "size_pt": max(sizes),
                "bold": bold,
                "italic": italic,
                "color": color,
                "font_raw": main.get("font", ""),
                "latin": latin,
                "east_asian": ea,
                "height_pt": round(y1 - y0, 2),
                "synthetic_bullet": False,
            })
    lines.sort(key=lambda l: (round(l["bbox"][1], 1), l["bbox"][0]))
    return lines


def is_symbol_only(line: dict) -> str | None:
    """整行只有一个项目符号（PDF 里符号常独立成行）→ 返回该符号。"""
    t = line["text"].strip()
    if len(t) <= 2 and t and all(c in BULLET_CHARS for c in t):
        return t
    return None


def strip_inline_bullet(line: dict) -> tuple[bool, str]:
    """行首带「• 文字」形式 → 返回 (True, 去符号文字)。"""
    t = line["text"]
    if len(t) > 1 and t[0] in BULLET_CHARS and t[1] in " \u00a0\t":
        return True, t[2:].strip()
    return False, t


def coalesce_bullets(lines: list[dict]) -> list[dict]:
    """把「独立符号行 + 紧跟的文字行」合并为一个列表项。

    判定依据是几何关系而不是猜测：符号行 x0 < 文字行 x0，
    两者差值就是悬挂缩进的宽度。
    """
    out: list[dict] = []
    i = 0
    while i < len(lines):
        ln = lines[i]
        sym = is_symbol_only(ln)
        if sym and i + 1 < len(lines):
            nxt = lines[i + 1]
            indent = nxt["bbox"][0] - ln["bbox"][0]
            # 文字行缩进更深，且垂直位置基本对齐 → 符号属于这一项
            if indent > 1.5 and abs(nxt["bbox"][1] - ln["bbox"][1]) < ln["height_pt"] * 1.2:
                merged = dict(nxt)
                merged["synthetic_bullet"] = True
                merged["bullet_symbol"] = sym
                merged["hanging_pt"] = round(indent, 1)
                merged["list_left_pt"] = round(nxt["bbox"][0], 1)
                out.append(merged)
                i += 2
                continue
        out.append(ln)
        i += 1
    return out


# ------------------------------------------------------------------ 版面节奏

def analyze_rhythm(lines: list[dict]) -> dict:
    """从实测行距分布推出：段内行距基准、段落切分阈值、正文行距。

    关键：段内行距一定比段间距小，所以「段内基准」要取行距分布的**小分位**
    （p25），不能取中位数。段落密集的页面段间距可能占多数，
    此时中位数会落在段间距上，行距被高估近一倍，重建后内容直接溢出。
    """
    gaps: list[float] = []
    for i in range(1, len(lines)):
        g = lines[i]["bbox"][1] - lines[i - 1]["bbox"][3]
        if g > 0:
            gaps.append(round(g, 2))

    if not gaps:
        return {"ref_gap": 2.0, "para_threshold": 5.0, "line_pt": 14.0}

    gs = sorted(gaps)
    ref = gs[min(int(len(gs) * 0.25), len(gs) - 1)]      # 低分位 = 段内行距基准
    threshold = max(ref * 2.5, ref + 2.5)                # 超过它就算段间隔

    tight = []
    for i in range(1, len(lines)):
        dy = lines[i]["bbox"][1] - lines[i - 1]["bbox"][1]
        gap = lines[i]["bbox"][1] - lines[i - 1]["bbox"][3]
        if 0 < gap <= threshold and 0 < dy:
            tight.append(round(dy, 2))
    line_pt = sorted(tight)[len(tight) // 2] if tight else round(ref + 12.0, 2)

    return {
        "ref_gap": round(ref, 2),
        "para_threshold": round(threshold, 2),
        "line_pt": round(line_pt, 2),
        "sample_gaps": len(gs),
    }


# ------------------------------------------------------------------ 段落合并

def merge_paragraphs(lines: list[dict], body_size: float, rhythm: dict) -> list[dict]:
    paragraphs: list[dict] = []
    cur: dict | None = None
    pthr = rhythm["para_threshold"]

    for ln in lines:
        inline_bullet, text = strip_inline_bullet(ln)
        is_bullet = inline_bullet or ln.get("synthetic_bullet", False)
        if ln.get("synthetic_bullet"):
            text = ln["text"]
        is_heading = (ln["size_pt"] >= body_size * 1.12
                      and len(text) <= 120 and not is_bullet)

        if cur is None:
            cur = _new_para(ln, text, is_bullet, is_heading)
            continue

        gap = ln["bbox"][1] - cur["_last_bottom"]
        same_size = abs(ln["size_pt"] - cur["size_pt"]) <= FONT_SIZE_EPS
        same_style = (ln["bold"] == cur["bold"]
                      and ln["color"] == cur["color"]
                      and ln["latin"] == cur["latin"])
        same_indent = abs(ln["bbox"][0] - cur["_indent"]) <= 3.0

        # 同一个列表项内部续行：也要求缩进一致
        continues = (gap <= pthr and same_size and same_style
                     and same_indent and not is_bullet and not is_heading
                     and cur["kind"] != "heading")

        if continues:
            joiner = "" if (has_cjk(cur["text"][-1:]) or has_cjk(text[:1])) else " "
            cur["text"] = (cur["text"] + joiner + text).strip()
            cur["_last_bottom"] = ln["bbox"][3]
            cur["lines"] += 1
            cur["bbox"] = [
                min(cur["bbox"][0], ln["bbox"][0]),
                min(cur["bbox"][1], ln["bbox"][1]),
                max(cur["bbox"][2], ln["bbox"][2]),
                max(cur["bbox"][3], ln["bbox"][3]),
            ]
            continue

        paragraphs.append(cur)
        cur = _new_para(ln, text, is_bullet, is_heading)

    if cur is not None:
        paragraphs.append(cur)
    return paragraphs


def _new_para(ln: dict, text: str, is_bullet: bool, is_heading: bool) -> dict:
    return {
        "kind": "bullet" if is_bullet else ("heading" if is_heading else "paragraph"),
        "text": text,
        "bbox": list(ln["bbox"]),
        "size_pt": ln["size_pt"],
        "bold": ln["bold"],
        "italic": ln["italic"],
        "color": ln["color"],
        "latin": ln["latin"],
        "east_asian": ln["east_asian"],
        "hanging_pt": ln.get("hanging_pt"),
        "list_left_pt": ln.get("list_left_pt"),
        "lines": 1,                      # 该段占几行，供 build 做高度预算
        "_last_bottom": ln["bbox"][3],
        "_indent": ln["bbox"][0],
    }


def finalize(paragraphs: list[dict], rhythm: dict, left_margin: float) -> list[dict]:
    """清理内部字段，补出给 Word 用的排版参数。

    段后间距的归属很容易搞反：两段之间的空隙，语义上属于「前一段的段后」。
    若误记到后一段上，标题后面就会贴住正文，整篇随之下移，分页跟着漂。
    """
    items: list[dict] = []

    for p in paragraphs:
        last_bottom = p.pop("_last_bottom", None)
        p.pop("_indent", None)
        bbox = p.pop("bbox", None)
        y_top = bbox[1] if bbox else 0.0
        y_bottom = bbox[3] if bbox else 0.0
        x0 = bbox[0] if bbox else None

        style = {
            "size_pt": p["size_pt"],
            "bold": p["bold"],
            "italic": p["italic"],
            "color": p["color"],
            "latin": p["latin"],
            "east_asian": p["east_asian"],
            "line_pt": rhythm["line_pt"],       # 实测正文行距，避免分页漂移
            "space_after_pt": 0.0,              # 稍后按实测空隙回填
        }

        entry = {"kind": p["kind"], "text": p["text"], "style": style,
                 "lines": p.get("lines", 1),
                 "_y_top": y_top, "_y_bottom": y_bottom,
                 "_y": y_top}          # 排序键，由 main 用完后移除

        if p["kind"] == "bullet":
            hanging = p.get("hanging_pt") or 18.0
            list_left = p.get("list_left_pt")
            left = (round(list_left - left_margin, 1)
                    if list_left and list_left > left_margin else 18.0)
            entry["indent_pt"] = max(left, 0.0)
            entry["hanging_pt"] = -abs(hanging)
        elif p["kind"] == "heading":
            entry["level"] = 1
            if x0 is not None and x0 > left_margin + 6:
                style["left_indent_pt"] = round(x0 - left_margin, 1)

        items.append(entry)

    # 回填：第 i 段与第 i+1 段之间的空隙 → 第 i 段的段后
    pthr = rhythm["para_threshold"]
    for i in range(len(items) - 1):
        gap = items[i + 1]["_y_top"] - items[i]["_y_bottom"]
        if gap > pthr:
            items[i]["style"]["space_after_pt"] = round(
                min(gap - rhythm["ref_gap"], 36.0), 1)

    for it in items:
        it.pop("_y_top", None)
        it.pop("_y_bottom", None)
    return items


# ------------------------------------------------------------------ 表格与图片

def extract_tables(page) -> list[dict]:
    tables: list[dict] = []
    try:
        finder = page.find_tables()
    except Exception:
        return tables
    for t in getattr(finder, "tables", []):
        try:
            data = t.extract()
        except Exception:
            continue
        if not data:
            continue
        rows = [[(c or "").strip().replace("\n", " ") for c in row] for row in data]
        if not any(any(c for c in r) for r in rows):
            continue
        tables.append({
            "kind": "table",
            "rows": rows,
            "header_row": True,
            "bbox": [round(float(v), 1) for v in t.bbox],
            "row_count": len(rows),
        })
    return tables


def extract_images(page, index: int, outdir: Path | None) -> list[dict]:
    items: list[dict] = []
    try:
        infos = page.get_image_info(xrefs=True)
    except Exception:
        return items
    for n, info in enumerate(infos):
        bbox = [float(v) for v in info.get("bbox", (0, 0, 0, 0))]
        entry = {
            "kind": "image",
            "bbox": [round(v, 1) for v in bbox],
            "width_pt": round(bbox[2] - bbox[0], 1),
            "height_pt": round(bbox[3] - bbox[1], 1),
            "source_size_px": [info.get("width"), info.get("height")],
            "xref": info.get("xref"),
        }
        if outdir is not None:
            try:
                pix = page.get_pixmap(clip=load_fitz().Rect(bbox), dpi=300)
                f = outdir / f"p{index + 1:03d}_img{n + 1:02d}.png"
                pix.save(str(f))
                entry["path"] = str(f)
            except Exception:
                pass
        items.append(entry)
    return items


# ------------------------------------------------------------------ 版式推断

def content_bounds(lines: list[dict], page) -> dict:
    """单页的内容边界（pt）。"""
    if not lines:
        return {"x0": 0.0, "x1": 0.0, "y0": 0.0, "y1": 0.0,
                "w": round(page.rect.width, 2), "h": round(page.rect.height, 2)}
    return {
        "x0": min(l["bbox"][0] for l in lines),
        "x1": max(l["bbox"][2] for l in lines),
        "y0": min(l["bbox"][1] for l in lines),
        "y1": max(l["bbox"][3] for l in lines),
        "w": round(page.rect.width, 2),
        "h": round(page.rect.height, 2),
    }


def infer_margins(bounds: list[dict]) -> dict:
    """跨页聚合推断页边距。

    关键：不能用「某一页最后一行的位置」当底边距。内容少的页天然排不满，
    那反映的是内容长度，不是页边距。正确做法是取所有页的极值：
        上边距 = min(所有页内容顶)
        下边距 = 页高 - max(所有页内容底)
        左/右边距 = 内容最左 / 页宽 - 内容最右
    否则底边距会被高估，可用高度被压缩，内容随即溢出到下一页。
    """
    if not bounds:
        return {"left": 72.0, "right": 72.0, "top": 72.0, "bottom": 72.0}
    usable = [b for b in bounds if b["x1"] > b["x0"] or b["y1"] > b["y0"]]
    if not usable:
        usable = bounds
    w = usable[0]["w"]
    h = max(b["h"] for b in usable)
    return {
        "left": round(max(min(b["x0"] for b in usable), 0.0), 1),
        "right": round(max(w - max(b["x1"] for b in usable), 0.0), 1),
        "top": round(max(min(b["y0"] for b in usable), 0.0), 1),
        "bottom": round(max(h - max(b["y1"] for b in usable), 0.0), 1),
    }


def body_size_of(lines: list[dict]) -> float:
    """正文基准字号：按字符数加权的中位字号。"""
    if not lines:
        return 10.5
    weighted = sorted(((l["size_pt"], max(len(l["text"]), 1)) for l in lines),
                      key=lambda x: x[0])
    total = sum(w for _, w in weighted)
    acc = 0
    for size, w in weighted:
        acc += w
        if acc >= total / 2:
            return size
    return weighted[-1][0]


# ------------------------------------------------------------------ 主流程

def parse_page_spec(spec: str | None, total: int) -> list[int]:
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
    ap = argparse.ArgumentParser(description="PDF 结构抽取 → blocks.json")
    ap.add_argument("pdf")
    ap.add_argument("--out", required=True)
    ap.add_argument("--pages", default=None)
    ap.add_argument("--images", default=None, help="图片导出目录（不给则不导出）")
    ap.add_argument("--verbose", action="store_true", help="打印版面节奏参数")
    args = ap.parse_args()

    fits = load_fitz()
    src = Path(args.pdf)
    if not src.exists():
        print(f"[错误] 不存在: {src}", file=sys.stderr)
        return 2

    img_dir = Path(args.images) if args.images else None
    if img_dir:
        img_dir.mkdir(parents=True, exist_ok=True)

    doc = fits.open(str(src))
    try:
        indices = parse_page_spec(args.pages, doc.page_count)

        # 第一遍：收集每页的几何与节奏事实
        per_page: list[dict] = []
        for i in indices:
            page = doc[i]
            lines = coalesce_bullets(spans_to_lines(page))
            per_page.append({
                "index": i,
                "page": page,
                "lines": lines,
                "bounds": content_bounds(lines, page),
                "body": body_size_of(lines),
                "rhythm": analyze_rhythm(lines),
            })

        # 页边距必须跨页聚合后再用（见 infer_margins 注释）
        margins = infer_margins([p["bounds"] for p in per_page])

        all_blocks: list[dict] = []
        page_meta: list[dict] = []

        for entry in per_page:
            i = entry["index"]
            page = entry["page"]
            lines = entry["lines"]
            body = entry["body"]
            rhythm = entry["rhythm"]

            paras = merge_paragraphs(lines, body, rhythm)
            blocks = finalize(paras, rhythm, margins["left"])
            tables = extract_tables(page)
            images = extract_images(page, i, img_dir)

            flow = [{"_y": b.get("_y", 0.0), "b": b} for b in blocks]
            for t in tables:
                flow.append({"_y": t["bbox"][1], "b": t})
            for im in images:
                flow.append({"_y": im["bbox"][1], "b": im})
            flow.sort(key=lambda x: x["_y"])
            page_blocks = [f["b"] for f in flow]
            for b in page_blocks:
                b.pop("_y", None)

            if i != indices[0]:
                page_blocks.insert(0, {"kind": "page_break", "text": ""})

            all_blocks.extend(page_blocks)
            kinds: dict[str, int] = {}
            for b in blocks:
                kinds[b["kind"]] = kinds.get(b["kind"], 0) + 1
            page_meta.append({
                "index": i + 1,
                "size_pt": [round(page.rect.width, 2), round(page.rect.height, 2)],
                "body_size_pt": body,
                "content_h_pt": round(entry["bounds"]["y1"] - entry["bounds"]["y0"], 1),
                "rhythm": rhythm,
                "line_count": len(lines),
                "kind_counts": kinds,
                "table_count": len(tables),
                "image_count": len(images),
            })

        first_page = per_page[0]["page"]
        first = page_meta[0]
        result = {
            "source": str(src),
            "page": {
                "width_pt": round(first_page.rect.width, 2),
                "height_pt": round(first_page.rect.height, 2),
                "margins_pt": margins,
                "latin": "Arial",
                "east_asian": "微软雅黑",
                "size_pt": first["body_size_pt"],
            },
            "page_meta": page_meta,
            "blocks": all_blocks,
        }
        dump_json(result, args.out)

        print(f"抽取完成: {src.name}")
        for pm in page_meta:
            k = pm["kind_counts"]
            print(f"  p{pm['index']}: {pm['line_count']:>3}行 内容高{pm['content_h_pt']:>6.1f}pt → "
                  f"段落{k.get('paragraph', 0)} 标题{k.get('heading', 0)} "
                  f"列表{k.get('bullet', 0)} 表格{pm['table_count']} 图{pm['image_count']}"
                  f"   正文{pm['body_size_pt']}pt")
        m = margins
        usable = result["page"]["height_pt"] - m["top"] - m["bottom"]
        print(f"  页面 {result['page']['width_pt']:.0f}x{result['page']['height_pt']:.0f}pt  "
              f"页边距 L{m['left']} R{m['right']} T{m['top']} B{m['bottom']}  "
              f"→ 可用高度 {usable:.1f}pt")
        if args.verbose:
            for pm in page_meta:
                r = pm["rhythm"]
                print(f"  p{pm['index']} 节奏: 段内行距={r['ref_gap']} "
                      f"段落阈值={r['para_threshold']} 行距={r['line_pt']}pt")
        print(f"\n[blocks] {args.out}")
        return 0
    finally:
        doc.close()


if __name__ == "__main__":
    raise SystemExit(main())
