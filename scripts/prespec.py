#!/usr/bin/env python3
"""prespec.py — 渲染前的页面高度 / 折行预检（纯算术，不启动 Word）。

为什么需要它
------------
「满页」的 PDF 只要多出几磅，末尾几行就会被挤到下一页；一旦有一页溢出，后面
所有页与原页的对应关系全部错乱，还得重新渲染、重新比对。而这种溢出完全可以在
渲染前算出来：

    每页需要高度 = Σ(各段折行数 × 行距) + (段数 − 1) × 段后间距

和「目标高度」（优先取原页实测内容高度，退而用页面可用高度）比一下，就能给出
ok / tight / overflow，以及该压哪个参数。实测这一步能把 Word 渲染迭代从 7 次
压到 2 次——省掉的是最贵的往返。

用法
----
    python prespec.py spec.json --preset page.json
    python prespec.py rebuild.md --preset page.json --json report.json
    python prespec.py spec.json --preset page.json -v      # 展开最紧张那页的逐段明细

退出码：0 = 全部 ok / tight（可渲染）；1 = 有页 overflow（应先改再渲染）。
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
from build_docx import estimate_lines, markdown_to_blocks  # noqa: E402

SAFETY = 0.995


def load_page(preset: str | None) -> dict:
    """与 build_docx.main 同序的页面参数解析。"""
    page = {"width_pt": 595.28, "height_pt": 841.89,
            "margins_pt": {"top": 56.7, "bottom": 56.7, "left": 56.7, "right": 56.7},
            "latin": "Arial", "east_asian": "微软雅黑", "size_pt": 10.5}
    if preset:
        page.update(json.loads(Path(preset).read_text(encoding="utf-8")))
    return page


def load_spec(spec_path: Path, page: dict) -> dict:
    if spec_path.suffix.lower() in (".md", ".markdown", ".txt"):
        return markdown_to_blocks(spec_path.read_text(encoding="utf-8"), page)
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    if spec.get("page"):
        page.update(spec["page"])
    return spec


def split_pages(blocks: list[dict]) -> list[list[dict]]:
    pages: list[list[dict]] = []
    cur: list[dict] = []
    for b in blocks:
        if b.get("kind") == "page_break":
            pages.append(cur)
            cur = []
        else:
            cur.append(b)
    pages.append(cur)
    return pages


def block_height(b: dict, avail_w: float, body_size: float) -> tuple[float, int, str]:
    """返回 (该块高度 pt, 估算行数, 备注)。"""
    kind = b.get("kind", "paragraph")
    style = b.get("style") or {}
    size = float(style.get("size_pt") or body_size)
    line_pt = float(style.get("line_pt") or size * 1.4)
    indent = float(b.get("indent_pt") or style.get("left_indent_pt") or 0.0)

    if kind == "table":
        rows = b.get("rows") or []
        n = max(len(rows), 1)
        return n * size * 1.35, n, f"table {len(rows)}行"
    if kind == "image":
        h = float(b.get("height_pt") or 0.0)
        return h, 0, "image" + ("" if h else "（高度未知，按 0 计）")

    declared = b.get("lines")
    if isinstance(declared, int) and declared > 1:
        n = declared
        note = "declared"
    else:
        n = estimate_lines(b.get("text", ""), size, max(avail_w - indent, 20.0))
        note = "est"
    return n * line_pt, n, note


def analyse(spec: dict, page: dict) -> tuple[list[dict], float, float]:
    margins = page.get("margins_pt") or {}
    top = float(margins.get("top", 72.0))
    bottom = float(margins.get("bottom", 72.0))
    left = float(margins.get("left", 72.0))
    right = float(margins.get("right", 72.0))
    avail = max(float(page.get("height_pt", 841.89)) - top - bottom, 1.0)
    avail_w = max(float(page.get("width_pt", 595.28)) - left - right, 1.0)
    body_size = float(page.get("size_pt", 10.5))

    targets = {pm.get("index"): pm.get("content_h_pt")
               for pm in spec.get("page_meta", [])}

    rows: list[dict] = []
    for pi, blocks in enumerate(split_pages(spec.get("blocks", [])), start=1):
        flow = [b for b in blocks
                if b.get("kind") in ("paragraph", "heading", "bullet", "table", "image")]
        if not flow:
            rows.append({"page": pi, "blocks": 0, "lines": 0, "fixed_pt": 0.0,
                         "var_pt": 0.0, "need_pt": 0.0, "target_pt": 0.0,
                         "avail_pt": round(avail, 1), "verdict": "empty",
                         "detail": []})
            continue

        fixed = 0.0
        var = 0.0
        lines = 0
        detail: list[dict] = []
        for b in flow:
            h, n, note = block_height(b, avail_w, body_size)
            fixed += h
            lines += n
            var += float((b.get("style") or {}).get("space_after_pt", 0.0))
            detail.append({"kind": b.get("kind", "paragraph"), "lines": n,
                           "h_pt": round(h, 1), "note": note,
                           "text": (b.get("text") or "")[:64]})
        need = fixed + var
        raw_target = targets.get(pi)
        target = min(float(raw_target), avail) if raw_target else avail
        target *= SAFETY

        if need <= target:
            verdict = "ok"
        elif need <= avail * SAFETY:
            verdict = "tight"
        else:
            verdict = "overflow"

        # 与 build_docx.apply_height_budget 同优先级的建议：先压段后间距，再压行距
        if verdict == "ok":
            hint = ""
        elif var > 0 and fixed < target:
            hint = f"space_after × {max((target - fixed) / var, 0.0):.3f}"
        else:
            hint = f"line_pt × {max(target / max(fixed, 1e-6), 0.5):.3f}（并置 space_after=0）"

        rows.append({"page": pi, "blocks": len(flow), "lines": lines,
                     "fixed_pt": round(fixed, 1), "var_pt": round(var, 1),
                     "need_pt": round(need, 1), "target_pt": round(target, 1),
                     "avail_pt": round(avail, 1), "verdict": verdict,
                     "hint": hint, "detail": detail})
    return rows, avail, avail_w


def main() -> int:
    ap = argparse.ArgumentParser(description="渲染前页面高度/折行预检")
    ap.add_argument("spec", help="blocks.json 或 .md")
    ap.add_argument("--preset", default=None, help="页面参数 JSON")
    ap.add_argument("--json", default=None, help="把报告写成 JSON")
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="展开最紧张那一页的逐段明细")
    args = ap.parse_args()

    spec_path = Path(args.spec)
    if not spec_path.exists():
        print(f"[错误] 不存在: {spec_path}", file=sys.stderr)
        return 2

    page = load_page(args.preset)
    spec = load_spec(spec_path, page)
    rows, avail, avail_w = analyse(spec, page)

    print("渲染前预检（prespec）")
    print(f"  spec   : {spec_path}")
    print(f"  页面   : {page['width_pt']:.1f} × {page['height_pt']:.1f} pt")
    print(f"  可用   : 高 {avail:.1f} pt  宽 {avail_w:.1f} pt")
    print("-" * 78)
    print(f"{'页':>4} {'判定':<9}{'块':>4}{'行':>5}{'固定':>9}{'间距':>8}"
          f"{'需要':>9}{'目标':>9}  建议")
    for r in rows:
        if r["verdict"] == "empty":
            print(f"{r['page']:>4} {'empty':<9}{'':>4}{'':>5}{'':>9}{'':>8}{'':>9}{'':>9}")
            continue
        print(f"{r['page']:>4} {r['verdict']:<9}{r['blocks']:>4}{r['lines']:>5}"
              f"{r['fixed_pt']:>9.1f}{r['var_pt']:>8.1f}{r['need_pt']:>9.1f}"
              f"{r['target_pt']:>9.1f}  {r.get('hint', '')}")
    print("-" * 78)

    bad = [r for r in rows if r["verdict"] == "overflow"]
    tight = [r for r in rows if r["verdict"] == "tight"]
    print(f"结论: ok={sum(1 for r in rows if r['verdict'] == 'ok')} "
          f"tight={len(tight)} overflow={len(bad)}  共 {len(rows)} 页")
    if bad:
        print("  溢出页即使把段后间距压到 0 也放不下 —— 先改内容或版式，再渲染。")
    elif tight:
        print("  有紧张页：build_docx 会自动等比压段后间距，预计不影响分页。")

    if args.verbose:
        cand = bad or tight or rows
        if cand:
            worst = max(cand, key=lambda r: r["need_pt"])
            print(f"\n最紧张页 p{worst['page']} 逐段明细（{worst['blocks']} 块）:")
            for d in worst["detail"]:
                print(f"  {d['kind']:<10} {d['lines']:>3}行 {d['h_pt']:>7.1f}pt "
                      f"({d['note']})  {d['text']}")

    if args.json:
        Path(args.json).write_text(
            json.dumps({"spec": str(spec_path), "page": page, "avail_pt": avail,
                        "avail_w_pt": avail_w, "pages": rows},
                       ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"报告: {args.json}")

    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())