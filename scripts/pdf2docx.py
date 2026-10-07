#!/usr/bin/env python3
"""pdf2docx.py — 一条命令跑完整个重建流程（对外入口）。

    python pdf2docx.py in.pdf out.docx [--workdir W] [--pages 1-3] [--verify]

流程：分诊 → 抽取/取证 → 重建 → 渲染 → 验收，全部串起来。

关于图片型 PDF 的能力边界（重要，别误期待）：
  文字型/混合型可以一次跑到位；图片型/扫描件在无人介入时只能得到
  基于 OCR 的**初稿**（OCR 会丢字、吞空格、写错专业词）。
  要拿到接近原件质量的图片型结果，需要 Agent 在 `--vision` 断点介入手工判读，
  或直接改用 SKILL.md 里的「视觉判读 → Markdown 重建」路径。
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

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from _runtime import check_deps, detect_renderer  # noqa: E402

import subprocess  # noqa: E402


def run(cmd: list[str], label: str) -> int:
    print(f"\n▶ {label}")
    proc = subprocess.run(cmd)
    return proc.returncode


def py() -> str:
    return sys.executable


def step_probe(pdf: Path, work: Path) -> dict:
    import json
    out = work / "probe.json"
    run([py(), str(HERE / "probe.py"), str(pdf), "--json", str(out)],
        "Phase 0 · 分诊")
    return json.loads(out.read_text(encoding="utf-8"))


def step_text_route(pdf: Path, docx: Path, work: Path, pages: str | None) -> str | None:
    blocks = work / "blocks.json"
    cmd = [py(), str(HERE / "extract.py"), str(pdf), "--out", str(blocks),
           "--images", str(work / "images")]
    if pages:
        cmd += ["--pages", pages]
    if run(cmd, "Phase 1 · 抽取结构") != 0:
        return None
    if run([py(), str(HERE / "build_docx.py"), str(blocks), str(docx)],
           "Phase 3 · 生成 docx") != 0:
        return None
    return str(blocks)


def step_image_route(pdf: Path, docx: Path, work: Path, pages: str | None) -> str | None:
    """图片型/扫描件：OCR 出文字与坐标，拼一个初稿 spec。"""
    ocr_json = work / "ocr.json"
    cmd = [py(), str(HERE / "ocr.py"), str(pdf), "--out", str(ocr_json),
           "--dpi", "200", "--workdir", str(work / "pages")]
    if pages:
        cmd += ["--pages", pages]
    if run(cmd, "Phase 1 · OCR 取文字与坐标") != 0:
        return None

    run([py(), str(HERE / "render_pages.py"), str(pdf),
         "--outdir", str(work / "pages"), "--tile"], "Phase 1 · 渲染页面素材")

    spec = work / "blocks.json"
    if run([py(), str(HERE / "ocr_to_spec.py"), str(ocr_json), str(pdf),
            "--out", str(spec)], "Phase 2 · OCR 结果转结构化规格") != 0:
        return None
    if run([py(), str(HERE / "build_docx.py"), str(spec), str(docx)],
           "Phase 3 · 生成 docx") != 0:
        return None
    return str(spec)


def step_verify(pdf: Path, docx: Path, work: Path, pages: str | None) -> int:
    out_pdf = work / "rendered.pdf"
    if run([py(), str(HERE / "render_docx.py"), str(docx), str(out_pdf)],
           "Phase 4a · 渲染回 PDF") != 0:
        print("  渲染失败，跳过视觉验收；内容抽取结果仍然有效。")
        return 1
    cmd = [py(), str(HERE / "compare.py"), str(pdf), str(out_pdf),
           "--outdir", str(work / "diff")]
    if pages:
        cmd += ["--pages", pages]
    return run(cmd, "Phase 4b · 逐页比对验收")


def main() -> int:
    ap = argparse.ArgumentParser(description="PDF → docx 高保真重建（一条命令）")
    ap.add_argument("pdf")
    ap.add_argument("docx")
    ap.add_argument("--workdir", default=None, help="中间产物目录，默认 <docx同级>/_pdf2docx")
    ap.add_argument("--pages", default=None, help="只处理指定页，如 1-3")
    ap.add_argument("--verify", action="store_true", default=True,
                    help="渲染回图并逐页比对（默认开启）")
    ap.add_argument("--no-verify", dest="verify", action="store_false")
    args = ap.parse_args()

    pdf = Path(args.pdf).resolve()
    docx = Path(args.docx).resolve()
    if not pdf.exists():
        print(f"[错误] 输入不存在: {pdf}", file=sys.stderr)
        return 2

    # 依赖门槛
    deps = check_deps()
    if not deps["ok"]:
        missing = ", ".join(d["pip"] for d in deps["missing_required"])
        print(f"[错误] 缺少必需依赖: {missing}", file=sys.stderr)
        print(f"  请先执行: {py()} {HERE / 'bootstrap.py'} --install", file=sys.stderr)
        return 3

    work = Path(args.workdir) if args.workdir else docx.parent / "_pdf2docx"
    work.mkdir(parents=True, exist_ok=True)

    print("=" * 68)
    print(f"输入: {pdf}")
    print(f"输出: {docx}")
    print(f"中间产物: {work}")
    print("=" * 68)

    info = step_probe(pdf, work)
    kind = info.get("kind_summary", {})
    route = info.get("route", {}).get("route", "A/text")
    print(f"\n判定: {kind}  →  路线 {route}")

    image_like = all(k in ("image", "scanned", "vector") for k in kind) and kind
    if image_like:
        print("\n注意：这是图片型/扫描件 PDF。无人介入时只能得到 OCR 初稿，")
        print("      质量受限于 OCR（会丢字、吞空格、写错专业词）。")
        print("      要拿到高保真结果，请按 SKILL.md 的「视觉判读 → Markdown 重建」路径做，")
        print("      并把 OCR 结果仅当作坐标参考。")
        spec = step_image_route(pdf, docx, work, args.pages)
    else:
        spec = step_text_route(pdf, docx, work, args.pages)

    if not spec:
        print("\n[失败] 重建中断。", file=sys.stderr)
        return 1

    print(f"\n✓ 已生成: {docx}")

    if args.verify:
        rend = detect_renderer()
        if not rend["kind"]:
            print(f"\n⚠ 无可用渲染器（{rend.get('reason')}），跳过 Phase 4 验收。")
            print("  安装 Word 或 LibreOffice 后可再次运行以获得比对报告。")
        else:
            rc = step_verify(pdf, docx, work, args.pages)
            print("\n" + "=" * 68)
            print("完成。" + ("验收全部通过。" if rc == 0 else
                            "验收有未通过项，请查看上面的比对表与 diff 目录配图。"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
