#!/usr/bin/env python3
"""render_docx.py — docx → pdf 的统一渲染入口（跨平台，自动选渲染器）。

分发包不该绑定单一渲染器：
  - 有 Microsoft Word  → 走 Word COM（render_docx.ps1），保真度最高
  - 没有 Word          → 走 LibreOffice headless
  - 两者都没有         → 明确报错并给出修复建议，而不是静默产出错误结果

用法:
  python render_docx.py in.docx out.pdf
  python render_docx.py in.docx out.pdf --renderer libreoffice
"""
from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _runtime import detect_renderer, scripts_dir  # noqa: E402


def render_with_word(docx: Path, pdf: Path, word_path: str) -> tuple[bool, str]:
    """Windows + Word：调用同目录的 PowerShell 脚本。"""
    ps1 = scripts_dir() / "render_docx.ps1"
    if not ps1.exists():
        return False, f"缺少脚本 {ps1}"

    shell = shutil.which("pwsh") or shutil.which("powershell")
    if not shell:
        return False, "未找到 pwsh/powershell，无法驱动 Word COM"

    cmd = [shell, "-NoProfile", "-File", str(ps1),
           "-InFile", str(docx), "-OutFile", str(pdf)]
    # 显式 utf-8 + errors=replace：Windows 下父进程可能按 GBK 解码子进程输出，
    # 一旦子进程打印了非 GBK 字节就会抛 UnicodeDecodeError 把整个渲染判为失败。
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=600)
    output = (proc.stdout or "") + (proc.stderr or "")
    ok = pdf.exists() and proc.returncode == 0
    return ok, output.strip()


def render_with_libreoffice(docx: Path, pdf: Path, soffice: str) -> tuple[bool, str]:
    """LibreOffice headless 转换。"""
    out_dir = pdf.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", str(out_dir), str(docx)]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=600)
    except subprocess.TimeoutExpired:
        return False, "LibreOffice 转换超时"

    produced = out_dir / (docx.stem + ".pdf")
    if produced.exists() and produced != pdf:
        try:
            if pdf.exists():
                pdf.unlink()
            shutil.move(str(produced), str(pdf))
        except Exception as exc:
            return False, f"重命名输出失败: {exc}"

    ok = pdf.exists()
    output = ((proc.stdout or "") + (proc.stderr or "")).strip()
    return ok, output


def main() -> int:
    ap = argparse.ArgumentParser(description="docx → pdf（自动选渲染器）")
    ap.add_argument("docx")
    ap.add_argument("pdf")
    ap.add_argument("--renderer", choices=["word", "libreoffice", "auto"], default="auto")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    src = Path(args.docx).resolve()
    dst = Path(args.pdf).resolve()
    if not src.exists():
        print(f"[错误] 输入不存在: {src}", file=sys.stderr)
        return 2

    dst.parent.mkdir(parents=True, exist_ok=True)

    info = detect_renderer()
    if args.renderer != "auto":
        import os
        os.environ["PDF2DOCX_RENDERER"] = args.renderer
        info = detect_renderer()

    kind, path = info.get("kind"), info.get("path")
    if not kind:
        print(f"[错误] 没有可用的渲染器：{info.get('reason')}", file=sys.stderr)
        print("  安装 Microsoft Word 或 LibreOffice，", file=sys.stderr)
        print("  或用环境变量 PDF2DOCX_SOFFICE 指定 soffice 路径。", file=sys.stderr)
        return 3

    if kind == "word":
        ok, output = render_with_word(src, dst, path)
    else:
        ok, output = render_with_libreoffice(src, dst, path)

    if ok and not args.quiet:
        size = dst.stat().st_size if dst.exists() else 0
        print(f"RENDER_OK renderer={kind} bytes={size} out={dst}")
    elif not ok:
        print(f"RENDER_FAIL renderer={kind}\n{output}", file=sys.stderr)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
