#!/usr/bin/env python3
"""_runtime.py — 运行时探测层（分发包的关键）。

本技能要能被复制到任意机器上运行，因此：
  - 不写死任何本机绝对路径；一切通过「环境变量 → 常见安装位置 → PATH」逐级探测
  - 渲染器不绑定单一实现：Word COM（Windows）与 LibreOffice 二选一，自动降级
  - 依赖缺失时给出明确的修复命令，而不是抛一句 ImportError

支持的环境变量覆盖：
  PDF2DOCX_WORD      指定 WINWORD.EXE 路径
  PDF2DOCX_SOFFICE   指定 soffice / libreoffice 可执行文件路径
  PDF2DOCX_RENDERER  强制渲染器：word | libreoffice | none
"""
from __future__ import annotations

import importlib
import importlib.util
import os
import platform
import shutil
import sys
from pathlib import Path

# ---------------------------------------------------------------- 目录

def skill_root() -> Path:
    """技能根目录（本文件在 <root>/scripts/ 下）。"""
    return Path(__file__).resolve().parent.parent


def scripts_dir() -> Path:
    return skill_root() / "scripts"


def assets_dir() -> Path:
    return skill_root() / "assets"


# ---------------------------------------------------------------- 依赖

# 模块名 -> 用途；python-docx 的导入名是 docx，需要单独映射
REQUIRED_DEPS = [
    ("pymupdf", "PDF 解析与页面渲染", "pymupdf"),
    ("PIL", "图像处理", "Pillow"),
    ("docx", "生成 .docx", "python-docx"),
]

OPTIONAL_DEPS = [
    ("pdfplumber", "复杂表格抽取（可选增强）", "pdfplumber"),
    ("rapidocr_onnxruntime", "图片型/扫描件 OCR", "rapidocr-onnxruntime"),
    ("cv2", "版面图像分析（可选增强）", "opencv-python-headless"),
]


def _has_module(name: str) -> bool:
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def check_deps(include_optional: bool = True) -> dict:
    missing_required, missing_optional, present = [], [], []
    for mod, purpose, pip_name in REQUIRED_DEPS:
        (present if _has_module(mod) else missing_required).append(
            {"module": mod, "pip": pip_name, "purpose": purpose})
    if include_optional:
        for mod, purpose, pip_name in OPTIONAL_DEPS:
            (present if _has_module(mod) else missing_optional).append(
                {"module": mod, "pip": pip_name, "purpose": purpose})
    return {
        "ok": not missing_required,
        "present": present,
        "missing_required": missing_required,
        "missing_optional": missing_optional,
        "python": sys.executable,
        "python_version": platform.python_version(),
    }


def pip_install_command(names: list[str], use_mirror: bool = False) -> str:
    quoted = " ".join(names)
    cmd = f'"{sys.executable}" -m pip install {quoted}'
    if use_mirror:
        cmd += " -i https://pypi.tuna.tsinghua.edu.cn/simple"
    return cmd


# ---------------------------------------------------------------- 渲染器探测

def _existing(paths: list[str]) -> str | None:
    for p in paths:
        if p and Path(p).exists():
            return str(Path(p))
    return None


def find_word() -> str | None:
    """Windows 上的 Microsoft Word 可执行文件。"""
    if platform.system() != "Windows":
        return None

    env = os.environ.get("PDF2DOCX_WORD")
    if env:
        return _existing([env])

    found = shutil.which("WINWORD.EXE") or shutil.which("winword.exe")
    if found:
        return found

    roots = [
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
    ]
    candidates: list[str] = []
    for root in roots:
        if not root:
            continue
        base = Path(root) / "Microsoft Office"
        if base.exists():
            # root/Office16/WINWORD.EXE、root/Office15/WINWORD.EXE 等
            candidates += [str(p) for p in base.glob("*/WINWORD.EXE")]
            candidates += [str(p) for p in base.glob("*/*/WINWORD.EXE")]
    return _existing(candidates)


def find_libreoffice() -> str | None:
    """跨平台的 soffice / libreoffice 可执行文件。"""
    env = os.environ.get("PDF2DOCX_SOFFICE") or os.environ.get("LIBREOFFICE_PATH")
    if env:
        hit = _existing([env])
        if hit:
            return hit

    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found

    system = platform.system()
    candidates: list[str] = []
    if system == "Windows":
        roots = [os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]
        for root in roots:
            candidates += [
                rf"{root}\LibreOffice\program\soffice.exe",
                rf"{root}\LibreOffice 7\program\soffice.exe",
            ]
            if os.path.isdir(root):
                candidates += [str(p) for p in Path(root).glob("LibreOffice*/program/soffice.exe")]
    elif system == "Darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    else:
        candidates += ["/usr/bin/soffice", "/usr/local/bin/soffice",
                       "/usr/bin/libreoffice", "/snap/bin/libreoffice"]
    return _existing(candidates)


def detect_renderer() -> dict:
    """选择一个可用的 docx→pdf 渲染器。

    优先 Word：验收要回答「这份 docx 在 Word 里长什么样」，
    Word 本体渲染才没有中间环节的偏差。没有 Word 时退回 LibreOffice。
    """
    forced = (os.environ.get("PDF2DOCX_RENDERER") or "").strip().lower()

    if forced == "none":
        return {"kind": None, "path": None, "reason": "已被 PDF2DOCX_RENDERER=none 禁用"}
    if forced == "word":
        path = find_word()
        return {"kind": "word", "path": path} if path else {
            "kind": None, "path": None, "reason": "指定用 Word，但未找到 WINWORD.EXE"}
    if forced == "libreoffice":
        path = find_libreoffice()
        return {"kind": "libreoffice", "path": path} if path else {
            "kind": None, "path": None, "reason": "指定用 LibreOffice，但未找到 soffice"}

    word = find_word()
    if word:
        return {"kind": "word", "path": word}
    lo = find_libreoffice()
    if lo:
        return {"kind": "libreoffice", "path": lo}
    return {"kind": None, "path": None,
            "reason": "未找到 Word 或 LibreOffice"}


# ---------------------------------------------------------------- 自检报告

def env_report() -> dict:
    deps = check_deps()
    return {
        "skill_root": str(skill_root()),
        "platform": platform.platform(),
        "python": deps["python"],
        "python_version": deps["python_version"],
        "deps_ok": deps["ok"],
        "missing_required": [d["pip"] for d in deps["missing_required"]],
        "missing_optional": [d["pip"] for d in deps["missing_optional"]],
        "renderer": detect_renderer(),
    }


def print_env_report() -> int:
    r = env_report()
    print("pdf-to-docx 运行环境自检")
    print(f"  技能目录  : {r['skill_root']}")
    print(f"  平台      : {r['platform']}")
    print(f"  Python    : {r['python_version']}  ({r['python']})")
    if r["missing_required"]:
        print(f"  ✗ 必需依赖缺失: {', '.join(r['missing_required'])}")
        print(f"    修复: {pip_install_command(r['missing_required'])}")
    else:
        print("  ✓ 必需依赖齐全")
    if r["missing_optional"]:
        print(f"  · 可选依赖未装: {', '.join(r['missing_optional'])}")
        print(f"    安装以解锁图片型/复杂表格: {pip_install_command(r['missing_optional'])}")
    rend = r["renderer"]
    if rend["kind"]:
        print(f"  ✓ 渲染器  : {rend['kind']}  ({rend['path']})")
    else:
        print(f"  ✗ 渲染器缺失: {rend.get('reason')}")
        print("    影响: render_docx 无法把 docx 渲染回图，Phase 4 验收只能出内容比对。")
        print("    修复: 安装 Microsoft Word 或 LibreOffice，")
        print("          或用 PDF2DOCX_SOFFICE 指定 soffice 路径。")
    return 0 if (r["deps_ok"] and rend["kind"]) else 1


if __name__ == "__main__":
    raise SystemExit(print_env_report())
