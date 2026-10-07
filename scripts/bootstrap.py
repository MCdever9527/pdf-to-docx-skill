#!/usr/bin/env python3
"""bootstrap.py — 一键安装与自检（分发包的入口）。

把技能目录复制到任意机器后，先跑这个：
    python scripts/bootstrap.py            # 只自检，不改动环境
    python scripts/bootstrap.py --install  # 缺什么装什么
    python scripts/bootstrap.py --install --with-ocr   # 连 OCR 一起装

它做三件事：
  1. 检查 Python 依赖，给出（或执行）精确的安装命令
  2. 探测 docx→pdf 渲染器（Word / LibreOffice），并说明缺失后果
  3. 写出 env.json，让后续脚本免去重复探测
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _runtime import (  # noqa: E402
    REQUIRED_DEPS,
    OPTIONAL_DEPS,
    env_report,
    pip_install_command,
    skill_root,
)


def pip_install(packages: list[str], mirror: bool) -> bool:
    if not packages:
        return True
    cmd = [sys.executable, "-m", "pip", "install", "--no-input", *packages]
    if mirror:
        cmd += ["-i", "https://pypi.tuna.tsinghua.edu.cn/simple"]
    print(f"  $ {' '.join(cmd)}")
    proc = subprocess.run(cmd)
    return proc.returncode == 0


def main() -> int:
    ap = argparse.ArgumentParser(description="pdf-to-docx 安装与自检")
    ap.add_argument("--install", action="store_true", help="实际执行安装")
    ap.add_argument("--with-ocr", action="store_true",
                    help="同时安装 OCR 与图像分析（图片型/扫描件必需）")
    ap.add_argument("--mirror", action="store_true", help="使用清华 PyPI 镜像")
    args = ap.parse_args()

    report = env_report()
    ok = True

    print("=" * 68)
    print("pdf-to-docx  安装 / 自检")
    print("=" * 68)
    print(f"技能目录 : {report['skill_root']}")
    print(f"平台     : {report['platform']}")
    print(f"Python   : {report['python_version']}  {report['python']}")
    print("-" * 68)

    # 1) 依赖
    # env_report() 已经在这两项里归约成 pip 名称列表（见 _runtime.env_report），
    # 这里不能再按 dict 取值，否则会 TypeError: string indices must be integers。
    missing_req = list(report["missing_required"])
    missing_opt = list(report["missing_optional"])
    if not missing_req:
        print("✓ 必需依赖齐全")
    else:
        ok = False
        print(f"✗ 缺少必需依赖: {', '.join(missing_req)}")
        for mod, purpose, pip_name in REQUIRED_DEPS:
            if pip_name in missing_req:
                print(f"    - {pip_name}: {purpose}")
        if args.install:
            print("  正在安装…")
            if pip_install(missing_req, args.mirror):
                print("  ✓ 安装完成")
                ok = True
            else:
                print("  ✗ 安装失败")
                print(f"    请手动执行: {pip_install_command(missing_req, args.mirror)}")
        else:
            print(f"    修复: {pip_install_command(missing_req, args.mirror)}")

    if missing_opt:
        print(f"· 可选依赖未装: {', '.join(missing_opt)}")
        print("    影响: 图片型/扫描件 PDF 无法 OCR；复杂表格抽取能力下降。")
        if args.with_ocr:
            if args.install:
                print("  正在安装…")
                if pip_install(missing_opt, args.mirror):
                    print("  ✓ 安装完成")
                else:
                    print(f"  ✗ 失败，请手动执行: {pip_install_command(missing_opt, args.mirror)}")
            else:
                print(f"    加 --install 一起装: {pip_install_command(missing_opt, args.mirror)}")
        else:
            print(f"    安装: {pip_install_command(missing_opt, args.mirror)}")

    # 2) 渲染器
    print("-" * 68)
    rend = report["renderer"]
    if rend["kind"]:
        print(f"✓ 渲染器: {rend['kind']}")
        print(f"    {rend['path']}")
    else:
        ok = False
        print(f"✗ 渲染器缺失: {rend.get('reason')}")
        print("    影响: 无法把 .docx 渲染回 PDF，Phase 4 的视觉验收会退化。")
        print("    修复: 安装 Microsoft Word 或 LibreOffice；")
        print("          或设置环境变量 PDF2DOCX_SOFFICE 指向 soffice 可执行文件。")

    # 3) 落地 env.json
    env_path = skill_root() / "env.json"
    try:
        env_path.write_text(json.dumps(report, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        print("-" * 68)
        print(f"环境快照已写入: {env_path}")
    except Exception as exc:
        print(f"（env.json 写入失败，不影响使用: {exc}）")

    print("=" * 68)
    print("结论:", "可以开始使用" if ok else "还有待解决的项，见上")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
