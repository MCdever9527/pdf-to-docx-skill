#!/usr/bin/env python3
"""_kit.py — pdf-to-docx 共用底座：PDF/DOCX 单位换算、字体映射、OOXML 助手。

设计原则：只依赖本机 Python 运行时已安装的包（pymupdf / python-docx / Pillow），
不调用任何外部独立程序。渲染回图走本机 Word COM（见 render_docx.ps1）。
"""
from __future__ import annotations

import json
import re
import unicodedata
from pathlib import Path

# ---------------------------------------------------------------- PDF 库引导

def load_fitz():
    """PyMuPDF 在新版改名 pymupdf，旧名 fitz 仍可用。两个都试。"""
    try:
        import pymupdf as fitz  # >= 1.24 推荐名
        return fitz
    except ImportError:
        pass
    try:
        import fitz  # 旧名
        return fitz
    except ImportError as exc:  # pragma: no cover
        raise SystemExit(
            "缺少 PyMuPDF。请先执行： python -m pip install pymupdf"
        ) from exc


# ---------------------------------------------------------------- 单位换算

PT_PER_INCH = 72.0
DXA_PER_PT = 20.0        # Word twips
EMU_PER_PT = 12700.0     # OOXML EMU


def pt_to_dxa(pt: float) -> int:
    return int(round(pt * DXA_PER_PT))


def pt_to_emu(pt: float) -> int:
    return int(round(pt * EMU_PER_PT))


def pt_to_px(pt: float, dpi: float) -> float:
    return pt * dpi / PT_PER_INCH


def round_half(x: float) -> float:
    """PDF 里字号常见 10.48/12.05 这种噪声值，吸附到 0.5 网格。"""
    return round(x * 2.0) / 2.0


# ---------------------------------------------------------------- 颜色

def int_to_hex_rgb(color: int) -> str:
    """PyMuPDF span['color'] 是 sRGB 整数 → '#RRGGBB'。"""
    if color is None:
        return "#000000"
    return "#{:06X}".format(int(color) & 0xFFFFFF)


def hex_to_rgb_tuple(hex_color: str) -> tuple[float, float, float]:
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4))  # type: ignore[return-value]


# ---------------------------------------------------------------- 字体

_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")

# PDF 内嵌字体名 → (拉丁字体, 东亚字体)；None 表示交给默认
_FONT_MAP = [
    (("timesnewroman", "times", "nimbusroman", "liberationserif", "georgia", "garamond"), "Times New Roman", "宋体"),
    (("courier", "nimbusmono", "liberationmono", "consol"), "Consolas", "宋体"),
    (("calibri", "carlito"), "Calibri", "微软雅黑"),
    (("cambria", "caladea"), "Cambria", "宋体"),
    (("helvetica", "arial", "liberationsans", "nimbussans", "roboto"), "Arial", "微软雅黑"),
    (("segoe",), "Segoe UI", "微软雅黑"),
    (("simsun", "songti", "stsong", "nsimsun"), "Times New Roman", "宋体"),
    (("simhei", "heiti", "stheiti"), "Arial", "黑体"),
    (("msyh", "yahei", "microsoftyahei"), "Arial", "微软雅黑"),
    (("kaiti", "kai"), "Times New Roman", "楷体"),
    (("fangsong", "fs"), "Times New Roman", "仿宋"),
    (("dengxian", "等线"), "Arial", "等线"),
]


def normalize_font_name(pdf_font: str) -> str:
    """剥掉子集前缀与样式后缀，只留族名（小写无空格）。"""
    name = _SUBSET_PREFIX.sub("", pdf_font or "")
    name = re.sub(r"[,\-](Bold|Italic|Oblique|MT|PS|MTExtra|Regular)$", "", name, flags=re.I)
    name = name.replace(" ", "").replace("_", "")
    return name.lower()


def map_font(pdf_font: str) -> tuple[str, str]:
    """PDF 字体名 → (latin, eastAsia)。命不中就给一套安全默认。"""
    low = normalize_font_name(pdf_font)
    for keys, latin, ea in _FONT_MAP:
        if any(k in low for k in keys):
            return latin, ea
    # 含 CJK 关键词的收尾
    if any(k in low for k in ("hei", "song", "kai", "ming", "gothic")):
        return "Arial", "微软雅黑"
    return "Arial", "微软雅黑"


def font_flags_to_style(flags: int) -> tuple[bool, bool]:
    """PyMuPDF flags 位：1=上标 2=斜体 4=衬线 8=等宽 16=粗体。"""
    italic = bool(flags & 2)
    bold = bool(flags & 16)
    return bold, italic


def is_cjk(ch: str) -> bool:
    if not ch:
        return False
    code = ord(ch)
    return (
        0x4E00 <= code <= 0x9FFF
        or 0x3400 <= code <= 0x4DBF
        or 0xF900 <= code <= 0xFAFF
        or 0x3040 <= code <= 0x30FF
        or 0xAC00 <= code <= 0xD7AF
    )


def has_cjk(text: str) -> bool:
    return any(is_cjk(c) for c in text if not c.isspace())


def clean_text(text: str) -> str:
    """去掉软连字符、零宽字符、PDF 常见的连字噪声。"""
    if not text:
        return ""
    text = text.replace("\u00ad", "").replace("\u200b", "").replace("\ufeff", "")
    text = text.replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\ufb01", "fi").replace("\ufb02", "fl")
    return unicodedata.normalize("NFC", text)


# ---------------------------------------------------------------- 几何

def bbox_of(block) -> tuple[float, float, float, float]:
    return tuple(float(v) for v in block[:4])  # type: ignore[return-value]


def overlap_ratio(a: tuple[float, float, float, float],
                  b: tuple[float, float, float, float]) -> float:
    """a 被 b 覆盖的面积占 a 的比例。用于判断文字是否落在表格/图片框内。"""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    area = max((ax1 - ax0) * (ay1 - ay0), 1e-6)
    return inter / area


def union_bbox(boxes) -> tuple[float, float, float, float]:
    xs0, ys0, xs1, ys1 = [], [], [], []
    for b in boxes:
        xs0.append(b[0]); ys0.append(b[1]); xs1.append(b[2]); ys1.append(b[3])
    return (min(xs0), min(ys0), max(xs1), max(ys1))


# ---------------------------------------------------------------- IO

def dump_json(obj, path: str | Path) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    return p


def load_json(path: str | Path):
    return json.loads(Path(path).read_text(encoding="utf-8"))
