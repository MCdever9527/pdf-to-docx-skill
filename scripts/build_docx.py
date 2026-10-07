#!/usr/bin/env python3
"""build_docx.py — 把结构规格重建为 .docx（Phase 3）。

规格来源可以是：
  A) blocks.json —— 由 extract.py / ocr.py+vision 生成的结构化块
  B) Markdown     —— 人工或 Agent 直接写的重建稿（推荐用于图片型 PDF）

设计要点（决定了「格式保得住」）：
  1. 页面尺寸与页边距按 PDF 实测值写入 section，不套 Word 默认值
  2. 字体同时写 w:rFonts 的 ascii/hAnsi(拉丁) 与 eastAsia(东亚)，避免中文回退乱字体
  3. 段落缩进/行距/段前后一律由 pt 换算为 twips/EMU，不做四舍五入到整数倍
  4. 表格生成真 OOXML 表格（不是空格对齐），带边框与表头重复
  5. 色块标题用段落底纹 w:shd + 段落下边框实现，按 OOXML schema 顺序插入 pPr

用法:
  python build_docx.py spec.md  out.docx --preset page.json
  python build_docx.py blocks.json out.docx
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path

for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _kit import map_font, pt_to_dxa  # noqa: E402

from docx import Document  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Pt, RGBColor  # noqa: E402

# ------------------------------------------------------------------ OOXML 助手


def _el(tag: str, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def _ordered_insert(pPr, element, order: list[str]):
    """按 OOXML schema 顺序把 element 插进 pPr，否则 Word 报「文件已损坏」。"""
    tag = element.tag.split("}")[-1]
    try:
        idx = order.index(tag)
    except ValueError:
        pPr.append(element)
        return
    for child in pPr:
        ctag = child.tag.split("}")[-1]
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(element)
            return
    pPr.append(element)


# pPr 子元素合法顺序（节选，覆盖本引擎会写的全部元素）
_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]


def set_east_asian_font(run, latin: str, east_asian: str) -> None:
    """python-docx 的 run.font.name 只写 ascii/hAnsi；中文必须补 eastAsia。"""
    run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), east_asian)
    rFonts.set(qn("w:cs"), latin)


def shade_paragraph(paragraph, fill_hex: str) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    shd = _el("w:shd", val="clear", color="auto", fill=fill_hex.lstrip("#"))
    _ordered_insert(pPr, shd, _PPR_ORDER)


def bottom_border(paragraph, color_hex: str, size_eighth_pt: int = 4,
                  space: int = 1) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = _el("w:bottom", val="single", sz=size_eighth_pt,
                 space=space, color=color_hex.lstrip("#"))
    pBdr.append(bottom)
    _ordered_insert(pPr, pBdr, _PPR_ORDER)


def set_paragraph_geometry(paragraph, *, space_before_pt=None, space_after_pt=None,
                           line_pt=None, line_multiple=None,
                           left_indent_pt=None, first_line_pt=None) -> None:
    pf = paragraph.paragraph_format
    if space_before_pt is not None:
        pf.space_before = Pt(space_before_pt)
    if space_after_pt is not None:
        pf.space_after = Pt(space_after_pt)
    if line_multiple is not None:
        pf.line_spacing = line_multiple
    elif line_pt is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_pt)
    if left_indent_pt is not None:
        pf.left_indent = Pt(left_indent_pt)
    if first_line_pt is not None:
        pf.first_line_indent = Pt(first_line_pt)


# ------------------------------------------------------------------ 样式与版式

def apply_page_setup(doc: Document, page: dict) -> None:
    """按 PDF 实测几何写页面设置：尺寸、页边距、页眉页脚距离。"""
    for section in doc.sections:
        _apply_to_section(section, page)


def _apply_to_section(section, page: dict) -> None:
    if page.get("width_pt"):
        section.page_width = Pt(page["width_pt"])
    if page.get("height_pt"):
        section.page_height = Pt(page["height_pt"])
    m = page.get("margins_pt") or {}
    if m.get("left") is not None:
        section.left_margin = Pt(m["left"])
    if m.get("right") is not None:
        section.right_margin = Pt(m["right"])
    if m.get("top") is not None:
        section.top_margin = Pt(m["top"])
    if m.get("bottom") is not None:
        section.bottom_margin = Pt(m["bottom"])
    if m.get("header") is not None:
        section.header_distance = Pt(m["header"])
    if m.get("footer") is not None:
        section.footer_distance = Pt(m["footer"])


def patch_default_font(doc: Document, latin: str, east_asian: str, size_pt: float) -> None:
    """改文档级默认字体，避免未显式设字体的 run 回退到 Calibri+宋体。"""
    styles = doc.styles.element
    for rPrDefault in styles.findall(qn("w:docDefaults") + "/" + qn("w:rPrDefault")):
        rPr = rPrDefault.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            rPrDefault.append(rPr)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), latin)
        rFonts.set(qn("w:hAnsi"), latin)
        rFonts.set(qn("w:eastAsia"), east_asian)
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), str(int(round(size_pt * 2))))


# ------------------------------------------------------------------ 页眉页脚

def _field_run(paragraph, instr: str, size_pt: float | None = None,
               latin: str = "Times New Roman", east_asian: str = "宋体"):
    """插入 Word 域（如 PAGE / NUMPAGES）——域必须由 fldChar 序列构成。"""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr_el, sep, end):
        run._element.append(el)
    if size_pt:
        run.font.size = Pt(size_pt)
    set_east_asian_font(run, latin, east_asian)
    return run


_ALIGN_MAP = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _write_fields(para, fmt: str, size_pt, latin: str, ea: str,
                  bold: bool = False, color: str | None = None) -> None:
    """把含 {PAGE} / {NUMPAGES} 占位符的模板写进段落，普通文字走 run，占位符走域。"""
    for tok in re.split(r"(\{PAGE\}|\{NUMPAGES\})", fmt):
        if not tok:
            continue
        if tok == "{PAGE}":
            r = _field_run(para, " PAGE ", size_pt, latin, ea)
        elif tok == "{NUMPAGES}":
            r = _field_run(para, " NUMPAGES ", size_pt, latin, ea)
        else:
            r = para.add_run(tok)
            if size_pt:
                r.font.size = Pt(size_pt)
            set_east_asian_font(r, latin, ea)
        if bold:
            r.bold = True
        if color:
            r.font.color.rgb = RGBColor.from_string(color.lstrip("#"))


def set_header(section, spec_header: dict) -> None:
    """写页眉。文本里的 {PAGE}/{NUMPAGES} 会转成域，所以页码能自动跟随。

    字段：text(或 format) / size_pt / bold / color / align / border_bottom
    """
    if not spec_header:
        return
    header = section.header
    header.is_linked_to_previous = False
    para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    for r in list(para.runs):
        r._element.getparent().remove(r._element)

    fmt = spec_header.get("text") or spec_header.get("format") or ""
    _write_fields(para, fmt,
                  spec_header.get("size_pt"),
                  spec_header.get("latin", "Times New Roman"),
                  spec_header.get("east_asian", "宋体"),
                  spec_header.get("bold", False),
                  spec_header.get("color"))
    para.alignment = _ALIGN_MAP.get(spec_header.get("align", "center"),
                                    WD_ALIGN_PARAGRAPH.CENTER)
    if spec_header.get("border_bottom"):
        bottom_border(para, spec_header["border_bottom"],
                      spec_header.get("border_size", 4))


def set_footer_page_number(section, spec_footer: dict) -> None:
    """写页脚页码：{format, size_pt, align}；format 如 'Page {PAGE} of {NUMPAGES}'。"""
    if not spec_footer:
        return
    footer = section.footer
    footer.is_linked_to_previous = False
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for r in list(para.runs):
        r._element.getparent().remove(r._element)

    _write_fields(para,
                  spec_footer.get("text") or spec_footer.get("format") or "{PAGE}",
                  spec_footer.get("size_pt"),
                  spec_footer.get("latin", "Times New Roman"),
                  spec_footer.get("east_asian", "宋体"),
                  spec_footer.get("bold", False),
                  spec_footer.get("color"))
    para.alignment = _ALIGN_MAP.get(spec_footer.get("align", "center"),
                                    WD_ALIGN_PARAGRAPH.CENTER)


# ------------------------------------------------------------------ 高度预算

def estimate_lines(text: str, size_pt: float, avail_width_pt: float) -> int:
    """估算段落折行后的行数。

    这一步是高度预算准确的前提：spec 里给的 lines 往往是「逻辑行数」，
    而 Word 会按版心宽度自动折行。若按 1 行估算，长段落会被严重低估，
    预算失效，内容一路溢出到后面所有页。
    """
    if not text:
        return 1
    cjk = sum(1 for ch in text if ord(ch) > 0x2E80)
    latin = max(len(text) - cjk, 0)
    # 经验字宽：拉丁约 0.5em，CJK 约 1.0em
    width = (latin * 0.5 + cjk * 1.0) * size_pt
    if avail_width_pt <= 1:
        return 1
    return max(1, math.ceil(width / avail_width_pt))


def apply_height_budget(spec: dict, page: dict, safety: float = 0.995) -> dict:
    """按页做高度预算，保证重建高度贴合原页实测高度。

    为什么必须有这一步：原 PDF 的某些页是「满页」，内容高度正好等于可用高度，
    重建时多出的一点点间距就会把末尾几行挤到下一页；一旦有一页溢出，
    后面所有页与原页的对应关系就全乱了。

    目标高度优先用「原页内容实测高度」。它比单纯用页面可用高度更准：
    内容少的页本就排不满，不该被拉伸，也不该被无谓压缩。

    压缩优先级：先压段后间距（视觉影响小），仍不够才动行距（影响观感）。
    """
    margins = page.get("margins_pt") or {}
    left = margins.get("left", 72.0)
    right = margins.get("right", 72.0)
    avail = page.get("height_pt", 841.89) - margins.get("top", 72.0) - \
        margins.get("bottom", 72.0)
    avail_w = max(page.get("width_pt", 595.28) - left - right, 1.0)

    # 每页的原页实测内容高度
    targets = {pm["index"]: pm.get("content_h_pt") for pm in spec.get("page_meta", [])}

    # 按分页标记切页
    pages: list[list[dict]] = []
    current: list[dict] = []
    for block in spec.get("blocks", []):
        if block.get("kind") == "page_break":
            pages.append(current)
            current = []
        else:
            current.append(block)
    pages.append(current)

    report = []
    for pi, blocks in enumerate(pages):
        flow = [b for b in blocks if b.get("kind") in ("paragraph", "heading", "bullet")]
        if not flow:
            continue

        # 逐段按「折行后的实际行数」累加高度，这是预算准不准的关键
        fixed = 0.0
        for b in flow:
            size = b["style"].get("size_pt") or 11.0
            line_pt = b["style"].get("line_pt") or (size * 1.4)
            indent = b.get("indent_pt") or b["style"].get("left_indent_pt") or 0.0
            declared = b.get("lines")
            if isinstance(declared, int) and declared > 1:
                n = declared
            else:
                n = estimate_lines(b.get("text", ""), size,
                                   max(avail_w - indent, 20.0))
            fixed += n * line_pt

        var = sum(b["style"].get("space_after_pt", 0.0) for b in flow)
        need = fixed + var

        # 目标：原页实测高度；拿不到就退回可用高度
        target = targets.get(pi + 1) or avail
        target = min(target, avail) * safety

        if need <= target * 1.005:
            report.append({"page": pi + 1, "need_pt": round(need, 1),
                           "target_pt": round(target, 1), "action": "none"})
            continue

        entry = {"page": pi + 1, "need_pt": round(need, 1),
                 "target_pt": round(target, 1)}

        if var > 0 and fixed < target:
            k = max((target - fixed) / var, 0.0)
            for b in flow:
                b["style"]["space_after_pt"] = round(
                    b["style"].get("space_after_pt", 0.0) * k, 2)
            entry.update({"action": "shrink_space_after", "factor": round(k, 3)})
        else:
            # 连行高都超了：按比例压行距（最后手段，会改变观感）
            k = max(target / max(fixed, 1e-6), 0.5)
            for b in flow:
                if b["style"].get("line_pt"):
                    b["style"]["line_pt"] = round(b["style"]["line_pt"] * k, 2)
                b["style"]["space_after_pt"] = 0.0
            entry.update({"action": "shrink_line_pt", "factor": round(k, 3)})

        report.append(entry)

    spec["_budget_report"] = report
    return spec



BULLET_RE = re.compile(r"^(\s*)[-*+•·]\s+(.*)$")
NUMBERED_RE = re.compile(r"^(\s*)(\d+)[.)、]\s+(.*)$")


def _add_tab(paragraph):
    """在 run 文本里，制表符必须是 <w:tab/> 元素，直接塞 \\t 字符不生效。"""
    run = paragraph.add_run()
    run._element.append(OxmlElement("w:tab"))
    return run


def add_runs(paragraph, text: str, base: dict) -> None:
    """支持 **粗体** / *斜体* 的内联标记，逐段生成 run；\\t 转成真制表符。"""
    tokens = re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text)
    for tok in tokens:
        if not tok:
            continue
        bold, italic, body = base["bold"], base["italic"], tok
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            bold, body = True, tok[2:-2]
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            italic, body = True, tok[1:-1]

        # 按制表符切分：文字走 run，间隔走 w:tab
        for i, seg in enumerate(body.split("\t")):
            if i > 0:
                _add_tab(paragraph)
            if not seg:
                continue
            run = paragraph.add_run(seg)
            run.bold = bold
            run.italic = italic
            if base.get("underline"):
                run.underline = True
            if base.get("color"):
                run.font.color.rgb = RGBColor.from_string(base["color"].lstrip("#"))
            if base.get("size_pt"):
                run.font.size = Pt(base["size_pt"])
            set_east_asian_font(run, base.get("latin", "Arial"),
                                base.get("east_asian", "微软雅黑"))


def build_from_blocks(doc: Document, spec: dict) -> dict:
    stats = {"paragraphs": 0, "headings": 0, "tables": 0, "images": 0,
             "bullets": 0, "sections": 1}
    page_cfg = spec.get("page") or {}
    header_cfg = spec.get("header") or {}
    footer_cfg = spec.get("footer") or {}

    # 首页的页眉页脚；后续分节默认链接到上一节，因此自动继承
    if doc.sections:
        set_header(doc.sections[0], header_cfg)
        set_footer_page_number(doc.sections[0], footer_cfg)

    for block in spec.get("blocks", []):
        kind = block.get("kind", "paragraph")
        text = block.get("text", "") or ""
        style = dict(block.get("style") or {})

        if kind == "table":
            rows = block.get("rows") or []
            if not rows:
                continue
            ncols = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=ncols)
            table.style = "Table Grid"
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for ri, row in enumerate(rows):
                for ci in range(ncols):
                    cell = table.cell(ri, ci)
                    cell.text = ""
                    para = cell.paragraphs[0]
                    val = row[ci] if ci < len(row) else ""
                    add_runs(para, str(val), {
                        "bold": ri == 0 and block.get("header_row", True),
                        "italic": False,
                        "size_pt": style.get("size_pt", 10),
                        "latin": style.get("latin", "Arial"),
                        "east_asian": style.get("east_asian", "微软雅黑"),
                    })
            if block.get("header_row", True):
                tr = table.rows[0]._tr
                trPr = tr.get_or_add_trPr()
                trPr.append(_el("w:tblHeader", val="true"))
            # 列宽（pt）：不指定时交给 Word 自适应
            widths = block.get("col_widths_pt")
            if widths:
                for ci in range(min(len(widths), ncols)):
                    w = Pt(widths[ci])
                    for row in table.rows:
                        row.cells[ci].width = w
            stats["tables"] += 1
            continue

        if kind == "image":
            path = block.get("path")
            if path and Path(path).exists():
                width = Pt(block["width_pt"]) if block.get("width_pt") else None
                doc.add_picture(str(path), width=width)
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                stats["images"] += 1
            continue

        if kind == "page_break":
            # 用「分节」而不是「分页符」：每页独立控制尺寸与边距，
            # 某一页内容略超时也不会把后面所有页一起顶出去。
            section = doc.add_section(WD_SECTION.NEW_PAGE)
            _apply_to_section(section, page_cfg)
            try:
                section.header.is_linked_to_previous = True
                section.footer.is_linked_to_previous = True
            except Exception:
                pass
            stats["sections"] += 1
            continue

        explicit_break = text.startswith("#PAGEBREAK#")
        if explicit_break:
            doc.add_page_break()
            text = text.replace("#PAGEBREAK#", "")

        level = block.get("level") or 0
        para = doc.add_paragraph()
        if kind == "bullet":
            base = {"bold": False, "italic": False,
                    "size_pt": style.get("size_pt", 10.5),
                    "latin": style.get("latin", "Arial"),
                    "east_asian": style.get("east_asian", "微软雅黑")}
            add_runs(para, "•  " + text, base)
            set_paragraph_geometry(para,
                                   left_indent_pt=block.get("indent_pt", 18),
                                   first_line_pt=block.get("hanging_pt", -12),
                                   space_after_pt=block.get("space_after_pt", 2),
                                   line_pt=block.get("line_pt"))
            stats["bullets"] += 1
            continue

        base = {
            "bold": bool(style.get("bold")) or level >= 1,
            "italic": bool(style.get("italic")),
            "underline": bool(style.get("underline")),
            "size_pt": style.get("size_pt"),
            "color": style.get("color"),
            "latin": style.get("latin", "Arial"),
            "east_asian": style.get("east_asian", "微软雅黑"),
        }
        add_runs(para, text, base)
        set_paragraph_geometry(
            para,
            space_before_pt=style.get("space_before_pt", 0),
            space_after_pt=style.get("space_after_pt", 6),
            line_pt=style.get("line_pt"),
            line_multiple=style.get("line_multiple"),
            left_indent_pt=style.get("left_indent_pt"),
            first_line_pt=style.get("first_line_pt"),
        )
        if style.get("align"):
            para.alignment = {
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                "left": WD_ALIGN_PARAGRAPH.LEFT,
            }.get(style["align"], WD_ALIGN_PARAGRAPH.LEFT)
        if style.get("shade"):
            shade_paragraph(para, style["shade"])
        if style.get("border_bottom"):
            bottom_border(para, style["border_bottom"],
                          style.get("border_size", 4))
        if level >= 1:
            stats["headings"] += 1
        else:
            stats["paragraphs"] += 1

    return stats


# ------------------------------------------------------------------ Markdown 模式

def markdown_to_blocks(md: str, base: dict | None = None) -> dict:
    """Markdown → blocks。标题层级会按基准字号推出合理字号，
    否则 md 路线下所有标题都会塌成正文大小。"""
    base = base or {}
    body_size = float(base.get("size_pt") or 10.5)
    line_pt = round(body_size * 1.35, 2)
    # 标题相对正文的字号倍率（经验值，覆盖 H1..H6）
    scale = {1: 2.36, 2: 1.55, 3: 1.27, 4: 1.15, 5: 1.08, 6: 1.04}
    blocks: list[dict] = []
    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip()

        if not stripped.strip():
            i += 1
            continue

        # 表格
        if "|" in stripped and i + 1 < len(lines) and re.match(r"^\s*\|?[\s:|-]+\|", lines[i + 1]):
            rows: list[list[str]] = []
            header = [c.strip() for c in stripped.strip("|").split("|")]
            rows.append(header)
            i += 2
            while i < len(lines) and "|" in lines[i]:
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            blocks.append({"kind": "table", "rows": rows, "header_row": True})
            continue

        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            lvl = len(m.group(1))
            size = round(body_size * scale.get(lvl, 1.04), 1)
            blocks.append({
                "kind": "heading", "level": lvl, "text": m.group(2).strip(),
                "lines": 1,
                "style": {
                    "size_pt": size,
                    "bold": lvl <= 3,
                    "latin": base.get("latin", "Arial"),
                    "east_asian": base.get("east_asian", "微软雅黑"),
                    "line_pt": round(size * 1.28, 2),
                    "space_after_pt": round(size * 0.42, 1),
                },
            })
            i += 1
            continue

        m = BULLET_RE.match(line)
        if m:
            blocks.append({
                "kind": "bullet", "level": 0, "text": m.group(2).strip(),
                "lines": 1,
                "indent_pt": 18 + 18 * (len(m.group(1)) // 2),
                "hanging_pt": -18,
                "style": {"size_pt": body_size, "bold": False,
                          "latin": base.get("latin", "Arial"),
                          "east_asian": base.get("east_asian", "微软雅黑"),
                          "line_pt": line_pt, "space_after_pt": 2.0},
            })
            i += 1
            continue

        m = NUMBERED_RE.match(line)
        if m:
            blocks.append({
                "kind": "bullet", "level": 0,
                "text": f"{m.group(2)}. {m.group(3).strip()}",
                "lines": 1, "indent_pt": 18, "hanging_pt": -18,
                "style": {"size_pt": body_size, "bold": False,
                          "latin": base.get("latin", "Arial"),
                          "east_asian": base.get("east_asian", "微软雅黑"),
                          "line_pt": line_pt, "space_after_pt": 2.0},
            })
            i += 1
            continue

        # 段落：合并软换行
        buf = [stripped]
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(
                r"^\s*(#{1,6}\s|[-*+•]\s|\d+[.)、]\s|\|)", lines[i]):
            buf.append(lines[i].strip())
            i += 1
        blocks.append({
            "kind": "paragraph", "level": 0, "text": " ".join(buf),
            "lines": 1,
            "style": {"size_pt": body_size, "bold": False,
                      "latin": base.get("latin", "Arial"),
                      "east_asian": base.get("east_asian", "微软雅黑"),
                      "line_pt": line_pt, "space_after_pt": round(body_size * 0.6, 1)},
        })

    return {"blocks": blocks}


# ------------------------------------------------------------------ 主流程

def main() -> int:
    ap = argparse.ArgumentParser(description="结构规格 → .docx 重建")
    ap.add_argument("spec", help="blocks.json 或 .md")
    ap.add_argument("out", help="输出 .docx")
    ap.add_argument("--preset", default=None,
                    help="页面设置 JSON（width_pt/height_pt/margins_pt/字体）")
    args = ap.parse_args()

    spec_path = Path(args.spec)
    if not spec_path.exists():
        print(f"[错误] 不存在: {spec_path}", file=sys.stderr)
        return 2

    # 先定页面参数，因为 md 路线要靠它推标题字号
    page = {"width_pt": 595.28, "height_pt": 841.89,
            "margins_pt": {"top": 56.7, "bottom": 56.7, "left": 56.7, "right": 56.7},
            "latin": "Arial", "east_asian": "微软雅黑", "size_pt": 10.5}
    if args.preset:
        page.update(json.loads(Path(args.preset).read_text(encoding="utf-8")))

    if spec_path.suffix.lower() in (".md", ".markdown", ".txt"):
        spec = markdown_to_blocks(spec_path.read_text(encoding="utf-8"), page)
        spec["source"] = "markdown"
    else:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        if spec.get("page"):
            page.update(spec["page"])
        elif args.preset:
            pass
        else:
            page.update(spec.get("page") or {})

    spec["page"] = page

    doc = Document()
    apply_page_setup(doc, page)
    patch_default_font(doc, page.get("latin", "Arial"),
                       page.get("east_asian", "微软雅黑"),
                       page.get("size_pt", 10.5))

    spec = apply_height_budget(spec, page)
    stats = build_from_blocks(doc, spec)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    print(f"已生成: {out_path}")
    print(f"  段落={stats['paragraphs']}  标题={stats['headings']}  "
          f"列表={stats['bullets']}  表格={stats['tables']}  图片={stats['images']}  "
          f"分节={stats['sections']}")

    shrunk = [r for r in spec.get("_budget_report", []) if r.get("action") != "none"]
    for r in shrunk:
        print(f"  ⚠ 高度预算 p{r['page']}: 需 {r['need_pt']}pt / 目标 {r['target_pt']}pt "
              f"→ {r['action']} ×{r['factor']}")
    print(f"  页面={page['width_pt']:.0f}x{page['height_pt']:.0f}pt  "
          f"页边距 L{page['margins_pt']['left']:.0f} "
          f"R{page['margins_pt']['right']:.0f} "
          f"T{page['margins_pt']['top']:.0f} B{page['margins_pt']['bottom']:.0f}pt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
