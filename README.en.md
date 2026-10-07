# PDF → Word · High-Fidelity Reconstruction

> Not a PDF *converter*. A PDF **reconstructor**.
> An AI-agent skill, and a standalone CLI pipeline.

A PDF is a **snapshot of laid-out output** — a pile of glyphs painted at coordinates.
Any tool that "reads text out and stuffs it into Word" throws away paragraphs, lists,
tables, margins and line spacing. The result is *right words, wrong layout*.

This project takes a different path: **read the layout visually → rebuild real OOXML
semantics → render back to images and diff page by page.**

- For agents: `SKILL.md` defines the full methodology and workflow
- For humans: `python scripts/pdf2docx.py in.pdf out.docx`

[中文说明](README.md)

---

## Measured results

A 6-page text-based résumé (612×792pt), rebuilt and rendered back for page-by-page comparison:

| Page | Text similarity | Pixel diff | Verdict |
|---|---|---|---|
| 1 | **1.000** | 10.8% | pass |
| 2 | **1.000** | 9.8% | pass |
| 3 | **1.000** | 7.1% | pass |
| 4 | 0.998 | 7.2% | pass |
| 5 | 0.956 | 8.5% | warn |
| 6 | 0.992 | 3.8% | pass |

**Page count 6 → 6. Text similarity 0.956–1.000. Zero failures.**

Pixel diff can never be zero — line breaks, font rasterization and micro-adjustments all
contribute. So **text similarity is the hard metric; pixel diff is reference only.**

---

## The core idea

Most PDF→Word tools fail the same way: they treat a PDF as a *data source* instead of a *picture*.

| # | Mechanism | Why it matters |
|---|---|---|
| 1 | **Read the PDF visually** | Layout is visual information. "Light-grey 3-column table", "hanging-indent list" — only visible by looking |
| 2 | **Reconstruct, don't transcode** | Rebuild paragraphs, headings, lists and *real* tables semantically |
| 3 | **Write OOXML directly** | Precise control over page size, margins, font size, line spacing, space-after, eastAsia fonts |
| 4 | **Render → diff → iterate** | Render back to images, compare against the original, fix what drifted |

Step 4 is the one most often skipped — and it's the watershed for "does the format survive".

---

## Features

- **Type-aware triage**: classifies `text` / `mixed` / `image` (scans) / `vector` and picks the matching route
- **Dual-channel for image PDFs**: built-in OCR (RapidOCR, pure Python, no external binary) supplies coordinates; visual reading supplies semantics
- **Built-in height budget**: detects "full pages" whose content hits the ceiling and shrinks space-after so pagination doesn't drift
- **Quantitative verification**: per-page text similarity, pixel diff, and side-by-side comparison images (original / result / diff heatmap)
- **Zero hard-coded paths**: renderer auto-degrades Word → LibreOffice → skip visual verification
- **No external converters**: does not shell out to pdf2zh, pandoc, or similar

---

## Install

Python 3.9+. Copy this repo anywhere, then:

```bash
python scripts/bootstrap.py                                  # check only
python scripts/bootstrap.py --install                        # install what's missing
python scripts/bootstrap.py --install --with-ocr --mirror    # include OCR (China mirror)
```

| Tier | Packages | If missing |
|---|---|---|
| Required | `pymupdf` `Pillow` `python-docx` | Won't run |
| Optional | `pdfplumber` `rapidocr-onnxruntime` `opencv-python-headless` | Weaker table extraction; no OCR for scans |

The renderer (used to render docx back to PDF for verification) is auto-detected:

```
Microsoft Word (COM, Windows, highest fidelity)
   ↓ absent
LibreOffice (headless, cross-platform)
   ↓ absent
Skip visual verification — extraction and docx generation still work
```

Override via env vars: `PDF2DOCX_WORD`, `PDF2DOCX_SOFFICE`, `PDF2DOCX_RENDERER`.

---

## Usage

```bash
python scripts/pdf2docx.py input.pdf output.docx
```

Chains triage → extraction → rebuild → render → verify, keeps intermediates in `_pdf2docx/`,
and prints a comparison table.

```bash
--pages 1-3     # process a subset first
--workdir W     # intermediate directory
--no-verify     # skip render verification
```

Step-by-step commands are documented in the [Chinese README](README.md#分步使用要精细控制时).

---

## Image-based / scanned PDFs

The hardest case, and where most of the effort went.

**OCR and vision each have hard limits — they must be fused.**

| Channel | Strong at | Weak at |
|---|---|---|
| Visual reading (agent looks at the page) | Hierarchy, columns, table membership, reading order; fixing terminology and mixed CJK/Latin | No precise coordinates, font sizes or colors |
| OCR / geometry extraction | Accurate coordinates and sizes | Drops glyphs, swallows spaces, mangles technical terms |

Measured example: OCR read `John Ka-Kit Leung (梁嘉傑)` as `JohnKa-Kit Leung(梁嘉)` —
**dropped 「傑」 and swallowed spaces** — while looking at the page reads it correctly
but cannot tell you it sits at `y=75.8`.

> **Fusion rule: vision decides *what it is*; OCR/geometry decides *where and how big*.**

For image PDFs, `pdf2docx.py` produces an **OCR draft** and states its quality boundary
explicitly. Approaching original quality requires an agent to follow the
"visual reading → Markdown rebuild" path in `SKILL.md`.

---

## Verification criteria

| Metric | Answers | Target |
|---|---|---|
| `text_similarity` | Is content complete and correct? | ≥ 0.995 ideal; ≥ 0.98 acceptable |
| `pixel_hot_ratio` | How close is the layout? | ≤ 0.20 normal |
| Page count | Did pagination drift? | Must match |

---

## Real pitfalls (all fixed in code)

Every one of these was located **quantitatively**, not guessed:

- **Don't infer the bottom margin from the last line of one page.** Sparse pages can't fill the page; that reflects content length, not margins. Measured: single-page inference gave 136pt vs the true 72pt — 56pt of usable height lost, content overflowed. **Aggregate extremes across pages instead.**
- **Use a low percentile (p25), not the median, for the intra-paragraph line-gap baseline.** On dense pages inter-paragraph gaps can dominate; the median lands on them. Measured: line spacing computed as 24.72pt vs true 14.88pt — 100pt+ of phantom height.
- **Line spacing ≠ line height × 1.32.** Measured 17.03 vs true 14.88 — enough to push a full page onto the next. Measure the **delta between adjacent line `y0`s**.
- **Space-after belongs to the *previous* paragraph.** Reverse it and headings stick to their body text.
- **PDFs often put the bullet glyph on its own line.** Detecting lists by leading character alone yields orphan `•` rows; merge by geometry (bullet row `x0` smaller; the delta is the hanging indent).
- **Use section breaks, not page breaks.** Once one page overflows, page breaks cascade the misalignment through the whole document; sections cut the chain.
- **In PowerShell, don't name a parameter `$Input`** — it's an automatic variable and you'll get an empty value.
- **Under `pwsh -File`, don't rely on `ExportAsFixedFormat` argument binding** — 12 optional args bind unreliably; use `SaveAs2($path, 17)`.
- **PyMuPDF ≥1.28 `get_pixmap(dpi=)` accepts int only** — a float raises TypeError.
- **Windows console mojibake**: call `sys.stdout.reconfigure(encoding="utf-8")` at script start.

---

## Known limitations

- Reading order in complex multi-column layouts with cross-column tables can still be wrong; image-based PDFs need page-by-page review.
- Math is flattened to linear text; no equation-object rebuild.
- Cross-page table splitting follows Word's own behavior and may differ from the original.
- For graphics-heavy pages (drawings, posters), "editable" vs "visually faithful" is a genuine trade-off.

---

## License

[MIT](LICENSE)
