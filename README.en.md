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

**Text-based**: a 6-page résumé (612×792pt), rebuilt and rendered back for page-by-page comparison:

| Page | Text similarity | Pixel diff | Verdict |
|---|---|---|---|
| 1 | **1.000** | 10.8% | pass |
| 2 | **1.000** | 9.8% | pass |
| 3 | **1.000** | 7.1% | pass |
| 4 | 0.998 | 7.2% | pass |
| 5 | 0.956 | 8.5% | warn |
| 6 | 0.992 | 3.8% | pass |

**Page count 6 → 6. Text similarity 0.956–1.000. Zero failures.**

**Image-based (scan)**: a 5-page US bankruptcy notice (552×765pt, **no text layer at all**):

| Page | Text similarity | Pixel diff |
|---|---|---|
| 1–5 | N/A (no baseline) | 8.3% / 8.7% / 9.0% / 8.5% / 4.1% |

Page numbers, the `Page N of 5` running header, bordered boxes, a grey shaded block,
underlined `D.I. NNN` citations, tab-aligned label/value pairs, a two-column signature
block and a footnote were all reproduced. The original's page 6 is a blank scan back.

Pixel diff can never be zero — line breaks, font rasterization and micro-adjustments all
contribute. So **text similarity is the hard metric (when a text layer exists); pixel diff
is reference only.**

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

**But step 1 hides a speed trap**: treating "looking at pages" as the *primary* information
channel. Looking is expensive (every look is a multimodal round-trip); it should be used for
**spot checks**, not wholesale reading. See "Efficiency discipline" below.

---

## Features

- **Type-aware triage**: classifies `text` / `mixed` / `image` (scans) / `vector` and picks the matching route
- **Dual-channel for image PDFs**: built-in OCR (RapidOCR, pure Python, no external binary) supplies coordinates; visual reading supplies semantics
- **Contact-sheet proofing**: `verify_sheet.py` tiles every line you need to eyeball into one
  image — measured 35 crop-and-look calls down to 3
- **Pre-render budget check**: `prespec.py` computes by pure arithmetic which page will
  overflow, saving an entire render iteration
- **Built-in height budget**: detects "full pages" whose content hits the ceiling and shrinks space-after so pagination doesn't drift
- **Quantitative verification**: per-page text similarity, pixel diff, and side-by-side comparison images (original / result / diff heatmap); sources with no text layer get `N/A` for similarity and are judged on layout alone
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
| Optional | `pdfplumber` `rapidocr-onnxruntime` `opencv-python-headless` `numpy` | Weaker table extraction; no OCR for scans |

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

### Step by step (when you need fine control)

```bash
# 0. Triage: what kind of PDF is this, which route applies
python scripts/probe.py in.pdf --json probe.json

# 1. Text-based: extract structure directly (incl. tables and images)
python scripts/extract.py in.pdf --out blocks.json --images imgs/

# 2. Image-based: run BOTH channels in parallel — page images + OCR text/coords
python scripts/render_pages.py in.pdf --outdir pages/ --dpi 160 --tile
python scripts/ocr.py in.pdf --out ocr.json --dpi 200
python scripts/ocr_to_spec.py ocr.json in.pdf --out blocks.json   # OCR draft

# 3. Proofing: tile the lines you must eyeball into ONE image (don't crop one by one)
python scripts/verify_sheet.py in.pdf --auto 1,3 --outdir sh/
python scripts/verify_sheet.py in.pdf --at "1:93:112:P1 caption" --outdir sh/
#   → sh/sheet_001.png (with #NN gauge) + sh/verify_map.json (index → page/y/label)

# 4. Pre-check: compute the height budget before rendering (exit 1 = a page overflows)
python scripts/prespec.py blocks.json --preset page.json -v

# 5. Build
python scripts/build_docx.py blocks.json out.docx
python scripts/build_docx.py rebuild.md  out.docx --preset page.json

# 6. Verify: render back to PDF and diff page by page
python scripts/render_docx.py out.docx out.pdf
python scripts/compare.py in.pdf out.pdf --outdir diff/
```

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

**Do that proofing in batches via `verify_sheet.py`.** OCR mangles technical terms, so a
human-grade look is required — but crop everything once and read it once. Counter-example
from a real 5-page run: cropping and reading line by line cost 35 calls, six of them spent
on a single word (`Liquidating` vs `Liqudating`) — and the earliest low-resolution crop got
it **wrong**.

---

## Efficiency discipline

A full 5-page scanned-document rebuild measured **50m18s, 118 tool calls, ~40 model turns**.
The retrospective was unambiguous — the bottleneck was not any slow script, it was **too many
round-trips**:

| Phase | Calls | Share |
|---|---|---|
| Triage + reading full pages | 13 | 11% |
| **Measurement + line-by-line proofing** | **~63** (28 crops + 35 looks) | **53%** |
| Modelling | 3 | 3% |
| **Render iterations** | **~25** (7 Word renders) | **21%** |
| Wrap-up | 5 | 4% |

Four rules (by payoff); the first three are now scripts:

1. **Let OCR produce the full draft first**; your eyes only adjudicate the handful of tokens
   it gets wrong.
2. **Proof from a contact sheet** (`verify_sheet.py`): 35 looks → 3.
3. **Write the measurement script once.** Emit everything in one pass (per-page line bands,
   per-block line counts, line/paragraph spacing distributions, left/right extremes, font
   size inference). Query afterwards — never re-derive.
4. **Run `prespec.py` before rendering.** Overflow is an arithmetic problem, not a rendering
   problem. Render iterations: 7 → 2.

**One thing you must not skip**: visual evidence needs **enough resolution, and one correct
look**. The same text read as `Liqudating` at low resolution and `Liquidating` at high
resolution. The lesson is not "look less" — it is "look properly, in bulk".

---

## Verification criteria

| Metric | Answers | Target |
|---|---|---|
| `text_similarity` | Is content complete and correct? | ≥ 0.995 ideal; ≥ 0.98 acceptable |
| `pixel_hot_ratio` | How close is the layout? | ≤ 0.20 normal |
| Page count | Did pagination drift? | Must match |

> **For image-based / scanned sources `text_similarity` is always `N/A`**: the original has no
> text layer, so there is no baseline. `compare.py` marks it `N/A` and judges on **layout
> closeness alone** (≤10% pass, ≤25% warn). If you see `0.000` instead of `N/A`, you're on an
> old script — that is a "cannot evaluate" masquerading as "content is entirely wrong".

`compare.py` emits a side-by-side image per page: **left = original, middle = result,
right = diff heatmap**.

---

## Real pitfalls (all fixed in code)

Every one of these was located **quantitatively**, not guessed:

- **Don't infer the bottom margin from the last line of one page.** Sparse pages can't fill the page; that reflects content length, not margins. Measured: single-page inference gave 136pt vs the true 72pt — 56pt of usable height lost, content overflowed. **Aggregate extremes across pages instead.**
- **Use a low percentile (p25), not the median, for the intra-paragraph line-gap baseline.** On dense pages inter-paragraph gaps can dominate; the median lands on them. Measured: line spacing computed as 24.72pt vs true 14.88pt — 100pt+ of phantom height.
- **Line spacing ≠ line height × 1.32.** Measured 17.03 vs true 14.88 — enough to push a full page onto the next. Measure the **delta between adjacent line `y0`s**.
- **Solve space-after from the whole line-box span, not from ink tops.** Ink tops depend on the first glyph (capitals, parentheses, ascender-less letters all differ); one file yielded 12.8 / 14.7 / 15.7. Solving `lines × line_pt + (blocks − 1) × space_after = line-box span` converges on 12.8.
- **Space-after belongs to the *previous* paragraph.** Reverse it and headings stick to their body text.
- **PDFs often put the bullet glyph on its own line.** Detecting lists by leading character alone yields orphan `•` rows; merge by geometry (bullet row `x0` smaller; the delta is the hanging indent).
- **Use section breaks, not page breaks.** Once one page overflows, page breaks cascade the misalignment through the whole document; sections cut the chain.
- **Scanned pages usually drift relative to each other — never define the text block from one page.** Page numbers are centred, so each page's offset can be solved from its page-number centre; once aligned, every element's left edge should land on an **evenly spaced indent ladder**. Measured: odd/even pages off by ±23.5pt; after alignment the ladder landed cleanly on 44/62/80/98/116 (18pt = 0.25″ steps).
- **Word does not kern by default; the source document often does.** Each line renders ~0.5% wider, and source lines that hug the right edge (under 1pt of slack) then break a word early, gain a line, and overflow the page. Fix: add `<w:kern w:val="8"/>` to `w:rPr` (**and to docDefaults**). If that isn't enough, widen the text block a few pt — **matching the text width matters more than matching the margin**.
- **The template's leftover empty paragraph inside a table cell destroys box geometry.** `cell.paragraphs[0]` carries the Normal style (1.15 line spacing + 8pt space-after); leaving it inflates a bordered box (measured 84pt → 126pt). Delete it before filling.
- **Paragraphs that span pages must be split at the original's break point.** Original PDFs break mid-paragraph; putting the whole paragraph in the earlier section pushes its last lines to the next page and misaligns everything after.
- **Two-column signature blocks as a table will jump pages.** When Word can't split a row it moves the whole table to the next page; use paragraphs with a left indent plus a tab stop instead.
- **Labels and values are tab-aligned, not space-aligned.** The original uses Word's default 0.5″ tab grid (from the left margin): `Objection Deadline:` → 144pt, `Responses:` → 108pt, `Status:` needs two spaces to reach 108pt. Underline only the label, never the padding spaces.
- **In PowerShell, don't name a parameter `$Input`** — it's an automatic variable and you'll get an empty value.
- **Under `pwsh -File`, don't rely on `ExportAsFixedFormat` argument binding** — 12 optional args bind unreliably; use `SaveAs2($path, 17)`.
- **PyMuPDF ≥1.28 `get_pixmap(dpi=)` accepts int only** — a float raises TypeError.
- **Windows console mojibake**: call `sys.stdout.reconfigure(encoding="utf-8")` at script start.

---

## Layout

```
.
├── SKILL.md                  # Agent skill definition (methodology / workflow / criteria / pitfalls)
├── README.md                 # Chinese README
├── README.en.md              # English README (this file)
├── CHANGELOG.md
├── LICENSE                   # MIT
└── scripts/
    ├── _kit.py               # Base: unit conversion, font mapping, OOXML helpers
    ├── _runtime.py           # Runtime probing: deps, renderer, paths (distribution-critical)
    ├── bootstrap.py          # One-shot install & self-check
    ├── pdf2docx.py           # Public entry point: whole pipeline in one command
    ├── probe.py              # Phase 0  Triage
    ├── extract.py            # Phase 1  Text-based structure extraction
    ├── render_pages.py       # Phase 1  Image-based: page images for visual reading
    ├── ocr.py                # Phase 1  Image-based: OCR text + coordinates
    ├── ocr_to_spec.py        # Phase 2  OCR result → structured spec
    ├── verify_sheet.py       # Phase 2  Contact sheet: lines to proof → one image
    ├── prespec.py            # Phase 3  Pre-render height/line-wrap check (no Word)
    ├── build_docx.py         # Phase 4  Build docx (with height budget)
    ├── render_docx.py        # Phase 5  Render back to PDF (picks renderer)
    ├── render_docx.ps1       # Phase 5  Word COM branch
    ├── compare.py            # Phase 5  Per-page diff and heatmaps
    ├── measure_ink.py        # Metrics: page ink extremes and line bands
    ├── diag_lines.py         # Dev: line geometry and spacing distributions
    ├── diag_fit.py           # Dev: per-page content height comparison
    └── make_scan_fixture.py  # Dev: rasterize any PDF into a scan fixture
```

---

## Known limitations

- Reading order in complex multi-column layouts with cross-column tables can still be wrong; image-based PDFs need page-by-page review.
- Math is flattened to linear text; no equation-object rebuild.
- Cross-page table splitting follows Word's own behavior and may differ from the original.
- For graphics-heavy pages (drawings, posters), "editable" vs "visually faithful" is a genuine trade-off.
- Lines that hug the right edge to within a fraction of a point are extremely sensitive to
  renderer differences; a machine or font-version change may shift a line or two. That is
  inherent to text layout engines, not a reconstruction error.
- **Check the page count yourself**: `compare.py` only diffs page-by-page when both sides
  agree on length. A blank scan back on one side, or one page missing on the other, misaligns
  the whole table — settle the page count first, then look at similarity.

---

## License

[MIT](LICENSE)