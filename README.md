# omni-previewer.yazi

**English** | [简体中文](./README.zh-CN.md)

All-in-one document previewer for [Yazi](https://github.com/sxyazi/yazi), focused on Windows + Office formats.

## Features

### PDF

- Direct `pdftoppm` rendering scaled to the preview pane — oversized pages (e.g. giant mind-map exports) no longer hit the image-size limit.
- `J`/`K` page navigation; status-bar page indicator shows `p.N/M`.
- Page-image cache only (the source PDF is never duplicated).

### Word documents (`docx`, `doc`, `docm`, `dotx`, `dotm`, `rtf`)

- **Progressive preview**: extracted text appears instantly, then a rendered page image replaces it automatically once conversion finishes.
- Page-faithful rendering via Microsoft Word → PDF → `pdftoppm` (layout, images, tables preserved).
- `J`/`K` page navigation, `F6` toggle between page-image and text mode.
- Text mode extracts paragraphs and renders tables as bordered grids (`docx_text.py`, python-docx).
- A hidden, reusable Word automation service keeps warm-document conversion fast (~0.5 s) and exits after 10 minutes idle.
- Safety: read-only open, macros force-disabled (`AutomationSecurity=3`), ActiveX/`vbaProject`/reachable external links rejected, existing Word sessions are never touched.
- Bounded cache (`%LOCALAPPDATA%\yazi\docx-pages`): 64 MiB total, ≤3 page images per document, entries older than 3 days pruned.

### PowerPoint (`pptx`, `pptm`, `ppt`, `ppsx`, `ppsm`, `pps`, `potx`, `potm`, `pot`)

- Slide images via Microsoft PowerPoint → PDF → `pdftoppm`, sharing the Word service process (slides count = pages).
- `J`/`K` slide navigation; status-bar `p.N/M` indicator.
- Safety: read-only + `WithWindow=False` open, `AutomationSecurity=3`, `vbaProject`/ActiveX/external links rejected. If you already have PowerPoint open, the plugin attaches read-only and never quits your instance.
- Legacy `.ppt`/`.pps`/`.pot` (CFB binary) supported via the same COM path.

### Excel workbooks (`xlsx`, `xls`)

- Bordered, aligned table preview with correct CJK width handling.
- Real cell styles translated to ANSI: fill colors, font colors, bold.
- Merged cells rendered across their span; centered cells stay centered.
- Legacy `.xls` converted once via Excel COM to a cached `.xlsx`, then full-fidelity rendering.
- Per-file output cache for instant scrolling (`%LOCALAPPDATA%\yazi\preview-cache`).

### CSV

- Bordered table preview with a bold header row, via the shared `table.py` grid renderer.
- Encoding sniffed: UTF-8 (with/without BOM) and GB18030.

### Markdown (`md`, `markdown`)

- **Rendered view** by default; `F6` toggles rendered / raw source.
- Headings, bold/italic/strikethrough, inline code, links, lists, block quotes, horizontal rules.
- Fenced code blocks keep full Pygments syntax highlighting (e.g. `asm`, `python`, `rust`).
- Markdown tables render as bordered grids with row separators; tables too wide for the pane fall back to a per-record layout instead of breaking.
- Math: `$...$` and `$$...$$` are converted to Unicode text (fractions, roots, Greek letters, super/subscripts); complex constructs (matrices, `\begin{...}` environments, etc.) render as images **inline with the text** (half-block truecolor, scrolls naturally).
- PlantUML fences render as **Unicode text diagrams** via `plantuml.jar -utxt` (transparent-PNG → inline-image fallback for diagram types utxt cannot express).
- Mermaid `flowchart`/`graph`/`sequenceDiagram` render via `mmdc` (Puppeteer/Chromium) as inline images; without `mmdc` they fall back to a text arrow rendering.
- Images (`![alt](path)`, relative paths resolved against the document) render **inline as half-block truecolor text** — they scroll with the document. Remote URLs show a placeholder.
- `==highlight==`, `**bold**`, `~~strike~~`, inline code, and list items with hanging-indent wrapping.

## Requirements

- Windows + Microsoft Word installed (for Word pipeline); PowerPoint for `ppt*`; Excel for `.xls` conversion
- Python 3 with `openpyxl`, `python-docx`, `pywin32`, `psutil` (`render.py`/`xlsx.py`/`table.py`/`docx_text.py`/`md.py` run via `python.exe`)
- [Poppler](https://github.com/oschwartz10612/poppler-windows) (`pdftoppm.exe`)
- [Pandoc](https://pandoc.org/) (text fallback, via the `docx-preview.yazi` plugin)
- Optional: Java + `plantuml.jar` (PlantUML text rendering), `matplotlib` + `pillow` (complex math / inline images), `mmdc` (Mermaid diagram PNG, `npm i -g @mermaid-js/mermaid-cli`), `pygments` (code-block highlighting)

## Installation

```sh
ya pkg add Mahjong404/omni-previewer.yazi
```

`yazi.toml`:

```toml
[[plugin.prepend_previewers]]
url = "*.{docx,DOCX,doc,DOC,docm,DOCM,dotx,DOTX,dotm,DOTM,dot,DOT,rtf,RTF,pptx,PPTX,pptm,PPTM,ppt,PPT,ppsx,PPSX,ppsm,PPSM,pps,PPS,potx,POTX,potm,POTM,pot,POT}"
run = "omni-previewer"

[[plugin.prepend_previewers]]
url = "*.{xlsx,XLSX,xls,XLS,csv,CSV,md,markdown}"
run = "omni-previewer"

[[plugin.prepend_previewers]]
url = "*.{pdf,PDF}"
run = "omni-previewer"

[[plugin.prepend_preloaders]]
url = "*.{docx,DOCX,doc,DOC,docm,DOCM,dotx,DOTX,dotm,DOTM,dot,DOT,rtf,RTF,pdf,PDF,pptx,PPTX,pptm,PPTM,ppt,PPT,ppsx,PPSX,ppsm,PPSM,pps,PPS,potx,POTX,potm,POTM,pot,POT}"
run = "omni-previewer"
```

`keymap.toml` (optional, page-image/text toggle):

```toml
[[mgr.prepend_keymap]]
on = "<F6>"
run = "plugin omni-previewer"
desc = "Toggle Word page-image / text preview"
```

Adjust the `PYTHON` constant at the top of `main.lua` to your interpreter.

## Roadmap

See [docs/ROADMAP.md](./docs/ROADMAP.md) for the phased plan.

- `.ipynb` notebook rendering; `.git` repo info (log/status summary) support
- Archive (`zip`/`rar`) preview enhancements: speed + richer listing
- macOS/Linux support: the Word pipeline currently relies on Windows COM; a cross-platform port would use headless LibreOffice (`soffice --convert-to pdf`). The table pipelines are already portable.

## License

MIT
