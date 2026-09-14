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

### Excel workbooks (`xlsx`)

- Bordered, aligned table preview with correct CJK width handling.
- Real cell styles translated to ANSI: fill colors, font colors, bold.
- Merged cells rendered across their span; centered cells stay centered.
- Per-file output cache for instant scrolling (`%LOCALAPPDATA%\yazi\preview-cache`).

### CSV

- Bordered table preview with a bold header row, via the shared `table.py` grid renderer.
- Encoding sniffed: UTF-8 (with/without BOM) and GB18030.

## Requirements

- Windows + Microsoft Word installed (for Word pipeline)
- Python 3 with `openpyxl`, `python-docx`, `pywin32`, `psutil` (`render.py`/`xlsx.py`/`table.py`/`docx_text.py` run via `python.exe`)
- [Poppler](https://github.com/oschwartz10612/poppler-windows) (`pdftoppm.exe`)
- [Pandoc](https://pandoc.org/) (text fallback, via the `docx-preview.yazi` plugin)

## Installation

```sh
ya pkg add Mahjong404/omni-previewer.yazi
```

`yazi.toml`:

```toml
[[plugin.prepend_previewers]]
url = "*.{docx,DOCX,doc,DOC,docm,DOCM,dotx,DOTX,dotm,DOTM,rtf,RTF}"
run = "omni-previewer"

[[plugin.prepend_previewers]]
url = "*.{xlsx,XLSX,csv,CSV}"
run = "omni-previewer"

[[plugin.prepend_preloaders]]
url = "*.{docx,DOCX,doc,DOC,docm,DOCM,dotx,DOTX,dotm,DOTM,rtf,RTF}"
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

- More formats: `pptx`, `ppt`, `xls` quick previews
- `.ipynb` notebook rendering; `.git` repo info (log/status summary) support
- Archive (`zip`/`rar`) preview enhancements: speed + richer listing
- macOS/Linux support: the Word pipeline currently relies on Windows COM; a cross-platform port would use headless LibreOffice (`soffice --convert-to pdf`). The table pipelines are already portable.

## License

MIT
