# omni-previewer.yazi

All-in-one document previewer for [Yazi](https://github.com/sxyazi/yazi), focused on Windows + Office formats.

## Features

### Word documents (`docx`, `doc`, `docm`, `dotx`, `dotm`, `rtf`)

- **Progressive preview**: extracted text appears instantly, then a rendered page image replaces it automatically once conversion finishes.
- Page-faithful rendering via Microsoft Word → PDF → `pdftoppm` (layout, images, tables preserved).
- `J`/`K` page navigation, `F6` toggle between page-image and text mode.
- A hidden, reusable Word automation service keeps warm-document conversion fast (~0.5 s) and exits after 10 minutes idle.
- Safety: read-only open, macros force-disabled (`AutomationSecurity=3`), ActiveX/`vbaProject`/reachable external links rejected, existing Word sessions are never touched.
- Bounded cache (`%LOCALAPPDATA%\yazi\docx-pages`): 128 MiB total, ≤3 page images per document, entries older than 3 days pruned.

### Excel workbooks (`xlsx`)

- Bordered, aligned table preview with correct CJK width handling.
- Real cell styles translated to ANSI: fill colors, font colors, bold.
- Merged cells rendered across their span; centered cells stay centered.
- Per-file output cache for instant scrolling (`%LOCALAPPDATA%\yazi\xlsx-cache`).

## Requirements

- Windows + Microsoft Word installed (for Word pipeline)
- Python 3 with `openpyxl`, `pywin32`, `psutil` (`render.py`/`xlsx.py` run via `python.exe`)
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
url = "*.{xlsx,XLSX}"
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

- More formats: `pptx`, `ppt`, `xls`, `csv`, `md` quick previews
- `.ipynb` notebook rendering; `.git` repo info (log/status summary) support
- PDF preview enhancements: faster first paint, richer page features
- Archive (`zip`/`rar`) preview enhancements: speed + richer listing
- `docx-preview` fallback independence (inline text extraction)

## License

MIT
