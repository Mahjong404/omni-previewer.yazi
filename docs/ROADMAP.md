# Roadmap — omni-previewer

Target: fast previews for docx/doc/xlsx/pptx/ppt/xls/csv/md, plus .ipynb/.git
support and pdf/zip/rar display enhancements. Windows-first; cross-platform
support comes last, but every phase keeps the portability seams clean.

## Status legend

- ✅ done · 🚧 in progress · ⬜ planned

## Phase 0 — Abstraction groundwork ✅

Small refactors that change no behavior but create the seams later work needs:

| Seam | Where | Purpose |
|---|---|---|
| Table-driven dispatch | `main.lua` `PAGE_EXTS` / `GRID_SCRIPTS` | New format = one table entry |
| `convert_to_pdf()` | `render.py` | Office→PDF backend seam; LibreOffice backend slots in here |
| Platform layer | `render.py` header comment | CACHE dir, msvcrt lock, pipe/mutex, subprocess flags — the only block POSIX port touches |

## Phase 1 — Zero-new-dependency formats 🚧

| Format | Backend | Status |
|---|---|---|
| CSV | `table.py` grid renderer, csv module, utf-8-sig/gb18030 sniff | ✅ |
| DOCX text-mode tables | `docx_text.py` (python-docx) → boxed grids via `table.py` | ✅ |
| MD | built-in `code` previewer already covers it | ✅ (no work needed) |

## Phase 2 — Office pipeline extension ⬜

| Format | Backend | Notes |
|---|---|---|
| PPTX | PowerPoint COM → `ExportAsFixedFormat(pdf)` → existing page pipeline | Reuses cache/service/notify wholesale |
| PPT | same PowerPoint service (COM opens legacy binary) | extension entry only after pptx works |
| XLS | Excel COM → cell reads → `table.py` grid | xlrd is the pure-Python alternative but unmaintained |

Phase 2 generalizes the Word service into an Office service (per-app dispatch)
and renames "pages" concepts to cover slides.

## Phase 3 — Pure-parser formats (naturally portable) ⬜

| Format | Backend |
|---|---|
| IPYNB | stdlib `json`: markdown cells rendered, code cells highlighted, outputs truncated |
| .git | `git log --oneline -10` + `status --short` + branch info panel on hover |

## Phase 4 — Enhancements to existing formats ⬜

- **PDF**: `pdftotext` text mode + F6 toggle for PDFs + outline/TOC extraction.
- **ZIP/RAR**: unified `7z l` backend — structured listing with size/ratio/count.

## Phase 5 — Cross-platform ⬜

Replace three seams only:

1. `convert_to_pdf` → `soffice --headless --convert-to pdf` (fidelity < Word COM).
2. Cache dir → `~/.cache/yazi`.
3. Pipe/mutex/lock → Unix socket / `fcntl.flock` / no `CREATE_NO_WINDOW`.

`table.py`/`docx_text.py`/`xlsx.py` are already platform-neutral.
