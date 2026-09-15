# Roadmap — omni-previewer

Target: fast previews for docx/doc/xlsx/pptx/ppt/xls/csv/md, plus .ipynb/.git
support and pdf/zip/rar display enhancements. Windows-first; cross-platform
support comes last, but every phase keeps the portability seams clean.

## Status legend

- ✅ done · 🚧 in progress · ⬜ planned · ⏸ deferred

## Phase 0 — Abstraction groundwork ✅

| Seam | Where | Purpose |
|---|---|---|
| Table-driven dispatch | `main.lua` `PAGE_EXTS` / `GRID_SCRIPTS` / `MD_EXTS` | New format = one table entry |
| `convert_to_pdf()` | `render.py` | Office→PDF backend seam; LibreOffice backend slots in here |
| Platform layer | `render.py` header comment | CACHE dir, msvcrt lock, pipe/mutex, subprocess flags — the only block POSIX port touches |

## Phase 1 — Zero-new-dependency formats ✅

| Format | Backend | Status |
|---|---|---|
| CSV | `table.py` grid renderer, csv module, utf-8-sig/gb18030 sniff | ✅ |
| DOCX text-mode tables | `docx_text.py` (python-docx) → boxed grids via `table.py` | ✅ |
| MD | `md.py` full renderer: headings/tables/code/math/diagrams + inline media overlays | ✅ (far beyond "built-in code previewer") |

## Phase 2 — Office pipeline extension ✅

| Format | Backend | Status |
|---|---|---|
| PPTX/PPT/PPS/POT | PowerPoint COM → `Presentation.Export` 直出 PNG（PDF 回退） | ✅ |
| XLS | Excel COM → 一次性转 `.xlsx` 缓存 → `xlsx.py` 全保真渲染 | ✅ |
| PDF | `pdftoppm` 按预览区尺寸渲染页图 | ✅ |

Word 服务已泛化为 Office 服务（`--serve` 单例 + `Global\yazi-docx-preview-svc`
mutex + `\\.\pipe\yazi-docx-svc` 命名管道，按 kind 分派 word/ppt）。
进程生命周期：`.owned-{name}-{pid}-{created}` 台账 + `sweep_stale_servers()`
兜底回收——owner python 已死但 Office 进程仍在的孤儿会被强杀。

## Phase 3 — Pure-parser formats ⏸ (deferred)

| Format | Backend |
|---|---|
| IPYNB | stdlib `json`: markdown cells rendered, code cells highlighted, outputs truncated |
| .git | `git log --oneline -10` + `status --short` + branch info panel on hover |

## Phase 4 — Enhancements to existing formats ⏸ (deferred)

- **PDF**: `pdftotext` text mode + F6 toggle for PDFs + outline/TOC extraction.
- **ZIP/RAR**: unified `7z l` backend — structured listing with size/ratio/count.

## Phase 5 — Cross-platform ⬜

Replace three seams only:

1. `convert_to_pdf` → `soffice --headless --convert-to pdf` (fidelity < Word COM).
2. Cache dir → `~/.cache/yazi`.
3. Pipe/mutex/lock → Unix socket / `fcntl.flock` / no `CREATE_NO_WINDOW`.

`table.py`/`docx_text.py`/`xlsx.py`/`md.py` are already platform-neutral.
