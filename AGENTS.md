# Repository Rules

## Git commits

- Commit author/committer must be the repo owner: `Mahjong404 <surtrmuelsyse@163.com>`.
- NEVER add `Co-Authored-By: Devin`, `Generated with Devin`, or any
  `devin-ai-integration[bot]` identity to commit messages.
- Do not push to remotes unless the user explicitly asks.

## Layout

- `main.lua` — Yazi plugin entry (peek/preload/seek/entry), ext-driven dispatch.
- `render.py` — page-image pipeline (Office → PDF → JPEG) plus platform layer.
- `table.py` — ANSI grid renderer + text-table adapters (csv, docx tables).
- `xlsx.py` — XLSX adapter (openpyxl, colors/merges).
- `docx_text.py` — DOCX text extractor for text-mode preview.
- `doc_text.py` — legacy OLE text extractor (.doc/.dot/.rtf) via the Word server's Content.Text.
- `pptx_text.py` — PPTX slide-outline extractor (OOXML only) for text-mode/fallback preview.
- `shell_thumb.py` — Windows Explorer thumbnail fast path (ctypes IThumbnailCache, cache-only).
- `md.py` — Markdown renderer (headings/tables/code/math/diagrams), emits a `{text, media}` manifest.

## Testing

Verify changes in a real Yazi PTY (winpty + `yazi.exe`), not only unit-level:
check the preview pane content, the `N left` task counter, and Lua errors.
