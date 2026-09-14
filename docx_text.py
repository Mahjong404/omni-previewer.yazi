import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import table
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

BOLD, DIM, RESET = "\x1b[1m", "\x1b[90m", "\x1b[0m"
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def para_line(p, out):
    text = p.text.strip()
    style = (p.style.name if p.style is not None else "") or ""
    if not text:
        out.append("")
        return
    if style.startswith(("Heading", "标题")):
        out.append(BOLD + text + RESET)
    else:
        out.append(text)


def table_block(tbl, out, max_width=0):
    rows = []
    prev_tcs = set()
    for row in tbl.rows[: table.MAX_ROWS]:
        cells, spans, tcs = [], [], set()
        for cell in row.cells:
            tc = cell._tc
            if tcs and tc in tcs:
                spans[-1] += 1
                continue
            tcs.add(tc)
            if tc in prev_tcs:
                text = ""
            else:
                text = " ".join(" ".join(p.text.split()) for p in cell.paragraphs).strip()
            cells.append(text)
            spans.append(1)
        prev_tcs = tcs
        rows.append((cells[: table.MAX_COLS], spans[: table.MAX_COLS]))
    out.extend(table.grid_lines(rows, max_width=max_width).splitlines())


def extract(path, max_width=0):
    doc = Document(path)
    out = []
    for el in doc.element.body:
        if el.tag == W_NS + "p":
            para_line(Paragraph(el, doc), out)
        elif el.tag == W_NS + "tbl":
            table_block(Table(el, doc), out, max_width)
    lines, i = [], 0
    while i < len(out):
        if out[i] == "" and (not lines or lines[-1] == ""):
            i += 1
            continue
        lines.append(out[i])
        i += 1
    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines)


if __name__ == "__main__":
    try:
        text = extract(sys.argv[1], int(sys.argv[3]) if len(sys.argv) > 3 else 0)
        if len(sys.argv) > 2:
            table.write_cache(text, sys.argv[2])
        sys.stdout.write(text)
    except Exception as exc:
        print("DOCX text extraction failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
