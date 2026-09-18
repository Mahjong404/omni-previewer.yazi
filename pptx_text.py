"""Immediate PPTX text preview: slide titles and body text straight from the
OOXML package, no PowerPoint instance. Same CLI/cache/stdout contract as
docx_text.py - the plugin shows this outline while (or instead of) the
high-fidelity slide images render."""
import os
import sys
import zipfile
from xml.etree import ElementTree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import table

BOLD, DIM, RESET = "\x1b[1m", "\x1b[90m", "\x1b[0m"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
MAX_SLIDES = 300

TITLE_TYPES = {"title", "ctrTitle", "subTitle"}


def slide_order(z):
    """Slide part names in presentation order via sldIdLst -> rels."""
    try:
        pres = ElementTree.fromstring(z.read("ppt/presentation.xml"))
        rels = ElementTree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
        targets = {rel.get("Id"): rel.get("Target")
                   for rel in rels.findall(REL + "Relationship")}
        order = []
        for sld_id in pres.iter(P + "sldId"):
            target = targets.get(sld_id.get(R + "id"))
            if target:
                name = "ppt/" + target.lstrip("/")
                if name in z.namelist():
                    order.append(name)
        if order:
            return order
    except Exception:
        pass
    names = [n for n in z.namelist()
             if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    return sorted(names, key=lambda n: int("".join(c for c in n.rsplit("slide", 1)[1] if c.isdigit()) or 0))


def shape_text(sp):
    """Text of one <p:sp>: paragraphs of <a:t> runs, one line per <a:p>."""
    lines = []
    for para in sp.iter(A + "p"):
        line = "".join(t.text or "" for t in para.iter(A + "t")).strip()
        if line:
            lines.append(line)
    return lines


def is_title(sp):
    ph = sp.find(".//" + P + "ph")
    return ph is not None and ph.get("type", "body") in TITLE_TYPES


def slide_lines(data):
    root = ElementTree.fromstring(data)
    titles, body = [], []
    for sp in root.iter(P + "sp"):
        for line in shape_text(sp):
            (titles if is_title(sp) else body).append(line)
    return titles, body


def extract(path, max_width=0):
    with zipfile.ZipFile(path) as z:
        names = slide_order(z)[:MAX_SLIDES]
        out = []
        for i, name in enumerate(names):
            titles, body = slide_lines(z.read(name))
            out.append(DIM + "Slide " + str(i + 1) + RESET)
            for line in titles:
                out.append("  " + BOLD + line + RESET)
            for line in body:
                out.append("  " + line)
            out.append("")
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out)


if __name__ == "__main__":
    try:
        text = extract(sys.argv[1], int(sys.argv[3]) if len(sys.argv) > 3 else 0)
        if len(sys.argv) > 2:
            table.write_cache(text, sys.argv[2])
        sys.stdout.write(text)
    except Exception as exc:
        print("PPTX text extraction failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
