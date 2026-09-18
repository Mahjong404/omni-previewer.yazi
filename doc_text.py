"""Legacy Word text preview (.doc/.dot/.rtf and other OLE formats).

No OOXML fast path exists for OLE binaries, so the text comes from the
persistent Word server: open the document read-only and take Content.Text -
a character stream, no pagination or export (~1-2s on a warm instance).
Same CLI/cache/stdout contract as docx_text.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import render
import table


def extract(path):
    reply = render.server_request(
        {"kind": "word", "text": True, "source": str(path)}, 60)
    out = []
    for line in (reply.get("text") or "").split("\n"):
        line = line.strip()
        if line or (out and out[-1]):
            out.append(line)
    while out and not out[-1]:
        out.pop()
    return "\n".join(out)


if __name__ == "__main__":
    try:
        text = extract(sys.argv[1])
        if len(sys.argv) > 2:
            table.write_cache(text, sys.argv[2])
        sys.stdout.write(text)
    except Exception as exc:
        print("DOC text extraction failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
