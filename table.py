import csv
import io
import os
import re
import sys
import time
import unicodedata

MAX_ROWS, MAX_COLS, CELL_MAX = 300, 20, 40
DIM, RESET, BOLD, TITLE = "\x1b[90m", "\x1b[0m", "\x1b[1m", "\x1b[1;36m"
ANSI = re.compile(r"\x1b\[[0-9;]*m")


def width(text):
    text = ANSI.sub("", text)
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)


def fit(text, w):
    if width(text) <= w:
        return text + " " * (w - width(text))
    out, n, i = "", 0, 0
    while i < len(text):
        m = ANSI.match(text, i)
        if m:
            out += m.group(0)
            i = m.end()
            continue
        c = text[i]
        cw = 2 if unicodedata.east_asian_width(c) in "WF" else 1
        if n + cw > w - 1:
            break
        out += c
        n += cw
        i += 1
    return out + RESET + "…" + " " * (w - n - 1)


def fit_center(text, w):
    t = fit(text, w).rstrip()
    lead = max(0, (w - width(t)) // 2)
    return " " * lead + t + " " * max(0, w - lead - width(t))


def wrap_cell(text, w):
    """Hard-wrap text to display width w, preserving ANSI spans; returns padded lines."""
    lines = []
    for seg in str(text).split("\n"):
        cur, n, i = "", 0, 0
        while i < len(seg):
            m = ANSI.match(seg, i)
            if m:
                cur += m.group(0)
                i = m.end()
                continue
            c = seg[i]
            cw = 2 if unicodedata.east_asian_width(c) in "WF" else 1
            if n + cw > w:
                lines.append(cur)
                cur, n = "", 0
            cur += c
            n += cw
            i += 1
        lines.append(cur)
    return [l + " " * max(0, w - width(l)) for l in lines] or [""]


def grid_lines(rows, header=None, max_width=0, rowsep=False):
    """rows: list of (cells, spans) — cells: list[str], spans: list[int].
    Plain text grid with box-drawing borders; cells wrap instead of truncating;
    header=True renders row 0 bold with a divider. max_width caps total width.
    rowsep=True draws a separator between data rows."""
    buf = io.StringIO()
    real = sys.stdout
    sys.stdout = buf
    try:
        _grid_body(rows, header, max_width, rowsep)
    finally:
        sys.stdout = real
    return buf.getvalue()


def _grid_body(rows, header=None, max_width=0, rowsep=False):
    rows = [(cells, spans) for cells, spans in rows if any(cells)]
    while rows and not any(rows[-1][0]):
        rows.pop()
    if not rows:
        print("(empty)")
        return
    ncols = min(max(sum(s) for _, s in rows), MAX_COLS)
    colw = [1] * ncols
    for cells, spans in rows:
        c = 0
        for text, span in zip(cells, spans):
            if span == 1 and c < ncols:
                colw[c] = min(max(colw[c], width(text)), CELL_MAX)
            c += span
    dropped = 0
    if max_width:
        # Keep the leftmost run of columns that fit at natural width; a
        # column whose content cannot fully display is dropped with a note
        # rather than squeezed into an unreadable strip.
        keep, used = 0, 1
        for c in range(ncols):
            if used + colw[c] + 3 <= max_width:
                used += colw[c] + 3
                keep += 1
            else:
                break
        if keep < ncols:
            dropped = ncols - max(keep, 1)
            ncols = max(keep, 1)
    if dropped:
        colw = colw[:ncols]
        rows = [_trim_row(cells, spans, ncols) for cells, spans in rows]
    avail = max_width - 3 * ncols - 1
    if max_width and sum(colw) > avail:
        while sum(colw) > max(avail, 4 * ncols):
            i = colw.index(max(colw))
            if colw[i] <= 4:
                break
            colw[i] -= 1

    def border(left, mid, right):
        print(DIM + left + mid.join("─" * (w + 2) for w in colw) + right + RESET)

    border("┌", "┬", "┐")
    for ri, (cells, spans) in enumerate(rows):
        wrapped, c = [], 0
        for text, span in zip(cells, spans):
            w = sum(colw[c:c + span]) + 3 * (span - 1)
            wrapped.append((wrap_cell(text, w), span, w))
            c += span
        height = max(len(wl) for wl, _, _ in wrapped)
        bold = bool(header) and ri == 0
        for h in range(height):
            out = DIM + "│" + RESET
            for wl, span, w in wrapped:
                body = wl[h] if h < len(wl) else " " * w
                if span > 1 and h == 0:
                    pad = max(0, (w - width(wl[0].rstrip())) // 2)
                    body = " " * pad + wl[0].rstrip() + " " * max(0, w - pad - width(wl[0].rstrip()))
                if bold:
                    body = BOLD + body + RESET
                out += " " + body + " " + DIM + "│" + RESET
            print(out)
        if header and ri == 0:
            border("├", "┼", "┤")
        elif rowsep and ri < len(rows) - 1:
            border("├", "┼", "┤")
    border("└", "┴", "┘")
    if dropped:
        print(DIM + f"  ⋯ {dropped} column(s) hidden — pane too narrow" + RESET)


def _trim_row(cells, spans, keep):
    """Drop cells beyond column `keep`; clip a span crossing the boundary."""
    out_c, out_s, c = [], [], 0
    for text, span in zip(cells, spans):
        if c >= keep:
            break
        out_c.append(text)
        out_s.append(min(span, keep - c))
        c += span
    return out_c, out_s


def csv_rows(path):
    raw = open(path, "rb").read(1 << 20)
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            sample = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        sample = raw.decode("utf-8", errors="replace")
    reader = csv.reader(io.StringIO(sample))
    rows = []
    for row in reader:
        rows.append(([" ".join(str(c).split()) for c in row], [1] * len(row)))
        if len(rows) >= MAX_ROWS:
            break
    return rows


def prune_cache(d, keep=100, max_age=3 * 86400):
    try:
        files = [os.path.join(d, f) for f in os.listdir(d) if f.endswith(".ansi")]
        now = time.time()
        alive = []
        for f in files:
            try:
                if now - os.path.getmtime(f) > max_age:
                    os.remove(f)
                else:
                    alive.append((os.path.getmtime(f), f))
            except OSError:
                pass
        for _, f in sorted(alive, reverse=True)[keep:]:
            try:
                os.remove(f)
            except OSError:
                pass
    except Exception:
        pass


def write_cache(text, cache):
    try:
        d = os.path.dirname(cache)
        os.makedirs(d, exist_ok=True)
        tmp = cache + "." + str(os.getpid()) + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            f.write(text)
        os.replace(tmp, cache)
        prune_cache(d)
    except Exception:
        pass


def main():
    path = sys.argv[1]
    max_width = int(sys.argv[3]) if len(sys.argv) > 3 else 0
    rows = csv_rows(path)
    out = grid_lines(rows, header=True, max_width=max_width)
    if len(sys.argv) > 2:
        write_cache(out, sys.argv[2])
    sys.stdout.write(out)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("Table preview failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
