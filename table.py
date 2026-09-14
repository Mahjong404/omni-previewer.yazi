import csv
import io
import os
import sys
import time
import unicodedata

MAX_ROWS, MAX_COLS, CELL_MAX = 300, 20, 40
DIM, RESET, BOLD, TITLE = "\x1b[90m", "\x1b[0m", "\x1b[1m", "\x1b[1;36m"


def width(text):
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)


def fit(text, w):
    if width(text) <= w:
        return text + " " * (w - width(text))
    out, n = "", 0
    for c in text:
        cw = 2 if unicodedata.east_asian_width(c) in "WF" else 1
        if n + cw > w - 1:
            break
        out += c
        n += cw
    return out + "…" + " " * (w - n - 1)


def fit_center(text, w):
    t = fit(text, w).rstrip()
    lead = max(0, (w - width(t)) // 2)
    return " " * lead + t + " " * max(0, w - lead - width(t))


def grid_lines(rows, header=None):
    """rows: list of (cells, spans) — cells: list[str], spans: list[int].
    Plain text grid with box-drawing borders; header=True renders row 0 bold with a divider."""
    buf = io.StringIO()
    real = sys.stdout
    sys.stdout = buf
    try:
        _grid_body(rows, header)
    finally:
        sys.stdout = real
    return buf.getvalue()


def _grid_body(rows, header=None):
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

    def border(left, mid, right):
        print(DIM + left + mid.join("─" * (w + 2) for w in colw) + right + RESET)

    def emit(text, w, bold=False):
        t = (BOLD + text + RESET) if bold else text
        return " " + t + " " + DIM + "│" + RESET

    border("┌", "┬", "┐")
    for ri, (cells, spans) in enumerate(rows):
        out = DIM + "│" + RESET
        c = 0
        for text, span in zip(cells, spans):
            w = sum(colw[c:c + span]) + 3 * (span - 1)
            body = fit_center(text, w) if span > 1 else fit(text, w)
            out += emit(body, w, bold=bool(header) and ri == 0)
            c += span
        print(out)
        if header and ri == 0:
            border("├", "┼", "┤")
    border("└", "┴", "┘")


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
    rows = csv_rows(path)
    out = grid_lines(rows, header=True)
    if len(sys.argv) > 2:
        write_cache(out, sys.argv[2])
    sys.stdout.write(out)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print("Table preview failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
