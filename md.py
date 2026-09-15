import hashlib
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import table

DIM, RESET, BOLD, ITAL, UND = "\x1b[90m", "\x1b[0m", "\x1b[1m", "\x1b[3m", "\x1b[4m"
CYAN, STRIKE = "\x1b[36m", "\x1b[9m"
BOLD_TXT = "\x1b[1;38;5;229m"
CODE = "\x1b[48;5;253m\x1b[30m"
HL = "\x1b[48;5;238m\x1b[38;5;222m"
MAX_LINES = 2000
PLANTUML_JAR = r"C:\software\CLI\plantuml\plantuml.jar"
ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")


def _disp_w(ch):
    return 2 if unicodedata.east_asian_width(ch) in "WF" else 1


def img_size(path):
    """(w, h) of a PNG/JPEG/GIF without external deps; None if unknown."""
    try:
        with open(path, "rb") as f:
            head = f.read(32)
            if head.startswith(b"\x89PNG"):
                w, h = struct.unpack(">II", head[16:24])
                return w, h
            if head[:6] in (b"GIF87a", b"GIF89a"):
                w, h = struct.unpack("<HH", head[6:10])
                return w, h
            if head.startswith(b"\xff\xd8"):
                f.seek(2)
                while True:
                    b = f.read(1)
                    if not b:
                        return None
                    if b != b"\xff":
                        continue
                    marker = f.read(1)
                    while marker == b"\xff":
                        marker = f.read(1)
                    if marker in (b"\xd8", b"\xd9") or b"\xd0" <= marker <= b"\xd7":
                        continue
                    seg = f.read(2)
                    if len(seg) < 2:
                        return None
                    seglen = struct.unpack(">H", seg)[0]
                    if marker in (b"\xc0", b"\xc1", b"\xc2", b"\xc3", b"\xc5", b"\xc6", b"\xc7",
                                  b"\xc9", b"\xca", b"\xcb", b"\xcd", b"\xce", b"\xcf"):
                        data = f.read(5)
                        h, w = struct.unpack(">HH", data[1:5])
                        return w, h
                    f.seek(seglen - 2, 1)
    except Exception:
        return None
    return None


def media_path(cache_base, kind, src, ext=".png"):
    """Content-hashed media cache so re-renders/edits reuse generated assets."""
    d = os.path.join(os.path.dirname(cache_base) or ".", "md-media")
    os.makedirs(d, exist_ok=True)
    h = hashlib.sha256(("v2\n" + kind + "\n" + src).encode("utf-8")).hexdigest()[:20]
    return os.path.join(d, h + ext)


def wrap_ansi(s, width):
    """Wrap an ANSI-styled line at display width; returns list of lines."""
    if width <= 0 or table.width(s) <= width:
        return [s]
    lines, cur, w, i = [], "", 0, 0
    while i < len(s):
        m = ANSI_RE.match(s, i)
        if m:
            cur += m.group(0)
            i = m.end()
            continue
        cw = _disp_w(s[i])
        if w + cw > width:
            lines.append(cur)
            cur, w = "", 0
        cur += s[i]
        w += cw
        i += 1
    lines.append(cur)
    return lines

GREEK = {"alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "zeta": "ζ",
         "eta": "η", "theta": "θ", "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ",
         "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ",
         "upsilon": "υ", "phi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
         "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π",
         "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω"}
OPS = {"sum": "Σ", "prod": "Π", "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮",
       "infty": "∞", "partial": "∂", "nabla": "∇", "pm": "±", "mp": "∓", "times": "×",
       "div": "÷", "cdot": "·", "ast": "∗", "star": "⋆", "circ": "∘", "bullet": "∙",
       "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠", "equiv": "≡",
       "approx": "≈", "cong": "≅", "sim": "∼", "simeq": "≃", "propto": "∝", "in": "∈",
       "notin": "∉", "subset": "⊂", "supset": "⊃", "subseteq": "⊆", "supseteq": "⊇",
       "cup": "∪", "cap": "∩", "setminus": "∖", "emptyset": "∅", "varnothing": "∅",
       "forall": "∀", "exists": "∃", "neg": "¬", "land": "∧", "wedge": "∧", "lor": "∨",
       "vee": "∨", "to": "→", "rightarrow": "→", "leftarrow": "←", "Rightarrow": "⇒",
       "Leftarrow": "⇐", "leftrightarrow": "↔", "Leftrightarrow": "⇔", "mapsto": "↦",
       "uparrow": "↑", "downarrow": "↓", "oplus": "⊕", "ominus": "⊖", "otimes": "⊗",
       "perp": "⊥", "parallel": "∥", "angle": "∠", "degree": "°", "hbar": "ℏ", "ell": "ℓ",
       "Re": "ℜ", "Im": "ℑ", "aleph": "ℵ", "wp": "℘", "ldots": "…", "cdots": "⋯",
       "vdots": "⋮", "ddots": "⋱", "prime": "′", "therefore": "∴", "because": "∵"}
SUP = str.maketrans("0123456789+-=()niabcedfghjklmoprtuvx", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱᵃᵇᶜᵉᵈᶠᵍʰʲᵏˡᵐᵒᵖʳᵗᵘᵛˣ")
SUB = str.maketrans("0123456789+-=()aeoxhklmnpstijruv", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜᵢⱼᵣᵤᵥ")


def _braced(s, i):
    if i >= len(s) or s[i] != "{":
        return None, i
    depth = 0
    for j in range(i, len(s)):
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
    return None, i


def _frac_sqrt_pass(s):
    m = re.search(r"\\(d?frac|t?frac|sqrt)", s)
    if not m:
        return s
    i = m.end()
    while i < len(s) and s[i] in " \t":
        i += 1
    if m.group(1) == "sqrt":
        idx = ""
        if i < len(s) and s[i] == "[":
            k = s.find("]", i)
            if k > 0:
                idx, i = s[i + 1:k], k + 1
        arg, end = _braced(s, i)
        if arg is None:
            return s[:m.start()] + "√" + s[m.end():]
        rep = (idx.translate(SUP) + "√(" + arg + ")") if idx else "√(" + arg + ")"
        return s[:m.start()] + rep + s[end:]
    a, end = _braced(s, i)
    while end < len(s) and s[end] in " \t":
        end += 1
    b, end2 = _braced(s, end)
    if a is None or b is None:
        return s[:m.start()] + " " + s[m.end():]
    return s[:m.start()] + a + "⁄" + b + s[end2:]


ACCENTS = {"hat": "̂", "bar": "̄", "overline": "̅", "vec": "⃗",
           "dot": "̇", "ddot": "̈", "tilde": "̃", "breve": "̆", "check": "̌"}
FONTS = ("boldsymbol", "bm", "mathbf", "mathbfit", "mathit", "mathrm", "mathsf",
         "mathtt", "mathcal", "mathbb", "mathfrak", "text", "operatorname")


def _accent(s):
    for name, mark in ACCENTS.items():
        s = re.sub(r"\\" + name + r"\s*\{([^{}]*)\}",
                   lambda m: "".join(c + mark for c in m.group(1)), s)
        s = re.sub(r"\\" + name + r"\s+(\w)", lambda m: m.group(1) + mark, s)
    return s


def math_unicode(s):
    prev = None
    while prev != s:
        prev = s
        s = _frac_sqrt_pass(s)
    s = re.sub(r"\\(?:" + "|".join(FONTS) + r")\s*\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\(?:" + "|".join(FONTS) + r")\s+(\w)", r"\1", s)
    s = _accent(s)
    s = re.sub(r"\\(sum|prod|int|iint|iiint|oint|coprod|bigcup|bigcap|bigoplus|bigotimes)_\{([^{}]*)\}\^\{([^{}]*)\}",
               lambda m: OPS.get(m.group(1), m.group(1)) + "_" + m.group(2).translate(SUB) + "^" + m.group(3).translate(SUP), s)
    s = re.sub(r"\\(left|right|bigl|bigr|Bigl|Bigr|bigg|Bigg|big|Big|limits|displaystyle|quad|qquad|,|;|!| )", " ", s)
    s = re.sub(r"\\([A-Za-z]+)", lambda m: GREEK.get(m.group(1), OPS.get(m.group(1), m.group(1))), s)
    s = re.sub(r"\^\{([^{}]*)\}", lambda m: "".join(ch.translate(SUP) for ch in m.group(1)), s)
    s = re.sub(r"_\{([^{}]*)\}", lambda m: "".join(ch.translate(SUB) for ch in m.group(1)), s)
    s = re.sub(r"\^([A-Za-z0-9+\-=()])", lambda m: m.group(1).translate(SUP), s)
    s = re.sub(r"_([A-Za-z0-9+\-=()])", lambda m: m.group(1).translate(SUB), s)
    s = s.replace("{", "").replace("}", "").replace("\\", " ").replace("~", " ")
    return re.sub(r" {2,}", " ", s).strip()


_TEX_DOC = (r"\documentclass{article}\usepackage{amsmath,amssymb,mathtools}"
            r"\pagestyle{empty}\begin{document}$%s$\end{document}")


def _math_png_latex(latex, out_path):
    """Real LaTeX → dvipng (full amsmath incl. environments); transparent bg."""
    latex_exe, dvipng = shutil.which("latex"), shutil.which("dvipng")
    if not (latex_exe and dvipng):
        return False
    tmpdir = out_path + ".tex.d"
    try:
        os.makedirs(tmpdir, exist_ok=True)
        tex_path = os.path.join(tmpdir, "m.tex")
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(_TEX_DOC % latex)
        r = subprocess.run([latex_exe, "-interaction=nonstopmode", "-halt-on-error", "m.tex"],
                           cwd=tmpdir, capture_output=True, timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        dvi = os.path.join(tmpdir, "m.dvi")
        if r.returncode != 0 or not os.path.exists(dvi):
            return False
        r = subprocess.run([dvipng, "-T", "tight", "-D", "220", "-bg", "Transparent",
                            "-fg", "rgb 1 1 1", "-o", out_path, dvi],
                           cwd=tmpdir, capture_output=True, timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return r.returncode == 0 and os.path.exists(out_path)
    except Exception:
        return False
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def math_png(latex, out_path):
    if os.path.exists(out_path):
        return True
    if _math_png_latex(latex, out_path):
        return True
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        fig = plt.figure(figsize=(0.1, 0.1))
        fig.text(0, 0, f"${latex}$", fontsize=14, color="white")
        fig.savefig(out_path, dpi=200, transparent=True, bbox_inches="tight", pad_inches=0.05)
        plt.close(fig)
        return os.path.exists(out_path)
    except Exception:
        return False


def mmdc_png(src, out_path):
    if os.path.exists(out_path):
        return True
    mmdc = shutil.which("mmdc")
    if not mmdc:
        return False
    try:
        tmp = out_path + ".mmd"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(src)
        r = subprocess.run([mmdc, "-i", tmp, "-o", out_path, "-b", "transparent",
                            "-s", "2", "-w", "1400", "-q"],
                           capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        os.path.exists(tmp) and os.remove(tmp)
        return r.returncode == 0 and os.path.exists(out_path)
    except Exception:
        return False


def plantuml_utxt(src, cache_file):
    if os.path.exists(cache_file):
        try:
            return open(cache_file, encoding="utf-8").read().rstrip("\n").split("\n")
        except Exception:
            return None
    if not os.path.exists(PLANTUML_JAR):
        return None
    body = src.strip()
    if "@start" not in body:
        body = "@startuml\n" + body + "\n@enduml"
    try:
        r = subprocess.run(["java", "-jar", PLANTUML_JAR, "-utxt", "-charset", "UTF-8", "-pipe"],
                           input=body.encode("utf-8"), capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if r.returncode != 0 or not r.stdout:
            return None
        text = r.stdout.decode("utf-8", errors="replace")
        if "Error" in text.split("\n", 1)[0]:
            return None
        table.write_cache(text, cache_file)
        return text.rstrip("\n").split("\n")
    except Exception:
        return None


MM_NODE = re.compile(r"([A-Za-z0-9_一-鿿]+)\s*(?:\(\(([^)]*)\)\)|\(([^)]*)\)|\[\[([^\]]*)\]\]|\[([^\]]*)\]|\{\{([^}]*)\}\}|\{([^}]*)\})?")
MM_NODE_PAT = r"([A-Za-z0-9_一-鿿]+)\s*(?:\(\([^)]*\)\)|\([^)]*\)|\[\[[^\]]*\]\]|\[[^\]]*\]|\{[^}]*\}|\{[^\}]*\}|\"[^\"]*\")?"
MM_EDGE = re.compile(MM_NODE_PAT + r"\s*([-=.o]{2,}[xo>]?)\s*(?:\|([^|]*)\|)?\s*" + MM_NODE_PAT)
MM_CHAIN_TEXT = re.compile(r"--\s*([^-|>].*?)\s*-->")
MM_SEQ = re.compile(r"([A-Za-z0-9_一-鿿]+)\s*(--?>>?|--?x|--?\))\s*([A-Za-z0-9_一-鿿]+)\s*:?\s*(.*)")


def _mm_labels(src):
    labels = {}
    for m in MM_NODE.finditer(src):
        inner = next((g for g in m.groups()[1:] if g and g.strip()), None)
        if inner:
            labels[m.group(1)] = inner.strip().strip('"')
    return labels


def mermaid_text(src):
    lines = [l.strip() for l in src.split("\n") if l.strip() and not l.strip().startswith("%%")]
    if not lines:
        return None
    head = lines[0].lower()
    labels = _mm_labels(src)

    def name(n):
        return labels.get(n, n)

    if head.startswith("sequencediagram"):
        out = []
        for l in lines[1:]:
            m = MM_SEQ.match(l)
            if m:
                arrow = "──▶" if not m.group(2).startswith("--") else "┄┄▶"
                tail = ("  " + m.group(4)) if m.group(4) else ""
                out.append(f"  {name(m.group(1))} {arrow} {name(m.group(3))}{tail}")
            elif l.lower().startswith("note"):
                out.append(DIM + "  ⓘ " + l.split(":", 1)[-1].strip() + RESET)
            elif re.match(r"(?i)^(alt|else|opt|loop|par|critical|break|rect)\b", l):
                out.append(CYAN + "  ┌ " + l + RESET)
            elif l.lower() == "end":
                out.append(CYAN + "  └" + RESET)
        return out or None
    if re.match(r"(?i)^(flowchart|graph)\b", head):
        out = []
        for l in lines[1:]:
            if re.match(r"(?i)^(subgraph|end|style|classdef|class |click|linkstyle|direction)\b", l):
                continue
            l = MM_CHAIN_TEXT.sub(r"-->\|\1\|", l)
            m = MM_EDGE.search(l)
            if m:
                arrow = m.group(2)
                sym = "──▶" if arrow.endswith(">") and "=" not in arrow and "." not in arrow else \
                      ("═▶" if "=" in arrow else ("┄┄▶" if "." in arrow else ("──" + arrow[-1] if arrow[-1] in "xo" else "───")))
                mid = (m.group(3) or "").strip()
                if sym.endswith("▶"):
                    conn = ("── " + mid + " ──▶") if mid else sym
                    out.append(f"  {name(m.group(1))} {conn} {name(m.group(4))}")
                else:
                    out.append(f"  {name(m.group(1))} {sym} {name(m.group(4))}")
        return out or None
    return None


def plantuml_png(src, out_path):
    if os.path.exists(out_path):
        return True
    if not os.path.exists(PLANTUML_JAR):
        return False
    try:
        body = src.strip()
        if "@start" not in body:
            body = "@startuml\n" + body + "\n@enduml"
        if "background" not in body:
            head, _, tail = body.partition("\n")
            body = head + "\nskinparam backgroundColor transparent\n" + tail
        tmp = out_path + ".puml"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(body)
        r = subprocess.run(["java", "-jar", PLANTUML_JAR, "-tpng", "-charset", "UTF-8", tmp],
                           capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        produced = tmp[:-5] + ".png"
        if r.returncode == 0 and os.path.exists(produced):
            os.replace(produced, out_path)
            os.remove(tmp)
            return True
        return False
    except Exception:
        return False


def highlight(code, lang):
    try:
        from pygments import highlight as hl
        from pygments.formatters import Terminal256Formatter
        from pygments.lexers import get_lexer_by_name, TextLexer
        lexer = get_lexer_by_name(lang or "") if lang else TextLexer()
        return hl(code, lexer, Terminal256Formatter(style="monokai")).rstrip("\n").split("\n")
    except Exception:
        return code.rstrip("\n").split("\n")


TOK = "\x00{}\x00"


def render_inline(s):
    spans = []

    def keep(text):
        spans.append(text)
        return TOK.format(len(spans) - 1)

    s = re.sub(r"`([^`]+)`", lambda m: keep(CODE + " " + m.group(1) + " " + RESET), s)
    s = re.sub(r"\$([^$\n]+)\$", lambda m: keep(ITAL + (math_unicode(m.group(1)) if not re.search(r"\\(begin|end|matrix|cases|aligned|split)\b", m.group(1)) else m.group(1)) + RESET), s)
    s = re.sub(r"!\[([^\]]*)\]\(([^)\s]+)\)", lambda m: keep(DIM + "[image: " + (m.group(1) or m.group(2)) + "]" + RESET), s)
    s = re.sub(r"\[\[([^\]]+)\]\]", lambda m: keep(UND + CYAN + m.group(1) + RESET), s)
    s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", lambda m: keep(UND + CYAN + m.group(1) + RESET + DIM + "(" + m.group(2) + ")" + RESET), s)
    s = re.sub(r"==([^=]+)==", lambda m: keep(HL + " " + m.group(1) + " " + RESET), s)
    s = re.sub(r"\*\*([^*]+)\*\*|__([^_]+)__", lambda m: keep(BOLD_TXT + (m.group(1) or m.group(2)) + RESET), s)
    s = re.sub(r"~~([^~]+)~~", lambda m: keep(STRIKE + m.group(1) + RESET), s)
    s = re.sub(r"(?<![\w*])\*([^*\n]+)\*(?![\w*])|(?<![\w_])_([^_\n]+)_(?![\w_])",
               lambda m: keep(ITAL + (m.group(1) or m.group(2)) + RESET), s)
    return re.sub(re.escape("\x00") + r"(\d+)" + re.escape("\x00"), lambda m: spans[int(m.group(1))], s)


def parse_table(lines, i, max_width=0):
    j = i
    block = []
    while j < len(lines) and lines[j].strip().startswith("|"):
        block.append(lines[j].strip())
        j += 1
    if len(block) < 2 or not re.match(r"^\|[\s:|-]+\|?$", block[1]):
        return None, i

    def cells_of(line):
        inner = line.strip().strip("|")
        return [re.sub(r"\\([\[\]|])", r"\1", c).strip() for c in inner.split("|")]

    header = [render_inline(c) for c in cells_of(block[0])]
    data = [[render_inline(c) for c in cells_of(b)] for b in block[2:]]
    ncols = len(header)
    natural = [0] * ncols
    for r in [header] + data:
        for c in range(min(ncols, len(r))):
            natural[c] = max(natural[c], table.width(r[c]))
    dropped = 0
    if max_width:
        # Drop rightmost columns that cannot fit; never emit an over-wide grid.
        keep, used = 0, 1
        for c in range(ncols):
            w = min(max(natural[c], 4), table.CELL_MAX)
            if used + w + 3 <= max_width:
                used += w + 3
                keep += 1
            else:
                break
        if keep < ncols:
            dropped = ncols - max(keep, 1)
            keep = max(keep, 1)
            header, ncols = header[:keep], keep
            data = [r[:keep] for r in data]
            natural = natural[:keep]
    grid = [(header, [1] * ncols)]
    grid += [(r, [1] * len(r)) for r in data]
    out = table.grid_lines(grid, header=True, max_width=max_width, rowsep=True).splitlines()
    if dropped:
        out.append(DIM + f"  ⋯ {dropped} column(s) hidden — pane too narrow" + RESET)
    return out, j


def render(md_path, cache_base, max_width=0, max_height=0):
    try:
        raw = open(md_path, "rb").read()
        text = raw.decode("utf-8-sig", errors="replace")
    except Exception:
        text = ""
    lines = text.replace("\r\n", "\n").split("\n")[:MAX_LINES]
    base = os.path.dirname(md_path)
    media = []
    out = []
    pre = set()  # line indexes that must not be re-wrapped (code/diagrams/grids)
    i = 0
    fence = re.compile(r"^(\s*)(`{3,}|~{3,})\s*([\w+-]*)\s*$")

    def emit_pre(ls):
        pre.update(range(len(out), len(out) + len(ls)))
        out.extend(ls)

    def emit_media(path, caption):
        """Reserve placeholder rows; main.lua overlays the real image via
        image_show into that sub-rect (native resolution, scrolls with text)."""
        out.append(DIM + caption + RESET)
        size = img_size(path) or (0, 0)
        w_px, h_px = size
        if w_px > 0:
            rows = max(2, round(h_px * (max_width or 60) / w_px / 2))
        else:
            rows = 8
        if max_height:
            rows = min(rows, max(4, max_height - 2))
        rows = min(rows, 30)
        line = len(out)
        out.extend([""] * rows)
        media.append({"line": line, "lines": rows, "path": path})

    while i < len(lines):
        line = lines[i]
        m = fence.match(line)
        if m:
            lang = (m.group(3) or "").lower()
            j = i + 1
            body = []
            while j < len(lines) and not lines[j].strip().startswith(m.group(2)[0] * 3):
                body.append(lines[j])
                j += 1
            src = "\n".join(body)
            if not lang and "@start" in src:
                lang = "plantuml"
            if lang == "mermaid":
                png = media_path(cache_base, "mermaid", src)
                if mmdc_png(src, png):
                    emit_media(png, "◆ mermaid")
                else:
                    txt = mermaid_text(src)
                    if txt:
                        out.append(DIM + "◆ mermaid" + RESET)
                        emit_pre(txt)
                    else:
                        out.append(DIM + "[mermaid diagram — renderer unavailable]" + RESET)
                        emit_pre([DIM + "│ " + RESET + l for l in highlight(src, "mermaid")])
            elif lang in ("plantuml", "puml"):
                txt = plantuml_utxt(src, media_path(cache_base, "plantuml", src, ".utxt"))
                if txt:
                    out.append(DIM + "◆ plantuml" + RESET)
                    emit_pre(txt)
                else:
                    png = media_path(cache_base, "plantuml", src)
                    if plantuml_png(src, png):
                        emit_media(png, "◆ plantuml")
                    else:
                        out.append(DIM + "[plantuml diagram — renderer unavailable]" + RESET)
                        emit_pre([DIM + "│ " + RESET + l for l in highlight(src, "java")])
            else:
                out.append(DIM + "```" + lang + RESET)
                emit_pre([DIM + "│ " + RESET + l for l in highlight(src, lang)])
                out.append(DIM + "```" + RESET)
            i = j + 1 if j < len(lines) else j
            continue
        st = line.strip()
        if st.startswith("$$"):
            j = i
            body = st[2:]
            while not body.rstrip().endswith("$$") and j + 1 < len(lines):
                j += 1
                body += "\n" + lines[j]
            latex = body.rstrip().removesuffix("$$").strip()
            uni = math_unicode(latex) if latex else ""
            degrade = re.search(r"\\(begin|text|operatorname|hat|vec|overline|underline|underbrace|overbrace|bar|tilde|dot|ddot|stackrel|xrightarrow|xleftarrow|overset|underset|binom|mod|bmod|boxed|color|tag|label|lefteqn|displaylines|substack|mathclap|intertext|shortintertext)", latex or "")
            if uni and not degrade:
                out.append("    " + ITAL + uni + RESET)
            elif latex:
                png = media_path(cache_base, "math", latex)
                if math_png(latex, png):
                    emit_media(png, "◈ " + re.sub(r"\s+", " ", latex)[:60])
                else:
                    out.append("    " + ITAL + (uni or latex) + RESET)
            i = j + 1
            continue
        if st.startswith("|") and "|" in st[1:]:
            tbl, ni = parse_table(lines, i, max_width)
            if tbl:
                emit_pre(tbl)
                i = ni
                continue
        mimg = re.match(r"^!\[([^\]]*)\]\(([^)\s]+)\)\s*$", st)
        if mimg:
            target = mimg.group(2)
            if re.match(r"(?i)^https?://", target):
                out.append(DIM + f"[image: {target}]" + RESET)
            else:
                p = target if os.path.isabs(target) else os.path.join(base, target.replace("/", os.sep))
                if os.path.exists(p):
                    emit_media(p, "▦ " + (mimg.group(1) or os.path.basename(p)))
                else:
                    out.append(DIM + f"[image missing: {target}]" + RESET)
            i += 1
            continue
        hm = re.match(r"^(#{1,6})\s+(.*)$", line)
        if hm:
            level = len(hm.group(1))
            style = BOLD + CYAN if level <= 2 else BOLD
            out.append(style + ("█ " if level <= 2 else "▌ ") + render_inline(hm.group(2)) + RESET)
            i += 1
            continue
        if re.match(r"^\s*([-*_])(\s*\1){2,}\s*$", line):
            w = max(10, min((max_width or 60) - 1, 60))
            out.append(DIM + "─" * w + RESET)
            i += 1
            continue
        if st.startswith(">"):
            inner = re.sub(r"^\s*>\s?", "", line)
            out.append(DIM + "│ " + RESET + ITAL + render_inline(inner) + RESET)
            i += 1
            continue
        lm = re.match(r"^(\s*)([-*+]|\d+\.)\s+(.*)$", line)
        if lm:
            hang_w = len(lm.group(1)) + len(lm.group(2)) + 1
            body = render_inline(lm.group(3))
            wl = wrap_ansi(body, max(10, (max_width or 80) - hang_w))
            out.append(lm.group(1) + CYAN + lm.group(2) + RESET + " " + wl[0])
            out.extend(" " * hang_w + l for l in wl[1:])
            i += 1
            continue
        out.append(render_inline(line) if st else "")
        i += 1
    # Wrap every wrappable line to the pane width so manifest line numbers map
    # 1:1 to screen rows (main.lua displays with Wrap.NO and overlays media).
    if max_width:
        idx_map, wrapped = {}, []
        for n, l in enumerate(out):
            idx_map[n] = len(wrapped)
            wrapped.extend([l] if n in pre else wrap_ansi(l, max_width))
        for m in media:
            m["line"] = idx_map.get(m["line"], m["line"])
        out = wrapped
    return {"text": "\n".join(out), "media": media}


if __name__ == "__main__":
    try:
        path = sys.argv[1]
        cache = sys.argv[2] if len(sys.argv) > 2 else None
        max_width = int(sys.argv[3]) if len(sys.argv) > 3 else 0
        max_height = int(sys.argv[4]) if len(sys.argv) > 4 else 0
        cache_base = cache[:-5] if cache and cache.endswith(".ansi") else (cache or "md")
        manifest = render(path, cache_base, max_width, max_height)
        blob = json.dumps(manifest, ensure_ascii=False)
        if cache:
            table.write_cache(blob, cache)
        sys.stdout.write(blob)
    except Exception as exc:
        print("MD preview failed: " + str(exc), file=sys.stderr)
        sys.exit(1)
