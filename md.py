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
    h = hashlib.sha256(("v4\n" + kind + "\n" + src).encode("utf-8")).hexdigest()[:20]
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
         "varepsilon": "ε", "varphi": "φ", "vartheta": "ϑ", "varrho": "ϱ",
         "varsigma": "ς", "varkappa": "ϰ", "varpi": "ϖ", "ell": "ℓ",
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
       "vdots": "⋮", "ddots": "⋱", "prime": "′", "therefore": "∴", "because": "∵",
       "mid": "∣", "iff": "⟺", "implies": "⟹", "impliedby": "⟸", "nmid": "∤"}
# TeX puts thick space around relations and medium space around binary ops
# regardless of source spacing; atoms for these carry their own padding.
_REL = {"leq", "le", "geq", "ge", "neq", "ne", "equiv", "approx", "sim",
        "simeq", "cong", "propto", "in", "notin", "ni", "subset", "supset",
        "subseteq", "supseteq", "to", "rightarrow", "leftarrow", "Rightarrow",
        "Leftarrow", "Leftrightarrow", "leftrightarrow", "iff", "implies",
        "impliedby", "mapsto", "mid", "nmid", "parallel", "perp", "models",
        "vdash", "dashv", "prec", "succ", "preceq", "succeq", "ll", "gg",
        "asymp", "doteq", "approxeq"}
_BINOPS = {"pm", "mp", "times", "div", "cdot", "cup", "cap", "setminus",
           "oplus", "ominus", "otimes", "odot", "uplus", "sqcap", "sqcup",
           "vee", "wedge", "land", "lor", "ast", "star", "circ", "bullet",
           "diamond", "wr", "amalg", "bigtriangleup", "bigtriangledown"}
SUP = str.maketrans("0123456789+-=()niabcedfghjklmoprtuvxABDEGHIJKLMNOPRTUVW",
                    "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱᵃᵇᶜᵉᵈᶠᵍʰʲᵏˡᵐᵒᵖʳᵗᵘᵛˣᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂ")
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
           "dot": "̇", "ddot": "̈", "tilde": "̃", "breve": "̆", "check": "̌",
           "overrightarrow": "⃗", "overleftarrow": "⃖",
           "overleftrightarrow": "⃡", "widetilde": "̃", "widehat": "̂"}


def _mk_boldit():
    """\\boldsymbol → math bold-italic glyphs (𝒙, 𝜶). SMP chars; if the
    terminal font lacks them, revert the _m_cmd branch to _m_atom(_m_lit(g))."""
    m = {}
    for i in range(26):
        m[chr(ord("A") + i)] = chr(0x1D468 + i)
        m[chr(ord("a") + i)] = chr(0x1D482 + i)
    for i in range(10):
        m[str(i)] = chr(0x1D7CE + i)
    caps = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"
    lows = "αβγδεζηθικλμνξοπρςστυφχψω"
    for i, ch in enumerate(caps):
        m[ch] = chr(0x1D71C + i + (1 if i >= 17 else 0))  # skip ϴ slot
    for i, ch in enumerate(lows):
        m[ch] = chr(0x1D736 + i)
    m.update({"ϵ": chr(0x1D750), "ϑ": chr(0x1D751), "ϰ": chr(0x1D752),
              "ϕ": chr(0x1D753), "ϱ": chr(0x1D754), "ϖ": chr(0x1D755),
              "∂": chr(0x1D74F), "∇": chr(0x1D735)})
    return {ord(k): v for k, v in m.items()}  # str.translate needs ordinal keys


_BOLDIT = _mk_boldit()
FONTS_BOLD = ("boldsymbol", "bm", "mathbf", "mathbfit", "pmb")
FONTS = ("mathit", "mathrm", "mathsf", "mathtt", "mathcal", "mathfrak",
         "text", "operatorname")
BB = {"C": "ℂ", "D": "𝔻", "E": "𝔼", "F": "𝔽", "H": "ℍ", "K": "𝕂",
      "N": "ℕ", "P": "ℙ", "Q": "ℚ", "R": "ℝ", "Z": "ℤ"}


def _accent(s):
    for name, mark in ACCENTS.items():
        s = re.sub(r"\\" + name + r"\s*\{([^{}]*)\}",
                   lambda m: "".join(c + mark for c in m.group(1)), s)
        s = re.sub(r"\\" + name + r"\s+(\w)", lambda m: m.group(1) + mark, s)
    return s


# ---------- 2-D math typesetting ----------
# Renders LaTeX math as terminal text rows (folio-style: the formula is
# text, not a raster image). A _MBox is a block of rows plus the index of
# its baseline row; horizontal composition aligns baselines like TeX.


class _MBox:
    __slots__ = ("lines", "base", "big", "side")

    def __init__(self, lines, base=0, big=False, side=False):
        self.lines = lines
        self.base = base
        self.big = big      # limits stack above/below (Σ), not as scripts
        self.side = side    # limits attach to the side (∫)

    @property
    def w(self):
        return max((sum(_disp_w(c) for c in l) for l in self.lines), default=0)

    @property
    def h(self):
        return len(self.lines)


def _m_pad(s, w):
    return s + " " * max(0, w - sum(_disp_w(c) for c in s))


def _m_center(b, w):
    out = []
    for l in b.lines:
        left = max(0, (w - sum(_disp_w(c) for c in l)) // 2)
        out.append(" " * left + _m_pad(l, w - left))
    return out


def _m_atom(t):
    return _MBox([t], 0)


def _m_hcat(boxes):
    boxes = [b for b in boxes if b is not None and b.h]
    if not boxes:
        return _MBox([""], 0)
    base = max(b.base for b in boxes)
    h = base + max(b.h - b.base for b in boxes)
    rows = []
    for r in range(h):
        line = ""
        for b in boxes:
            i = r - (base - b.base)
            seg = b.lines[i] if 0 <= i < b.h else ""
            line += _m_pad(seg, b.w)
        rows.append(line.rstrip())
    return _MBox(rows, base)


def _m_frac(num, den):
    w = max(num.w, den.w) + 2
    return _MBox(_m_center(num, w) + ["─" * w] + _m_center(den, w), num.h)


SUP_CH = set("0123456789+-=()niabcedfghjklmoprtuvxABDEGHIJKLMNOPRTUVW")
SUB_CH = set("0123456789+-=()aeoxhklmnpstijruv")


def _m_script(base, sub, sup):
    if base.h == 1:
        ok = True
        t = base.lines[0]
        if sub and sub.h == 1 and set(sub.lines[0]) <= SUB_CH:
            t += sub.lines[0].translate(SUB)
        elif sub:
            ok = False
        if sup and sup.h == 1 and set(sup.lines[0]) <= SUP_CH:
            t += sup.lines[0].translate(SUP)
        elif sup:
            ok = False
        if ok:
            return _MBox([t], 0)
    off = base.w
    rows = []
    if sup:
        rows += [" " * off + l for l in sup.lines]
    rows += list(base.lines)
    if sub:
        rows += [" " * off + l for l in sub.lines]
    b = _MBox(rows, base.base + (sup.h if sup else 0))
    return b


def _m_limits(base, sub, sup):
    if base.side:
        return _m_script(base, sub, sup)
    w = max(base.w, sub.w if sub else 0, sup.w if sup else 0)
    rows = []
    if sup:
        rows += _m_center(sup, w)
    rows += _m_center(base, w)
    if sub:
        rows += _m_center(sub, w)
    return _MBox(rows, base.base + (sup.h if sup else 0))


def _m_attach(base, sub, sup):
    if base.big:
        return _m_limits(base, sub, sup)
    return _m_script(base, sub, sup)


def _m_sqrt(index, inner):
    if inner.h == 1:
        pre = (index.lines[0].translate(SUP) if index else "") + "√"
        return _MBox([pre + "(" + inner.lines[0] + ")"], 0)
    w = inner.w + 1
    rows = ["  " + "─" * w]
    for k, l in enumerate(inner.lines):
        rows.append(("╲╱ " if k == inner.base else "   ") + l)
    if index:
        rows[0] = index.lines[0] + rows[0][len(index.lines[0]):] \
            if len(index.lines[0]) <= 2 else "ⁿ" + rows[0][1:]
    return _MBox(rows, inner.base + 1)


_DELIMS = {"(": ("⎛", "⎜", "⎝", "⎜"), ")": ("⎞", "⎟", "⎠", "⎟"),
           "[": ("⎡", "⎢", "⎣", "⎢"), "]": ("⎤", "⎥", "⎦", "⎥"),
           "{": ("⎧", "⎨", "⎩", "⎪"), "}": ("⎫", "⎬", "⎭", "⎪"),
           "|": ("│", "│", "│", "│"), "‖": ("‖", "‖", "‖", "‖"),
           "⟨": ("⟨", "│", "⟨", "│"), "⟩": ("⟩", "│", "⟩", "│"),
           ".": ("", "", "", "")}
_DELIM_NAMES = {"(": "(", ")": ")", "[": "[", "]": "]", "{": "{", "}": "}",
                "|": "|", ".": ".", "vert": "|", "Vert": "‖", "lvert": "|",
                "rvert": "|", "langle": "⟨", "rangle": "⟩", "lbrace": "{",
                "rbrace": "}", "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈",
                "rceil": "⌉", "\\|": "‖"}
_FLAT_DELIM = {"(": "(", ")": ")", "[": "[", "]": "]", "{": "{", "}": "}",
               "|": "|", "‖": "‖", "⟨": "⟨", "⟩": "⟩", ".": ""}


def _m_tall_col(kind, h):
    if h <= 2:
        t, b = _DELIMS[kind][0], _DELIMS[kind][2]
        return [t] + [b] * (h - 1)
    t, mid, b, fill = _DELIMS[kind]
    center = (h - 1) // 2
    return [t if i == 0 else b if i == h - 1 else mid if i == center else fill
            for i in range(h)]


def _m_wrap(l, inner, r):
    if inner.h == 1:
        return _MBox([_FLAT_DELIM[l] + inner.lines[0] + _FLAT_DELIM[r]], 0)
    left = _m_tall_col(l, inner.h)
    right = _m_tall_col(r, inner.h)
    w = inner.w
    rows = [left[i] + _m_pad(inner.lines[i], w) + right[i] for i in range(inner.h)]
    return _MBox(rows, inner.base)


def _m_toks(s):
    # TeX math mode ignores newlines and collapses whitespace runs; \\ is a
    # separate row-break token and unaffected by this normalization.
    s = re.sub(r"[ \t\n\r]+", " ", s)
    toks, i, n = [], 0, len(s)
    while i < n:
        c = s[i]
        if c == "\\":
            if s.startswith("\\\\", i):
                toks.append("\\\\")
                i += 2
                continue
            m = re.match(r"\\begin\s*\{([A-Za-z*]+)\}", s[i:])
            if m:
                toks.append(("begin", m.group(1)))
                i += m.end()
                continue
            m = re.match(r"\\end\s*\{[A-Za-z*]+\}", s[i:])
            if m:
                toks.append(("end",))
                i += m.end()
                continue
            m = re.match(r"\\([A-Za-z]+)", s[i:])
            if m:
                toks.append(("cmd", m.group(1)))
                i += m.end()
                continue
            toks.append(("cmd", s[i + 1:i + 2] or " "))
            i += 2
            continue
        if c in "{}^_&[]":
            toks.append(c)
            i += 1
            continue
        j = i
        while j < n and s[j] not in "\\{}^_&[]":
            j += 1
        toks.append(s[i:j])
        i = j
    return toks


def _m_parse(toks, i, stops=()):
    cells = []
    while i < len(toks):
        t = toks[i]
        if t in stops or (isinstance(t, tuple) and t[0] == "end"):
            break
        if t == "{":
            b, i = _m_group(toks, i)
            cells.append(b)
            continue
        if t == "^" or t == "_":
            base = cells.pop() if cells else _m_atom("")
            i += 1
            a, i = _m_arg(toks, i)
            sub, sup = (a, None) if t == "_" else (None, a)
            if i < len(toks) and toks[i] in ("^", "_"):
                t2 = toks[i]
                i += 1
                a2, i = _m_arg(toks, i)
                if t2 == "^":
                    sup = a2
                else:
                    sub = a2
            cells.append(_m_attach(base, sub, sup))
            continue
        if t == "&" or t == "\\\\":
            break
        if isinstance(t, tuple):
            if t[0] == "begin":
                b, i = _m_env(toks, i)
            else:
                b, i = _m_cmd(toks, i)
            cells.append(b)
            continue
        if t.strip():
            cells.append(_m_atom(re.sub(r"\s*(<=|>=|<|>|=)\s*", r" \1 ", t)))
        i += 1
    return _m_hcat(cells), i


def _m_group(toks, i):
    b, i = _m_parse(toks, i + 1, ("}",))
    return b, i + 1


def _m_arg(toks, i):
    if i >= len(toks):
        return _m_atom(""), i
    t = toks[i]
    if t == "{":
        return _m_group(toks, i)
    if isinstance(t, tuple) and t[0] == "begin":
        return _m_env(toks, i)
    if isinstance(t, tuple):
        return _m_cmd(toks, i)
    if t in ("^", "_", "&", "\\\\"):
        return _m_atom(""), i
    if len(t) > 1:
        toks[i] = t[1:]
        return _m_atom(t[0]), i
    return _m_atom(t), i + 1


_BIGOPS = {"sum": "Σ", "prod": "Π", "coprod": "∐", "bigcup": "⋃", "bigcap": "⋂",
           "bigoplus": "⨁", "bigotimes": "⨂", "bigvee": "⋁", "bigwedge": "⋀",
           "bigsqcup": "⨆", "bigodot": "⨀", "biguplus": "⨄"}
_INTS = {"int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮", "oiint": "∯",
         "oiiint": "∰", "varoint": "∮", "ointctrclockwise": "∳"}
_BIGLIM = {"lim", "limsup", "liminf", "sup", "inf", "max", "min",
           "argmax", "argmin", "det", "gcd"}
_FNAMES = {"log", "ln", "lg", "exp", "sin", "cos", "tan", "sec", "csc", "cot",
           "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "coth", "sech",
           "csch", "sign", "rank", "tr", "diag", "dim", "ker", "deg", "hom",
           "Hom", "End", "Aut", "Pr", "mod", "bmod", "gcd", "lcm", "min", "max"}
_TEXTFONTS = {"text", "mathrm", "mathbf", "mathit", "mathsf", "mathtt", "mbox",
              "hbox", "operatorname", "textbf", "textit", "textrm", "textnormal",
              "textup", "mathcal", "mathscr",
              "mathfrak", "mathbfit", "displaystyle", "textstyle"}
_SPACES = {",": " ", ";": " ", ":": " ", " ": " ", "quad": "  ", "qquad": "    ",
           "!": "", "enspace": " ", "thinspace": " ", "medspace": " ",
           "thickspace": " ", "negthinspace": "", "negmedspace": "",
           "negthickspace": "", "~": " ", "nbsp": " "}
_SKIPARG = {"tag", "label", "nonumber", "notag", "eqref", "ref", "pageref",
            "kern", "hspace", "vspace", "mkern", "mspace", "rule", "phantom",
            "hphantom", "vphantom", "mathclap", "mathllap", "mathrlap",
            "clap", "llap", "rlap", "smash", "lefteqn", "color", "textcolor",
            "colorbox", "fcolorbox", "pagecolor", "definecolor", "size",
            "strut", "mathstrut", "noalign", "hline", "cline", "hdashline",
            "vspace", "hspace", "raisebox", "displaylimits", "nolimits",
            "limits", "displaystyle", "scriptstyle", "scriptscriptstyle",
            "intertext", "shortintertext", "allowbreak", "numberwithin",
            "ensuremath", "cr", "noalign", "relax"}


def _m_lit(b):
    """Flatten a box to a single literal line (for \\text{...} args)."""
    return re.sub(r" {2,}", " ", " ".join(l.strip() for l in b.lines)).strip()


def _m_accent(name, inner):
    if name in ("overline", "overbar"):
        if inner.h == 1:
            return _m_atom("".join(c + "̅" for c in inner.lines[0]))
        return _MBox(["‾" * inner.w] + inner.lines, inner.base + 1)
    if name == "underline":
        return _MBox(inner.lines + ["▁" * inner.w], inner.base)
    mark = ACCENTS.get(name, "̂")
    if inner.h == 1:
        return _m_atom("".join(c + mark for c in inner.lines[0]))
    lines = list(inner.lines)
    lines[inner.base] = "".join(c + mark for c in inner.lines[inner.base])
    return _MBox(lines, inner.base)


def _m_arrow(name, sub, sup):
    ch = "→" if "right" in name else "←"
    w = max(4, (sup.w if sup else 0) + 2, (sub.w if sub else 0) + 2)
    arrow = "─" * (w - 1) + ch if "right" in name else ch + "─" * (w - 1)
    rows = _m_center(sup, w) + [arrow] if sup else [arrow]
    if sub:
        rows += _m_center(sub, w)
    return _MBox(rows, sup.h if sup else 0)


def _m_optarg(toks, i):
    """[...] optional arg → literal string or None."""
    if i >= len(toks) or toks[i] != "[":
        return None, i
    buf, j = "", i + 1
    while j < len(toks) and toks[j] != "]":
        t = toks[j]
        buf += t[1] if isinstance(t, tuple) else t
        j += 1
    return buf, j + 1


def _m_delim(toks, i):
    """Delimiter after \\left/\\right/\\big*: a cmd token or the first
    char of a literal run (spaces before it are skipped, rest pushed back)."""
    if i >= len(toks):
        return ".", i
    t = toks[i]
    if isinstance(t, tuple):
        return _DELIM_NAMES.get(t[1], "."), i + 1
    stripped = t.lstrip()
    if not stripped:
        return ".", i + 1
    if stripped != t or len(stripped) > 1:
        toks[i] = stripped[1:]
        return _DELIM_NAMES.get(stripped[0], "."), i
    return _DELIM_NAMES.get(t, "."), i + 1


def _m_cmd(toks, i):
    name = toks[i][1]
    i += 1
    if name in ("frac", "dfrac", "tfrac", "cfrac"):
        a, i = _m_arg(toks, i)
        b, i = _m_arg(toks, i)
        return _m_frac(a, b), i
    if name == "sqrt":
        idx, i = _m_optarg(toks, i)
        c, i = _m_arg(toks, i)
        return _m_sqrt(_m_atom(idx) if idx else None, c), i
    if name in ("binom", "dbinom", "tbinom"):
        a, i = _m_arg(toks, i)
        b, i = _m_arg(toks, i)
        w = max(a.w, b.w)
        return _m_wrap("(", _MBox(_m_center(a, w) + _m_center(b, w), a.h), ")"), i
    if name == "left":
        l, i = _m_delim(toks, i)
        inner, i = _m_parse(toks, i, (("cmd", "right"),))
        r = "."
        if i < len(toks) and toks[i] == ("cmd", "right"):
            i += 1
            r, i = _m_delim(toks, i)
        inner.base = (inner.h - 1) // 2 if inner.h > 1 else 0
        return _m_wrap(l, inner, r), i
    if name in ("right", "middle"):
        _, i = _m_delim(toks, i)  # stray closer - drop its delimiter
        return _m_atom(""), i
    if name in ("bigl", "bigr", "Bigl", "Bigr", "biggl", "biggr", "Biggl",
                "Biggr", "big", "Big", "bigg", "Bigg", "bigl", "bigr"):
        ch, i = _m_delim(toks, i)
        return _m_atom(_FLAT_DELIM.get(ch, ch)), i
    if name in _BIGOPS or name in _INTS:
        b = _MBox([_BIGOPS.get(name) or _INTS[name]], 0, big=True,
                  side=name in _INTS)
        return b, i
    if name in _BIGLIM:
        pad = " " if i < len(toks) and toks[i] == "{" else ""
        return _MBox([name + pad], 0, big=True), i
    if name in _FNAMES:
        pad = " " if i < len(toks) and toks[i] == "{" else ""
        return _m_atom(name + pad), i
    if name in ("boldsymbol", "bm", "pmb"):
        g, i = _m_arg(toks, i)
        return _MBox([l.translate(_BOLDIT) for l in g.lines], g.base), i
    if name in _TEXTFONTS:
        g, i = _m_arg(toks, i)
        return _m_atom(_m_lit(g)), i
    if name == "mathbb":
        g, i = _m_arg(toks, i)
        return _m_atom("".join(BB.get(c, c) for c in _m_lit(g))), i
    if name in ACCENTS or name in ("overline", "underline", "overbar"):
        g, i = _m_arg(toks, i)
        return _m_accent(name, g), i
    if name in ("overset", "stackrel", "underset"):
        a, i = _m_arg(toks, i)
        b, i = _m_arg(toks, i)
        w = max(a.w, b.w) + 2
        if name == "underset":
            return _MBox(_m_center(b, w) + _m_center(a, w), b.h - 1), i
        top = _m_center(a, w)
        return _MBox(top + _m_center(b, w), len(top)), i
    if name in ("overbrace", "underbrace"):
        g, i = _m_arg(toks, i)
        w = g.w
        if name == "overbrace":
            return _MBox(["⏞" * w] + _m_center(g, w), g.h), i
        return _MBox(_m_center(g, w) + ["⏟" * w], g.base), i
    if name in ("xrightarrow", "xleftarrow"):
        opt, i = _m_optarg(toks, i)
        sup, i = _m_arg(toks, i)
        return _m_arrow(name, _m_atom(opt) if opt else None, sup), i
    if name in ("pmod", "mod"):
        g, i = _m_arg(toks, i)
        return _m_atom("(mod " + _m_lit(g) + ")"), i
    if name == "boxed":
        g, i = _m_arg(toks, i)
        w = g.w + 2
        rows = ["┌" + "─" * w + "┐"]
        for l in g.lines:
            rows.append("│ " + _m_pad(l, g.w) + " │")
        rows.append("└" + "─" * w + "┘")
        return _MBox(rows, g.base + 1), i
    if name == "substack":
        g, i = _m_arg(toks, i)
        return g, i
    if name == "not":
        g, i = _m_arg(toks, i)
        lit = _m_lit(g)
        return _m_atom(lit[:1] + "̸" + lit[1:] if lit else "̸"), i
    if name in _SPACES:
        return _m_atom(_SPACES[name]), i
    if name in _SKIPARG:
        if name in ("kern", "hspace", "vspace", "mkern", "mspace", "rule",
                    "raisebox", "phantom", "hphantom", "vphantom", "mathclap",
                    "mathllap", "mathrlap", "clap", "llap", "rlap", "smash",
                    "lefteqn", "color", "textcolor", "colorbox", "fcolorbox",
                    "tag", "label", "eqref", "ref", "pageref", "cline",
                    "intertext", "shortintertext", "ensuremath"):
            _, i = _m_optarg(toks, i)
            _, i = _m_arg(toks, i)
            if name in ("textcolor", "fcolorbox"):
                _, i = _m_arg(toks, i)
            if name in ("color",):
                _, i = _m_arg(toks, i)
        return _m_atom(""), i
    if name in GREEK:
        return _m_atom(GREEK[name]), i
    if name in OPS:
        pad = " " if name in _REL or name in _BINOPS else ""
        return _m_atom(pad + OPS[name] + pad), i
    if name in ("{", "}"):
        return _m_atom(name), i
    if name == "|":
        return _m_atom("‖"), i
    if name == "'":
        return _m_atom("′"), i
    return _m_atom(name), i


_ENVS = {"pmatrix": ("(", ")"), "bmatrix": ("[", "]"), "Bmatrix": ("{", "}"),
         "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"), "matrix": (".", "."),
         "smallmatrix": ("(", ")"), "cases": ("{", "."), "rcases": (".", "}"),
         "dcases": ("{", "."), "aligned": (".", "."), "align": (".", "."),
         "array": (".", "."), "gathered": (".", "."), "split": (".", "."),
         "eqnarray": (".", "."), "multlined": (".", "."), "bmod": (".", ".")}


def _m_env(toks, i):
    env = toks[i][1]
    i += 1
    if env == "array" and i < len(toks) and toks[i] == "{":
        _, i = _m_group(toks, i)  # column spec {cc} - not typeset
    rows, depth = [[]], 1
    while i < len(toks) and depth:
        t = toks[i]
        if isinstance(t, tuple) and t[0] == "begin":
            depth += 1
        elif isinstance(t, tuple) and t[0] == "end":
            depth -= 1
            if depth == 0:
                i += 1
                break
        if depth and (t == "\\\\" or t == ("cmd", "cr")):
            rows.append([])
            i += 1
            continue
        if depth:
            rows[-1].append(t)
            i += 1
    grid = []
    for r in rows:
        cells = [[]]
        for t in r:
            if t == "&":
                cells.append([])
            else:
                cells[-1].append(t)
        grid.append([_m_parse(c, 0)[0] for c in cells])
    grid = [r for r in grid if any(c.h and any(l.strip() for l in c.lines) for c in r)]
    if not grid:
        return _m_atom(""), i
    ncols = max(len(r) for r in grid)
    for r in grid:
        r += [_m_atom("")] * (ncols - len(r))
    colw = [max(r[ci].w for r in grid) for ci in range(ncols)]
    rightish = env in ("aligned", "align", "split", "eqnarray")
    leftish = env in ("cases", "dcases", "array", "multlined")

    def align(seg, ci):
        sw = sum(_disp_w(x) for x in seg)
        if rightish and ci % 2 == 0 and ci + 1 < ncols:
            return " " * max(0, colw[ci] - sw) + seg
        if leftish or rightish:
            return _m_pad(seg, colw[ci])
        left = max(0, (colw[ci] - sw) // 2)
        return " " * left + _m_pad(seg, colw[ci] - left)

    outlines, base_row = [], 0
    for ri, r in enumerate(grid):
        base = max(c.base for c in r)
        h = base + max(c.h - c.base for c in r)
        if ri == (len(grid) - 1) // 2:
            base_row = len(outlines) + base
        for rr in range(h):
            line = ""
            for ci, c in enumerate(r):
                k = rr - (base - c.base)
                seg = c.lines[k] if 0 <= k < c.h else ""
                line += align(seg, ci) + ("  " if ci < ncols - 1 else "")
            outlines.append(line.rstrip())
    inner = _MBox(outlines, base_row)
    l, r = _ENVS.get(env, (".", "."))
    return _m_wrap(l, inner, r), i


def math_2d(latex):
    """Typeset LaTeX math as text rows; None when unparseable/empty."""
    try:
        toks = _m_toks(latex)
        segs, cur, depth = [], [], 0
        for t in toks:
            if isinstance(t, tuple) and t[0] == "begin":
                depth += 1
            elif isinstance(t, tuple) and t[0] == "end":
                depth -= 1
            if t == "\\\\" and depth == 0:
                segs.append(cur)
                cur = []
            else:
                cur.append(t)
        segs.append(cur)
        lines = []
        for seg in segs:
            b, _ = _m_parse(seg, 0)
            lines += b.lines
        if not lines or not any(l.strip() for l in lines):
            return None
        return lines
    except Exception:
        return None


def math_unicode(s):
    s = s.replace("\\{", "\x01").replace("\\}", "\x02")
    prev = None
    while prev != s:
        prev = s
        s = _frac_sqrt_pass(s)
    s = re.sub(r"\\(?:" + "|".join(FONTS_BOLD) + r")\s*\{([^{}]*)\}",
               "\x1b[1m\\1\x1b[22m", s)
    s = re.sub(r"\\(?:" + "|".join(FONTS_BOLD) + r")\s+(\w)",
               "\x1b[1m\\1\x1b[22m", s)
    s = re.sub(r"\\mathbb\s*\{([^{}]*)\}",
               lambda m: "".join(BB.get(c, c) for c in m.group(1)), s)
    s = re.sub(r"\\mathbb\s+(\w)", lambda m: BB.get(m.group(1), m.group(1)), s)
    s = re.sub(r"\\(?:" + "|".join(FONTS) + r")\s*\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\(?:" + "|".join(FONTS) + r")\s+(\w)", r"\1", s)
    s = _accent(s)
    s = re.sub(r"\\boxed\s*\{([^{}]*)\}", r"⟦ \1 ⟧", s)
    s = re.sub(r"\\binom\s*\{([^{}]*)\}\s*\{([^{}]*)\}", r"C(\1,\2)", s)
    s = re.sub(r"\\(sum|prod|int|iint|iiint|oint|coprod|bigcup|bigcap|bigoplus|bigotimes)_\{([^{}]*)\}\^\{([^{}]*)\}",
               lambda m: OPS.get(m.group(1), m.group(1)) + m.group(2).translate(SUB) + m.group(3).translate(SUP), s)
    s = re.sub(r"\\(?:left|right|bigl|bigr|Bigl|Bigr|bigg|Bigg|big|Big|limits|displaystyle|quad|qquad)\b", " ", s)
    s = re.sub(r"\\[,;! ]", " ", s)
    s = re.sub(r"\\([A-Za-z]+)", lambda m: GREEK.get(m.group(1), OPS.get(m.group(1), m.group(1))), s)
    def _sup_of(g):
        return g.translate(SUP) if g and set(g) <= SUP_CH else "^(" + g + ")"
    def _sub_of(g):
        return g.translate(SUB) if g and set(g) <= SUB_CH else "_(" + g + ")"
    s = re.sub(r"\^\{([^{}]*)\}", lambda m: _sup_of(m.group(1)), s)
    s = re.sub(r"_\{([^{}]*)\}", lambda m: _sub_of(m.group(1)), s)
    s = re.sub(r"\^([A-Za-z0-9+\-=()*])", lambda m: _sup_of(m.group(1)), s)
    s = re.sub(r"_([A-Za-z0-9+\-=()])", lambda m: _sub_of(m.group(1)), s)
    s = re.sub(r"\b(min|max|lim|sup|inf|det|ker|dim|deg|rank|gcd|arg|Pr|hom)\s*\{", r"\1 {", s)
    s = s.replace("{", "").replace("}", "").replace("\\", " ").replace("~", " ")
    s = s.replace("\x01", "{").replace("\x02", "}")
    return re.sub(r" {2,}", " ", s).strip()


def math_flat(s, multiline=False):
    """Degrade math to flat text when the 2-D layout won't fit: environments
    collapse to their delimiters, & → ' ', and \\ → '\\n' (blocks) or '; '
    (inline contexts such as table cells where newlines are impossible)."""
    def envdelim(m):
        l, r = _ENVS.get(m.group(2), (".", "."))
        return (l if m.group(1) == "begin" else r).replace(".", "")
    s = re.sub(r"\\(begin|end)\s*\{([A-Za-z*]+)\}", envdelim, s)
    s = re.sub(r"\\cr\b", "; ", s)
    s = s.replace("\\\\", "\x03").replace("&", " ")
    s = re.sub(r"[ \t\r\n]+", " ", s)
    s = math_unicode(s)
    return s.replace("\x03", "\n" if multiline else "; ")


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
        r = subprocess.run([mmdc, "-i", tmp, "-o", out_path, "-t", "dark",
                            "-b", "#1e1e2e", "-s", "2", "-w", "1400", "-q"],
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
    s = re.sub(r"\$([^$\n]+)\$", lambda m: keep(ITAL + (math_unicode(m.group(1)) if not re.search(r"\\(begin|end)\b", m.group(1)) else math_flat(m.group(1))) + RESET), s)
    s = re.sub(r"!\[([^\]]*)\]\(([^)]*)\)", lambda m: keep(DIM + "[image: " + (m.group(1) or m.group(2)) + "]" + RESET), s)
    s = re.sub(r"\[\[([^\]]+)\]\]", lambda m: keep(UND + CYAN + m.group(1) + RESET), s)
    s = re.sub(r"\[([^\]]+)\]\(([^)]*)\)", lambda m: keep(UND + CYAN + m.group(1) + RESET + DIM + "(" + m.group(2).strip() + ")" + RESET), s)
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
        # | inside $...$ is math (determinant/norm bars), not a cell boundary.
        inner = line.strip().strip("|")
        cells, cur, in_math = [], "", False
        k = 0
        while k < len(inner):
            ch = inner[k]
            if ch == "\\" and k + 1 < len(inner):
                cur += inner[k:k + 2]
                k += 2
                continue
            if ch == "$":
                in_math = not in_math
            if ch == "|" and not in_math:
                cells.append(cur)
                cur = ""
            else:
                cur += ch
            k += 1
        cells.append(cur)
        return [re.sub(r"\\([\[\]|])", r"\1", c).strip() for c in cells]

    header = [render_inline(c) for c in cells_of(block[0])]
    data = [[render_inline(c) for c in cells_of(b)] for b in block[2:]]
    ncols = len(header)
    natural = [0] * ncols
    for r in [header] + data:
        for c in range(min(ncols, len(r))):
            natural[c] = max(natural[c], table.width(r[c]))
    dropped = 0
    if max_width:
        # Keep every column that fits at minimum width 4; grid_lines then
        # shrinks+wraps cells into the remaining space. Only columns that
        # truly cannot fit at all are dropped.
        keep, used = 0, 1
        for c in range(ncols):
            if used + 4 + 3 <= max_width:
                used += 4 + 3
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
    jobs = {}    # path -> (kind, src): PNG producers run in parallel post-parse
    slots = []   # (insert_idx, path, entry, fallback): blank rows spliced later
    i = 0
    fence = re.compile(r"^(\s*)(`{3,}|~{3,})\s*([\w+-]*)\s*$")
    PRE_MARK = "\x02"  # sentinel: line is pre-rendered, must not be re-wrapped

    def emit_pre(ls):
        out.extend(PRE_MARK + l for l in ls)

    def queue_png(kind, src):
        path = media_path(cache_base, kind, src)
        jobs.setdefault(path, (kind, src))
        return path

    def emit_media(path, caption=None, fallback=None):
        """Reserve placeholder rows (spliced after generation); main.lua
        overlays the real image into that sub-rect via image_show."""
        if caption:
            out.append(DIM + caption + RESET)
        entry = {"line": 0, "lines": 0, "path": path}
        slots.append((len(out), path, entry, fallback))
        media.append(entry)

    def emit_rich(text, prefix=""):
        """Text line; complex inline math ($..$ with environments) is
        promoted to a rendered media block instead of raw LaTeX."""
        segs = re.split(r"(\$[^$\n]+\$)", text)
        if len(segs) == 1:
            out.append(prefix + render_inline(text))
            return
        buf = ""

        def flush():
            nonlocal buf
            if buf.strip():
                out.append(prefix + render_inline(buf))
            buf = ""

        for s in segs:
            if len(s) > 2 and s.startswith("$") and s.endswith("$") and \
                    re.search(r"\\(begin|end)\b", s):
                flush()
                rows = math_2d(s[1:-1])
                if rows:
                    emit_pre(["    " + l for l in rows])
                else:
                    buf += math_unicode(s[1:-1])
            else:
                buf += s
        flush()

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
                if shutil.which("mmdc"):
                    emit_media(queue_png("mermaid", src), "◆ mermaid")
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
                elif os.path.exists(PLANTUML_JAR):
                    emit_media(queue_png("plantuml", src), "◆ plantuml")
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
        qm = re.match(r"^\s*>\s?", line)
        if st.startswith("$$") or (qm and qm.group(0) and
                                   line[qm.end():].strip().startswith("$$")):
            quote = DIM + "│ " + RESET if qm and not st.startswith("$$") else ""
            j = i
            body = st[2:] if not quote else line[qm.end():].strip()[2:]
            while not body.rstrip().endswith("$$") and j + 1 < len(lines):
                j += 1
                body += "\n" + (re.sub(r"^\s*>\s?", "", lines[j]) if quote
                                else lines[j])
            latex = body.rstrip().removesuffix("$$").strip()
            rows = math_2d(latex) if latex else None
            if rows and (not max_width or
                         max(sum(_disp_w(c) for c in l) for l in rows) <= max_width - 4):
                emit_pre([quote + "    " + l for l in rows])
            else:
                uni = math_flat(latex, multiline=True) if latex else ""
                for raw in uni.splitlines() or [""]:
                    for l in wrap_ansi(raw, max(10, (max_width or 80) - 4)) or [""]:
                        out.append(quote + "    " + (ITAL + l + RESET if uni
                                   else DIM + l + RESET))
                if not uni and latex:
                    out.append(quote + "    " + DIM
                               + re.sub(r"\s+", " ", latex).strip() + RESET)
            i = j + 1
            continue
        if st.startswith("|") and "|" in st[1:]:
            tbl, ni = parse_table(lines, i, max_width)
            if tbl:
                emit_pre(tbl)
                i = ni
                continue
        mimg = re.match(r"^!\[([^\]]*)\]\(([^)]*)\)\s*$", st)
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
            body = lm.group(3)
            if re.search(r"\$[^$\n]*\\(?:begin|end)\b[^$\n]*\$", body):
                emit_rich(body, lm.group(1) + CYAN + lm.group(2) + RESET + " ")
            else:
                hang_w = len(lm.group(1)) + len(lm.group(2)) + 1
                wl = wrap_ansi(render_inline(body), max(10, (max_width or 80) - hang_w))
                out.append(lm.group(1) + CYAN + lm.group(2) + RESET + " " + wl[0])
                out.extend(" " * hang_w + l for l in wl[1:])
            i += 1
            continue
        if st:
            emit_rich(line)
        else:
            out.append("")
        i += 1
    # Generate all queued media in parallel (mmdc/latex/plantuml are
    # independent subprocesses; serialized would multiply first-open time).
    if jobs:
        from concurrent.futures import ThreadPoolExecutor

        def _gen(item):
            path, (kind, src) = item
            if os.path.exists(path):
                return
            try:
                {"mermaid": mmdc_png, "plantuml": plantuml_png}[kind](src, path)
            except Exception:
                pass

        with ThreadPoolExecutor(max_workers=4) as ex:
            list(ex.map(_gen, jobs.items()))
    # Splice placeholder rows (or a readable fallback line) at each slot.
    # Ascending order + cumulative offset: slot indexes were recorded before
    # any rows were inserted, so each splice shifts positions below it.
    offset = 0
    for idx, path, entry, fb in sorted(slots, key=lambda t: t[0]):
        pos = idx + offset
        size = img_size(path) if os.path.exists(path) else None
        if size and size[0] > 0:
            rows = max(2, round(size[1] * (max_width or 60) / size[0] / 2))
            if max_height:
                rows = min(rows, max(4, max_height - 2))
            rows = min(rows, 30)
            out[pos:pos] = [""] * rows
            entry["line"], entry["lines"] = pos, rows
        else:
            rows = 1
            out[pos:pos] = ["    " + DIM + (fb or "[image unavailable]") + RESET]
            media.remove(entry)
        offset += rows
    # Wrap every wrappable line to the pane width so manifest line numbers map
    # 1:1 to screen rows (main.lua displays with Wrap.NO and overlays media).
    if max_width:
        idx_map, wrapped = {}, []
        for n, l in enumerate(out):
            idx_map[n] = len(wrapped)
            if l.startswith(PRE_MARK):
                wrapped.append(l[1:])
            else:
                wrapped.extend(wrap_ansi(l, max_width))
        for m in media:
            m["line"] = idx_map.get(m["line"], m["line"])
        out = wrapped
    else:
        out = [l[1:] if l.startswith(PRE_MARK) else l for l in out]
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
