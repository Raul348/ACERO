"""Mini-conversor de fórmulas en texto a ecuaciones de Word (OMML, m:oMath).

Soporta: fracciones a/b, raíces √(...), subíndices (f'm, Vm, Pg, Vc...), superíndices (², ³, ^n),
productos con "·", letras griegas (α, φ, Σ, γ, Δ, δ, μ, σ) y paréntesis. Si la cadena no se puede
convertir, omml() devuelve None y la memoria la escribe como texto.
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

SUPERS = {"²": "2", "³": "3", "⁴": "4", "⁻": "-", "¹": "1"}
# variables con subíndice (base, subíndice); la clave es la cadena como se escribe en las fórmulas
SUBS = {
    "f'm": ("f'", "m"), "f'c": ("f'", "c"), "f'b": ("f'", "b"), "f't": ("f'", "t"), "v'm": ("v'", "m"),
    "fy": ("f", "y"), "Vm1": ("V", "m1"), "Ve1": ("V", "e1"), "Mu1": ("M", "u1"), "Me1": ("M", "e1"),
    "Vm": ("V", "m"), "Ve": ("V", "e"), "VE": ("V", "E"), "Vu": ("V", "u"), "Mu": ("M", "u"), "Me": ("M", "e"),
    "Pg": ("P", "g"), "Pm": ("P", "m"), "Pc": ("P", "c"), "Pt": ("P", "t"), "Pp": ("P", "p"),
    "Vc": ("V", "c"), "Lm": ("L", "m"), "Nc": ("N", "c"), "Ts": ("T", "s"), "Em": ("E", "m"),
    "Gm": ("G", "m"), "Ac": ("A", "c"), "An": ("A", "n"), "As": ("A", "s"), "Asf": ("A", "sf"),
    "Ast": ("A", "st"), "Acf": ("A", "cf"), "Acs": ("A", "cs"), "Av": ("A", "v"), "Fa": ("F", "a"),
    "Ms": ("M", "s"), "fm": ("f", "m"), "fa": ("f", "a"), "C1": ("C", "1"), "Tp": ("T", "p"),
    "Tl": ("T", "l"), "hn": ("h", "n"), "CT": ("C", "T"), "R0": ("R", "0"), "Ia": ("I", "a"),
    "Ip": ("I", "p"), "Fi": ("F", "i"), "Pi": ("P", "i"), "hi": ("h", "i"), "Pj": ("P", "j"),
    "hj": ("h", "j"), "Hj": ("H", "j"), "αi": ("α", "i"), "VEi": ("V", "Ei"), "Vei": ("V", "ei"),
    "Fj": ("F", "j"), "tn": ("t", "n"), "s1": ("s", "1"), "s2": ("s", "2"), "s3": ("s", "3"),
    "s4": ("s", "4"), "δe": ("δ", "e"), "Δe": ("Δ", "e"), "hs": ("h", "s"), "At": ("A", "t"),
    "wd": ("w", "d"), "wl": ("w", "l"), "Hi-1": ("H", "i-1"), "σm": ("σ", "m"), "Lpaño": ("L", "paño"),
    "Lmáx": ("L", "máx"), "Lconf": ("L", "conf"), "Fa1": ("F", "a1"), "Fa2": ("F", "a2"), "dmín": ("d", "mín"),
}
_SUB_KEYS = sorted(SUBS, key=len, reverse=True)
D_OPEN = "<m:d><m:dPr><m:begChr m:val=\"(\"/><m:endChr m:val=\")\"/></m:dPr><m:e>"
D_CLOSE = "</m:e></m:d>"


# ------------------------------------------------------------------ tokens
def _tokenizar(s: str) -> list[str]:
    toks = []
    i = 0
    while i < len(s):
        c = s[i]
        if c.isspace():
            toks.append(" ")
            while i < len(s) and s[i].isspace():
                i += 1
            continue
        if c in "()/^·+-=≤≥<>,√*×":
            toks.append(c)
            i += 1
            continue
        if c in SUPERS:
            j = i
            while j < len(s) and s[j] in SUPERS:
                j += 1
            toks.append("^")
            toks.append("".join(SUPERS[x] for x in s[i:j]))
            i = j
            continue
        m = re.match(r"[0-9]+(?:\.[0-9]+)?(?:e[-+]?[0-9]+)?", s[i:])
        if m and c.isdigit():
            toks.append(m.group(0))
            i += len(m.group(0))
            continue
        m = re.match(r"[A-Za-zÁÉÍÓÚáéíóúñÑα-ωΑ-Ω'_0-9]+", s[i:])
        if m:
            toks.append(m.group(0))
            i += len(m.group(0))
            continue
        toks.append(c)
        i += 1
    return toks


# ------------------------------------------------------------------ nodos
def _r(texto: str, normal: bool = False) -> str:
    rpr = "<m:rPr><m:sty m:val=\"p\"/></m:rPr>" if normal else ""
    return f"<m:r>{rpr}<m:t xml:space=\"preserve\">{escape(texto)}</m:t></m:r>"


def _ident(tok: str) -> str:
    if len(tok) > 1 and tok[0] in "ΣΔ" and tok not in SUBS:
        return _r(tok[0]) + _ident(tok[1:])
    if tok in SUBS:
        b, sb = SUBS[tok]
        return f"<m:sSub><m:e>{_r(b)}</m:e><m:sub>{_r(sb)}</m:sub></m:sSub>"
    if re.fullmatch(r"[0-9.e+-]+", tok):
        return _r(tok, normal=True)
    if len(tok) > 1 and tok.isalpha() and tok.lower() in ("mín", "máx", "min", "max", "mín", "dmín"):
        return _r(tok, normal=True)
    return _r(tok)


class _Parser:
    """Parser recursivo: expr := term (op term)* ; term := factor ('/' factor)* ; factor := atom ('^' atom)?"""

    def __init__(self, toks: list[str]):
        self.t = [x for x in toks]
        self.i = 0
        self.grupo = False

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else None

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def parse(self) -> str:
        out = self.expr(stop=None)
        if self.i < len(self.t):
            raise ValueError("tokens sobrantes")
        return out

    def expr(self, stop) -> str:
        parts = []
        while self.peek() is not None and self.peek() != stop:
            if self.peek() == ")" and stop is None:
                raise ValueError("paréntesis")
            parts.append(self.term(stop))
            tok = self.peek()
            if tok in ("+", "-", "=", "≤", "≥", "<", ">", ",", "·", "×", "*"):
                self.take()
                sym = {"*": "·"}.get(tok, tok)
                parts.append(_r(f" {sym} " if sym in "=≤≥<>" else sym, normal=True))
            elif tok == " ":
                self.take()
                parts.append(_r(" ", normal=True))
        return "".join(parts)

    def term(self, stop) -> str:
        num, gnum = self.factor(stop)
        while self.peek() == "/":
            self.take()
            den, gden = self.factor(stop)
            if gnum:
                num = num[len(D_OPEN):-len(D_CLOSE)]
            if gden:
                den = den[len(D_OPEN):-len(D_CLOSE)]
            num, gnum = f"<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>", False
        return num

    def factor(self, stop) -> tuple[str, bool]:
        self.grupo = False
        base = self.atom(stop)
        grupo = self.grupo
        if self.peek() == "^":
            grupo = False
            self.take()
            exp = self.atom(stop, quitar_parentesis=True)
            base = f"<m:sSup><m:e>{base}</m:e><m:sup>{exp}</m:sup></m:sSup>"
        return base, grupo

    def atom(self, stop, quitar_parentesis: bool = False) -> str:
        tok = self.take()
        if tok is None:
            raise ValueError("fin inesperado")
        if tok == " ":
            return self.atom(stop, quitar_parentesis)
        if tok == "(":
            dentro = self.expr(stop=")")
            if self.take() != ")":
                raise ValueError("falta )")
            if quitar_parentesis:
                return dentro
            self.grupo = True
            return D_OPEN + dentro + D_CLOSE
        if tok == "√":
            if self.peek() == "(":
                self.take()
                dentro = self.expr(stop=")")
                if self.take() != ")":
                    raise ValueError("falta )")
            else:
                dentro = self.atom(stop)
            return f"<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/><m:e>{dentro}</m:e></m:rad>"
        if tok in ("-", "+"):
            return _r(tok, normal=True) + self.atom(stop)
        if tok in (")", "/", "^"):
            raise ValueError(f"token inesperado {tok}")
        return _ident(tok)


def omml(formula: str) -> str | None:
    """Devuelve el XML <m:oMath> de la fórmula o None si no se puede convertir."""
    if not formula or not formula.strip():
        return None
    try:
        cuerpo = _Parser(_tokenizar(formula.strip())).parse()
    except (ValueError, IndexError, RecursionError):
        return None
    return f"<m:oMath xmlns:m=\"{M_NS}\" xmlns:w=\"{W_NS}\">{cuerpo}</m:oMath>"
