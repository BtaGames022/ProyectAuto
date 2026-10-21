"""Utilidades de bajo nivel para leer/escribir el XML de un .docx preservando su formato."""
import copy
import re
import unicodedata
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_P, W_R, W_T, W_TC, W_TR = qn("w:p"), qn("w:r"), qn("w:t"), qn("w:tc"), qn("w:tr")
W_BR, W_TAB, W_CR = qn("w:br"), qn("w:tab"), qn("w:cr")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def norm(s: str) -> str:
    """Mayúsculas sin tildes ni espacios sobrantes (para comparar encabezados)."""
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.upper().split())


# ── Lectura ───────────────────────────────────────────────────────────────────
def par_text(p) -> str:
    out = []
    for el in p.iter():
        if el.tag == W_T:
            out.append(el.text or "")
        elif el.tag in (W_BR, W_CR):
            out.append("\n")
        elif el.tag == W_TAB:
            out.append("\t")
    return "".join(out)


def cell_text(tc) -> str:
    """Texto de una celda: líneas no vacías unidas por salto de línea."""
    lines = []
    for p in tc.findall(W_P):
        for ln in par_text(p).split("\n"):
            if ln.strip():
                lines.append(ln.strip())
    return "\n".join(lines)


def phys_cells(tr):
    return tr.findall(W_TC)


def es_negrita(rpr) -> bool:
    if rpr is None:
        return False
    b = rpr.find(qn("w:b"))
    if b is None:
        return False
    return b.get(qn("w:val"), "true").lower() not in ("0", "false", "off")


def par_runs(p):
    """[(texto, negrita)] de los runs con texto de un párrafo."""
    out = []
    for r in p.iter(W_R):
        txt = "".join(t.text or "" for t in r.findall(W_T))
        if txt:
            out.append((txt, es_negrita(r.find(qn("w:rPr")))))
    return out


# ── Escritura ─────────────────────────────────────────────────────────────────
def _new_par(ppr, rpr, texto):
    p = OxmlElement("w:p")
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    if texto:
        p.append(_new_run(rpr, texto))
    return p


def _new_run(rpr, texto):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = texto
    r.append(t)
    return r


def _first_text_par(tc):
    paras = tc.findall(W_P)
    return next((p for p in paras if par_text(p).strip()), paras[0])


def _first_run_rpr(p):
    for r in p.iter(W_R):
        if "".join(t.text or "" for t in r.findall(W_T)).strip():
            return r.find(qn("w:rPr"))
    return None


def set_cell_text(tc, texto: str) -> bool:
    """Reemplaza el texto de una celda conservando el formato (pPr/rPr) del
    primer párrafo con texto. No hace nada si el texto ya coincide."""
    texto = (texto or "").strip()
    if cell_text(tc) == texto:
        return False
    paras = tc.findall(W_P)
    if not paras:
        return False
    tpl = _first_text_par(tc)
    ppr = tpl.find(qn("w:pPr"))
    rpr = _first_run_rpr(tpl)
    if rpr is None:
        rpr = _first_run_rpr(paras[0])
    ppr_c = copy.deepcopy(ppr) if ppr is not None else None
    rpr_c = copy.deepcopy(rpr) if rpr is not None else None
    for p in paras:
        tc.remove(p)
    for linea in (texto.split("\n") if texto else [""]):
        tc.append(_new_par(ppr_c, rpr_c, linea))
    return True


def plantillas_celda(tc):
    """(pPr, rPr_negrita, rPr_normal) tomados de la propia celda."""
    p = _first_text_par(tc)
    ppr = p.find(qn("w:pPr"))
    bold = norm_ = None
    for par in tc.findall(W_P):
        for r in par.iter(W_R):
            if not "".join(t.text or "" for t in r.findall(W_T)).strip():
                continue
            rpr = r.find(qn("w:rPr"))
            if es_negrita(rpr):
                bold = bold if bold is not None else rpr
            else:
                norm_ = norm_ if norm_ is not None else rpr
    base = bold if bold is not None else norm_
    if bold is None and base is not None:
        bold = copy.deepcopy(base)
        bold.append(OxmlElement("w:b"))
    if norm_ is None and base is not None:
        norm_ = copy.deepcopy(base)
        for b in norm_.findall(qn("w:b")):
            norm_.remove(b)
    return ppr, bold, norm_


def reescribir_celda(tc, parrafos, centrado_bold=False):
    """Rehace los párrafos de una celda. `parrafos`: lista de [(texto, negrita)];
    [] = párrafo en blanco. Usa los formatos de la celda original."""
    ppr, rbold, rnorm = plantillas_celda(tc)
    ppr_c = copy.deepcopy(ppr) if ppr is not None else None
    for p in tc.findall(W_P):
        tc.remove(p)
    if not parrafos:
        parrafos = [[]]
    for par in parrafos:
        p = _new_par(ppr_c, None, "")
        for texto, negrita in par:
            rpr = rbold if (negrita or centrado_bold) else rnorm
            p.append(_new_run(rpr if rpr is not None else rbold, texto))
        tc.append(p)


# ── Reemplazo con regex respetando runs ───────────────────────────────────────
def _set_run_text(r, texto):
    ts = r.findall(W_T)
    if not ts:
        if not texto:
            return
        t = OxmlElement("w:t")
        r.append(t)
        ts = [t]
    ts[0].text = texto
    ts[0].set(XML_SPACE, "preserve")
    for t in ts[1:]:
        r.remove(t)


def reemplazar_regex_par(p, regex, fn) -> bool:
    """Aplica regex→fn(match) sobre el texto de un párrafo sin perder el formato
    de los runs. Devuelve True si hubo cambios."""
    runs = list(p.iter(W_R))
    texts = ["".join(t.text or "" for t in r.findall(W_T)) for r in runs]
    full = "".join(texts)
    matches = list(regex.finditer(full))
    if not matches:
        return False
    for m in reversed(matches):
        s, e = m.span()
        repl = fn(m)
        offs, pos = [], 0
        for t in texts:
            offs.append(pos)
            pos += len(t)
        a = b = None
        for i, t in enumerate(texts):
            if not t:
                continue
            if a is None and offs[i] <= s < offs[i] + len(t):
                a = i
            if offs[i] < e <= offs[i] + len(t):
                b = i
        if a is None or b is None:
            continue
        if a == b:
            texts[a] = texts[a][:s - offs[a]] + repl + texts[a][e - offs[a]:]
        else:
            texts[a] = texts[a][:s - offs[a]] + repl
            for i in range(a + 1, b):
                texts[i] = ""
            texts[b] = texts[b][e - offs[b]:]
    for r, t in zip(runs, texts):
        if "".join(x.text or "" for x in r.findall(W_T)) != t:
            _set_run_text(r, t)
    return True


def set_par_text(p, texto: str):
    """Pone el texto de un párrafo en su primer run con texto (conserva formato)."""
    runs = list(p.iter(W_R))
    con_texto = [r for r in runs if r.findall(W_T)]
    if not con_texto:
        p.append(_new_run(None, texto))
        return
    _set_run_text(con_texto[0], texto)
    for r in con_texto[1:]:
        _set_run_text(r, "")


def set_bloque_parrafos(pars: list, lineas: list):
    """Asigna `lineas` a una lista de párrafos consecutivos: reutiliza los
    existentes, clona el último si faltan y elimina los sobrantes."""
    if not pars:
        return
    lineas = lineas or [""]
    for i, linea in enumerate(lineas):
        if i < len(pars):
            set_par_text(pars[i], linea)
        else:
            nuevo = copy.deepcopy(pars[-1])
            set_par_text(nuevo, linea)
            pars[-1].addnext(nuevo)
            pars.append(nuevo)
    for p in pars[len(lineas):]:
        p.getparent().remove(p)
