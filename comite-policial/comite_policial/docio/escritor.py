"""Escritura de los cambios de la aplicación sobre el documento original,
conservando su formato. Siempre se parte del documento base (el importado),
por lo que la operación es repetible y no acumula deformaciones."""
import copy
import os
import tempfile
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from comite_policial.core.recuadros import (
    contenido, lineas_planas_izquierda, parrafos_derecha)
from comite_policial.docio.importador import abrir
from comite_policial.docio.mapa import detectar, bloque_texto_x
from comite_policial.docio.semana import (
    SEMANA_RE, ACUM_RE, funcion_semana, funcion_acum)
from comite_policial.docio.xml import (
    W_P, W_TR, cell_text, phys_cells, par_text, set_cell_text,
    reescribir_celda, reemplazar_regex_par, set_par_text, set_bloque_parrafos, norm)


# ── Tablas ────────────────────────────────────────────────────────────────────
def _sincronizar_filas(seg, filas: list):
    """Ajusta el nº de filas de datos clonando/eliminando y escribe los valores."""
    data = list(seg.data)
    while len(data) > len(filas):
        tr = data.pop()
        tr.getparent().remove(tr)
    while len(data) < len(filas) and data:
        nuevo = copy.deepcopy(data[-1])
        data[-1].addnext(nuevo)
        data.append(nuevo)
    for tr, valores in zip(data, filas):
        _escribir_fila(tr, valores)


def _escribir_fila(tr, valores):
    for tc, v in zip(phys_cells(tr), valores):
        set_cell_text(tc, str(v))


def _aplicar_tablas(segs, tablas: dict):
    for key, datos in tablas.items():
        seg = segs.get(key)
        if seg is None:
            continue
        _sincronizar_filas(seg, [list(r) for r in datos["rows"]])
        if seg.total is not None and datos.get("total"):
            _escribir_fila(seg.total, datos["total"])


# ── Semana / turno ────────────────────────────────────────────────────────────
def _dentro_de(el, tablas_excluidas: set) -> bool:
    p = el.getparent()
    while p is not None:
        if p in tablas_excluidas:
            return True
        p = p.getparent()
    return False


def _aplicar_semana_turno(doc, recuadro_tbls, semana, semana_imp, turno, turno_imp):
    excl = set(recuadro_tbls)
    if semana and semana != semana_imp:
        f1, f2 = funcion_semana(semana), funcion_acum(semana)
        if f1:
            for p in doc.element.body.iter(W_P):
                if _dentro_de(p, excl):
                    continue
                reemplazar_regex_par(p, SEMANA_RE, f1)
                reemplazar_regex_par(p, ACUM_RE, f2)
    if turno and turno != turno_imp:
        for p in doc.element.body.iter(W_P):
            if _dentro_de(p, excl):
                continue
            t = par_text(p).strip()
            if norm(t).startswith("TURNO") and ":" in t:
                import re
                reemplazar_regex_par(
                    p, re.compile(r"(?<=:)\s*.+$"), lambda m: " " + turno)
                break


# ── Recuadros ─────────────────────────────────────────────────────────────────
def _reescribir_izq(tbl, rec):
    izq = phys_cells(tbl.findall(W_TR)[0])[0]
    parrafos = [[(ln, True)] if ln else [] for ln in lineas_planas_izquierda(rec)]
    reescribir_celda(izq, parrafos, centrado_bold=True)


def _reescribir_der(tbl, rec):
    der = phys_cells(tbl.findall(W_TR)[0])[1]
    reescribir_celda(der, parrafos_derecha(rec))


def _aplicar_cambios(tbl, rec, forzar=False):
    snap = rec.get("_snap")
    actual = contenido(rec)
    if forzar or not snap or snap.get("izq") != actual["izq"]:
        _reescribir_izq(tbl, rec)
    if forzar or not snap or snap.get("der") != actual["der"]:
        _reescribir_der(tbl, rec)


def _aplicar_recuadros(doc, rec_tbls, recuadros):
    from comite_policial.exporters.recuadros_export import (
        crear_recuadro_word, TITULO_SECCION)
    usados, elems = set(), []
    for rec in recuadros:
        o = rec.get("_orig")
        if o is not None and 0 <= o < len(rec_tbls) and o not in usados:
            usados.add(o)
            el = rec_tbls[o]
            _aplicar_cambios(el, rec)
        else:
            c = rec.get("_clone_of")
            base = (rec_tbls[c] if c is not None and 0 <= c < len(rec_tbls)
                    else (rec_tbls[0] if rec_tbls else None))
            if base is not None:
                el = copy.deepcopy(base)
                _aplicar_cambios(el, rec, forzar=True)
            else:
                el = None
                pass
        elems.append(el)

    body = doc.element.body
    if rec_tbls:
        gap_tpl = None
        sib = rec_tbls[0].getnext()
        if sib is not None and sib.tag == W_P and not par_text(sib).strip():
            gap_tpl = sib
        marcadores = []
        for slot in rec_tbls:
            m = OxmlElement("w:p")
            slot.addprevious(m)
            slot.getparent().remove(slot)
            marcadores.append(m)
        ultimo = None
        for k, el in enumerate(elems):
            if k < len(marcadores):
                marcadores[k].addprevious(el)
                ultimo = el
            else:
                gap = (copy.deepcopy(gap_tpl) if gap_tpl is not None
                       else OxmlElement("w:p"))
                ultimo.addnext(gap)
                gap.addnext(el)
                ultimo = el
        for m in marcadores:
            m.getparent().remove(m)
        # (los párrafos en blanco sobrantes se conservan para no alterar el espaciado)
    elif recuadros:
        # El documento no tenía recuadros: se crean desde cero al final
        doc.add_heading(TITULO_SECCION, level=1)
        for rec in recuadros:
            crear_recuadro_word(doc, rec)
            doc.add_paragraph()


# ── API ───────────────────────────────────────────────────────────────────────
def escribir(base_bytes: bytes, estado: dict, estado_import: dict, destino: str):
    """Aplica `estado` (datos actuales de la app) sobre el documento base y lo
    guarda en `destino` (reemplazo atómico)."""
    doc = abrir(base_bytes)
    segs, rec_tbls = detectar(doc)

    _aplicar_tablas(segs, estado.get("tablas", {}))
    _aplicar_semana_turno(doc, rec_tbls,
                          estado.get("semana", ""), estado_import.get("semana", ""),
                          estado.get("turno", ""), estado_import.get("turno", ""))

    texto_x = estado.get("texto_x")
    if texto_x is not None and texto_x.strip() != estado_import.get("texto_x", "").strip():
        pars = bloque_texto_x(doc)
        if pars:
            set_bloque_parrafos(pars, [l for l in texto_x.split("\n")])

    _aplicar_recuadros(doc, rec_tbls, estado.get("recuadros", []))

    carpeta = os.path.dirname(os.path.abspath(destino))
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=carpeta)
    os.close(fd)
    try:
        doc.save(tmp)
        os.replace(tmp, destino)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
