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
    SEMANA_RE, ACUM_RE, SERV_FECHAS_RE, funcion_semana, funcion_acum)
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
        filas = [list(r) for r in datos["rows"]]
        if key == "migr" and seg.ncell == 7:
            # el documento es del formato anterior (sin EXTENAS): se omite esa columna
            filas = [r[:6] + r[7:] for r in filas]
        if key == "homi_d":
            # En el Word la variación solo figura en la fila TOTAL; si las filas
            # de datos venían en blanco, se respetan (la app las calcula aparte).
            en_blanco = all(not cell_text(phys_cells(tr)[-1]).strip() for tr in seg.data)
            if en_blanco:
                filas = [r[:-1] + [""] for r in filas]
        _sincronizar_filas(seg, filas)
        if seg.total is not None and datos.get("total"):
            total = list(datos["total"])
            if key == "migr" and seg.ncell == 7:
                total = total[:6] + total[7:]
            if key == "homi_d" and str(total[-1]).endswith("%"):
                total[-1] = str(total[-1]).replace(".", ",")   # el Word usa coma: -11,9%
            _escribir_fila(seg.total, total)



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
# Cada recuadro es una fila (w:tr) de 2 celdas. Una tabla puede traer varios.
def _reescribir_izq(tr, rec):
    izq = phys_cells(tr)[0]
    parrafos = [[(ln, True)] if ln else [] for ln in lineas_planas_izquierda(rec)]
    reescribir_celda(izq, parrafos, centrado_bold=True)


def _reescribir_der(tr, rec):
    der = phys_cells(tr)[1]
    reescribir_celda(der, parrafos_derecha(rec))


def _aplicar_cambios(tr, rec, forzar=False):
    snap = rec.get("_snap")
    actual = contenido(rec)
    if forzar or not snap or snap.get("izq") != actual["izq"]:
        _reescribir_izq(tr, rec)
    if forzar or not snap or snap.get("der") != actual["der"]:
        _reescribir_der(tr, rec)


def _tabla_nueva(tr):
    """Copia de la tabla de `tr` que conserva solo esa fila (un recuadro por tabla)."""
    tbl = tr.getparent()
    pos = tbl.findall(W_TR).index(tr)
    nueva = copy.deepcopy(tbl)
    for k, fila in enumerate(nueva.findall(W_TR)):
        if k != pos:
            nueva.remove(fila)
    return nueva


def _tabla_de(tr):
    """La propia tabla si tiene una sola fila; si no, una copia de una fila."""
    tbl = tr.getparent()
    return tbl if len(tbl.findall(W_TR)) == 1 else _tabla_nueva(tr)


def _tablas_distintas(filas):
    out = []
    for tr in filas:
        if tr.getparent() not in out:
            out.append(tr.getparent())
    return out


def _aplicar_recuadros(doc, rec_rows, recuadros):
    from comite_policial.exporters.recuadros_export import (
        crear_recuadro_word, TITULO_SECCION)

    # Misma lista, mismo orden (solo ediciones): se modifica cada fila en el lugar
    if rec_rows and [r.get("_orig") for r in recuadros] == list(range(len(rec_rows))):
        for tr, rec in zip(rec_rows, recuadros):
            _aplicar_cambios(tr, rec)
        return

    usados, elems = set(), []
    for rec in recuadros:
        o = rec.get("_orig")
        if o is not None and 0 <= o < len(rec_rows) and o not in usados:
            usados.add(o)
            _aplicar_cambios(rec_rows[o], rec)
            el = _tabla_de(rec_rows[o])
        else:
            c = rec.get("_clone_of")
            base = (rec_rows[c] if c is not None and 0 <= c < len(rec_rows)
                    else (rec_rows[0] if rec_rows else None))
            if base is not None:
                el = _tabla_nueva(base)
                _aplicar_cambios(el.findall(W_TR)[0], rec, forzar=True)
            else:
                el = None
        elems.append(el)

    body = doc.element.body
    rec_tbls = _tablas_distintas(rec_rows)
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
    segs, rec_rows = detectar(doc)

    _aplicar_tablas(segs, estado.get("tablas", {}))
    _aplicar_semana_turno(doc, _tablas_distintas(rec_rows),
                          estado.get("semana", ""), estado_import.get("semana", ""),
                          estado.get("turno", ""), estado_import.get("turno", ""))

    fechas = estado.get("servicio_fechas")
    if fechas is not None and fechas.strip() != estado_import.get("servicio_fechas", "").strip():
        for p in doc.element.body.iter(W_P):
            if "SERVICIO ESPECIAL DE FISCALIZACION" in norm(par_text(p)):
                reemplazar_regex_par(p, SERV_FECHAS_RE,
                                     lambda m: m.group(1) + fechas.strip())
                break

    texto_x = estado.get("texto_x")
    if texto_x is not None and texto_x.strip() != estado_import.get("texto_x", "").strip():
        pars = bloque_texto_x(doc)
        if pars:
            set_bloque_parrafos(pars, [l for l in texto_x.split("\n")])

    _aplicar_recuadros(doc, rec_rows, estado.get("recuadros", []))

    carpeta = os.path.dirname(os.path.abspath(destino))
    fd, tmp = tempfile.mkstemp(suffix=".docx", dir=carpeta)
    os.close(fd)
    try:
        doc.save(tmp)
        os.replace(tmp, destino)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
