"""Localiza en el documento las tablas que corresponden a cada sección de la app."""
from docx.oxml.ns import qn
from comite_policial.docio.xml import (W_TR, cell_text, phys_cells, norm)

W_TBL, W_P = qn("w:tbl"), qn("w:p")


class Segmento:
    """Zona de una tabla: filas de datos + fila de total (+ fuente)."""
    def __init__(self, tbl, key, data, total, fuente, ncell):
        self.tbl, self.key = tbl, key
        self.data, self.total, self.fuente, self.ncell = data, total, fuente, ncell


def _filas(tbl):
    return tbl.findall(W_TR)


def _primer_texto(tr) -> str:
    cs = phys_cells(tr)
    return norm(cell_text(cs[0])) if cs else ""


def _escanear(tbl, key, hdr_rows, ncell) -> Segmento:
    rows = _filas(tbl)
    i = max(hdr_rows) + 1
    data, total, fuente = [], None, None
    while i < len(rows):
        cs = phys_cells(rows[i])
        ft = norm(cell_text(cs[0])) if cs else ""
        if ft.startswith("TOTAL"):
            total = rows[i]
            i += 1
            break
        if ft.startswith("FUENTE") or len(cs) != ncell:
            break
        data.append(rows[i])
        i += 1
    if i < len(rows) and _primer_texto(rows[i]).startswith("FUENTE"):
        fuente = rows[i]
    return Segmento(tbl, key, data, total, fuente, ncell)


def detectar(doc):
    """Devuelve (segmentos: {clave: Segmento}, recuadros: [tr]).

    Cada recuadro es una fila (w:tr) de 2 celdas; una misma tabla puede traer
    varios recuadros (una fila por recuadro)."""
    segs, recuadros = {}, []
    body = doc.element.body
    for tbl in body.iter(W_TBL):
        if tbl.getparent() is not body:        # ignorar tablas anidadas
            continue
        rows = _filas(tbl)
        if not rows:
            continue
        r0 = [norm(cell_text(c)) for c in phys_cells(rows[0])]
        if not r0:
            continue
        # Recuadro de la sección X (IX en informes anteriores): filas de 2 celdas, derecha empieza con "Fecha"
        if len(r0) == 2 and all(len(phys_cells(r)) == 2 for r in rows):
            if all(norm(cell_text(phys_cells(r)[1])).startswith("FECHA") for r in rows):
                recuadros.extend(rows)
            continue
        h, n = r0[0], len(r0)
        if h == "TIPO" and n == 3:
            segs.setdefault("armas", _escanear(tbl, "armas", [0], 3))
        elif h == "DROGA" and n == 3:
            primera = _primer_texto(rows[1]) if len(rows) > 1 else ""
            if "drogas_g" in segs or "PLANTA" in primera or "(UNIDAD)" in primera:
                # tabla propia de plantas / unidades (formato actual)
                segs.setdefault("drogas_p", _escanear(tbl, "drogas_p", [0], 3))
            else:
                segs.setdefault("drogas_g", _escanear(tbl, "drogas_g", [0], 3))
                # formato antiguo: ambas tablas dentro de una sola
                for j in range(1, len(rows)):
                    if _primer_texto(rows[j]) == "DROGA":
                        segs.setdefault("drogas_p", _escanear(tbl, "drogas_p", [j], 3))
                        break
        elif h == "HOMICIDIOS" and n == 3 and r0[1].startswith("SEMANA"):
            segs.setdefault("homi_a", _escanear(tbl, "homi_a", [0], 3))
        elif h == "HOMICIDIOS" and n == 3 and "01.ENE" in r0[1]:
            segs.setdefault("homi_c", _escanear(tbl, "homi_c", [0], 3))
        elif h == "HOMICIDIOS" and n == 4 and r0[3].startswith("VAR"):
            # III-D comparativo interanual (antes venía pegado como imagen)
            segs.setdefault("homi_d", _escanear(tbl, "homi_d", [0], 4))
        elif h == "MODALIDAD":
            segs.setdefault("homi_b", _escanear(tbl, "homi_b", [0, 1], 5))
        elif h == "UNIDAD" and n in (7, 8):          # 8 = con la columna EXTENAS
            segs.setdefault("migr", _escanear(tbl, "migr", [0], n))
        elif h == "UNIDAD" and n == 9:
            # Servicio especial de fiscalización de extranjeros (tabla nueva)
            segs.setdefault("servicio_ext", _escanear(tbl, "servicio_ext", [0], 9))
        elif h == "FECHA" and n == 9:
            segs.setdefault("flujo", _escanear(tbl, "flujo", [0], 9))
        elif h == "ESTADO" and n == 3:
            segs.setdefault("veh_e", _escanear(tbl, "veh_e", [0], 3))
        elif h.startswith("TIPO DE VEHICULO"):
            segs.setdefault("veh_t", _escanear(tbl, "veh_t", [0], 3))
        elif h == "DELITOS SEXUALES" and n == 3:
            segs.setdefault("brisex", _escanear(tbl, "brisex", [0], 3))
        elif h.startswith("BANDAS DESARTICULADAS") and n == 4:
            segs.setdefault("bandas", _escanear(tbl, "bandas", [0], 4))
        elif h == "DELITO" and n == 4:
            segs.setdefault("det_a", _escanear(tbl, "det_a", [0], 4))
        elif h == "PREFECTURA" and n == 4 and r0[3].startswith("SEMANA"):
            segs.setdefault("det_b", _escanear(tbl, "det_b", [0], 4))
        elif h == "PREFECTURA" and n == 4 and "ANO" in r0[3]:
            segs.setdefault("det_c", _escanear(tbl, "det_c", [0], 4))
    return segs, recuadros


def bloque_texto_x(doc):
    """Párrafos (elementos w:p) del texto libre bajo el último encabezado
    'XI. DILIGENCIAS DE RELEVANCIA' (o 'X.' en informes anteriores)."""
    body = doc.element.body
    hijos = list(body)
    idx = None
    for i, el in enumerate(hijos):
        if el.tag == W_P and norm("".join(
                t.text or "" for t in el.iter(qn("w:t")))).startswith(("X. DILIGENCIAS", "XI. DILIGENCIAS")):
            idx = i
    if idx is None:
        return []
    pars = []
    for el in hijos[idx + 1:]:
        if el.tag == W_TBL:
            break
        if el.tag == W_P:
            txt = "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
            if txt:
                pars.append(el)
            elif pars:
                break
    return pars
