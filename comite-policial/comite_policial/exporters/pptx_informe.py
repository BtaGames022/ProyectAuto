"""Informe semanal en PowerPoint con el diseño exacto del comité.

No se dibuja nada desde cero: se abre `plantillas/plantilla_comite.pptx` (el
informe real sin las fotos de la semana) y se reescriben solo los textos y las
celdas, de modo que fondo, fuentes, colores, bordes y posiciones son los del
original. Las láminas de diligencias se clonan de una lámina prototipo, una
por recuadro de la sección X.

La plantilla la prepara `tools/construir_plantilla_pptx.py`, que deja nombres
canónicos: cada tabla es «T:<clave de la app>», cada foto un marco «FOTO:<clave>»
y el prototipo de diligencia «DILIG_CABECERA» / «DILIG_CUERPO». Así el exportador
no depende de cómo se llamen los objetos en el informe original.

Este módulo no depende de la interfaz: recibe el estado como diccionarios.
"""
from __future__ import annotations

import copy
import datetime
import os
import re
import tempfile

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from comite_policial.core.fechas import MESES_ABREV
from comite_policial.core.recuadros import campos_def
from comite_policial.docio.semana import periodo_cerrado, semana_corta
from comite_policial.exporters.pptx_fotos import (
    REGION_FOTOS, celdas_fotos, colocar_fotos, motivo, poner_foto)

ALTO_LAMINA = 6858000
ANCHO_LAMINA = 12192000
PIE_LAMINA = 300000          # franja del pie que ninguna tabla debe invadir

# ── Opciones de cada tabla de la plantilla ───────────────────────────────────
# clave de la app → (filas de cabecera, hay total, hay fuente, texto libre)
# «texto libre»: celdas con listas o frases largas; la fila crece con el texto.
TABLAS = {
    "armas":        (1, True, True, False),
    "drogas_g":     (1, True, True, False),
    "drogas_p":     (1, True, True, False),
    "homi_a":       (1, True, True, False),
    "homi_b":       (2, True, True, False),
    "homi_c":       (1, True, True, False),
    "homi_d":       (1, True, True, False),
    "migr":         (1, True, False, True),
    "servicio_ext": (1, False, False, True),
    "flujo":        (1, True, True, False),
    "bandas":       (1, False, True, True),
    "veh_e":        (1, True, True, False),
    "veh_t":        (1, True, True, False),
    "det_b":        (1, True, True, False),
    "det_c":        (1, True, True, False),
}
# Tablas que deben encajar entre su posición y lo que hay debajo (fotos, otra tabla)
AJUSTAR_ABAJO = {"drogas_p", "flujo", "migr", "servicio_ext", "bandas", "det_b", "det_c"}

# Alturas mínimas (EMU) a las que se puede comprimir cada tipo de fila
PISOS = {"cab": 400000, "dato": 200000, "total": 200000, "fuente": 180000}

DIAS = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO"]


class PlantillaIncompatible(ValueError):
    """La plantilla no trae los nombres que espera el exportador."""


# ═════════════════════════════════════════════════════════════════════════════
# TEXTO
# ═════════════════════════════════════════════════════════════════════════════
_INVALIDOS = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def _limpio(texto) -> str:
    """Texto apto para XML: quita caracteres de control que romperían el archivo
    (por ejemplo los que llegan al pegar desde Word o desde un PDF)."""
    return _INVALIDOS.sub("", str(texto if texto is not None else ""))


def _set_par(p, texto: str):
    """Pone `texto` en un <a:p> conservando el formato de su primer run."""
    texto = _limpio(texto)
    runs = p.findall(qn("a:r"))
    if runs:
        runs[0].find(qn("a:t")).text = texto
        for r in runs[1:]:
            p.remove(r)
        for br in p.findall(qn("a:br")):
            p.remove(br)
        return
    r = etree.Element(qn("a:r"))
    rpr = etree.SubElement(r, qn("a:rPr"))
    fin = p.find(qn("a:endParaRPr"))
    if fin is not None:
        for k, v in fin.attrib.items():
            rpr.set(k, v)
        for hijo in fin:
            rpr.append(copy.deepcopy(hijo))
    etree.SubElement(r, qn("a:t")).text = texto
    if fin is not None:
        fin.addprevious(r)
    else:
        p.append(r)


def _texto_par(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("a:t")))


def _set_celda(tc, texto):
    """Reemplaza el contenido de una celda; cada línea es un párrafo."""
    cuerpo = tc.find(qn("a:txBody"))
    ps = cuerpo.findall(qn("a:p"))

    # molde: el párrafo con texto real (los vacíos/«\xa0» pueden tener otro color)
    def _txt(p):
        return _texto_par(p).replace("\xa0", "").strip()
    base = next((p for p in ps if _txt(p)), None)
    if base is None:
        base = next((p for p in ps if p.find(qn("a:r")) is not None), ps[0])
    lineas = _limpio(texto).split("\n") or [""]
    for p in ps:
        cuerpo.remove(p)
    for ln in lineas:
        p = copy.deepcopy(base)
        _set_par(p, ln)
        cuerpo.append(p)


def _tam_celda(tc) -> float:
    for rpr in tc.iter(qn("a:rPr")):
        if rpr.get("sz"):
            return int(rpr.get("sz")) / 100
    return 10.0


def _lineas(texto: str, pt: float, ancho_emu: int, ancho_car: float = 0.53) -> int:
    """Líneas que ocupa `texto` a `pt` puntos en `ancho_emu` (estimación Arial).

    `ancho_car`: ancho medio de un carácter en «em»: 0.53 para texto corrido;
    ~0.75 para mayúsculas en negrita (títulos)."""
    util = max(ancho_emu / 12700 - 14.4, 20)
    por_linea = max(util / (ancho_car * pt), 4)
    total = 0
    for parrafo in str(texto).split("\n"):
        n, actual = 1, 0
        for palabra in parrafo.split(" "):
            largo = len(palabra) + (1 if actual else 0)
            if actual and actual + largo > por_linea:
                n += 1
                actual = len(palabra)
            else:
                actual += largo
        total += n
    return total


# ═════════════════════════════════════════════════════════════════════════════
# TEXTOS CON FECHAS (títulos, rótulos, encabezados)
# ═════════════════════════════════════════════════════════════════════════════
def titulo_servicio(fechas: str) -> str:
    """'19 Y 20.OCT.026' -> 'LUNES 19 Y MARTES 20.OCT.026' (título del servicio especial)."""
    f = " ".join(str(fechas or "").split()).upper()
    if not f or any(d in f for d in DIAS):
        return f
    m = re.search(r"\.([A-ZÁÉÍÓÚ]{3})\.(\d{3})\s*$", f)
    if not m or m.group(1) not in MESES_ABREV:
        return f
    mes, anio = MESES_ABREV.index(m.group(1)) + 1, 2000 + int(m.group(2))

    def con_dia(mm):
        try:
            d = datetime.date(anio, mes, int(mm.group(0)))
        except ValueError:
            return mm.group(0)
        return f"{DIAS[d.weekday()]} {d.day}"
    return re.sub(r"\d{1,2}", con_dia, f[:m.start()]) + f[m.start():]


class _Contexto:
    """Textos que cambian cada semana, ya compuestos."""

    def __init__(self, estado, ano_ant, ano_act, periodo):
        self.semana = _limpio(estado.get("semana", "")).strip()
        self.corta = semana_corta(self.semana)
        self.fin = self.semana.split(" AL ")[-1] if " AL " in self.semana else self.semana
        self.ano_ant, self.ano_act = ano_ant, ano_act
        per = (periodo or "").strip()
        self.periodo = (periodo_cerrado(self.semana) if not per or per.lower() == "auto" else per)
        self.servicio = titulo_servicio(_limpio(estado.get("servicio_fechas", "")))

    def rotulo_periodo(self, ano):
        return f"{self.periodo} {ano}".strip()

    def cabeceras(self) -> dict:
        """Celdas de cabecera con la semana o el año: {clave: {(fila, col): texto}}."""
        sem, fecha = f"SEMANA {self.corta}", f"{self.ano_act} A LA FECHA"
        return {
            "armas":    {(0, 1): sem, (0, 2): fecha},
            "drogas_p": {(0, 1): sem, (0, 2): fecha},
            "drogas_g": {(0, 1): sem, (0, 2): fecha},
            "homi_a":   {(0, 1): f"SEMANA DEL {self.corta}", (0, 2): f"AÑO {self.ano_act} A LA FECHA"},
            "homi_c":   {(0, 1): f"01.ENE AL {self.fin}", (0, 2): f"% AÑO {self.ano_act}"},
            "homi_b":   {(0, 1): f"SEMANA DEL {self.corta}", (0, 4): f"AÑO {self.ano_act} A LA FECHA"},
            "homi_d":   {(0, 0): "HOMICIDIOS", (0, 1): self.rotulo_periodo(self.ano_ant),
                         (0, 2): self.rotulo_periodo(self.ano_act),
                         (0, 3): f"VAR. % {self.ano_ant} / {self.ano_act}"},
            "veh_e":    {(0, 1): sem, (0, 2): fecha},
            "veh_t":    {(0, 1): sem, (0, 2): fecha},
            "det_b":    {(0, 3): sem},
            "det_c":    {(0, 0): "PREFECTURA", (0, 3): f"AÑO {self.ano_act} A LA FECHA"},
        }

    def reescribir(self, texto: str, anterior: str):
        """(posición, texto nuevo) para un párrafo de título o rótulo, o None.

        El texto nuevo reemplaza al párrafo desde `posición`: así un título como
        «FISCALIZACIONES A EXTRANJEROS   Semana del 19 …» conserva su primera parte."""
        t = texto
        if re.match(r"\s*servicio\s+especial\s+de\s+fiscalizaci", anterior, re.I):
            return 0, self.servicio          # segunda línea del título: día(s) del servicio
        reglas = (
            (r"\s*homicidios\s+del\s+01\.ENE", f"HOMICIDIOS DEL 01.ENE al {self.fin}"),
            (r"\s*modalidad\s+en\s+homicidios\s+semana",
             f"MODALIDAD EN HOMICIDIOS SEMANA DEL {self.corta.replace(' AL ', ' al ')}"),
            (r"\s*comparativo\b",
             f"COMPARATIVO {self.periodo.replace(' - ', '-')} {self.ano_ant}/{self.ano_act}"
             .replace("  ", " ")),
            (r"\s*detenidos\s+\d{4}\s+a\s+la\s+fecha", f"DETENIDOS {self.ano_act} A LA FECHA"),
        )
        for patron, nuevo in reglas:
            if re.match(patron, t, re.I):
                return 0, nuevo
        m = re.search(r"semana\s+del\s+\d", t, re.I)
        return (m.start(), f"Semana del {self.corta}") if m else None


def _reescribir_desde(p, inicio: int, nuevo: str):
    """Reemplaza el texto del párrafo desde la posición `inicio`, conservando el
    formato del run donde empieza y todo lo que quede antes."""
    nuevo = _limpio(nuevo)
    pos = 0
    runs = p.findall(qn("a:r"))
    for r in runs:
        t = r.find(qn("a:t"))
        txt = t.text or ""
        if pos + len(txt) > inicio or r is runs[-1]:
            t.text = txt[:max(inicio - pos, 0)] + nuevo
            for hermano in list(r.itersiblings()):
                if hermano.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
                    p.remove(hermano)
            return
        pos += len(txt)
    if not runs:
        _set_par(p, nuevo)


def _poner_nota(laminas, texto: str):
    """Nota «(*) Droga Sintética…» de la lámina de droga (forma «NOTA:sintetica»).

    Se reescribe con lo que escribió el usuario, un párrafo por línea. Si queda vacía
    se quita la forma, para que no sobreviva la nota de otra semana."""
    lineas = [ln.rstrip() for ln in _limpio(texto).splitlines() if ln.strip()]
    for sl in laminas:
        for sh in list(sl.shapes):
            if not sh.name.startswith("NOTA:"):
                continue
            if not lineas:
                sh._element.getparent().remove(sh._element)
                continue
            cuerpo = sh._element.txBody
            ps = cuerpo.findall(qn("a:p"))
            base = ps[0]
            for p in ps[1:]:
                cuerpo.remove(p)
            _set_par(base, lineas[0])
            ref = base
            for ln in lineas[1:]:
                nuevo = copy.deepcopy(base)
                _set_par(nuevo, ln)
                ref.addnext(nuevo)
                ref = nuevo
            # no debe pisar lo que hay debajo (la segunda tabla de droga): si no cabe, baja la letra
            debajo = [o.top for o in sl.shapes
                      if o._element is not sh._element and o.top is not None and o.width is not None
                      and o.top > sh.top + 100000 and o.left < sh.left + sh.width
                      and o.left + o.width > sh.left]
            limite = min(debajo + [ALTO_LAMINA - PIE_LAMINA]) - 40000
            pt = pt0 = _tam_celda(sh._element)
            while True:
                n = sum(_lineas(ln, pt, sh.width) for ln in lineas)
                alto = int(n * pt * 1.2 * 12700 + 91440)
                if sh.top + alto <= limite or pt <= 7:
                    break
                pt -= 1
            if pt != pt0:
                for el in list(cuerpo.iter(qn("a:rPr"))) + list(cuerpo.iter(qn("a:endParaRPr"))):
                    el.set("sz", str(int(pt * 100)))
            sh.height = max(sh.height, alto)


def _actualizar_textos(laminas, ctx: _Contexto):
    for sl in laminas:
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            ps = sh._element.txBody.findall(qn("a:p"))
            textos = [_texto_par(p) for p in ps]
            for i, p in enumerate(ps):
                cambio = ctx.reescribir(textos[i], textos[0] if i == 1 else "")
                if cambio is not None and (textos[i].strip() or cambio[1]):
                    _reescribir_desde(p, *cambio)


# ═════════════════════════════════════════════════════════════════════════════
# TABLAS
# ═════════════════════════════════════════════════════════════════════════════
def _fila_clon(tr):
    nueva = copy.deepcopy(tr)
    for ext in nueva.findall(qn("a:extLst")):
        nueva.remove(ext)
    return nueva


def _llenar_tabla(shape, datos, total, *, n_cab=1, hay_total=True, hay_fuente=True,
                  cab=None, texto_libre=False, limite=None):
    """Vuelca `datos` (+ `total`) en la tabla conservando el estilo de cada fila.

    Agrega o quita filas de datos clonando la última fila de datos.
    `cab`: {(fila, col): texto} para encabezados que llevan la semana o el año.
    `limite`: y máxima (EMU) que puede alcanzar la tabla; si no cabe se comprime.
    """
    tbl = shape._element.graphic.graphicData.tbl
    trs = tbl.findall(qn("a:tr"))
    ini = n_cab
    fin = len(trs) - (1 if hay_total else 0) - (1 if hay_fuente else 0)
    n_plantilla, n_datos = fin - ini, len(datos)

    if n_datos > n_plantilla and n_plantilla > 0:
        ref = trs[fin - 1]
        for _ in range(n_datos - n_plantilla):
            nueva = _fila_clon(ref)
            ref.addnext(nueva)
            ref = nueva
    elif n_datos < n_plantilla:
        for tr in trs[ini + n_datos:fin]:
            tbl.remove(tr)
    trs = tbl.findall(qn("a:tr"))

    ncols = len(tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol")))

    def escribir(tr, valores, es_total=False):
        for ci, tc in enumerate(tr.findall(qn("a:tc"))):
            if tc.get("hMerge") or tc.get("vMerge") or ci >= ncols:
                continue
            if es_total and ci == 0:
                continue                  # el rótulo «TOTAL …» es el de la plantilla
            if ci < len(valores):
                _set_celda(tc, valores[ci])
                pr = tc.find(qn("a:tcPr"))
                if texto_libre and pr is not None:
                    pr.set("anchor", "ctr")

    for i, fila in enumerate(datos):
        escribir(trs[ini + i], list(fila))
    if hay_total and total is not None:
        escribir(trs[ini + n_datos], list(total), es_total=True)
    for (r, c), txt in (cab or {}).items():
        _set_celda(trs[r].findall(qn("a:tc"))[c], txt)

    presupuesto = (limite - shape.top) if limite else None
    _ajustar_altos(shape, tbl, n_cab, hay_total, hay_fuente, texto_libre, presupuesto)


def _poner_cabeceras(forma, cab):
    trs = forma._element.graphic.graphicData.tbl.findall(qn("a:tr"))
    for (r, c), txt in (cab or {}).items():
        _set_celda(trs[r].findall(qn("a:tc"))[c], txt)


def _roles(n_filas, n_cab, hay_total, hay_fuente):
    n_datos = n_filas - n_cab - (1 if hay_total else 0) - (1 if hay_fuente else 0)
    return (["cab"] * n_cab + ["dato"] * n_datos + (["total"] if hay_total else [])
            + (["fuente"] if hay_fuente else []))


def _alto_necesario(tr, anchos) -> int:
    """Alto (EMU) que pide el texto de la fila más larga de sus celdas."""
    necesario = 0
    for ci, tc in enumerate(tr.findall(qn("a:tc"))):
        if tc.get("hMerge") or tc.get("vMerge"):
            continue
        span = int(tc.get("gridSpan", "1"))
        ancho = sum(anchos[ci:ci + span])
        pt = _tam_celda(tc)
        txt = "\n".join(_texto_par(p) for p in tc.iter(qn("a:p")))
        necesario = max(necesario, _lineas(txt, pt, ancho) * pt * 1.2 * 12700 + 60000)
    return int(necesario)


def _reducir_letra(trs, roles, anchos, minimo=700) -> bool:
    """Baja 1 pt la letra de las celdas de varias líneas (listas, nacionalidades) de
    las filas de datos y del total; las cifras no se tocan. False si ya no se puede."""
    cambio = False
    for tr, rol in zip(trs, roles):
        if rol not in ("dato", "total"):
            continue
        for ci, tc in enumerate(tr.findall(qn("a:tc"))):
            if tc.get("hMerge") or tc.get("vMerge"):
                continue
            ancho = sum(anchos[ci:ci + int(tc.get("gridSpan", "1"))])
            txt = "\n".join(_texto_par(p) for p in tc.iter(qn("a:p")))
            if _lineas(txt, _tam_celda(tc), ancho) < 2:
                continue
            for el in list(tc.iter(qn("a:rPr"))) + list(tc.iter(qn("a:endParaRPr"))):
                sz = el.get("sz")
                if sz and int(sz) > minimo:
                    el.set("sz", str(max(int(sz) - 100, minimo)))
                    cambio = True
    return cambio


def _comprimir(altos, roles, presupuesto):
    """Reparte el exceso sobre lo que cada fila puede ceder por encima de su piso."""
    exceso = sum(altos) - presupuesto
    if exceso <= 0:
        return altos
    holgura = [max(h - PISOS[r], 0) for h, r in zip(altos, roles)]
    if sum(holgura) <= 0:
        return altos
    f = min(exceso / sum(holgura), 1.0)
    return [h - int(hol * f) for h, hol in zip(altos, holgura)]


def _ajustar_altos(shape, tbl, n_cab, hay_total, hay_fuente, texto_libre, presupuesto):
    """Altura de filas: crece con el texto libre y se comprime al presupuesto."""
    trs = tbl.findall(qn("a:tr"))
    anchos = [int(g.get("w")) for g in tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol"))]
    roles = _roles(len(trs), n_cab, hay_total, hay_fuente)
    base = [int(tr.get("h")) for tr in trs]
    altos = list(base)

    if texto_libre:
        datos = [h for h, r in zip(base, roles) if r == "dato"]
        piso = min(datos) if datos else PISOS["dato"]

        def calcular(compacto):
            out = []
            for tr, rol, h0 in zip(trs, roles, base):
                if rol == "cab":
                    out.append(h0)
                    continue
                piso_fila = PISOS[rol] if compacto else (piso if rol == "dato" else h0)
                out.append(max(_alto_necesario(tr, anchos), piso_fila))
            return out

        altos = calcular(False)
        if presupuesto and sum(altos) > presupuesto:
            justo = calcular(True)                     # filas justas al texto
            while sum(justo) > presupuesto and _reducir_letra(trs, roles, anchos):
                justo = calcular(True)                 # y, si aún no cabe, letra más chica
            pleno = calcular(False)
            if sum(pleno) <= presupuesto:
                altos = pleno
            else:                                      # la holgura que sobra se reparte entre las filas
                f = 0.0
                if sum(pleno) > sum(justo):
                    f = max(0.0, min(1.0, (presupuesto - sum(justo)) / (sum(pleno) - sum(justo))))
                altos = [int(j + (p - j) * f) for j, p in zip(justo, pleno)]
    elif presupuesto:
        altos = _comprimir(altos, roles, presupuesto)

    for tr, h in zip(trs, altos):
        tr.set("h", str(int(h)))
    shape.height = sum(altos)


def _combinar_variacion(shape, n_datos):
    """III-D: la columna VAR. % va combinada en las filas de datos."""
    trs = shape._element.graphic.graphicData.tbl.findall(qn("a:tr"))
    for i, tr in enumerate(trs[1:1 + n_datos]):
        tc = tr.findall(qn("a:tc"))[3]
        for atr in ("rowSpan", "vMerge"):
            tc.attrib.pop(atr, None)
        if i == 0 and n_datos > 1:
            tc.set("rowSpan", str(n_datos))
        elif i > 0:
            tc.set("vMerge", "1")


def _indexar_tablas(prs) -> dict:
    """{clave: (lámina, forma)} de las tablas «T:<clave>» de la plantilla."""
    idx = {}
    for sl in prs.slides:
        for sh in sl.shapes:
            if getattr(sh, "has_table", False) and sh.has_table and sh.name.startswith("T:"):
                idx[sh.name[2:]] = (sl, sh)
    return idx


def _limite_inferior(sl, forma) -> int:
    """Y máxima de una tabla: lo primero que hay debajo (foto, texto, otra tabla)
    que se cruce con ella en horizontal, o el pie de la lámina."""
    izq, der = forma.left, forma.left + forma.width
    limite = ALTO_LAMINA - PIE_LAMINA
    for sh in sl.shapes:
        if sh._element is forma._element or sh.left is None or sh.width is None:
            continue
        if sh.top > forma.top + forma.height // 2 and sh.left < der and sh.left + sh.width > izq:
            limite = min(limite, sh.top - 60000)
    return limite


# ═════════════════════════════════════════════════════════════════════════════
# LÁMINAS
# ═════════════════════════════════════════════════════════════════════════════
def _clonar_lamina(prs, origen):
    """Duplica una lámina (fondo incluido) al final de la presentación."""
    nueva = prs.slides.add_slide(origen.slide_layout)
    arbol = nueva.shapes._spTree
    for el in list(arbol)[2:]:                 # fuera marcadores del layout
        arbol.remove(el)
    mapa = {}
    for rid, rel in origen.part.rels.items():
        if rel.reltype == RT.IMAGE:
            mapa[rid] = nueva.part.relate_to(rel.target_part, rel.reltype)
    bg = origen._element.cSld.find(qn("p:bg"))
    if bg is not None:
        nbg = copy.deepcopy(bg)
        for el in nbg.iter():
            for k in list(el.attrib):
                if k.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"):
                    el.set(k, mapa[el.get(k)])
        nueva._element.cSld.insert(0, nbg)
    for el in list(origen.shapes._spTree)[2:]:
        arbol.append(copy.deepcopy(el))
    return nueva


def _quitar_lamina(prs, lamina):
    ids = prs.slides._sldIdLst
    for sld in list(ids):
        if prs.slides.get(int(sld.get("id"))) is lamina:
            prs.part.drop_rel(sld.rId)
            ids.remove(sld)
            return


def _mover_al_final(prs, lamina):
    ids = prs.slides._sldIdLst
    for sld in list(ids):
        if prs.slides.get(int(sld.get("id"))) is lamina:
            ids.remove(sld)
            ids.append(sld)
            return


# ── Diligencias ──────────────────────────────────────────────────────────────
X_TXT = 467836           # borde izquierdo del texto
Y_ENC = 862781           # y de la cabecera (unidad y título)


def _parrafos_diligencia(rec) -> list:
    """Campos del recuadro como párrafos [(texto, negrita)], con una línea en
    blanco entre campos (así están las láminas originales)."""
    campos = rec.get("campos", {})
    out = []
    for etiqueta, clase in campos_def(rec):
        valor = str(campos.get(etiqueta, "")).strip()
        if not valor:
            continue
        if out:
            out.append([])
        lineas = [ln.rstrip() for ln in valor.splitlines() if ln.strip()]
        rot = re.sub(r"\s+\(\d+\)$", "", etiqueta)
        if clase == "lista":
            out.append([(f"{rot}:", True)])
            out.extend([(ln, False)] for ln in lineas)
        else:
            out.append([(f"{rot}: ", True), (lineas[0], False)])
            out.extend([(ln, False)] for ln in lineas[1:])
    return out


def _run(texto, negrita, pt, color="260000"):
    r = etree.Element(qn("a:r"))
    rpr = etree.SubElement(r, qn("a:rPr"), lang="es-ES", sz=str(int(pt * 100)),
                           b="1" if negrita else "0", dirty="0")
    relleno = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(relleno, qn("a:srgbClr"), val=color)
    etree.SubElement(rpr, qn("a:latin"), typeface="Arial")
    etree.SubElement(rpr, qn("a:cs"), typeface="Arial")
    etree.SubElement(r, qn("a:t")).text = _limpio(texto)
    return r


def _ppr(algn):
    ppr = etree.Element(qn("a:pPr"), algn=algn)
    etree.SubElement(ppr, qn("a:buNone"))
    return ppr


def _reconstruir_parrafos(cuerpo, algn, parrafos, pt, color):
    """parrafos: lista de listas [(texto, negrita)]; [] = línea en blanco."""
    for p in cuerpo.findall(qn("a:p")):
        cuerpo.remove(p)
    for runs in parrafos:
        p = etree.SubElement(cuerpo, qn("a:p"))
        p.append(_ppr(algn))
        for texto, negrita in runs:
            p.append(_run(texto, negrita, pt, color))
        fin = etree.SubElement(p, qn("a:endParaRPr"), lang="es-ES",
                               sz=str(int(pt * 100)), dirty="0")
        if not runs:
            etree.SubElement(etree.SubElement(fin, qn("a:solidFill")),
                             qn("a:srgbClr"), val=color)


def _alto_parrafos(parrafos, pt, ancho_emu, ancho_car=0.53):
    n = 0
    for runs in parrafos:
        n += _lineas("".join(t for t, _ in runs), pt, ancho_emu, ancho_car) if runs else 1
    return int(n * pt * 1.2 * 12700 + 91440)


def _cuerpo_unico(forma):
    tc = forma._element.graphic.graphicData.tbl.find(qn("a:tr")).find(qn("a:tc"))
    return tc.find(qn("a:txBody"))


def _color_texto(forma, defecto="260000") -> str:
    clr = _cuerpo_unico(forma).xpath(".//a:rPr/a:solidFill/a:srgbClr/@val")
    return clr[0] if clr else defecto


def _fijar_tabla_unica(forma, x, y, ancho, alto):
    tbl = forma._element.graphic.graphicData.tbl
    tbl.find(qn("a:tblGrid")).find(qn("a:gridCol")).set("w", str(int(ancho)))
    tbl.find(qn("a:tr")).set("h", str(int(alto)))
    forma.left, forma.top, forma.width, forma.height = int(x), int(y), int(ancho), int(alto)


def _lamina_diligencia(prs, proto, rec, avisos):
    sl = _clonar_lamina(prs, proto)
    enc = next(sh for sh in sl.shapes if sh.name == "DILIG_CABECERA")
    cue = next(sh for sh in sl.shapes if sh.name == "DILIG_CUERPO")
    color_enc, color_cue = _color_texto(enc), _color_texto(cue)
    fotos = [f for f in rec.get("fotos", []) if f]
    # con fotos el texto va en la columna izquierda; sin fotos puede ensancharse
    # si el relato es largo (se prefiere letra de 11 pt o más)
    anchos = ([REGION_FOTOS[0] - 250000 - X_TXT] if fotos
              else [8200000, 9800000, 11200000])

    unidad = rec.get("unidad", "").strip().upper()
    if rec.get("grupo", "").strip():
        unidad = f"{unidad} {rec['grupo'].strip().upper()}".strip()
    titulo = [ln.strip().upper() for ln in rec.get("titulo", "").splitlines() if ln.strip()]
    pars_enc = ([[(unidad, True)]] if unidad else []) + [[(t, True)] for t in titulo]
    pars_enc = pars_enc or [[(" ", True)]]
    pars = _parrafos_diligencia(rec) or [[("Sin datos.", False)]]

    for ancho in anchos:
        for pt_enc in (20, 18, 16, 14, 12):
            alto_enc = _alto_parrafos(pars_enc, pt_enc, ancho, 0.75)
            if alto_enc <= 1450000:
                break
        y_cuerpo = Y_ENC + alto_enc + 100000
        libre = ALTO_LAMINA - PIE_LAMINA - y_cuerpo
        for pt_cue in (14, 13, 12, 11, 10, 9, 8):
            alto = _alto_parrafos(pars, pt_cue, ancho)
            if alto <= libre:
                break
        if alto <= libre and pt_cue >= 11:
            break
    if alto > libre:
        avisos.append(f"Diligencia «{unidad[:40]}»: el texto no cabe en una lámina; "
                      "conviene acortarlo.")

    _reconstruir_parrafos(_cuerpo_unico(enc), "l", pars_enc, pt_enc, color_enc)
    _fijar_tabla_unica(enc, X_TXT, Y_ENC, ancho, alto_enc)
    _reconstruir_parrafos(_cuerpo_unico(cue), "just", pars, pt_cue, color_cue)
    _fijar_tabla_unica(cue, X_TXT, y_cuerpo, ancho, min(alto, libre))

    for (x, y, w, h), ruta in zip(celdas_fotos(len(fotos)), fotos):
        try:
            poner_foto(sl, ruta, x, y, w, h, modo="ajustar")
        except Exception as exc:
            avisos.append(f"No se pudo insertar la foto «{os.path.basename(ruta)}» "
                          f"de la diligencia «{unidad[:40]}»: {motivo(exc)}")
    return sl


# ═════════════════════════════════════════════════════════════════════════════
# API
# ═════════════════════════════════════════════════════════════════════════════
def exportar_informe(estado: dict, ruta_salida: str, plantilla: str,
                     fotos: dict | None = None, ano_ant="2025", ano_act="2026",
                     periodo="auto") -> list:
    """Genera el PPTX. Devuelve la lista de avisos (textos que no caben, etc.).

    estado = {"semana": "31.AGO AL 06.SEP.026",
              "servicio_fechas": "19 Y 20.OCT.026",
              "nota_sintetica": "(*) Droga Sintética: Ketamina – Éxtasis",
              "tablas": {clave: {"rows": [[...]], "total": [...] | None}},
              "recuadros": [dict, ...]}
    fotos  = {clave de un marco de la plantilla: ruta de la imagen}
    periodo = rótulo del comparativo («ENE - SEP»); «auto» lo deduce de la semana.
    """
    avisos: list = []
    fotos = fotos or {}
    ctx = _Contexto(estado, ano_ant, ano_act, periodo)
    tablas = estado.get("tablas", {})

    prs = Presentation(plantilla)
    indice = _indexar_tablas(prs)
    proto = next((sl for sl in prs.slides
                  if any(sh.name == "DILIG_CABECERA" for sh in sl.shapes)), None)
    if not indice or proto is None:
        raise PlantillaIncompatible(
            "La plantilla es de un formato anterior (no trae las tablas y marcos "
            "nombrados). Regenerarla con tools/construir_plantilla_pptx.py.")
    cierre = prs.slides[-1]
    if not estado.get("servicio_en_pptx", True) and "servicio_ext" in indice:
        sl_serv, _ = indice.pop("servicio_ext")
        _quitar_lamina(prs, sl_serv)
    limites = {k: _limite_inferior(sl, forma) for k, (sl, forma) in indice.items()}

    fijas = [sl for sl in prs.slides if sl is not proto and sl is not cierre]
    _actualizar_textos(fijas, ctx)
    if "nota_sintetica" in estado:               # sin la clave, queda el texto de la plantilla
        _poner_nota(fijas, estado["nota_sintetica"])

    cabeceras = ctx.cabeceras()
    for clave, (n_cab, hay_total, hay_fuente, libre) in TABLAS.items():
        if clave not in indice:
            continue
        forma = indice[clave][1]
        datos_t = tablas.get(clave)
        if not datos_t:                           # sin datos: igual se actualizan los encabezados
            _poner_cabeceras(forma, cabeceras.get(clave))
            continue
        filas = [list(f) for f in datos_t["rows"]]
        total = datos_t.get("total")
        if clave == "homi_d":                    # la variación va solo en el total (con coma)
            filas = [[f[0], f[1], f[2], ""] for f in filas]
            if total:
                total = list(total[:-1]) + [str(total[-1]).replace(".", ",")]
        _llenar_tabla(forma, filas, total, n_cab=n_cab, hay_total=hay_total,
                      hay_fuente=hay_fuente, cab=cabeceras.get(clave), texto_libre=libre,
                      limite=limites[clave] if clave in AJUSTAR_ABAJO else None)
        if clave == "homi_d":
            _combinar_variacion(forma, len(filas))

    # Avisos por desbordes de tablas
    for clave, (sl, forma) in indice.items():
        if forma.top + forma.height > ALTO_LAMINA - 100000:
            avisos.append(f"La tabla «{clave}» se sale de la lámina; reduce el texto de sus celdas.")

    colocar_fotos(prs, fotos, avisos)

    # Láminas de diligencias (una por recuadro)
    for rec in estado.get("recuadros", []):
        _lamina_diligencia(prs, proto, rec, avisos)
    _quitar_lamina(prs, proto)
    _mover_al_final(prs, cierre)

    prs.core_properties.title = f"Informe Comité Policial — {ctx.semana}"
    _guardar_seguro(prs, ruta_salida)
    return avisos


def _guardar_seguro(prs, destino: str):
    """Guarda en un archivo temporal, comprueba que se vuelve a abrir y recién
    entonces reemplaza el destino: un fallo nunca deja un PPTX a medias ni pisa
    uno bueno."""
    carpeta = os.path.dirname(os.path.abspath(destino))
    fd, tmp = tempfile.mkstemp(suffix=".pptx", dir=carpeta)
    os.close(fd)
    try:
        prs.save(tmp)
        n = len(Presentation(tmp).slides)
        if n < 2:
            raise ValueError("el archivo generado quedó incompleto")
        os.replace(tmp, destino)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
