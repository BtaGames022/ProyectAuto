"""Genera plantillas/plantilla_comite.pptx a partir de un informe PPT real.

Uso:
    python tools/construir_plantilla_pptx.py "PPT COMITE POLICIAL 19 AL 25.OCT.026.pptx"
    python tools/construir_plantilla_pptx.py informe.pptx --salida otra.pptx --prototipo 13

Conserva el diseño exacto del informe (fondos, tablas, textos, logos y las fotos
institucionales de la portada) y prepara lo que el exportador necesita:

  * Tablas: cada tabla se reconoce por su cabecera (mismo criterio que
    docio/mapa.py para el Word) y se renombra «T:<clave>», de modo que el
    exportador no depende de nombres como «Tabla 6» ni de números de lámina.
  * Tablas pegadas como imagen (III-D, IV Migraciones y IX Servicio especial)
    se rehacen como tablas reales, en el mismo lugar y con el mismo aspecto.
  * Fotos: cada foto de las láminas fijas se reemplaza por un marco invisible
    «FOTO:<clave>» en su misma posición y capa. La app lee esos marcos para el
    botón «Fotos»; no hay coordenadas escritas en el código.
  * Diligencias: se conserva una sola lámina prototipo (cabecera «DILIG_CABECERA»
    y cuerpo «DILIG_CUERPO») y la lámina de cierre; el resto se descarta.

Ejecutarlo solo cuando cambie el diseño del informe; la plantilla se versiona.
"""
import argparse
import copy
import sys
import unicodedata
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "plantillas" / "plantilla_comite.pptx"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

NAVY, CELESTE, BLANCO, NEGRO = "002060", "8DB3E2", "FFFFFF", "000000"

# Tablas del informe que vienen pegadas como imagen: lámina (1-based) -> nombre del
# objeto. Si el comité cambia el formato hay que revisar este mapa (el script avisa).
TABLAS_IMAGEN = {4: ("Imagen 14", "homi_d"), 5: ("Imagen 2", "migr"),
                 6: ("Imagen 5", "servicio_ext")}

# Sección que corresponde a las tablas de cada lámina: (prefijo de las fotos, rótulo)
SECCIONES = [
    ({"armas"}, "armas", "I. Armas"),
    ({"drogas_g", "drogas_p"}, "droga", "II. Droga"),
    ({"homi_a", "homi_b", "homi_c", "homi_d"}, "homi", "III. Homicidios"),
    ({"migr"}, "migr", "IV. Migraciones"),
    ({"servicio_ext"}, "servicio", "IX. Servicio especial"),
    ({"flujo"}, "flujo", "IV-B. Flujo migratorio"),
    ({"bandas", "veh_e", "veh_t"}, "vehban", "V y VII. Vehículos y bandas"),
    ({"det_b", "det_c"}, "det", "VIII. Detenidos"),
]


def norm(s) -> str:
    s = unicodedata.normalize("NFD", str(s or "").replace("\xa0", " "))
    return " ".join("".join(c for c in s if unicodedata.category(c) != "Mn").upper().split())


# ═════════════════════════════════════════════════════════════════════════════
# Reconocimiento de tablas
# ═════════════════════════════════════════════════════════════════════════════
def clave_tabla(sh):
    """Clave de la app para una tabla de la lámina (None si no se reconoce)."""
    t = sh.table
    n = len(t.columns)
    r0 = [norm(c.text) for c in t.rows[0].cells]
    h = r0[0]
    seg = norm(t.cell(1, 0).text) if len(t.rows) > 1 else ""
    if h == "TIPO" and n == 3:
        return "armas"
    if h == "DROGA" and n == 3:
        return "drogas_p" if ("PLANTA" in seg or "(UNIDAD)" in seg) else "drogas_g"
    if h == "HOMICIDIOS" and n == 3 and r0[1].startswith("SEMANA"):
        return "homi_a"
    if h == "HOMICIDIOS" and n == 3 and "01.ENE" in r0[1]:
        return "homi_c"
    if h == "MODALIDAD":
        return "homi_b"
    if h == "FECHA" and n == 9:
        return "flujo"
    if h == "ESTADO" and n == 3:
        return "veh_e"
    if h.startswith("TIPO DE VEHICULO"):
        return "veh_t"
    if h.startswith("BANDAS DESARTICULADAS") and n == 4:
        return "bandas"
    if h.startswith("PREFECT") and n == 4 and r0[3].startswith("SEMANA"):
        return "det_b"
    if h.startswith("PREFECT") and n == 4 and "ANO" in r0[3]:
        return "det_c"
    return None


# ═════════════════════════════════════════════════════════════════════════════
# Construcción de tablas nativas
# ═════════════════════════════════════════════════════════════════════════════
def _borde(tag, ancho=12700):
    return (f'<a:{tag} w="{ancho}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>'
            f'<a:srgbClr val="{NEGRO}"/></a:solidFill><a:prstDash val="solid"/><a:round/>'
            f'<a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>'
            f'</a:{tag}>')


def _tc(texto, sz, negrita, color, fondo, algn="ctr"):
    """Celda con el estilo de las tablas del comité: Arial, bordes negros finos."""
    rpr = (f'lang="es-CL" sz="{int(sz * 100)}" b="{1 if negrita else 0}" dirty="0"')
    fuente = (f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
              f'<a:latin typeface="Arial"/><a:cs typeface="Arial"/>')
    ps = "".join(
        f'<a:p><a:pPr algn="{algn}"><a:buNone/></a:pPr>'
        f'<a:r><a:rPr {rpr}>{fuente}</a:rPr><a:t>{escape(ln)}</a:t></a:r>'
        f'<a:endParaRPr {rpr}>{fuente}</a:endParaRPr></a:p>'
        for ln in (texto.split("\n") or [""]))
    relleno = f'<a:solidFill><a:srgbClr val="{fondo}"/></a:solidFill>' if fondo else "<a:noFill/>"
    return etree.fromstring(
        f'<a:tc xmlns:a="{NS_A}"><a:txBody><a:bodyPr/><a:lstStyle/>{ps}</a:txBody>'
        f'<a:tcPr marL="44450" marR="44450" marT="0" marB="0" anchor="ctr">'
        f'{_borde("lnL")}{_borde("lnR")}{_borde("lnT")}{_borde("lnB")}{relleno}</a:tcPr></a:tc>')


def _tabla_nueva(slide, nombre, x, y, anchos, filas, altos):
    """Inserta una tabla. `filas`: lista de filas; cada celda es (texto, sz, negrita,
    color, fondo, algn)."""
    gf = slide.shapes.add_table(len(filas), len(anchos), Emu(x), Emu(y),
                                Emu(sum(anchos)), Emu(sum(altos)))
    gf.name = nombre
    tbl = gf._element.graphic.graphicData.tbl
    for tr in tbl.findall(qn("a:tr")):
        tbl.remove(tr)
    tp = tbl.find(qn("a:tblPr"))                  # sin estilo: todo el formato es explícito
    for k in list(tp.attrib):
        del tp.attrib[k]
    for hijo in list(tp):
        tp.remove(hijo)
    for g, w in zip(tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol")), anchos):
        g.set("w", str(int(w)))
    for celdas, h in zip(filas, altos):
        tr = etree.SubElement(tbl, qn("a:tr"), h=str(int(h)))
        for c in celdas:
            tr.append(_tc(*c))
    return gf


def _repartir(total, pesos):
    """Reparte `total` EMU proporcionalmente a `pesos`, sumando exactamente `total`."""
    s = sum(pesos)
    out = [int(total * p / s) for p in pesos]
    out[-1] += total - sum(out)
    return out


def tabla_migraciones(slide, x, y, ancho):
    """IV. Migraciones: 8 columnas (con EXTENAS), unidad en celeste, TOTAL en azul."""
    cab = ["UNIDAD", "FISCALIZACIONES", "FISCALIZADOS O CONSULTADOS", "DENUNCIADOS",
           "NOTIFICACIÓN DE EXPULSIÓN", "DETENIDOS", "EXTENAS", "NACIONALIDAD DENUNCIADOS"]
    anchos = _repartir(ancho, [1217, 1290, 1290, 1160, 1240, 1000, 900, 1265])
    filas = [[(c, 9, True, BLANCO, NAVY) for c in cab]]
    for u in ("POLINT VPO", "POLINT VMA", "POLINT LAD", "POLINT SAN", "SEPI QTO", "SEPI IPA"):
        filas.append([(u, 12, True, NEGRO, CELESTE)]
                     + [("00", 12, True, NEGRO, BLANCO)] * 6
                     + [("", 9, True, NEGRO, BLANCO, "l")])
    filas.append([("TOTAL", 12, True, BLANCO, NAVY)]
                 + [("00", 12, True, BLANCO, NAVY)] * 7)
    altos = [520000] + [420000] * 6 + [340000]
    return _tabla_nueva(slide, "T:migr", x, y, anchos, filas, altos)


def tabla_servicio(slide, x, y, ancho):
    """IX. Servicio especial de fiscalización de extranjeros: 9 columnas, una fila."""
    cab = ["UNIDAD", "FISCALIZACIÓN", "FISCALIZADOS O CONSULTADOS", "DENUNCIADOS",
           "NOTIFICACIÓN DE EXPULSIÓN", "DETENIDOS", "EXTENA", "TOTAL DENUNCIADOS",
           "NACIONALIDAD DENUNCIADOS"]
    anchos = _repartir(ancho, [213, 219, 217, 209, 211, 210, 210, 210, 291])
    filas = [[(c, 9, True, BLANCO, NAVY) for c in cab],
             [("DEMIG VPO- DEMIG VMA- SEPISAN- DEMIGLAD- SEPIQTO- SEPIIPA", 11, True, NEGRO, BLANCO)]
             + [("0", 15, True, NEGRO, BLANCO)] * 7
             + [("", 10, True, NEGRO, BLANCO, "l")]]
    return _tabla_nueva(slide, "T:servicio_ext", x, y, anchos, filas, [700000, 1000000])


def tabla_comparativa(slide, modelo, x, y, ancho):
    """III-D: copia de la tabla III-C con una columna más (VAR. %) y esa columna
    combinada en las filas de datos; solo el TOTAL lleva la variación."""
    el = copy.deepcopy(modelo._element)
    for ext in list(el.iter(qn("a:extLst"))):         # ids únicos: PowerPoint los regenera
        ext.getparent().remove(ext)
    modelo._element.addnext(el)
    ids = [int(i) for i in slide._element.xpath("//p:cNvPr/@id")]
    cnv = el.xpath(".//p:cNvPr")[0]
    cnv.set("id", str(max(ids) + 1))
    cnv.set("name", "T:homi_d")
    tbl = el.graphic.graphicData.tbl

    cols = tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol"))
    cols[1].addnext(copy.deepcopy(cols[1]))
    trs = tbl.findall(qn("a:tr"))
    for tr in trs:
        tcs = tr.findall(qn("a:tc"))
        tcs[1].addnext(copy.deepcopy(tcs[1]))
    for g, w in zip(tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol")),
                    _repartir(ancho, [1200, 840, 840, 883])):
        g.set("w", str(w))

    ultima = trs[-1].findall(qn("a:tc"))                    # fila FUENTE: ocupa las 4 columnas
    ultima[0].set("gridSpan", "4")
    for tc in ultima[1:]:
        tc.set("hMerge", "1")
    datos = trs[1:-2]                                       # entre la cabecera y el TOTAL
    if len(datos) > 1:
        datos[0].findall(qn("a:tc"))[3].set("rowSpan", str(len(datos)))
        for tr in datos[1:]:
            tr.findall(qn("a:tc"))[3].set("vMerge", "1")
    var_total = trs[-2].findall(qn("a:tc"))[3]              # celda blanca con la variación
    pr = var_total.find(qn("a:tcPr"))
    for relleno in [c for c in pr if c.tag in (qn("a:noFill"), qn("a:solidFill"), qn("a:gradFill"),
                                              qn("a:blipFill"), qn("a:pattFill"), qn("a:grpFill"))]:
        pr.remove(relleno)
    etree.SubElement(etree.SubElement(pr, qn("a:solidFill")), qn("a:srgbClr"), val=BLANCO)
    for rpr in var_total.iter(qn("a:rPr")):
        for relleno in rpr.findall(qn("a:solidFill")):
            rpr.remove(relleno)
        negro = etree.Element(qn("a:solidFill"))
        etree.SubElement(negro, qn("a:srgbClr"), val=NEGRO)
        rpr.insert(0, negro)

    # FUENTE a 8 pt y una línea, como en la imagen original (a 10 pt se parte en dos y
    # la tabla invade el rótulo de la tabla de modalidad que va debajo)
    trs[-1].set("h", "200000")
    for rpr in list(trs[-1].iter(qn("a:rPr"))) + list(trs[-1].iter(qn("a:endParaRPr"))):
        rpr.set("sz", "800")

    forma = next(s for s in slide.shapes if s._element is el)
    forma.left, forma.top, forma.width = Emu(x), Emu(y), Emu(ancho)
    forma.height = Emu(sum(int(tr.get("h")) for tr in trs))
    return forma


# ═════════════════════════════════════════════════════════════════════════════
# Marcos de foto
# ═════════════════════════════════════════════════════════════════════════════
def soltar_imagen(sl, rid):
    """Quita la relación de una imagen ya retirada de la lámina (si nada más la usa),
    para que no quede empaquetada en el archivo."""
    if rid in sl.part.rels and not sl._element.xpath(f'//*[@r:embed="{rid}"]'):
        sl.part.drop_rel(rid)


def _marco_xml(id_, clave, descr, x, y, w, h):
    descr = escape(descr, {'"': "&quot;"})
    return etree.fromstring(
        f'<p:sp xmlns:p="{NS_P}" xmlns:a="{NS_A}"><p:nvSpPr>'
        f'<p:cNvPr id="{id_}" name="FOTO:{clave}" descr="{descr}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>'
        f'<a:p><a:endParaRPr lang="es-CL"/></a:p></p:txBody></p:sp>')


def _orden_lectura(figuras, tolerancia=250000):
    """Ordena las fotos como se leen: filas de arriba abajo, cada fila de izquierda a
    derecha (las que difieren poco en altura cuentan como la misma fila)."""
    figuras = sorted(figuras, key=lambda f: f.top)
    filas, actual = [], []
    for f in figuras:
        if actual and f.top - actual[0].top > tolerancia:
            filas.append(actual)
            actual = []
        actual.append(f)
    if actual:
        filas.append(actual)
    return [f for fila in filas for f in sorted(fila, key=lambda f: f.left)]


def marcos_de_foto(slide, prefijo, seccion, informe):
    """Reemplaza las fotos de la lámina por marcos invisibles; devuelve cuántos."""
    fotos = [s for s in slide.shapes if s.shape_type == 13]
    fotos = _orden_lectura(fotos)
    ids = [int(i) for i in slide._element.xpath("//p:cNvPr/@id")]
    sig = max(ids) + 1
    arbol = slide.shapes._spTree
    for i, pic in enumerate(fotos, 1):
        clave = f"{prefijo}_{i}"
        el = pic._element
        marco = _marco_xml(sig, clave, f"{seccion}|{i}|{len(fotos)}",
                           pic.left, pic.top, pic.width, pic.height)
        sig += 1
        el.addprevious(marco)                 # misma capa que la foto original
        arbol.remove(el)
        soltar_imagen(slide, el.blip_rId)
        informe.append(f"   {clave:10} {pic.width / 914400:4.1f} x {pic.height / 914400:4.1f} in")
    return len(fotos)


# ═════════════════════════════════════════════════════════════════════════════
# Láminas
# ═════════════════════════════════════════════════════════════════════════════
def quitar_lamina(prs, lamina):
    ids = prs.slides._sldIdLst
    for sld in list(ids):
        if prs.slides.get(int(sld.get("id"))) is lamina:
            prs.part.drop_rel(sld.rId)
            ids.remove(sld)
            return


def _es_diligencia(sl):
    return any(sh.has_text_frame and norm(sh.text_frame.text).startswith("DILIGENCIAS RELEVANTES")
               for sh in sl.shapes)


def _tablas_1x1(sl):
    return sorted((s for s in sl.shapes if getattr(s, "has_table", False) and s.has_table
                   and len(s.table.rows) == 1 and len(s.table.columns) == 1),
                  key=lambda s: s.top)


def elegir_prototipo(slides, numero):
    """Lámina prototipo de diligencia: la indicada, o la de cuerpo de texto más alto."""
    cand = [(n, sl) for n, sl in enumerate(slides, 1)
            if _es_diligencia(sl) and len(_tablas_1x1(sl)) == 2]
    assert cand, "no hay láminas de diligencias con cabecera y cuerpo (dos tablas 1x1)"
    if numero:
        for n, sl in cand:
            if n == numero:
                return n, sl
        raise SystemExit(f"la lámina {numero} no es una diligencia con cabecera y cuerpo")
    return max(cand, key=lambda c: _tablas_1x1(c[1])[1].height)


def preparar_prototipo(sl):
    cab, cuerpo = _tablas_1x1(sl)
    cab.name, cuerpo.name = "DILIG_CABECERA", "DILIG_CUERPO"
    for sh in list(sl.shapes):                         # las fotos de la semana no van
        if sh.shape_type == 13:
            rid = sh._element.blip_rId
            sh._element.getparent().remove(sh._element)
            soltar_imagen(sl, rid)
    for sh, textos in ((cab, ["UNIDAD", "TÍTULO DEL PROCEDIMIENTO"]), (cuerpo, ["Fecha: "])):
        cuerpo_tx = sh._element.graphic.graphicData.tbl.find(qn("a:tr")).find(qn("a:tc")).find(qn("a:txBody"))
        ps = cuerpo_tx.findall(qn("a:p"))
        for p in ps[len(textos):]:
            cuerpo_tx.remove(p)
        for p, t in zip(ps, textos):
            runs = p.findall(qn("a:r"))
            for r in runs[1:]:
                p.remove(r)
            if runs:
                runs[0].find(qn("a:t")).text = t


# ═════════════════════════════════════════════════════════════════════════════
def main(origen, salida, prototipo):
    prs = Presentation(origen)
    slides = list(prs.slides)
    informe = []
    assert len(slides) >= 4, "el informe tiene muy pocas láminas"
    cierre = slides[-1]
    n_dil = next((n for n, sl in enumerate(slides, 1) if n > 1 and _es_diligencia(sl)), None)
    assert n_dil, "no se encontró la sección «DILIGENCIAS RELEVANTES»"
    fijas = slides[1:n_dil - 1]
    informe.append(f"Láminas: portada + {len(fijas)} fijas + diligencias desde la {n_dil} + cierre ({len(slides)} en total)")

    # 1. Reconocer y renombrar las tablas nativas
    claves_por_lamina = {}
    for n, sl in enumerate(slides[:n_dil - 1], 1):
        for sh in sl.shapes:
            if getattr(sh, "has_table", False) and sh.has_table:
                k = clave_tabla(sh)
                if k:
                    assert not any(k in v for v in claves_por_lamina.values()), f"tabla «{k}» repetida"
                    sh.name = f"T:{k}"
                    claves_por_lamina.setdefault(n, set()).add(k)

    # 2. Tablas pegadas como imagen -> tablas nativas
    for n, (nombre, clave) in TABLAS_IMAGEN.items():
        sl = slides[n - 1]
        pic = next((s for s in sl.shapes if s.shape_type == 13 and s.name == nombre), None)
        assert pic is not None, (f"lámina {n}: no está la imagen «{nombre}» de la tabla {clave}; "
                                 "revisar TABLAS_IMAGEN en este script")
        x, y, w = pic.left, pic.top, pic.width
        posicion = list(sl.shapes._spTree).index(pic._element)
        if clave == "homi_d":
            modelo = next(s for s in sl.shapes if s.name == "T:homi_c")
            nueva = tabla_comparativa(sl, modelo, x, y, w)
        elif clave == "migr":
            nueva = tabla_migraciones(sl, x, y, w)
        else:
            nueva = tabla_servicio(sl, x, y, w)
        arbol = sl.shapes._spTree                      # la tabla ocupa la capa de la imagen
        arbol.remove(nueva._element)
        arbol.insert(posicion, nueva._element)
        arbol.remove(pic._element)
        soltar_imagen(sl, pic._element.blip_rId)
        claves_por_lamina.setdefault(n, set()).add(clave)
        informe.append(f" tabla {clave}: imagen «{nombre}» (lámina {n}) -> tabla nativa")

    # 2b. Nota «(*) …» de la lámina de droga: texto que cambia cada semana
    for n, sl in enumerate(slides[1:n_dil - 1], 2):
        for sh in sl.shapes:
            if sh.has_text_frame and norm(sh.text_frame.text).startswith("(*)"):
                sh.name = "NOTA:sintetica"
                informe.append(f" nota «{sh.text_frame.text.strip()[:40]}» (lámina {n}) -> editable")

    # 3. Marcos de foto en las láminas fijas
    for n, sl in enumerate(slides[1:n_dil - 1], 2):
        claves = claves_por_lamina.get(n, set())
        prefijo, rotulo = next(((p, r) for ks, p, r in SECCIONES if ks & claves),
                               (f"lam{n}", f"Lámina {n}"))
        informe.append(f" lámina {n} ({rotulo}): marcos de foto")
        if marcos_de_foto(sl, prefijo, rotulo, informe) == 0:
            informe.append("   (sin fotos)")

    # 4. Prototipo de diligencia y cierre; el resto de diligencias se descarta
    n_proto, proto = elegir_prototipo(slides, prototipo)
    preparar_prototipo(proto)
    informe.append(f" prototipo de diligencia: lámina {n_proto}")
    for sl in slides[n_dil - 1:-1]:
        if sl is not proto:
            quitar_lamina(prs, sl)

    faltan = {"armas", "drogas_g", "drogas_p", "homi_a", "homi_b", "homi_c", "homi_d", "migr",
              "servicio_ext", "flujo", "veh_e", "veh_t", "bandas", "det_b", "det_c"} - set().union(*claves_por_lamina.values())
    assert not faltan, f"no se reconocieron las tablas: {sorted(faltan)}"

    prs.core_properties.title = "Informe Comité Policial"
    prs.core_properties.author = ""
    prs.core_properties.last_modified_by = ""
    salida.parent.mkdir(exist_ok=True)
    prs.save(salida)
    print("\n".join(informe))
    print(f"\n{salida}  ({salida.stat().st_size / 1024:.0f} KB, {len(prs.slides)} láminas)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("origen", help="informe PPT real con el formato nuevo")
    ap.add_argument("--salida", type=Path, default=DESTINO)
    ap.add_argument("--prototipo", type=int, default=0,
                    help="n.º de la lámina de diligencia a usar como molde (por defecto, la de cuerpo más alto)")
    a = ap.parse_args()
    main(a.origen, a.salida, a.prototipo)
