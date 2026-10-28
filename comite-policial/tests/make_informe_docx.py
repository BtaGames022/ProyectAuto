"""Genera un informe semanal .docx sintético con la misma estructura que el real.

Sirve para probar la importación y la sincronización sin versionar el informe
verdadero. Cada tabla lleva los encabezados que busca `docio/mapa.py`, filas de
datos, TOTAL y FUENTE (celda combinada), y los recuadros de la sección IX son
tablas de 1 fila x 2 columnas con las etiquetas en negrita.

    python tests/make_informe_docx.py [ruta]
"""
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

SEMANA = "31.AGO AL 06.SEP.026"
FIN = "06.SEP.026"


def _sombra(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def _texto(celda, texto, negrita=False, centro=True, color=None, tam=9):
    celda.text = ""
    p = celda.paragraphs[0]
    if centro:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(texto)
    r.bold = negrita
    r.font.size = Pt(tam)
    r.font.name = "Arial"
    if color:
        r.font.color.rgb = RGBColor.from_string(color)


def tabla_spec(doc, spec):
    """spec: lista de (tipo, valores); tipo = enc | dat | tot | fte."""
    ncol = max(len(v) for k, v in spec if k != "fte")
    t = doc.add_table(rows=len(spec), cols=ncol)
    t.style = "Table Grid"
    for fila, (tipo, vals) in zip(t.rows, spec):
        if tipo == "fte":
            c = fila.cells[0].merge(fila.cells[-1])
            _texto(c, vals, False, centro=False, tam=8)
            continue
        for ci, (c, v) in enumerate(zip(fila.cells, vals)):
            if tipo in ("enc", "tot"):
                _texto(c, v, True, color="FFFFFF")
                _sombra(c, "002060")
            else:
                _texto(c, str(v), ci == 0, centro=ci != 0)
                if ci == 0:
                    _sombra(c, "DAE9F8")
    doc.add_paragraph()
    return t


def tabla(doc, encabezados, filas, total=None, fuente=None):
    spec = [("enc", e) for e in encabezados] + [("dat", f) for f in filas]
    if total:
        spec.append(("tot", total))
    if fuente:
        spec.append(("fte", fuente))
    return tabla_spec(doc, spec)


def recuadro(doc, unidad, grupo, detenidos, titulo, campos):
    """campos: lista de (etiqueta, valor, es_lista)."""
    t = doc.add_table(rows=1, cols=2)
    t.style = "Table Grid"
    izq, der = t.rows[0].cells
    izq.text = ""
    primero = True
    bloques = [unidad.upper(), grupo.upper(), (f"{int(detenidos):02d} DETENIDO"
               + ("" if int(detenidos) == 1 else "S")) if detenidos else "", titulo.upper()]
    for bloque in [b for b in bloques if b]:
        if not primero:
            izq.add_paragraph()
        for ln in bloque.split("\n"):
            p = izq.paragraphs[0] if primero else izq.add_paragraph()
            primero = False
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(ln)
            r.bold = True
            r.font.size = Pt(9.5)
            r.font.name = "Arial"
    der.text = ""
    primero = True
    previo_linea = False
    for etiqueta, valor, es_lista in campos:
        es_linea = etiqueta in ("Fecha", "Lugar")
        if not primero and not (es_linea and previo_linea):
            der.add_paragraph()
        previo_linea = es_linea
        lineas = valor.split("\n")
        p = der.paragraphs[0] if primero else der.add_paragraph()
        primero = False
        r = p.add_run(f"{etiqueta}:" if es_lista else f"{etiqueta}: ")
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.name = "Arial"
        resto = [] if es_lista else lineas[:1]
        if resto:
            r2 = p.add_run(resto[0])
            r2.font.size = Pt(9.5)
            r2.font.name = "Arial"
        for ln in (lineas if es_lista else lineas[1:]):
            q = der.add_paragraph()
            r3 = q.add_run(ln)
            r3.font.size = Pt(9.5)
            r3.font.name = "Arial"
    doc.add_paragraph()


def crear(destino: str, semana=SEMANA):
    fin = semana.split(" AL ")[-1]
    doc = Document()
    doc.add_heading("INFORME SEMANAL COMITÉ POLICIAL", 0)
    doc.add_paragraph(f"Semana del {semana}")
    doc.add_paragraph("Turno: PROF (A/C) Juan PEREZ")

    sem = f"SEMANA {semana}"
    doc.add_heading("I.- ARMAS DE LA SEMANA", 1)
    tabla(doc, [["TIPO", sem, "2026 A LA FECHA"]],
          [["Armas de Fuego", "6", "77"], ["Aire Comprimido", "4", "18"],
           ["Armas Hechizas", "0", "12"], ["Armas Fogueo", "2", "19"],
           ["Armas de Fantasía", "0", "2"]],
          ["TOTAL", "12", "128"], "FUENTE: Brain")

    doc.add_heading("II.- DROGA INCAUTADA", 1)
    cab = ["DROGA", sem, "2026 A LA FECHA"]
    tabla_spec(doc, [
        ("enc", cab),
        ("dat", ["Cannabis Sativa (g.)", "1.651,21", "1.292.399,17"]),
        ("dat", ["Clorhidrato Cocaína (g.)", "245,24", "516.460,69"]),
        ("dat", ["Cocaína Base (g.)", "463,59", "53.247,43"]),
        ("dat", ["Sintéticas (g.) (*)", "37,13", "4.131,54"]),
        ("tot", ["TOTAL", "2.397,17", "1.866.238,83"]),
        ("enc", cab),
        ("dat", ["Cannabis Sativa Plantas (unidad)", "0", "27.083"]),
        ("tot", ["TOTAL", "0", "27.083"]),
        ("fte", "FUENTE: Brain"),
    ])

    doc.add_heading("III.- HOMICIDIOS", 1)
    tabla(doc, [["HOMICIDIOS", f"SEMANA DEL {semana}", "AÑO 2026 A LA FECHA"]],
          [["BH Valparaíso", "0", "35"], ["BH Los Andes", "0", "5"],
           ["BH San Antonio", "2", "14"]],
          ["TOTAL GENERAL", "2", "54"], "FUENTE: Elaboración propia")
    tabla(doc, [["HOMICIDIOS", f"01.ENE AL {fin}", "% AÑO 2026"]],
          [["Con Detenido PDI", "27", "50,0%"], ["Esclarecido", "10", "18,5%"],
           ["Identificado", "8", "14,8%"], ["Investigación", "9", "16,7%"]],
          ["TOTAL GENERAL", "54", "100%"], "FUENTE: Elaboración propia")
    tabla(doc, [["MODALIDAD", f"SEMANA DEL {semana}", "", "", "AÑO 2026 A LA FECHA"],
                ["", "BH Valparaíso", "BH Los Andes", "BH San Antonio", ""]],
          [["Arma de Fuego", "0", "0", "0", "21"], ["Arma Cortante", "0", "0", "1", "26"],
           ["Elemento Contundente", "0", "0", "1", "3"],
           ["Otras Modalidades", "0", "0", "0", "4"]],
          ["TOTAL GENERAL", "0", "0", "2", "54"], "FUENTE: Elaboración propia")

    doc.add_heading("IV.- MIGRACIONES", 1)
    tabla(doc, [["UNIDAD", "FISCALIZACIONES", "FISCALIZADOS", "DENUNCIADOS",
                 "NOTIFICACIÓN DE EXPULSIÓN", "DETENIDOS", "NACIONALIDAD DENUNCIADOS"]],
          [["POLINT VPO", "1", "45", "13", "3", "0", "08 venezolanos (Art 32 N°3)\n01 colombiano"],
           ["POLINT VMA", "2", "45", "20", "0", "0", "13 venezolanos"],
           ["POLINT LAD", "2", "41", "2", "0", "0", "1 argentino"],
           ["POLINT SAN", "1", "5", "0", "0", "0", "0"],
           ["SEPI QTO", "1", "12", "5", "0", "0", "01 colombiano"],
           ["SEPI IPA", "1", "13", "0", "1", "0", "0"]],
          ["TOTAL", "8", "161", "40", "4", "0", "40"])
    tabla(doc, [["FECHA", "ENTRADOS LIBERTADORES", "ENTRADOS USPALLATA", "SALIDOS HORCONES",
                 "SALIDOS USPALLATA", "ENTRADOS", "SALIDOS", "TOTAL", "DETENIDOS"]],
          [["03.SEP.026", "0", "292", "0", "3", "292", "3", "295", "0"],
           ["04.SEP.026", "1.702", "478", "2.157", "995", "2.180", "3.152", "5.332", "0"],
           ["05.SEP.026", "1.274", "388", "1.617", "1.403", "1.662", "3.020", "4.682", "3"],
           ["06.SEP.026", "1.567", "725", "1.236", "271", "2.292", "1.507", "3.799", "0"]],
          ["TOTAL", "4.543", "1.883", "5.010", "2.672", "6.426", "7.682", "14.108", "3"],
          "FUENTE: DEMIG LAD")

    doc.add_heading("V.- VEHÍCULOS", 1)
    tabla(doc, [["ESTADO", sem, "2026 A LA FECHA"]],
          [["Incautado", "14", "85"], ["Recuperado", "2", "57"]],
          ["TOTAL", "16", "142"], "FUENTE: UU.DD REPOL VPO")
    tabla(doc, [["TIPO DE VEHÍCULO", sem, "2026 A LA FECHA"]],
          [["Camión", "1", "8"], ["Motocicleta", "0", "13"], ["Tractocamión", "0", "2"],
           ["Veh. Furgón", "0", "7"], ["Maquinaria", "0", "1"],
           ["Automóvil (Suv)", "15", "111"]],
          ["TOTAL", "16", "142"], "FUENTE: UU.DD REPOL VPO")

    doc.add_heading("VI.- DELITOS SEXUALES", 1)
    tabla(doc, [["DELITOS SEXUALES", sem, "AÑO 2026 A LA FECHA"]],
          [["Brisex VPO", "1", "20"], ["Brisex SAN", "0", "9"], ["Brisex LAD", "0", "4"]],
          ["TOTAL", "1", "33"], "FUENTE: UU.DD")

    doc.add_heading("VII.- BANDAS DESARTICULADAS", 1)
    tabla(doc, [["BANDAS DESARTICULADAS", "CANTIDAD", "INTEGRANTES", "DELITO"]],
          [["PROVPO", "1", "3", "ART 3 LEY 20.000"], ["PREVMA", "1", "3", "LEY 20.000"],
           ["PROSAN", "0", "0", "0"], ["PROLAD", "1", "2", "LEY 20.000"]],
          None, "FUENTE: UU.DD REPOL VPO")

    doc.add_heading("VIII.- DETENIDOS Y ARRESTADOS", 1)
    tabla(doc, [["DELITO", "ARRESTADO", "DETENIDO", sem]],
          [["Ley 20.000", "2", "10", "12"], ["Robo con violencia", "0", "7", "7"],
           ["Receptación", "1", "3", "4"]],
          ["TOTAL", "3", "20", "23"], "FUENTE: Brain")
    tabla(doc, [["PREFECTURA", "ARRESTADO", "DETENIDO", sem]],
          [["PREVMA", "5", "12", "17"], ["PROLAD", "7", "12", "19"],
           ["PROVPO", "0", "15", "15"], ["PROSAN", "9", "4", "13"],
           ["PRECOCEN", "0", "9", "9"], ["JENACIBER", "0", "0", "0"]],
          ["TOTAL", "21", "52", "73"], "FUENTE: Brain")
    tabla(doc, [["PREFECTURA", "ARRESTADO", "DETENIDO", "AÑO 2026 A LA FECHA"]],
          [["PREVMA", "249", "553", "802"], ["PROLAD", "144", "481", "625"],
           ["PROVPO", "145", "345", "490"], ["PROSAN", "159", "244", "403"],
           ["PRECOCEN", "0", "283", "283"], ["JENACIBER", "3", "24", "27"]],
          ["TOTAL", "700", "1.930", "2.630"], "FUENTE: Brain")

    doc.add_heading("IX.- CONCURRENCIAS Y DILIGENCIAS RELEVANTES", 1)
    casos = [
        ("Brigada de Robos Valparaíso", "", "9", "Robo en lugar no habitado\nLavado de activos",
         [("Fecha", "01.SEP.026.", False), ("Lugar", "Región Metropolitana y Los Lagos.", False),
          ("Detenidos", "09 detenidos de nacionalidad chilena.", False),
          ("Diligencias", "Se ejecutaron 11 órdenes de entrada y registro.", False),
          ("Especies incautadas", "-10 vehículos\n-27 teléfonos móviles", True),
          ("Resultados", "Siete quedaron en prisión preventiva.", False)]),
        ("Brigada de Investigación Criminal Los Andes", "Grupo MT-0", "3",
         "Infracción artículo 4 Ley 20.000",
         [("Fecha", "01.SEP.026.", False), ("Lugar", "San Esteban y Valparaíso.", False),
          ("Droga incautada", "264,29 gramos de cannabis sativa\n51,37 gramos de cocaína base", True),
          ("Resultados", "Prisión preventiva por tráfico.", False)]),
        ("Brigada de Homicidios San Antonio", "", "", "Homicidio con arma cortante",
         [("Fecha", "02.SEP.026.", False), ("Lugar", "San Antonio, Tejas Verdes.", False),
          ("Víctima", "Hombre, chileno, 44 años.", False),
          ("Hechos", "Riña en las inmediaciones de un domicilio.", False),
          ("Resultados", "Se continúan las diligencias.", False)]),
        ("Brigada Investigadora de Delitos en Recintos Portuarios Valparaíso", "", "2",
         "Infracción a Ley de Propiedad Industrial",
         [("Fecha", "02.SEP.026.", False), ("Lugar", "Estación Central, RM.", False),
          ("Diligencias", "Entrada y registro con Aduanas.", False)]),
        ("Brigada de Homicidios Valparaíso", "", "1", "Detención por homicidio con arma de fuego",
         [("Fecha", "03.SEP.026.", False), ("Lugar", "Cerro Yungay, Valparaíso.", False),
          ("Hechos", "Detención con orden vigente.", False)]),
        ("Brigada de Homicidios San Antonio", "", "", "Homicidio con elemento contundente",
         [("Fecha", "03.SEP.026.", False), ("Lugar", "Playa Grande, El Tabo.", False),
          ("Víctima", "Mujer, chilena, 42 años.", False)]),
        ("Brigada Antinarcóticos La Calera", "", "4", "Control terrestre antinarcóticos",
         [("Fecha", "04.SEP.026.", False), ("Lugar", "Peaje Puerto Oscuro.", False),
          ("Droga incautada", "-2 kilos 100 gramos de cannabis sativa", True)]),
        ("Brigada de Investigación Criminal Quillota", "", "1", "Hallazgo de vehículo",
         [("Fecha", "05.SEP.026.", False), ("Lugar", "Quillota.", False),
          ("Diligencia", "Se recuperó un vehículo con encargo.", False)]),
    ]
    for unidad, grupo, det, titulo, campos in casos:
        recuadro(doc, unidad, grupo, det, titulo, campos)

    doc.add_heading("X. DILIGENCIAS DE RELEVANCIA DE LA SEMANA", 1)
    doc.add_paragraph("Operativo coordinado con Fiscalía en la región.")
    doc.add_paragraph("Se incautaron especies por un valor aproximado de 281 millones.")

    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    doc.save(destino)
    return destino


if __name__ == "__main__":
    ruta = sys.argv[1] if len(sys.argv) > 1 else "tests/data/informe_ejemplo.docx"
    print(crear(ruta))
