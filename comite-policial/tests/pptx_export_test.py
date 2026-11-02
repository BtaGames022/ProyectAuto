"""Prueba del exportador PowerPoint (no necesita interfaz): python tests/pptx_export_test.py"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from PIL import Image
from pptx import Presentation

from comite_policial.docio.semana import periodo_cerrado, semana_corta
from comite_policial.exporters.pptx_fotos import leer_marcos
from comite_policial.exporters.pptx_informe import (
    ALTO_LAMINA, PIE_LAMINA, PlantillaIncompatible, exportar_informe, titulo_servicio)

PLANTILLA = os.path.join(os.path.dirname(__file__), "..", "plantillas", "plantilla_comite.pptx")

# ── Datos de ejemplo (semana distinta a la de la plantilla) ──────────────────
ESTADO = {
 "semana": "12.OCT AL 18.OCT.026",
 "servicio_fechas": "12 Y 13.OCT.026",
 "nota_sintetica": "(*) Droga Sintética: Ketamina – MDMA (éxtasis)\nTusi y otras sustancias",
 "tablas": {
  "armas": {"rows":[["Armas de Fuego","6","77"],["Armas Fogueo","2","19"],["Aire Comprimido","4","18"],["Armas Hechizas","0","12"],["Armas de Fantasía","0","2"],["Réplicas de Armas","1","1"]],"total":["TOTAL","13","129"]},
  "drogas_p": {"rows":[["Cannabis Sativa Plantas (unidad)","0","27.083"]],"total":["TOTAL","0","27.083"]},
  "drogas_g": {"rows":[["Cannabis Sativa (g.)","1.651,21","1.292.399,17"],["Clorhidrato Cocaína (g.)","245,24","516.460,69"],["Cocaína Base (g.)","463,59","53.247,43"],["Sintéticas (g.) (*)","37,13","4.131,54"]],"total":["TOTAL","2.397,17","1.866.238,83"]},
  "homi_a": {"rows":[["BH Valparaíso","0","35"],["BH Los Andes","0","5"],["BH San Antonio","2","14"]],"total":["TOTAL GENERAL","2","54"]},
  "homi_c": {"rows":[["Con Detenido PDI","27","50,0%"],["Esclarecido","10","18,5%"],["Identificado","8","14,8%"],["Investigación","9","16,7%"]],"total":["TOTAL, GENERAL","54","100%"]},
  "homi_b": {"rows":[["Arma de Fuego","0","0","0","21"],["Arma Cortante","0","0","1","26"],["Elemento Contundente","0","0","1","3"],["Otras Modalidades","0","0","0","4"]],"total":["TOTAL GENERAL","0","0","2","54"]},
  "homi_d": {"rows":[["Homicidios","55","47","-14.5%"],["Femicidios","2","3","+50.0%"],["Parricidios","2","2","0%"],["Infanticidios","0","0","0%"]],"total":["TOTAL, GENERAL","59","52","-11.9%"]},
  "migr": {"rows":[["POLINT VPO","1","45","13","3","0","01 argentino","08 venezolanos (Art 32 N°3)\n02 venezolanos (Art 109)\n01 colombiano (Art 119)"],
                   ["POLINT VMA","2","45","20","3","0","1 colombiano","01 chino, Art. 32 N°3 in clan\n02 chinos, Art. 109 trabajo sin autoriz\n13 venezolanos, Art. 32 N°3, In clan"],
                   ["POLINT LAD","2","41","2","0","0","0","1 argentino\n1 chino"],["POLINT SAN","1","5","0","0","0","0","0"],
                   ["SEPI QTO","1","12","5","0","0","0","01 colombiano\n04 venezolanos"],["SEPI IPA","1","13","0","1","0","0","0"]],"total":["TOTAL","8","161","40","8","0","1","40"]},
  "servicio_ext": {"rows":[["DEMIG VPO- DEMIG VMA- SEPISAN- DEMIGLAD- SEPIQTO- SEPIIPA","10","376","67","4","1","3","67","49 venezolanos\n5 haitianos\n7 colombianos"]],"total":None},
  "flujo": {"rows":[["03.SEP.026","0","292","0","3","292","3","295","0"],["04.SEP.026","1.702","478","2.157","995","2.180","3.152","5.332","0"],["05.SEP.026","1.274","388","1.617","1.403","1.662","3.020","4.682","3"],["06.SEP.026","1.567","725","1.236","271","2.292","1.507","3.799","0"]],"total":["TOTAL","4.543","1.883","5.010","2.672","6.426","7.682","14.108","3"]},
  "veh_e": {"rows":[["Incautado","14","85"],["Recuperado","2","57"]],"total":["TOTAL","16","142"]},
  "veh_t": {"rows":[["Camión","1","8"],["Motocicleta","0","13"],["Tractocamión","0","2"],["Veh. Furgón","0","7"],["Maquinaria","0","1"],["Automóvil (Suv)","15","111"]],"total":["TOTAL","16","142"]},
  "bandas": {"rows":[["PROVPO","1","3","ART 3 LEY 20.000"],["PROVPO","1","8","ASOCIACION DELICTUAL, SABOTAJE INFORMATICO Y LAVADO DE ACTIVOS"],["PREVMA","1","3","LEY 20.000"],["PROSAN","0","0","0"],["PROLAD","1","2","LEY 20.000"]],"total":None},
  "det_b": {"rows":[["PREVMA","5","12","17"],["PROLAD","7","12","19"],["PROVPO","0","15","15"],["PROSAN","9","4","13"],["PRECOCEN","0","9","9"],["JENACIBER","0","0","0"]],"total":["TOTAL","21","52","73"]},
  "det_c": {"rows":[["PREVMA","249","553","802"],["PROLAD","144","481","625"],["PROVPO","145","345","490"],["PROSAN","159","244","403"],["PRECOCEN","0","283","283"],["JENACIBER","3","24","27"]],"total":["TOTAL","700","1.930","2.630"]},
 },
 "recuadros": [
  {"tipo":"Homicidio / Lesiones graves","unidad":"Brigada de Homicidios San Antonio","grupo":"","detenidos":"","titulo":"Homicidio con arma cortante",
   "campos":{"Fecha":"02.SEP.026.","Lugar":"San Antonio, sector Tejas Verdes, localidad de Llolleo.","Detenido":"Hombre, chileno, 26 años, con antecedentes policiales.","Víctima":"Hombre, chileno, 44 años, con antecedentes por lesiones menos graves.","Hechos":"Tras las diligencias investigativas, se estableció que la víctima y el imputado, actual pareja de la expareja de esta, protagonizaron una riña en las inmediaciones de un domicilio, instancia en que el imputado habría lesionado a la víctima en tres oportunidades con un arma cortante, huyendo posteriormente del lugar.","Resultados":"Se continúan realizando diligencias pertinentes a fin de determinar el paradero del imputado."}},
  {"tipo":"Detención / Infracción Ley 20.000 (MT-0)","unidad":"Brigada Antinarcóticos y contra el Crimen Organizado La Calera","grupo":"","detenidos":"4","titulo":"Control Terrestre Antinarcóticos y Fiscalía Los Vilos",
   "campos":{"Fecha":"04.SEP.026.","Lugar":"Peaje Puerto Oscuro, ruta 5 norte","Detenidos":"2 Hombres, una menor Chilena, 1 Argentina, todos sin antecedentes.","Diligencias":"Funcionarios efectuaron control en peaje Puerto Oscuro, fiscalizando un vehículo con cuatro ocupantes, detectando mediante canes 2 kilos de cannabis sativa.","Droga incautada":"-2 kilos 100 gramos de cannabis sativa","Dinero incautado":"-$ 143.000 dinero en efectivo.","Especies incautadas":"-04 Teléfono celulares.\n-01 automóvil.","Resultados":"Los tres imputados adultos quedaron en prisión preventiva."}},
 ],
}


def textos(prs):
    for n, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if sh.has_text_frame:
                yield n, sh.text_frame.text
            if getattr(sh, "has_table", False) and sh.has_table:
                for fila in sh.table.rows:
                    for c in fila.cells:
                        yield n, c.text


def tabla(prs, clave):
    return next(sh for sl in prs.slides for sh in sl.shapes if sh.name == f"T:{clave}")


def main():
    # ── ayudantes de fecha ───────────────────────────────────────────────────
    assert semana_corta("19.OCT AL 25.OCT.026") == "19 AL 25.OCT.026"
    assert semana_corta("31.AGO AL 06.SEP.026") == "31.AGO AL 06.SEP.026"      # cambia de mes
    assert semana_corta("cualquier cosa") == "cualquier cosa"
    assert periodo_cerrado("19.OCT AL 25.OCT.026") == "ENE - SEP"
    assert periodo_cerrado("31.AGO AL 06.SEP.026") == "ENE - AGO"
    assert periodo_cerrado("29.DIC AL 04.ENE.027") == "ENE" and periodo_cerrado("") == ""
    assert titulo_servicio("19 Y 20.OCT.026") == "LUNES 19 Y MARTES 20.OCT.026"
    assert titulo_servicio("19.OCT.026") == "LUNES 19.OCT.026"
    assert titulo_servicio("LUNES 19 Y MARTES 20.OCT.026") == "LUNES 19 Y MARTES 20.OCT.026"
    assert titulo_servicio("") == "" and titulo_servicio("sin fecha") == "SIN FECHA"

    # ── marcos de foto de la plantilla ───────────────────────────────────────
    marcos = leer_marcos(PLANTILLA)
    claves = [m.clave for m in marcos]
    assert len(claves) == len(set(claves)) >= 20, claves
    for m in marcos:                                  # numeración 1..n por lámina
        assert 1 <= m.indice <= m.total and m.w > 0 and m.h > 0, m
    for lamina in {m.lamina for m in marcos}:
        idx = sorted(m.indice for m in marcos if m.lamina == lamina)
        assert idx == list(range(1, len(idx) + 1)), (lamina, idx)
    assert {"armas_1", "armas_2", "det_1", "det_2", "servicio_1"} <= set(claves)

    with tempfile.TemporaryDirectory() as tmp:
        foto = os.path.join(tmp, "foto.png")
        Image.new("RGB", (400, 200), (200, 30, 30)).save(foto)
        ruta = os.path.join(tmp, "informe.pptx")
        fotos = {"armas_1": foto, "det_2": foto, "droga_2": os.path.join(tmp, "perdida.jpg"),
                 "no_existe": foto}
        avisos = exportar_informe(ESTADO, ruta, PLANTILLA, fotos=fotos)
        prs = Presentation(ruta)
        todo = list(textos(prs))
        assert len(prs.slides) == 9 + len(ESTADO["recuadros"]) + 1, len(prs.slides)
        # los únicos avisos esperados: la foto perdida y la clave sin marco
        assert len(avisos) == 2 and any("droga_2" in a for a in avisos) \
            and any("no_existe" in a for a in avisos), avisos

        # nada de la semana de la plantilla (19 al 25.OCT) debe sobrevivir
        viejos = [t for _, t in todo if "25.OCT" in t or "LUNES 19" in t or "19 AL" in t.upper()]
        assert not viejos, viejos
        juntos = "\n".join(t for _, t in todo)
        for esperado in (
                "Semana del 12 AL 18.OCT.026", "SEMANA 12 AL 18.OCT.026", "SEMANA DEL 12 AL 18.OCT.026",
                "01.ENE AL 18.OCT.026", "HOMICIDIOS DEL 01.ENE al 18.OCT.026",
                "MODALIDAD EN HOMICIDIOS SEMANA DEL 12 al 18.OCT.026",
                "COMPARATIVO ENE-SEP 2025/2026", "ENE - SEP 2025", "ENE - SEP 2026",
                "VAR. % 2025 / 2026", "-11,9%", "DETENIDOS 2026 A LA FECHA",
                "LUNES 12 Y MARTES 13.OCT.026", "Réplicas de Armas", "EXTENAS",
                "49 venezolanos", "1.866.238,83",
                "BRIGADA DE HOMICIDIOS SAN ANTONIO", "Fecha: 02.SEP.026."):
            assert esperado in juntos, esperado
        # nota de droga sintética: la que escribió el usuario (dos párrafos), no la de la plantilla
        assert "MDMA (éxtasis)" in juntos and "Tusi y otras sustancias" in juntos
        assert "Ketamina – Éxtasis" not in juntos
        nota = next(sh for sl in prs.slides for sh in sl.shapes if sh.name == "NOTA:sintetica")
        assert len(nota.text_frame.paragraphs) == 2
        # la semana aparece también en el título de Migraciones, que mezcla dos textos
        assert any("FISCALIZACIONES A EXTRANJEROS" in t and "Semana del 12 AL 18.OCT.026" in t
                   for _, t in todo)

        # tablas: nombres canónicos, filas y combinaciones
        assert len(tabla(prs, "migr").table.rows) == 1 + 6 + 1
        assert len(tabla(prs, "migr").table.columns) == 8
        assert len(tabla(prs, "servicio_ext").table.rows) == 2
        assert len(tabla(prs, "armas").table.rows) == 1 + 6 + 1 + 1
        homi_d = tabla(prs, "homi_d")
        tc = homi_d._element.graphic.graphicData.tbl.findall(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}tr")[1].findall(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}tc")[3]
        assert tc.get("rowSpan") == "4"

        # fotos: en su marco (misma posición y tamaño), sin marcos vacíos ni formas ocultas
        for clave, n_lam in (("armas_1", 2), ("det_2", 9)):
            m = next(m for m in marcos if m.clave == clave)
            pics = [sh for sh in prs.slides[n_lam - 1].shapes if sh.shape_type == 13]
            assert any((p.left, p.top, p.width, p.height) == (m.x, m.y, m.w, m.h) for p in pics), clave
        assert not [sh for sl in prs.slides for sh in sl.shapes if sh.name.startswith("FOTO:")]

        # las láminas de diligencias van antes del cierre «PDI»
        assert any(t.strip() == "PDI" for n, t in todo if n == len(prs.slides))
        for sl in list(prs.slides)[9:-1]:
            assert not [sh for sh in sl.shapes if sh.has_text_frame and "UNIDAD" == sh.text_frame.text]
            for sh in sl.shapes:
                assert sh.top + sh.height <= ALTO_LAMINA, (sh.name, sh.top + sh.height)

        # sin recuadros ni tablas también debe exportar (sin restos de la semana de la plantilla)
        vacio = os.path.join(tmp, "vacio.pptx")
        exportar_informe({"semana": "12.OCT AL 18.OCT.026", "tablas": {}, "recuadros": []},
                         vacio, PLANTILLA)
        pv = Presentation(vacio)
        assert len(pv.slides) == 10
        assert not [t for _, t in textos(pv) if "Semana del 19" in t or "SEMANA 19" in t]

        # una nota muy larga baja la letra para no pisar la segunda tabla de droga
        def letra_nota(n_lineas):
            ruta_n = os.path.join(tmp, f"nota_{n_lineas}.pptx")
            exportar_informe(dict(ESTADO, nota_sintetica="\n".join(
                f"(*) sustancia número {i}" for i in range(n_lineas))), ruta_n, PLANTILLA)
            pn_ = Presentation(ruta_n)
            sh_ = next(sh for sl in pn_.slides for sh in sl.shapes if sh.name == "NOTA:sintetica")
            sz_ = {int(r.get("sz")) for r in sh_._element.iter(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")}
            return pn_, sh_, max(sz_)
        pn_, sh_, sz_normal = letra_nota(2)
        assert sz_normal == 1000                          # dos líneas caben con la letra de la plantilla
        pn_, sh_, sz_4 = letra_nota(4)                    # cuatro ya no caben a 10 pt: baja y encaja
        assert sz_4 < 1000 and tabla(pn_, "drogas_p").top >= sh_.top + sh_.height - 40000, sz_4
        assert letra_nota(14)[2] == 700                   # una nota absurda llega al mínimo sin romper nada

        # nota vacía: la forma se quita (no sobrevive la de otra semana); sin la clave, queda la de la plantilla
        vacia = os.path.join(tmp, "sin_nota.pptx")
        exportar_informe(dict(ESTADO, nota_sintetica="  \n "), vacia, PLANTILLA)
        pn = Presentation(vacia)
        assert not [sh for sl in pn.slides for sh in sl.shapes if sh.name == "NOTA:sintetica"]
        assert "Droga Sintética" not in "\n".join(t for _, t in textos(pn))
        sin_clave = os.path.join(tmp, "sin_clave.pptx")
        exportar_informe({k: v for k, v in ESTADO.items() if k != "nota_sintetica"}, sin_clave, PLANTILLA)
        assert "Droga Sintética: Ketamina – Éxtasis" in "\n".join(
            t for _, t in textos(Presentation(sin_clave)))

        # Migraciones con listas largas de nacionalidades: cabe en la lámina bajando la letra
        def migr_con(lineas_por_unidad):
            return {"semana": "12.OCT AL 18.OCT.026", "recuadros": [], "tablas": {
                "migr": {"rows": [[u, "1", "2", "3", "4", "5", "6",
                                   "\n".join(f"{i:02d} venezolanos" for i in range(n))]
                                  for u, n in zip(("POLINT VPO", "POLINT VMA", "POLINT LAD",
                                                   "POLINT SAN", "SEPI QTO", "SEPI IPA"),
                                                  lineas_por_unidad)],
                         "total": ["TOTAL", "6", "12", "18", "24", "30", "36", "84"]},
                "servicio_ext": {"rows": [["U1", "1", "2", "3", "4", "5", "6", "7", "x"],
                                          ["U2", "1", "2", "3", "4", "5", "6", "7", "y"],
                                          ["U3", "1", "2", "3", "4", "5", "6", "7", "z"]],
                                 "total": None}}}
        pl = os.path.join(tmp, "largo.pptx")
        av = exportar_informe(migr_con([9, 9, 2, 9, 2, 1]), pl, PLANTILLA)
        p2 = Presentation(pl)
        migr = tabla(p2, "migr")
        assert migr.top + migr.height <= ALTO_LAMINA - PIE_LAMINA + 1, migr.top + migr.height
        assert not [a for a in av if "migr" in a], av
        letra = {int(r.get("sz")) for r in migr._element.iter(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")}
        assert min(letra) < 900, letra                    # las nacionalidades bajaron de 9 pt
        assert len(tabla(p2, "servicio_ext").table.rows) == 1 + 3
        # un caso imposible (14 líneas en cada unidad) no se oculta: avisa que se sale
        av = exportar_informe(migr_con([14] * 6), os.path.join(tmp, "imposible.pptx"), PLANTILLA)
        assert any("migr" in a for a in av), av

    # plantilla de un formato anterior (sin tablas ni marcos con nombre): error claro
    with tempfile.TemporaryDirectory() as tmp:
        vieja = os.path.join(tmp, "vieja.pptx")
        Presentation().save(vieja)
        try:
            exportar_informe(ESTADO, os.path.join(tmp, "x.pptx"), vieja)
        except PlantillaIncompatible as exc:
            assert "construir_plantilla_pptx" in str(exc)
        else:
            raise AssertionError("debió rechazar una plantilla sin nombres canónicos")
        assert not os.path.exists(os.path.join(tmp, "x.pptx"))
    print("OK")


if __name__ == "__main__":
    main()
