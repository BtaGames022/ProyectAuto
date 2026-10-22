"""Prueba del exportador PowerPoint (no necesita interfaz): python tests/pptx_export_test.py"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from pptx import Presentation
from comite_policial.exporters.pptx_informe import exportar_informe

PLANTILLA = os.path.join(os.path.dirname(__file__), "..", "plantillas", "plantilla_comite.pptx")

# ── Datos de ejemplo (semana distinta a la de la plantilla) ──────────────────
ESTADO = {
 "semana": "12.OCT AL 18.OCT.026",
 "tablas": {
  "armas": {"rows":[["Armas de Fuego","6","77"],["Aire Comprimido","4","18"],["Armas Hechizas","0","12"],["Armas Fogueo","2","19"],["Armas de Fantasía","0","2"]],"total":["TOTAL","12","128"]},
  "drogas_p": {"rows":[["Cannabis Sativa Plantas (unidad)","0","27.083"]],"total":["TOTAL","0","27.083"]},
  "drogas_g": {"rows":[["Cannabis Sativa (g.)","1.651,21","1.292.399,17"],["Clorhidrato Cocaína (g.)","245,24","516.460,69"],["Cocaína Base (g.)","463,59","53.247,43"],["Sintéticas (g.) (*)","37,13","4.131,54"]],"total":["TOTAL","2.397,17","1.866.238,83"]},
  "homi_a": {"rows":[["BH Valparaíso","0","35"],["BH Los Andes","0","5"],["BH San Antonio","2","14"]],"total":["TOTAL GENERAL","2","54"]},
  "homi_c": {"rows":[["Con Detenido PDI","27","50,0%"],["Esclarecido","10","18,5%"],["Identificado","8","14,8%"],["Investigación","9","16,7%"]],"total":["TOTAL, GENERAL","54","100%"]},
  "homi_b": {"rows":[["Arma de Fuego","0","0","0","21"],["Arma Cortante","0","0","1","26"],["Elemento Contundente","0","0","1","3"],["Otras Modalidades","0","0","0","4"]],"total":["TOTAL GENERAL","0","0","2","54"]},
  "homi_d": {"rows":[["Homicidios","55","47","-14,5%"],["Femicidios","2","3","+50.0%"],["Parricidios","2","2","0%"],["Infanticidios","0","0","0%"]],"total":["TOTAL, GENERAL","59","52","-11,9%"]},
  "migr": {"rows":[["POLINT VPO","1","45","13","3","0","08 venezolanos (Art 32 N°3)\n02 venezolanos (Art 109)\n01 colombiano (Art 119)"],
                   ["POLINT VMA","2","45","20","3 venezolanos\n1 colombiano","0","01 chino, Art. 32 N°3 in clan\n02 chinos, Art. 109 trabajo sin autoriz\n13 venezolanos, Art. 32 N°3, In clan"],
                   ["POLINT LAD","2","41","2","0","0","1 argentino\n1 chino"],["POLINT SAN","1","5","0","0","0","0"],
                   ["SEPI QTO","1","12","5","0","0","01 colombiano\n04 venezolanos"],["SEPI IPA","1","13","0","1","0","0"]],"total":["TOTAL","8","161","40","8","0","40"]},
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
            if sh.has_table:
                for fila in sh.table.rows:
                    for c in fila.cells:
                        yield n, c.text


def main():
    with tempfile.TemporaryDirectory() as tmp:
        ruta = os.path.join(tmp, "informe.pptx")
        avisos = exportar_informe(ESTADO, ruta, PLANTILLA)
        prs = Presentation(ruta)
        todo = list(textos(prs))
        assert len(prs.slides) == 8 + len(ESTADO["recuadros"]) + 1, len(prs.slides)
        assert not avisos, avisos
        # nada de la semana de la plantilla debe sobrevivir
        viejos = [t for _, t in todo if "31.AGO" in t]
        assert not viejos, viejos
        juntos = "\n".join(t for _, t in todo)
        for esperado in ("Semana del 12.OCT AL 18.OCT.026", "SEMANA 12.OCT AL 18.OCT.026",
                         "01.ENE AL 18.OCT.026", "1.866.238,83", "VAR. % 2025 / 2026",
                         "BRIGADA DE HOMICIDIOS SAN ANTONIO", "Fecha: 02.SEP.026."):
            assert esperado in juntos, esperado
        # las láminas de diligencias van antes del cierre «PDI»
        assert any(t.strip() == "PDI" for _, t in todo if _ == len(prs.slides))
        # sin recuadros ni tablas también debe exportar
        vacio = os.path.join(tmp, "vacio.pptx")
        exportar_informe({"semana": "12.OCT AL 18.OCT.026", "tablas": {}, "recuadros": []},
                         vacio, PLANTILLA)
        assert len(Presentation(vacio).slides) == 9
    print("OK")


if __name__ == "__main__":
    main()
