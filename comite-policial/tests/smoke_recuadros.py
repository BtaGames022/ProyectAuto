import os, sys
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
import tempfile

from comite_policial.ui.app import App
from comite_policial.ui.recuadros import CASILLA_SI, CASILLA_NO
from comite_policial.core.recuadros import *

a = App(); a.update()
a._switch("recuadros"); a.update()

# Cargar el ejemplo de la imagen mediante el formulario real
a.rv_tipo.set("Detención / Infracción Ley 20.000 (MT-0)")
a._rec_cambiar_tipo()
a.rv_unidad.set("Brigada de Investigación Criminal Villa Alemana")
a.rv_grupo.set("Grupo MT-0"); a.rv_det.set("1")
a.rv_titulo.insert("1.0", "Asociación ilícita\nInfracción Ley 20.000")
w = a._rec_widgets
w["Fecha"][0].delete(0, "end"); w["Fecha"][0].insert(0, "31.AGO.026")
w["Lugar"][0].insert(0, "Troncos Viejos.")
w["Detenidos"][0].insert("1.0", "01 hombre, chileno, de 43 años de edad, con orden de detención vigente por asociación ilícita.")
w["Diligencias"][0].insert("1.0", "Equipo Micro Tráfico Cero se percató de un vehículo con encargo por robo de placas patentes.")
w["Especies incautadas"][0].insert("1.0", "-01 vehículo tipo SUV, marca Ford, modelo Territory.\n-02 placas Patentes.")
w["Avalúo droga incautada"][0].insert("1.0", "-19.300.000.-")
w["Resultados"][0].insert("1.0", "Fiscalía instruyó la devolución del vehículo a su propietario.")
a._rec_guardar(); a.update()
a._rec_duplicar(); a.update()
a.rv_tipo.set("Fiscalización migratoria"); a._rec_cambiar_tipo(); a.update()
print("n recuadros", len(a.recuadros))
rec = a.recuadros[0]
print(grupos_izquierda(rec))
for p in parrafos_derecha(rec): print(p)
# Selección de recuadros para el PowerPoint (casilla PPT)
assert len(a._estado_pptx()["recuadros"]) == len(a.recuadros)
a.recuadros[0]["en_pptx"] = False
assert len(a._estado_pptx()["recuadros"]) == len(a.recuadros) - 1
a._rec_refrescar_lista(0); a.update()
assert a.t_rec.raw_values(a.t_rec.get_children()[0])[0] == CASILLA_NO
print(a.v_rec_pptx.get())
a._rec_marcar_todas(False)
assert a._estado_pptx()["recuadros"] == []
a._rec_marcar_todas(True)
with tempfile.TemporaryDirectory() as tmp:
    ruta, avisos = a._gen_pptx(tmp)
    assert os.path.isfile(ruta), ruta
    print("pptx OK:", len(a.recuadros), "recuadros, avisos:", avisos)
a.destroy()
