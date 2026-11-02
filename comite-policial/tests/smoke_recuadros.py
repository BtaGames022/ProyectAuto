import os, sys
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
os.chdir(ROOT)
os.environ.setdefault("COMITE_TIPOS_JSON", os.path.join(
    __import__("tempfile").gettempdir(), "comite_tipos_prueba.json"))
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

# Tipo de procedimiento escrito a mano (no está en la lista)
from comite_policial.config.catalogos import TIPOS_RECUADRO
a._rec_nuevo(); a.update()
campos_antes = [d[0] for d in a._rec_defs_actuales()]
a.rv_unidad.set("Brigada de Prueba")
a.rv_tipo.set("Procedimiento   especial de prueba")      # texto libre
a._rec_guardar(); a.update()                                # sin pulsar Enter
nuevo = a.recuadros[-1]
assert nuevo["tipo"] == "Procedimiento especial de prueba", nuevo["tipo"]
assert [d[0] for d in TIPOS_RECUADRO[nuevo["tipo"]]] == campos_antes   # hereda los campos
assert nuevo["tipo"] in a._rec_cb.cget("values")
assert os.path.isfile(os.environ["COMITE_TIPOS_JSON"])
assert "Procedimiento especial de prueba" in open(os.environ["COMITE_TIPOS_JSON"], encoding="utf-8").read()
# mayúsculas distintas: usa el tipo existente en vez de duplicarlo
a._rec_nuevo(); a.rv_tipo.set("procedimiento especial de prueba"); a._rec_tipo_escrito()
assert a.rv_tipo.get() == "Procedimiento especial de prueba"
# texto vacío: vuelve al último tipo válido
a.rv_tipo.set("  "); a._rec_tipo_escrito()
assert a.rv_tipo.get().strip() in TIPOS_RECUADRO
a.recuadros.pop()
# Cuadro único de Detalle: formato fijo protege rótulos; tipo manual permite edición libre
a._rec_nuevo(); a.update()
assert not a._rec_det_libre
assert "Fecha" in a._rec_widgets and a._rec_widgets["Fecha"][0].winfo_class() == "Entry"
assert a._rec_texto.winfo_class() == "Text"
texto_fijo = a._rec_texto.get("1.0", "end-1c")
assert texto_fijo.startswith("Lugar:")
# Intento de borrar el rótulo "Lugar:" en un tipo predefinido -> se restaura
a._rec_texto.delete("1.0", "1.6")
a._rec_texto.edit_modified(True)
a._rec_det_modificado()
assert a._rec_texto.get("1.0", "end-1c").startswith("Lugar:")
# Escribir después del rótulo sí se conserva
a._rec_texto.insert("1.7", "Av. Alemania 123")
a._rec_texto.edit_modified(True)
a._rec_det_modificado()
assert a._rec_valores_campos()["Lugar"] == "Av. Alemania 123"

# Tipo nuevo manual -> permite reemplazar todo el recuadro de texto libremente
a.rv_unidad.set("Brigada de Prueba")
a.rv_tipo.set("Operativo libre de prueba")
a._rec_tipo_escrito()
assert a._rec_det_libre
a._rec_widgets["Fecha"][0].delete(0, "end")
a._rec_widgets["Fecha"][0].insert(0, "02.NOV.026")
a._rec_texto.delete("1.0", "end")
a._rec_texto.insert("1.0", "Sector: Feria Pinto\n\nFiscalizados: 40 personas\n\nDetalle:\n-02 detenidos\n-01 vehículo incautado")
a._rec_texto.edit_modified(True)
a._rec_det_modificado()
a._rec_guardar(); a.update()
libre = a.recuadros[-1]
assert libre["tipo"] == "Operativo libre de prueba"
assert libre["campos"]["Fecha"] == "02.NOV.026"
assert libre["campos"]["Sector"] == "Feria Pinto"
assert libre["campos"]["Fiscalizados"] == "40 personas"
assert "-02 detenidos" in libre["campos"]["Detalle"]
assert [d[0] for d in campos_def(libre)] == ["Fecha", "Sector", "Fiscalizados", "Detalle"]
a.recuadros.pop()
del TIPOS_RECUADRO["Operativo libre de prueba"]
del TIPOS_RECUADRO["Procedimiento especial de prueba"]
os.remove(os.environ["COMITE_TIPOS_JSON"])
with tempfile.TemporaryDirectory() as tmp:
    ruta, avisos = a._gen_pptx(tmp)
    assert os.path.isfile(ruta), ruta
    print("pptx OK:", len(a.recuadros), "recuadros, avisos:", avisos)
a.destroy()
