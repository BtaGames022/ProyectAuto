from comite_policial.ui.app import App
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
print(a._gen_word("output")[-30:].encode("ascii", "replace"))
print(a._gen_pptx("output")[-30:].encode("ascii", "replace"))
a.destroy()
