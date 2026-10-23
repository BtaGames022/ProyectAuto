"""Prueba de integración: importar en la app, editar y comprobar la sincronización."""
import shutil, zipfile
import os, sys
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
EJEMPLO = "tests/data/informe_ejemplo.docx"
if not os.path.isfile(EJEMPLO):
    print(f"SKIP: falta {EJEMPLO} (informe real, no se versiona)")
    sys.exit(0)
from tkinter import filedialog, messagebox
from comite_policial.ui.app import App
from comite_policial.docio.importador import importar

shutil.copy("tests/data/informe_ejemplo.docx", "tests/data/_work.docx")
filedialog.askopenfilename = lambda **k: "tests/data/_work.docx"
messagebox.showinfo = lambda *a, **k: None
messagebox.showwarning = lambda *a, **k: None

a = App(); a.update()
a._doc_importar(); a.update()
print("semana:", a.v_semana.get(), "| turno:", a.v_turno.get())
print("recuadros en app:", len(a.recuadros), "| filas det_a:", len(a.t_det_a.get_children()))
for k in ("armas", "flujo", "migr", "bandas"):
    t = getattr(a, {"armas": "t_armas", "flujo": "t_flujo", "migr": "t_migr", "bandas": "t_bandas"}[k])
    print(k, [t.raw_values(r)[:3] for r in t.get_children()][:2])

# Editar una celda como lo haría el usuario (callback de la tabla + hook global)
t = a.t_armas
fila = t.get_children()[0]
vals = t.raw_values(fila); vals[1] = "10"; t.item(fila, values=vals)
a._recalc_armas(); a._doc_cambio()
a.update(); a.after(1200, a.quit); a.mainloop()      # deja correr el temporizador de auto-guardado

r = importar("tests/data/_work.docx")
print("armas en el DOCX tras editar:", r["tablas"]["armas"]["rows"][0], r["tablas"]["armas"]["total"])
print("status:", a.v_status.get())

# Editar un recuadro desde la interfaz
a._switch("recuadros"); a.update()
a.t_rec.selection_set(a.t_rec.get_children()[0]); a._rec_editar()
a._rec_widgets["Lugar"][0].delete(0, "end"); a._rec_widgets["Lugar"][0].insert(0, "Lugar desde la app")
a._rec_guardar(); a.update(); a.after(1200, a.quit); a.mainloop()
r = importar("tests/data/_work.docx")
print("recuadro 0 lugar en el DOCX:", r["recuadros"][0]["campos"]["Lugar"], "| total recuadros:", len(r["recuadros"]))
a.destroy()
import os; os.remove("tests/data/_work.docx")
