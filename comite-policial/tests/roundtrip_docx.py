"""Prueba de ida y vuelta: importar -> modificar -> escribir -> reimportar."""
import copy, zipfile
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
from comite_policial.docio.importador import importar
from comite_policial.docio.escritor import escribir

SRC = "tests/data/informe_ejemplo.docx"
OUT = "tests/data/_salida.docx"
base = open(SRC, "rb").read()
imp = importar(SRC)


def xml(path):
    return zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")


# 1) Sin cambios: el contenido debe ser idéntico al original
escribir(base, copy.deepcopy(imp), imp, OUT)
a, b = xml(SRC), xml(OUT)
print("sin cambios -> XML identico:", a == b, len(a), len(b))
re1 = importar(OUT)
print("sin cambios -> estado igual:", re1["tablas"] == imp["tablas"], len(re1["recuadros"]) == len(imp["recuadros"]))

# 2) Con cambios
est = copy.deepcopy(imp)
est["tablas"]["armas"]["rows"][0][1] = "9"                       # editar celda
est["tablas"]["armas"]["total"][1] = "15"                        # total
est["tablas"]["det_a"]["rows"].append(["PRUEBA NUEVA", "1", "2", "3"])   # fila nueva
del est["tablas"]["bandas"]["rows"][-1]                          # fila menos
est["semana"] = "07.SEP AL 13.SEP.026"; est["turno"] = "PROF Juan PEREZ"
est["texto_x"] = "Texto nuevo de relevancia.\nSegunda linea."
r0 = est["recuadros"][0]; r0["campos"]["Lugar"] = "Lugar EDITADO"; r0["detenidos"] = "3"
dup = copy.deepcopy(est["recuadros"][1]); dup["_orig"] = None; dup["_snap"] = None; dup["_clone_of"] = 1
dup["unidad"] = "UNIDAD DUPLICADA"; est["recuadros"].insert(2, dup)   # nuevo recuadro
del est["recuadros"][-1]                                          # recuadro eliminado
est["recuadros"][5], est["recuadros"][6] = est["recuadros"][6], est["recuadros"][5]  # reordenar
escribir(base, est, imp, OUT)
r = importar(OUT)
print("armas:", r["tablas"]["armas"]["rows"][0], r["tablas"]["armas"]["total"])
print("det_a filas:", len(r["tablas"]["det_a"]["rows"]), r["tablas"]["det_a"]["rows"][-1])
print("bandas filas:", len(r["tablas"]["bandas"]["rows"]))
print("semana/turno:", r["semana"], "|", r["turno"])
print("texto X:", r["texto_x"].replace("\n", " / "))
print("recuadros:", len(r["recuadros"]), "(esperado 48)")
x0 = r["recuadros"][0]
print("rec0:", x0["campos"]["Lugar"], x0["detenidos"], "|", x0["unidad"])
print("dup:", r["recuadros"][2]["unidad"])
print("orden 5/6 intercambiado:", r["recuadros"][5]["campos"]["Fecha"], r["recuadros"][6]["campos"]["Fecha"],
      "vs orig", imp["recuadros"][5]["campos"]["Fecha"], imp["recuadros"][6]["campos"]["Fecha"])
x = xml(OUT)
print("encabezados semana actualizados:", x.count("07.SEP AL 13.SEP.026"), "| restos de la semana vieja:", x.count("06.SEP.026 "))
import docx
d = docx.Document(OUT); print("abre en python-docx OK, tablas:", len(d.tables))
os.remove(OUT)
