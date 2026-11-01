"""III-D (comparativo interanual): importación desde tabla y escritura de vuelta al Word.

python tests/homi_d_test.py
"""
import copy
import os
import sys
import tempfile
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tests"))
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from comite_policial.docio.importador import importar
from comite_policial.docio.escritor import escribir
import make_informe_docx


def xml(ruta):
    x = zipfile.ZipFile(ruta).read("word/document.xml").decode("utf-8")
    return x.split("?>", 1)[-1].strip()      # sin el prólogo (comillas y saltos distintos)


def comprobar(origen):
    base = open(origen, "rb").read()
    imp = importar(base)
    d = imp["tablas"].get("homi_d")
    assert d, "III-D no se importó"
    assert [r[0] for r in d["rows"]] == ["Homicidios", "Femicidios", "Parricidios", "Infanticidios"], d["rows"]
    assert d["rows"][0][1:3] == ["55", "47"], d["rows"][0]
    assert d["total"][0].startswith("TOTAL") and d["total"][3] == "-11,9%", d["total"]
    assert len(d["rows"][0]) == 4

    with tempfile.TemporaryDirectory() as tmp:
        salida = os.path.join(tmp, "salida.docx")
        # sin cambios: el XML queda idéntico
        escribir(base, copy.deepcopy(imp), imp, salida)
        assert xml(salida) == xml(origen)
        # con cambios: se escriben las cifras y la variación sigue solo en el total
        est = copy.deepcopy(imp)
        est["tablas"]["homi_d"]["rows"][0][1:] = ["60", "30", "-50,0%"]
        est["tablas"]["homi_d"]["total"] = ["TOTAL, GENERAL", "64", "35", "-45.3%"]   # la app usa punto
        escribir(base, est, imp, salida)
        r = importar(salida)["tablas"]["homi_d"]
        assert r["rows"][0][1:3] == ["60", "30"], r["rows"][0]
        assert r["rows"][0][3] == "", "la variación por fila debe seguir en blanco"
        assert r["total"][1:] == ["64", "35", "-45,3%"], r["total"]   # el Word usa coma
        assert len(r["rows"]) == 4


with tempfile.TemporaryDirectory() as tmp:
    sintetico = os.path.join(tmp, "informe.docx")
    make_informe_docx.crear(sintetico)
    comprobar(sintetico)

real = os.path.join("tests", "data", "informe_iiid.docx")
if os.path.isfile(real):          # informe real del comité (no se versiona)
    comprobar(real)
print("OK")
