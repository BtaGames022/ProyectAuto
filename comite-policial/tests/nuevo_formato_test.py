"""Formato nuevo del informe: EXTENAS en migraciones, Éxtasis en drogas, servicio
especial de extranjeros, recuadros con varias filas por tabla y tipos nuevos de
concurrencia.  Usa tests/data/informe_nuevo.docx (informe real, no se versiona).

python tests/nuevo_formato_test.py
"""
import copy
import os
import sys
import tempfile
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ORIGEN = "tests/data/informe_nuevo.docx"
if not os.path.isfile(ORIGEN):
    print(f"SKIP: falta {ORIGEN} (informe real, no se versiona)")
    sys.exit(0)

from comite_policial.config.catalogos import TIPOS_RECUADRO
from comite_policial.core import tipos_personalizados
from comite_policial.docio.importador import importar
from comite_policial.docio.escritor import escribir


def xml(ruta):
    x = zipfile.ZipFile(ruta).read("word/document.xml").decode("utf-8")
    return x.split("?>", 1)[-1].strip()


base = open(ORIGEN, "rb").read()
imp = importar(base)
t = imp["tablas"]

# ── Importación ──────────────────────────────────────────────────────────────
assert imp["semana"] == "19.OCT AL 25.OCT.026", imp["semana"]   # «19 AL 25.OCT.026»
assert len(t["migr"]["rows"]) == 6 and len(t["migr"]["rows"][0]) == 8
assert t["migr"]["rows"][0][:6] == ["POLINT VPO", "02", "84", "31", "01", "00"]
assert t["migr"]["total"][6] == "03" and t["migr"]["total"][7] == "76", t["migr"]["total"]
assert [r[0] for r in t["drogas_p"]["rows"]] == ["Cannabis Sativa Plantas (unidad)", "Éxtasis (unidad)"]
assert t["drogas_p"]["total"][1:] == ["3", "27.151"], t["drogas_p"]["total"]
assert "drogas_g" in t and len(t["drogas_g"]["rows"]) == 4
assert len(t["servicio_ext"]["rows"]) == 1 and len(t["servicio_ext"]["rows"][0]) == 9
assert t["servicio_ext"]["rows"][0][1:8] == ["10", "376", "67", "4", "1", "3", "67"]
assert imp["servicio_fechas"] == "19 Y 20.OCT.026", imp["servicio_fechas"]
n_rec = len(imp["recuadros"])
assert n_rec == 25, n_rec                     # incluye las 2 filas de una misma tabla
assert imp["recuadros"][4]["campos"]["Fecha"].startswith("19.OCT")

with tempfile.TemporaryDirectory() as tmp:
    salida = os.path.join(tmp, "s.docx")

    # ── Sin cambios: el contenido queda idéntico ─────────────────────────────
    escribir(base, copy.deepcopy(imp), imp, salida)
    assert xml(salida) == xml(ORIGEN), "sin cambios el XML debe quedar igual"

    # ── Ediciones en el lugar ────────────────────────────────────────────────
    est = copy.deepcopy(imp)
    est["tablas"]["migr"]["rows"][0][6] = "05"
    est["tablas"]["migr"]["total"][6] = "07"
    est["tablas"]["drogas_p"]["rows"][1][1:] = ["4", "63"]
    est["tablas"]["servicio_ext"]["rows"][0][1] = "12"
    est["tablas"]["servicio_ext"]["rows"].append(["SEPI QTO"] + ["1"] * 7 + ["2 venezolanos"])
    est["servicio_fechas"] = "21 Y 22.OCT.026"
    est["recuadros"][4]["campos"]["Lugar"] = "Lugar editado (fila de tabla multiple)"
    escribir(base, est, imp, salida)
    r = importar(salida)
    assert r["tablas"]["migr"]["rows"][0][6] == "05" and r["tablas"]["migr"]["total"][6] == "07"
    assert r["tablas"]["drogas_p"]["rows"][1][:3] == ["Éxtasis (unidad)", "4", "63"]
    assert len(r["tablas"]["servicio_ext"]["rows"]) == 2
    assert r["tablas"]["servicio_ext"]["rows"][1][0] == "SEPI QTO"
    assert r["servicio_fechas"] == "21 Y 22.OCT.026", r["servicio_fechas"]
    assert len(r["recuadros"]) == n_rec
    assert r["recuadros"][4]["campos"]["Lugar"] == "Lugar editado (fila de tabla multiple)"

    # ── Editar la 2.ª fila de la tabla de dos recuadros ─────────────────────
    multi = next(i for i, x in enumerate(imp["recuadros"])
                 if imp["recuadros"][i]["_orig"] == i and i > 0 and
                 x["_orig"] is not None and
                 "EQUIPO MT0" in imp["recuadros"][i - 1]["unidad"].upper())
    est = copy.deepcopy(imp)
    est["recuadros"][multi]["campos"]["Lugar"] = "Segunda fila editada"
    escribir(base, est, imp, salida)
    r = importar(salida)
    assert r["recuadros"][multi]["campos"]["Lugar"] == "Segunda fila editada"
    assert r["recuadros"][multi - 1]["campos"] == imp["recuadros"][multi - 1]["campos"]

    # ── Cambios de estructura: quitar, reordenar y duplicar ──────────────────
    est = copy.deepcopy(imp)
    del est["recuadros"][2]
    est["recuadros"][multi - 2], est["recuadros"][multi - 1] = (
        est["recuadros"][multi - 1], est["recuadros"][multi - 2])
    dup = copy.deepcopy(est["recuadros"][0])
    dup["_orig"], dup["_snap"], dup["_clone_of"] = None, None, 0
    dup["unidad"] = "UNIDAD DUPLICADA"
    est["recuadros"].append(dup)
    escribir(base, est, imp, salida)
    r = importar(salida)
    assert len(r["recuadros"]) == n_rec, len(r["recuadros"])     # -1 +1
    assert [x["unidad"] for x in r["recuadros"][:-1]] == [x["unidad"] for x in est["recuadros"][:-1]]
    assert r["recuadros"][-1]["unidad"] == "UNIDAD DUPLICADA"

    # ── Cambio de semana (fecha corta «19 AL 25.OCT.026») ────────────────────
    est = copy.deepcopy(imp)
    est["semana"] = "26.OCT AL 01.NOV.026"
    escribir(base, est, imp, salida)
    x = xml(salida)
    assert "19 AL 25.OCT.026" not in x and "19 al 25.OCT.026" not in x
    assert x.count("26.OCT AL 01.NOV.026") >= 5, x.count("26.OCT AL 01.NOV.026")
    assert importar(salida)["semana"] == "26.OCT AL 01.NOV.026"

# ── Tipos nuevos de concurrencia ─────────────────────────────────────────────
with tempfile.TemporaryDirectory() as tmp:
    ruta = os.path.join(tmp, "tipos.json")
    antes = set(TIPOS_RECUADRO)
    try:
        nuevos = tipos_personalizados.registrar(imp["recuadros"], ruta)
        assert nuevos, "el informe trae formatos que la app no tenía"
        assert all(n in TIPOS_RECUADRO for n in nuevos)
        assert os.path.isfile(ruta)
        assert tipos_personalizados.registrar(imp["recuadros"], ruta) == []   # sin duplicar
        for n in nuevos:                       # reinicio: se cargan desde el archivo
            del TIPOS_RECUADRO[n]
        tipos_personalizados.PERSONALIZADOS.clear()
        assert set(tipos_personalizados.cargar(ruta)) == set(nuevos)
    finally:
        for n in set(TIPOS_RECUADRO) - antes:
            del TIPOS_RECUADRO[n]
        tipos_personalizados.PERSONALIZADOS.clear()
# ── Interfaz: importar el informe y ver los datos en la app ────────────────
import shutil
from tkinter import filedialog, messagebox
from comite_policial.ui.app import App

cuadros = []
messagebox.showinfo = lambda *a, **k: cuadros.append(a[1])
messagebox.showwarning = lambda *a, **k: cuadros.append(a[1])
messagebox.showerror = lambda *a, **k: cuadros.append(a[1])
with tempfile.TemporaryDirectory() as tmp:
    copia = os.path.join(tmp, "informe.docx")
    shutil.copy(ORIGEN, copia)
    filedialog.askopenfilename = lambda **k: copia
    os.environ["COMITE_TIPOS_JSON"] = os.path.join(tmp, "tipos.json")
    tipos_personalizados.RUTA = os.environ["COMITE_TIPOS_JSON"]
    antes = set(TIPOS_RECUADRO)
    a = App()
    a.update()
    try:
        a._doc_importar()
        a.update()
        vals = lambda tree: [tree.raw_values(r) for r in tree.get_children()]
        migr = vals(a.t_migr)
        assert len(a.t_migr["columns"]) == 8
        assert migr[0][:6] == ["POLINT VPO", "02", "84", "31", "01", "00"], migr[0]
        assert migr[-1][0] == "TOTAL" and migr[-1][6] == "03" and migr[-1][7] == "76", migr[-1]
        assert [r[0] for r in vals(a.t_drogas_p)[:2]] == [
            "Cannabis Sativa Plantas (unidad)", "Éxtasis (unidad)"], vals(a.t_drogas_p)
        serv = vals(a.t_servicio)
        assert len(serv) == 1 and serv[0][1] == "10" and len(serv[0]) == 9, serv
        assert a.v_serv_fechas.get() == "19 Y 20.OCT.026"
        assert len(a.recuadros) == 25
        # la edición del total de migraciones conserva la última celda
        a._switch("migr"); a.update()
        from tkinter import ttk
        rh_migr = int(ttk.Style().lookup(a.t_migr._style_name, "rowheight") or 0)
        rh_armas = int(ttk.Style().lookup(a.t_armas._style_name, "rowheight") or 0)
        assert rh_migr >= 150 and rh_armas < 50, (rh_migr, rh_armas)
        fila = a.t_migr.get_children()[0]
        a.t_migr._abrir_editor(fila, 7)
        a.update()
        ed_nac = a.t_migr._entry_popup.winfo_children()[0]
        assert ed_nac.winfo_class() == "Text", ed_nac.winfo_class()
        assert ed_nac.get("1.0", "end-1c").count("\n") == 6
        ed_nac.event_generate("<Escape>")
        a.update()
        v = a.t_migr.raw_values(fila); v[6] = "4"; a.t_migr.item(fila, values=v)
        a._recalc_migr()
        assert a.t_migr.raw_values(a.t_migr.get_children()[-1])[7] == "76"
        a._switch("servicio"); a.update()
        rh_serv = int(ttk.Style().lookup(a.t_servicio._style_name, "rowheight") or 0)
        assert rh_serv >= 100, rh_serv
        assert serv[0][0].splitlines() == [
            "DEMIG VPO", "DEMIG VMA", "SEPISAN",
            "DEMIGLAD", "SEPIQTO", "SEPIIPA"], serv[0][0]
        assert serv[0][8].splitlines() == [
            "49 venezolanos", "5 haitianos", "7 colombianos",
            "4 argentinos", "1 ecuatoriano", "1 chileno"], serv[0][8]
        # tipos nuevos disponibles en el desplegable
        nuevos = set(TIPOS_RECUADRO) - antes
        assert nuevos and nuevos <= set(a._rec_cb.cget("values")), nuevos
        assert any("opciones nuevas de concurrencia" in c for c in cuadros), cuadros
        # agregar y eliminar filas en Plantas (unidades) + guardar en Word
        a._switch("drogas"); a.update()
        a._add_drogas_p()
        ch_dp = a.t_drogas_p.get_children()
        assert len(ch_dp) == 4, len(ch_dp)  # 2 originales + 1 nueva + TOTAL
        a.t_drogas_p.item(ch_dp[2], values=["LSD (unidad)", "2", "100"])
        a._recalc_drogas_p()
        assert a.t_drogas_p.raw_values(ch_dp[-1]) == ["TOTAL", "5", "27.251"]
        a._add_drogas_p()
        ch_dp = a.t_drogas_p.get_children()
        assert len(ch_dp) == 5
        a.t_drogas_p.selection_set(ch_dp[3])
        a._del_drogas_p()
        assert len(a.t_drogas_p.get_children()) == 4
        a._doc_guardar(auto=False)
        r_dp = importar(copia)
        assert [x[0] for x in r_dp["tablas"]["drogas_p"]["rows"]] == [
            "Cannabis Sativa Plantas (unidad)", "Éxtasis (unidad)", "LSD (unidad)"
        ], r_dp["tablas"]["drogas_p"]
        assert r_dp["tablas"]["drogas_p"]["total"][1:] == ["5", "27.251"]
        # exportar a PowerPoint con y sin la lámina de Servicio Especial
        from pptx import Presentation
        assert a.v_serv_pptx.get() is True
        ruta, avisos = a._gen_pptx(tmp)
        assert os.path.isfile(ruta)
        prs_con = Presentation(ruta)
        tb_dp = next(sh.table for sl in prs_con.slides for sh in sl.shapes if sh.name == "T:drogas_p")
        assert [tb_dp.cell(i, 0).text.strip() for i in range(1, 4)] == [
            "Cannabis Sativa Plantas (unidad)", "Éxtasis (unidad)", "LSD (unidad)"
        ], [tb_dp.cell(i, 0).text.strip() for i in range(len(tb_dp.rows))]
        assert [tb_dp.cell(4, 1).text.strip(), tb_dp.cell(4, 2).text.strip()] == ["5", "27.251"]
        assert any(sh.name == "T:servicio_ext" for sl in prs_con.slides for sh in sl.shapes)
        a.v_serv_pptx.set(False)
        ruta_sin, _ = a._gen_pptx(tmp)
        prs_sin = Presentation(ruta_sin)
        assert len(prs_sin.slides) == len(prs_con.slides) - 1
        assert not any(sh.name == "T:servicio_ext" for sl in prs_sin.slides for sh in sl.shapes)
    finally:
        a.destroy()
        for n in set(TIPOS_RECUADRO) - antes:
            del TIPOS_RECUADRO[n]
        tipos_personalizados.PERSONALIZADOS.clear()
print("OK")
