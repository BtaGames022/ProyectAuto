"""Regresión: la rueda del mouse sobre el editor de recuadros no debe cambiar
desplegables ni reconstruir el formulario (el bug «se repite la vista»)."""
import os, sys
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from comite_policial.ui.app import App


def rueda(app, widget, delta=-120):
    app.update()
    x = widget.winfo_rootx() + widget.winfo_width() // 2
    y = widget.winfo_rooty() + widget.winfo_height() // 2
    widget.event_generate("<MouseWheel>", delta=delta, rootx=x, rooty=y)
    app.update()


def combos(w, out):
    for c in w.winfo_children():
        if c.winfo_class() == "TCombobox":
            out.append(c)
        combos(c, out)
    return out


a = App()
a.geometry("1200x560")
a._switch("recuadros")
a.update()

reconstrucciones = []
original = a._rec_construir_campos
a._rec_construir_campos = lambda *x, **k: (reconstrucciones.append(1), original(*x, **k))[1]

tipo = a.rv_tipo.get()
cbs = combos(a._ws, [])
assert cbs, "no se encontraron desplegables"
for cb in cbs:
    antes = cb.get()
    for _ in range(6):
        rueda(a, cb)
    assert cb.get() == antes, f"la rueda cambió un desplegable: {antes!r} -> {cb.get()!r}"
assert a.rv_tipo.get() == tipo, "cambió el tipo de procedimiento"
assert not reconstrucciones, f"se reconstruyó el formulario {len(reconstrucciones)} veces"

# la rueda sobre un desplegable desplaza el panel en vez de quedarse inmóvil
a._scroll.al_inicio()
a.update()
y0 = a._canvas.yview()[0]
for _ in range(5):
    rueda(a, cbs[0])
assert a._canvas.yview()[0] > y0, "el panel no se desplazó"

# y sobre un Text sin desbordar también
txt = [w for w, c in a._rec_widgets.values() if w.winfo_class() == "Text"][0]
y1 = a._canvas.yview()[0]
for _ in range(3):
    rueda(a, txt)
assert a._canvas.yview()[0] > y1, "el panel no se desplazó sobre un Text"

# la vista previa no crece sin límite aunque el relato sea larguísimo
w = a._rec_widgets
w["Diligencias"][0].insert("1.0", "texto largo " * 4000)
a._rec_preview()
a.update()
assert int(a._pv_der.cget("height")) <= 24, a._pv_der.cget("height")
a.destroy()
print("OK")
