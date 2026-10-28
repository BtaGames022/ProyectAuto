"""Desplazamiento con rueda del mouse sin sorpresas.

Causa del error original: la rueda sobre un desplegable (Combobox) cambiaba su
valor y reconstruía el formulario en cada «tick», y sobre un Text o una tabla
se desplazaban a la vez el widget y el panel. Aquí un único enrutador decide
quién se desplaza:
  • Text / tabla / lista con contenido por recorrer → se desplazan ellos.
  • Cualquier otro widget (incluidos los desplegables) → el panel que lo contiene.
  • Un desplegable nunca cambia de valor con la rueda.
"""
import sys
import tkinter as tk
from tkinter import ttk

from comite_policial.config.theme import *

_NATIVOS = ("Text", "Treeview", "Listbox")


def _pasos(event) -> int:
    """Líneas a desplazar (positivo = hacia abajo), sea cual sea la plataforma."""
    d = getattr(event, "delta", 0)
    num = getattr(event, "num", 0)
    if num == 4:
        return -3
    if num == 5:
        return 3
    if not d:
        return 0
    if sys.platform == "darwin":
        return -1 if d > 0 else 1
    pasos = int(round(-d / 120))
    return pasos or (-1 if d > 0 else 1)


def _puede(widget, pasos) -> bool:
    try:
        y0, y1 = widget.yview()
    except (tk.TclError, AttributeError):
        return False
    return (pasos < 0 and y0 > 0.0) or (pasos > 0 and y1 < 1.0)


def _panel(widget):
    """Primer panel desplazable (canvas marcado) que contiene a `widget`."""
    w = widget
    while w is not None:
        if getattr(w, "_es_panel_scroll", False):
            return w
        w = w.master
    return None


def _enrutar(event):
    pasos = _pasos(event)
    if not pasos:
        return "break"
    try:
        w = event.widget.winfo_containing(event.x_root, event.y_root) or event.widget
    except (KeyError, tk.TclError):
        w = event.widget
    if w.winfo_class() in _NATIVOS and _puede(w, pasos):
        w.yview_scroll(pasos, "units")
        return "break"
    panel = _panel(w)
    if panel is not None and _puede(panel, pasos):
        panel.yview_scroll(pasos * 2, "units")
    return "break"


def instalar_rueda(root: tk.Misc):
    """Una sola vez por aplicación."""
    if getattr(root, "_rueda_ok", False):
        return
    root._rueda_ok = True
    eventos = ["<MouseWheel>"] if sys.platform in ("win32", "darwin") else \
        ["<MouseWheel>", "<Button-4>", "<Button-5>"]
    for ev in eventos:
        root.bind_all(ev, _enrutar)
        for clase in _NATIVOS + ("TCombobox", "TScrollbar", "Canvas"):
            root.bind_class(clase, ev, _enrutar)


class ScrollFrame(tk.Frame):
    """Panel con barra vertical; el contenido va en `.interior`.

    La barra ocupa siempre su columna: así aparecer/desaparecer no cambia el
    ancho del contenido ni provoca re-cálculos en cadena.
    """

    def __init__(self, parent, bg=BG, ancho_barra=14, **kw):
        super().__init__(parent, bg=bg, **kw)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.canvas._es_panel_scroll = True
        self.barra = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                                   style="Vertical.TScrollbar")
        self.canvas.configure(yscrollcommand=self.barra.set)
        self.barra.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.interior = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.interior, anchor="nw")
        self._pendiente = None
        self.interior.bind("<Configure>", self._programar)
        self.canvas.bind("<Configure>", self._programar)

    def _programar(self, _e=None):
        # varios <Configure> seguidos se funden en un único ajuste
        if self._pendiente is None:
            self._pendiente = self.after_idle(self._ajustar)

    def _ajustar(self):
        self._pendiente = None
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if w <= 1:
            return
        self.canvas.itemconfigure(self._win, width=w)
        alto = max(self.interior.winfo_reqheight(), h)
        self.canvas.configure(scrollregion=(0, 0, w, alto))
        if self.interior.winfo_reqheight() <= h:
            self.canvas.yview_moveto(0)

    def al_inicio(self):
        self.canvas.yview_moveto(0)
