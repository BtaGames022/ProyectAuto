"""Estilo global: valores por defecto de los widgets Tk y temas de ttk."""
import tkinter as tk
from tkinter import ttk

from comite_policial.config.theme import *


def aplicar_estilo(root: tk.Tk):
    """Debe ejecutarse antes de crear los widgets."""
    # ── Widgets clásicos: campos de texto con borde fino y foco azul ─────────
    o = root.option_add
    cuerpo = (FONT, SIZE_BASE)
    for clase in ("Entry", "Text", "Listbox", "Spinbox"):
        o(f"*{clase}.font", cuerpo)
        o(f"*{clase}.background", SURFACE)
        o(f"*{clase}.foreground", INK)
        o(f"*{clase}.relief", "flat")
        o(f"*{clase}.borderWidth", 0)
        o(f"*{clase}.highlightThickness", 1)
        o(f"*{clase}.highlightBackground", LINE_STRONG)
        o(f"*{clase}.highlightColor", NAVY)
        o(f"*{clase}.selectBackground", GOLD_SOFT)
        o(f"*{clase}.selectForeground", INK)
        o(f"*{clase}.insertBackground", NAVY)
    o("*Label.font", cuerpo)
    o("*Button.font", cuerpo)
    o("*Checkbutton.font", cuerpo)
    o("*TCombobox*Listbox.font", cuerpo)
    o("*TCombobox*Listbox.background", SURFACE)
    o("*TCombobox*Listbox.foreground", INK)
    o("*TCombobox*Listbox.selectBackground", NAVY)
    o("*TCombobox*Listbox.selectForeground", "#FFFFFF")

    # ── ttk ──────────────────────────────────────────────────────────────────
    s = ttk.Style(root)
    try:
        s.theme_use("clam")
    except tk.TclError:
        pass

    s.configure("TCombobox", fieldbackground=SURFACE, background=SKY,
                foreground=INK, bordercolor=LINE_STRONG, lightcolor=SURFACE,
                darkcolor=SURFACE, arrowcolor=NAVY, padding=4,
                selectbackground=SURFACE, selectforeground=INK)
    s.map("TCombobox",
          fieldbackground=[("readonly", SURFACE), ("disabled", BG)],
          bordercolor=[("focus", NAVY)],
          lightcolor=[("focus", NAVY)], darkcolor=[("focus", NAVY)],
          selectbackground=[("readonly", SURFACE)],
          selectforeground=[("readonly", INK)])

    for orient in ("Vertical", "Horizontal"):
        s.configure(f"{orient}.TScrollbar", troughcolor=BG, background=LINE_STRONG,
                    bordercolor=BG, lightcolor=LINE_STRONG, darkcolor=LINE_STRONG,
                    arrowcolor=INK_MID, arrowsize=12, gripcount=0, relief="flat")
        s.map(f"{orient}.TScrollbar", background=[("active", INK_SOFT)])
    # barras de las tablas (más finas)
    for orient in ("Vertical", "Horizontal"):
        s.configure(f"Thin.{orient}.TScrollbar", troughcolor=SURFACE,
                    background=LINE_STRONG, bordercolor=SURFACE,
                    lightcolor=LINE_STRONG, darkcolor=LINE_STRONG,
                    arrowcolor=INK_MID, arrowsize=10, relief="flat")

    s.configure("TCheckbutton", background=NAVY_DEEP, foreground="#C9D4EA",
                font=(FONT, SIZE_SMALL))
    s.map("TCheckbutton", background=[("active", NAVY_DEEP)],
          foreground=[("active", "#FFFFFF")])
    s.configure("Surface.TCheckbutton", background=SURFACE, foreground=INK,
                font=(FONT, SIZE_BASE))
    s.map("Surface.TCheckbutton", background=[("active", SURFACE)])
