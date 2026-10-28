"""Vistas adicionales: IV-B Flujo migratorio y X. Diligencias de relevancia (texto)."""
import tkinter as tk
from tkinter import messagebox
from comite_policial.config.theme import *
from comite_policial.core.calculos import *
from comite_policial.ui.widgets import *

COLS_FLUJO = ["FECHA", "ENTRADOS LIBERTADORES", "ENTRADOS USPALLATA",
              "SALIDOS HORCONES", "SALIDOS USPALLATA", "ENTRADOS", "SALIDOS",
              "TOTAL", "DETENIDOS"]


class ExtraViewsMixin:

    # ── IV-B. FLUJO MIGRATORIO ───────────────────────────────────────────────
    def _build_v_flujo(self):
        v = self._make_view("flujo")
        section_title(v, "IV-B", "Flujo migratorio, Complejo Fronterizo Los Libertadores",
                      "Libertadores, Uspallata y Horcones")
        _, c = card(v, "Flujo diario")
        hint_label(c, "Doble clic para editar · ENTRADOS, SALIDOS y TOTAL se calculan solos")
        widths = [110, 190, 170, 150, 150, 100, 100, 90, 100]
        self.t_flujo = tree_with_scroll(c, COLS_FLUJO, widths, height=7,
                                        editable=[0, 1, 2, 3, 4, 8],
                                        on_change=self._recalc_flujo)
        self.t_flujo.add_total(["TOTAL"] + ["0"] * 8)
        self.t_flujo.add_fuente("FUENTE: DEMIG LAD", len(COLS_FLUJO))

        bar = tk.Frame(c, bg=CARD_BG); bar.pack(anchor="w", pady=(8, 0))
        Boton(bar, "Agregar día", self._add_flujo, "success").pack(side="left", padx=(0, 8))
        Boton(bar, "Eliminar seleccionado", self._del_flujo, "danger").pack(side="left")

    def _add_flujo(self):
        t = self.t_flujo
        filas = [r for r in t.get_children()
                 if not {"total", "fuente"} & set(t.item(r, "tags"))]
        total = next((r for r in t.get_children() if "total" in t.item(r, "tags")), None)
        idx = list(t.get_children()).index(total) if total else "end"
        t.insert("", idx, values=["dd.MES.026"] + ["0"] * 8)
        self._recalc_flujo()
        self._doc_cambio()

    def _del_flujo(self):
        t = self.t_flujo
        for r in t.selection():
            if not {"total", "fuente"} & set(t.item(r, "tags")):
                t.delete(r)
        self._recalc_flujo()
        self._doc_cambio()

    def _recalc_flujo(self):
        t = self.t_flujo
        data = [r for r in t.get_children()
                if not {"total", "fuente"} & set(t.item(r, "tags"))]
        tot = [0] * 8
        for r in data:
            v = [str(x) for x in t.item(r, "values")]
            n = [safe_int(x) for x in v[1:]]
            n[4] = n[0] + n[1]            # ENTRADOS
            n[5] = n[2] + n[3]            # SALIDOS
            n[6] = n[4] + n[5]            # TOTAL
            t.item(r, values=[v[0]] + [fmt(x) for x in n])
            tot = [a + b for a, b in zip(tot, n)]
        for r in t.get_children():
            if "total" in t.item(r, "tags"):
                t.item(r, values=["TOTAL"] + [fmt(x) for x in tot])

    # ── X. DILIGENCIAS DE RELEVANCIA (texto libre) ───────────────────────────
    def _build_v_textox(self):
        v = self._make_view("textox")
        section_title(v, "X", "Diligencias de relevancia de la semana")
        _, c = card(v, "Resumen de la semana")
        hint_label(c, "Texto libre que va bajo el título «X. Diligencias de relevancia de la semana»")
        self.txt_x = tk.Text(c, height=14, wrap="word", padx=10, pady=10)
        self.txt_x.pack(fill="x", pady=(0, 4))
        self.txt_x.bind("<KeyRelease>", lambda e: self._doc_cambio())
