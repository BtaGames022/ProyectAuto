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
        section_title(v, "IV‑B", "FLUJO MIGRATORIO COMPLEJO FRONTERIZO LOS LIBERTADORES",
                      "Libertadores, Uspallata y Horcones")
        _, c = card(v, "Flujo diario")
        hint_label(c, "Doble clic para editar · ENTRADOS, SALIDOS y TOTAL se calculan solos")
        widths = [110, 150, 140, 130, 130, 100, 100, 100, 100]
        self.t_flujo = tree_with_scroll(c, COLS_FLUJO, widths, height=7,
                                        editable=[0, 1, 2, 3, 4, 8],
                                        on_change=self._recalc_flujo)
        self.t_flujo.add_total(["TOTAL"] + ["0"] * 8)
        self.t_flujo.add_fuente("FUENTE: DEMIG LAD", len(COLS_FLUJO))

        bar = tk.Frame(c, bg=CARD_BG); bar.pack(pady=8)
        action_btn(bar, "➕  Agregar día", self._add_flujo, SUCCESS, small=True).pack(side="left", padx=4)
        action_btn(bar, "🗑  Eliminar seleccionado", self._del_flujo, DANGER, small=True).pack(side="left", padx=4)

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
        section_title(v, "X", "DILIGENCIAS DE RELEVANCIA DE LA SEMANA")
        _, c = card(v, "Resumen de la semana")
        hint_label(c, "Texto libre que va bajo el título «X. Diligencias de relevancia de la semana»")
        self.txt_x = tk.Text(c, height=10, wrap="word", font=(FONT, 10),
                             relief="solid", bd=1, padx=8, pady=8)
        self.txt_x.pack(fill="x", padx=12, pady=10)
        self.txt_x.bind("<KeyRelease>", lambda e: self._doc_cambio())
