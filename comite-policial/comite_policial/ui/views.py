"""Vistas de cada sección del informe (I a VIII-C)."""
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.calculos import *
from comite_policial.ui.widgets import *


class ViewsMixin:
    """Construcción de las vistas de cada sección."""

    # =========================================================================
    # I. ARMAS
    # =========================================================================
    def _build_v_armas(self):
        v = self._make_view("armas")
        section_title(v, "I", "Armas de la semana")

        _, c = card(v, "Incautación de Armas — Semana vs. 2026 a la fecha")
        hint_label(c, "Doble clic sobre una celda numérica para editarla")

        cols   = ["TIPO DE ARMA", "SEMANA", "2026 A LA FECHA"]
        widths = [300, 160, 180]
        self.t_armas = tree_with_scroll(c, cols, widths, height=7,
                                        editable=[1, 2],
                                        on_change=self._recalc_armas)
        self._fill_armas_default()


    def _fill_armas_default(self):
        self.t_armas.clear()
        for tipo in TIPOS_ARMAS:
            self.t_armas.add_row([tipo, "0", "0"])
        self.t_armas.add_total(["TOTAL", "0", "0"])
        self.t_armas.add_fuente("FUENTE: Brain", 3)

    def _recalc_armas(self):
        rows = list(self.t_armas.get_children())
        ts, ty = 0, 0
        data_rows = [r for r in rows
                     if "total" not in self.t_armas.item(r,"tags")
                     and "fuente" not in self.t_armas.item(r,"tags")]
        for r in data_rows:
            vals = self.t_armas.item(r, "values")
            ts += safe_int(vals[1]); ty += safe_int(vals[2])
        # Actualizar fila total
        for r in rows:
            if "total" in self.t_armas.item(r, "tags"):
                self.t_armas.item(r, values=("TOTAL", fmt(ts), fmt(ty)))

    # =========================================================================
    # II. DROGAS
    # =========================================================================
    def _build_v_drogas(self):
        v = self._make_view("drogas")
        section_title(v, "II", "Droga incautada en la semana")

        _, c1 = card(v, "Sustancias en Gramos")
        hint_label(c1, "Doble clic para editar · Admite decimales (ej: 1651,21)")
        cols = ["DROGA", "SEMANA", "2026 A LA FECHA"]
        wid  = [320, 160, 180]
        self.t_drogas_g = tree_with_scroll(c1, cols, wid, height=6,
                                           editable=[1, 2],
                                           on_change=self._recalc_drogas_g)
        self._fill_drogas_g()

        _, c2 = card(v, "Plantas (unidades)")
        hint_label(c2, "Doble clic para editar nombre o cantidades · Los totales se recalculan automáticamente")
        self.t_drogas_p = tree_with_scroll(c2, cols, wid, height=6,
                                           editable=[0, 1, 2],
                                           on_change=self._recalc_drogas_p)
        self._fill_drogas_p()
        bar_p = tk.Frame(c2, bg=CARD_BG)
        bar_p.pack(anchor="w", pady=(8, 0))
        Boton(bar_p, "Agregar fila", self._add_drogas_p, "success").pack(side="left", padx=(0, 8))
        Boton(bar_p, "Eliminar seleccionada", self._del_drogas_p, "danger").pack(side="left")

        _, c3 = card(v, "Nota de droga sintética (*)")
        hint_label(c3, "Sale bajo la tabla de gramos en la lámina de droga del PowerPoint · "
                       "un párrafo por línea · vacía = la nota no se muestra")
        self.txt_nota_sint = tk.Text(c3, height=3, wrap="word", padx=10, pady=8,
                                     font=(FONT, SIZE_BASE))
        self.txt_nota_sint.insert("1.0", NOTA_SINTETICA)
        self.txt_nota_sint.pack(fill="x")

    def _fill_drogas_g(self):
        self.t_drogas_g.clear()
        for tipo in TIPOS_DROGAS_G:
            self.t_drogas_g.add_row([tipo, "0,00", "0,00"])
        self.t_drogas_g.add_total(["TOTAL", "0,00", "0,00"])
        self.t_drogas_g.add_fuente("FUENTE: Brain", 3)

    def _fill_drogas_p(self):
        self.t_drogas_p.clear()
        for tipo in TIPOS_DROGAS_P:
            self.t_drogas_p.add_row([tipo, "0", "0"])
        self.t_drogas_p.add_total(["TOTAL", "0", "0"])
        self.t_drogas_p.add_fuente("FUENTE: Brain", 3)

    def _recalc_drogas_g(self):
        rows = [r for r in self.t_drogas_g.get_children()
                if "total" not in self.t_drogas_g.item(r,"tags")
                and "fuente" not in self.t_drogas_g.item(r,"tags")]
        ts = sum(safe_float(self.t_drogas_g.item(r,"values")[1]) for r in rows)
        ty = sum(safe_float(self.t_drogas_g.item(r,"values")[2]) for r in rows)
        for r in self.t_drogas_g.get_children():
            if "total" in self.t_drogas_g.item(r,"tags"):
                self.t_drogas_g.item(r, values=("TOTAL", fmt(ts), fmt(ty)))

    def _add_drogas_p(self):
        t = self.t_drogas_p
        hijos = list(t.get_children())
        total = next((r for r in hijos if "total" in t.item(r, "tags")), None)
        idx = hijos.index(total) if total else "end"
        nueva = t.insert("", idx, values=("Nueva sustancia (unidad)", "0", "0"))
        self._recalc_drogas_p()
        self._doc_cambio()
        t.selection_set(nueva)
        t.focus(nueva)
        t.see(nueva)

    def _del_drogas_p(self):
        t = self.t_drogas_p
        datos = [r for r in t.get_children()
                 if not {"total", "fuente"} & set(t.item(r, "tags"))]
        for r in t.selection():
            if r not in datos:
                continue
            if len(datos) <= 1:
                t.item(r, values=("", "0", "0"))
            else:
                t.delete(r)
                datos.remove(r)
        self._recalc_drogas_p()
        self._doc_cambio()

    def _recalc_drogas_p(self):
        rows = [r for r in self.t_drogas_p.get_children()
                if "total" not in self.t_drogas_p.item(r,"tags")
                and "fuente" not in self.t_drogas_p.item(r,"tags")]
        ts = sum(safe_int(self.t_drogas_p.item(r,"values")[1]) for r in rows)
        ty = sum(safe_int(self.t_drogas_p.item(r,"values")[2]) for r in rows)
        for r in self.t_drogas_p.get_children():
            if "total" in self.t_drogas_p.item(r,"tags"):
                self.t_drogas_p.item(r, values=("TOTAL", fmt(ts), fmt(ty)))

    # =========================================================================
    # III-A. HOMICIDIOS POR BH
    # =========================================================================
    def _build_v_homi_a(self):
        v = self._make_view("homi_a")
        section_title(v, "III-A", "Homicidios por brigada")

        _, c = card(v, "Homicidios — Semana vs. Año a la Fecha")
        hint_label(c, "Doble clic para editar — Los totales se recalculan automáticamente")
        cols   = ["HOMICIDIOS", "SEMANA", "AÑO 2026 A LA FECHA"]
        widths = [280, 160, 200]
        self.t_homi_a = tree_with_scroll(c, cols, widths, height=6,
                                         editable=[1, 2],
                                         on_change=self._recalc_homi_a)
        for bh in BH_PREF:
            self.t_homi_a.add_row([bh, "0", "0"])
        self.t_homi_a.add_total(["TOTAL GENERAL", "0", "0"])
        self.t_homi_a.add_fuente(
            "FUENTE: Elaboración propia a partir de las concurrencias BH Repol Vpo", 3)


    def _recalc_homi_a(self):
        data = [r for r in self.t_homi_a.get_children()
                if "total" not in self.t_homi_a.item(r,"tags")
                and "fuente" not in self.t_homi_a.item(r,"tags")]
        ts = sum(safe_int(self.t_homi_a.item(r,"values")[1]) for r in data)
        ty = sum(safe_int(self.t_homi_a.item(r,"values")[2]) for r in data)
        for r in self.t_homi_a.get_children():
            if "total" in self.t_homi_a.item(r,"tags"):
                self.t_homi_a.item(r, values=("TOTAL GENERAL", fmt(ts), fmt(ty)))

    # =========================================================================
    # III-B. MODALIDAD EN HOMICIDIOS
    # =========================================================================
    def _build_v_homi_b(self):
        v = self._make_view("homi_b")
        section_title(v, "III-B", "Modalidad en homicidios")

        _, c = card(v, "Modalidad — Desglose por Brigada y Año")
        hint_label(c, "Doble clic para editar · Columnas: BH VPO · BH LAD · BH SAN · Año a la fecha")
        cols   = ["MODALIDAD", "BH Valparaíso", "BH Los Andes",
                  "BH San Antonio", "AÑO 2026 A LA FECHA"]
        widths = [220, 130, 130, 130, 180]
        self.t_homi_b = tree_with_scroll(c, cols, widths, height=7,
                                         editable=[1, 2, 3, 4],
                                         on_change=self._recalc_homi_b)
        for m in MODALIDADES:
            self.t_homi_b.add_row([m, "0", "0", "0", "0"])
        self.t_homi_b.add_total(["TOTAL GENERAL", "0", "0", "0", "0"])
        self.t_homi_b.add_fuente(
            "FUENTE: Elaboración propia a partir de las concurrencias BH Repol Vpo", 5)


    def _recalc_homi_b(self):
        data = [r for r in self.t_homi_b.get_children()
                if "total" not in self.t_homi_b.item(r,"tags")
                and "fuente" not in self.t_homi_b.item(r,"tags")]
        tots = [sum(safe_int(self.t_homi_b.item(r,"values")[i]) for r in data)
                for i in range(1, 5)]
        for r in self.t_homi_b.get_children():
            if "total" in self.t_homi_b.item(r,"tags"):
                self.t_homi_b.item(r, values=(
                    "TOTAL GENERAL", fmt(tots[0]), fmt(tots[1]),
                    fmt(tots[2]), fmt(tots[3])))

    # =========================================================================
    # III-C. RESULTADO INVESTIGATIVO
    # =========================================================================
    def _build_v_homi_c(self):
        v = self._make_view("homi_c")
        section_title(v, "III-C", "Resultado en homicidios",
                      "porcentaje sobre base 100")

        _, c = card(v, "Resultado investigativo (el porcentaje se calcula solo)")
        hint_label(c, "Edita la columna CASOS · El porcentaje se calcula solo")
        cols   = ["HOMICIDIOS", "CASOS (01.ENE - HOY)", "% AÑO 2026"]
        widths = [280, 200, 160]
        self.t_homi_c = tree_with_scroll(c, cols, widths, height=7,
                                         editable=[1],
                                         on_change=self._recalc_homi_c)
        for res in RESULTADOS:
            self.t_homi_c.add_row([res, "0", "0,0%"])
        self.t_homi_c.add_total(["TOTAL, GENERAL", "0", "100%"])
        self.t_homi_c.add_fuente(
            "FUENTE: Elaboración propia a partir de las concurrencias BH Repol Vpo", 3)


    def _recalc_homi_c(self):
        data = [r for r in self.t_homi_c.get_children()
                if "total" not in self.t_homi_c.item(r,"tags")
                and "fuente" not in self.t_homi_c.item(r,"tags")]
        total = sum(safe_int(self.t_homi_c.item(r,"values")[1]) for r in data)
        for r in data:
            vals = list(self.t_homi_c.item(r,"values"))
            c = safe_int(vals[1])
            vals[2] = porcentaje(c, total)
            self.t_homi_c.item(r, values=vals)
        for r in self.t_homi_c.get_children():
            if "total" in self.t_homi_c.item(r,"tags"):
                self.t_homi_c.item(r, values=("TOTAL, GENERAL", fmt(total), "100%"))

    # =========================================================================
    # III-D. COMPARATIVO INTERANUAL
    # =========================================================================
    def _build_v_homi_d(self):
        v = self._make_view("homi_d")
        section_title(v, "III-D",
                      f"Comparativo de homicidios {PERIODO} {ANO_ANT}/{ANO_ACT}")

        _, c = card(v, f"Comparativo {PERIODO} {ANO_ANT} vs {ANO_ACT} — Variación % automática")
        hint_label(c, f"Edita las columnas {ANO_ANT} y {ANO_ACT} · La variación se calcula sola")
        cols   = ["HOMICIDIOS", f"{PERIODO} {ANO_ANT}",
                  f"{PERIODO} {ANO_ACT}", f"% AÑO {ANO_ACT}"]
        widths = [220, 160, 160, 160]
        self.t_homi_d = tree_with_scroll(c, cols, widths, height=7,
                                         editable=[1, 2],
                                         on_change=self._recalc_homi_d)
        for cat in HOMI_COMP_CATS:
            self.t_homi_d.add_row([cat, "0", "0", "0%"])
        self.t_homi_d.add_total(["TOTAL, GENERAL", "0", "0", "0%"])
        self.v_semana.trace_add("write", lambda *a: self._actualizar_periodo_homi_d())
        self._actualizar_periodo_homi_d()

    def _actualizar_periodo_homi_d(self):
        from comite_policial.docio.semana import periodo_cerrado
        per = periodo_cerrado(self.v_semana.get()) or PERIODO
        cols = self.t_homi_d["columns"]
        self.t_homi_d.heading(cols[1], text=f"{per} {ANO_ANT}")
        self.t_homi_d.heading(cols[2], text=f"{per} {ANO_ACT}")


    def _recalc_homi_d(self):
        data = [r for r in self.t_homi_d.get_children()
                if "total" not in self.t_homi_d.item(r,"tags")
                and "fuente" not in self.t_homi_d.item(r,"tags")]
        ta, tb = 0, 0
        for r in data:
            vals = list(self.t_homi_d.item(r,"values"))
            a = safe_int(vals[1]); b = safe_int(vals[2])
            ta += a; tb += b
            vals[3] = variacion(a, b)
            self.t_homi_d.item(r, values=vals)
        for r in self.t_homi_d.get_children():
            if "total" in self.t_homi_d.item(r,"tags"):
                self.t_homi_d.item(r, values=(
                    "TOTAL, GENERAL", fmt(ta), fmt(tb), variacion(ta, tb)))

    # =========================================================================
    # IV. MIGRACIONES
    # =========================================================================
    def _build_v_migr(self):
        v = self._make_view("migr")
        section_title(v, "IV", "Migraciones y Policía Internacional")

        _, c = card(v, "Fiscalizaciones por Unidad")
        hint_label(c, "Doble clic para editar · EXTENAS y NACION. DENUNCIADOS admiten varias líneas (Enter agrega línea, Tab o clic fuera guarda)")
        cols   = ["UNIDAD", "FISC.", "FISCALIZADOS", "DENUNCIADOS",
                  "NOTIF. EXPULSIÓN", "DETENIDOS", "EXTENAS", "NACION. DENUNCIADOS"]
        widths = [110, 70, 115, 115, 140, 95, 150, 265]
        self.t_migr = tree_with_scroll(c, cols, widths, height=9,
                                       editable=[1, 2, 3, 4, 5, 6, 7],
                                       multiline=[6, 7],
                                       on_change=self._recalc_migr)
        for u in MIGR_UNITS:
            self.t_migr.add_row([u, "0", "0", "0", "0", "0", "0", ""])
        self.t_migr.add_total(["TOTAL", "0", "0", "0", "0", "0", "0", ""])


    def _recalc_migr(self):
        import re as _re
        def _sum_col(r, i):
            val = str(self.t_migr.item(r, "values")[i])
            if i == 6 and not val.strip().replace(".", "").isdigit():
                return sum(int(m) for m in _re.findall(r"\d+", val))
            return safe_int(val)
        data = [r for r in self.t_migr.get_children()
                if "total" not in self.t_migr.item(r,"tags")]
        tots = [sum(_sum_col(r, i) for r in data)
                for i in range(1, 7)]
        for r in self.t_migr.get_children():
            if "total" in self.t_migr.item(r,"tags"):
                # la última celda del total (nacionalidades) no se calcula: se respeta
                ult = self.t_migr.item(r, "values")[7]
                self.t_migr.item(r, values=(
                    "TOTAL", fmt(tots[0]), fmt(tots[1]), fmt(tots[2]),
                    fmt(tots[3]), fmt(tots[4]), fmt(tots[5]), ult))

    # =========================================================================
    # V. VEHÍCULOS
    # =========================================================================
    def _build_v_vehic(self):
        v = self._make_view("vehic")
        section_title(v, "V", "Vehículos incautados y recuperados",
                      "Región Policial Valparaíso")

        _, c1 = card(v, "Por Estado")
        hint_label(c1, "Doble clic para editar")
        cols = ["ESTADO", "SEMANA", "2026 A LA FECHA"]
        wid  = [280, 160, 180]
        self.t_veh_e = tree_with_scroll(c1, cols, wid, height=4,
                                        editable=[1, 2],
                                        on_change=self._recalc_veh_e)
        for e in VEH_ESTADO:
            self.t_veh_e.add_row([e, "0", "0"])
        self.t_veh_e.add_total(["TOTAL", "0", "0"])
        self.t_veh_e.add_fuente("FUENTE: UU.DD REPOL VPO", 3)

        _, c2 = card(v, "Por Tipo de Vehículo")
        hint_label(c2, "Doble clic para editar")
        cols2 = ["TIPO DE VEHÍCULO", "SEMANA", "2026 A LA FECHA"]
        self.t_veh_t = tree_with_scroll(c2, cols2, wid, height=8,
                                        editable=[1, 2],
                                        on_change=self._recalc_veh_t)
        for t in VEH_TIPOS:
            self.t_veh_t.add_row([t, "0", "0"])
        self.t_veh_t.add_total(["TOTAL", "0", "0"])
        self.t_veh_t.add_fuente("FUENTE: UU.DD REPOL VPO", 3)

    def _recalc_veh_e(self):
        self._recalc_simple(self.t_veh_e, "TOTAL", [1, 2])

    def _recalc_veh_t(self):
        self._recalc_simple(self.t_veh_t, "TOTAL", [1, 2])

    # =========================================================================
    # VI. BRISEX
    # =========================================================================
    def _build_v_brisex(self):
        v = self._make_view("brisex")
        section_title(v, "VI", "Delitos sexuales (Brisex)")

        _, c = card(v, "Brigadas de Delitos Sexuales")
        hint_label(c, "Doble clic para editar")
        cols   = ["DELITOS SEXUALES", "SEMANA", "AÑO 2026 A LA FECHA"]
        widths = [280, 160, 200]
        self.t_brisex = tree_with_scroll(c, cols, widths, height=6,
                                         editable=[1, 2],
                                         on_change=lambda: self._recalc_simple(
                                             self.t_brisex, "TOTAL", [1, 2]))
        for b in BRISEX_UNITS:
            self.t_brisex.add_row([b, "0", "0"])
        self.t_brisex.add_total(["TOTAL", "0", "0"])
        self.t_brisex.add_fuente("FUENTE: UU.DD", 3)

    # =========================================================================
    # VII. BANDAS
    # =========================================================================
    def _build_v_bandas(self):
        v = self._make_view("bandas")
        section_title(v, "VII", "Bandas desarticuladas")

        _, c = card(v, "Registro de Bandas por Prefectura")
        hint_label(c, "Doble clic para editar · La columna DELITO admite texto libre")
        cols   = ["PREFECTURA", "CANTIDAD", "INTEGRANTES", "DELITO"]
        widths = [160, 100, 120, 420]
        self.t_bandas = tree_with_scroll(c, cols, widths, height=7,
                                         editable=[1, 2, 3], multiline=[3])
        for p in BANDAS_PREFS:
            self.t_bandas.add_row([p, "0", "0", ""])

    # =========================================================================
    # VIII-A. DETENIDOS POR DELITO (filas dinámicas)
    # =========================================================================
    def _build_v_det_a(self):
        v = self._make_view("det_a")
        section_title(v, "VIII-A", "Detenidos y arrestados por delito")

        # Panel de ingreso
        _, pf = card(v, "Agregar delito")
        grid = tk.Frame(pf, bg=CARD_BG); grid.pack(fill="x", pady=4)

        tk.Label(grid, text="Delito:", bg=CARD_BG, fg=TEXT_DARK,
                 font=(FONT, 10)).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self._v_del_d  = tk.StringVar()
        tk.Entry(grid, textvariable=self._v_del_d, width=52,
                 font=(FONT, 10)).grid(
            row=0, column=1, padx=6, pady=4)

        tk.Label(grid, text="Arrestado:", bg=CARD_BG, fg=TEXT_DARK,
                 font=(FONT, 10)).grid(row=0, column=2, sticky="w", padx=(12, 4))
        self._v_del_a  = tk.StringVar(value="0")
        tk.Entry(grid, textvariable=self._v_del_a, width=7,
                 font=(FONT_MONO, 11), justify="center").grid(row=0, column=3, padx=4)

        tk.Label(grid, text="Detenido:", bg=CARD_BG, fg=TEXT_DARK,
                 font=(FONT, 10)).grid(row=0, column=4, sticky="w", padx=(12, 4))
        self._v_del_dt = tk.StringVar(value="0")
        tk.Entry(grid, textvariable=self._v_del_dt, width=7,
                 font=(FONT_MONO, 11), justify="center").grid(row=0, column=5, padx=4)

        brow = tk.Frame(pf, bg=CARD_BG); brow.pack(anchor="w", pady=(8, 0))
        Boton(brow, "Agregar fila",
              self._add_det_a, "success").pack(side="left", padx=(0, 8))
        Boton(brow, "Eliminar seleccionado",
              self._del_det_a, "danger").pack(side="left")

        # Tabla
        _, c = card(v, "Registro de Detenidos por Delito")
        hint_label(c, "Doble clic para editar filas existentes · Los totales se recalculan automáticamente")
        cols   = ["DELITO", "ARRESTADO", "DETENIDO", "SEMANA"]
        widths = [420, 110, 110, 110]
        self.t_det_a = tree_with_scroll(c, cols, widths, height=12,
                                        editable=[0, 1, 2],
                                        on_change=self._recalc_det_a)
        self.t_det_a.add_total(["TOTAL", "0", "0", "0"])
        self.t_det_a.add_fuente("FUENTE: Brain", 4)

    def _add_det_a(self):
        delito = self._v_del_d.get().strip()
        if not delito:
            messagebox.showwarning("Datos incompletos", "Ingrese el nombre del delito.")
            return
        arr = safe_int(self._v_del_a.get())
        det = safe_int(self._v_del_dt.get())
        # Insertar antes de la fila TOTAL
        children = list(self.t_det_a.get_children())
        total_row = next((r for r in children
                          if "total" in self.t_det_a.item(r,"tags")), None)
        idx = children.index(total_row) if total_row else "end"
        self.t_det_a.insert("", idx, values=(delito, fmt(arr), fmt(det), fmt(arr+det)))
        self._v_del_d.set(""); self._v_del_a.set("0"); self._v_del_dt.set("0")
        self._recalc_det_a()
        self._doc_cambio()

    def _del_det_a(self):
        sel = self.t_det_a.selection()
        if not sel: return
        for r in sel:
            if "total" not in self.t_det_a.item(r,"tags") and \
               "fuente" not in self.t_det_a.item(r,"tags"):
                self.t_det_a.delete(r)
        self._recalc_det_a()
        self._doc_cambio()

    def _recalc_det_a(self):
        data = [r for r in self.t_det_a.get_children()
                if "total" not in self.t_det_a.item(r,"tags")
                and "fuente" not in self.t_det_a.item(r,"tags")]
        # Actualizar columna total de cada fila
        ta, td = 0, 0
        for r in data:
            vals = list(self.t_det_a.item(r,"values"))
            a = safe_int(vals[1]); d = safe_int(vals[2])
            vals[3] = fmt(a + d)
            self.t_det_a.item(r, values=vals)
            ta += a; td += d
        for r in self.t_det_a.get_children():
            if "total" in self.t_det_a.item(r,"tags"):
                self.t_det_a.item(r, values=("TOTAL", fmt(ta), fmt(td), fmt(ta+td)))

    # =========================================================================
    # VIII-B. DETENIDOS POR PREFECTURA — SEMANA
    # =========================================================================
    def _build_v_det_b(self):
        v = self._make_view("det_b")
        section_title(v, "VIII-B", "Detenidos y arrestados por prefectura")

        _, c = card(v, "Semana — Arrestados y Detenidos por Prefectura")
        hint_label(c, "Doble clic para editar · Total semana = Arrestado + Detenido")
        cols   = ["PREFECTURA", "ARRESTADO", "DETENIDO", "SEMANA"]
        widths = [240, 130, 130, 160]
        self.t_det_b = tree_with_scroll(c, cols, widths, height=9,
                                        editable=[1, 2],
                                        on_change=self._recalc_det_b)
        for p in DET_PREFS:
            self.t_det_b.add_row([p, "0", "0", "0"])
        self.t_det_b.add_total(["TOTAL", "0", "0", "0"])
        self.t_det_b.add_fuente("FUENTE: Brain", 4)


    def _recalc_det_b(self):
        data = [r for r in self.t_det_b.get_children()
                if "total" not in self.t_det_b.item(r,"tags")
                and "fuente" not in self.t_det_b.item(r,"tags")]
        ta, td = 0, 0
        for r in data:
            vals = list(self.t_det_b.item(r,"values"))
            a = safe_int(vals[1]); d = safe_int(vals[2])
            vals[3] = fmt(a + d)
            self.t_det_b.item(r, values=vals)
            ta += a; td += d
        for r in self.t_det_b.get_children():
            if "total" in self.t_det_b.item(r,"tags"):
                self.t_det_b.item(r, values=("TOTAL", fmt(ta), fmt(td), fmt(ta+td)))

    # =========================================================================
    # VIII-C. DETENIDOS ACUMULADO AÑO
    # =========================================================================
    def _build_v_det_c(self):
        v = self._make_view("det_c")
        section_title(v, "VIII-C",
                      f"Detenidos {ANO_ACT} a la fecha, por prefectura")

        _, c = card(v, f"Acumulado {ANO_ACT} — Arrestados y Detenidos por Prefectura")
        hint_label(c, "Doble clic para editar")
        cols   = ["PREFECTURA", "ARRESTADO", "DETENIDO", f"AÑO {ANO_ACT} A LA FECHA"]
        widths = [240, 130, 130, 200]
        self.t_det_c = tree_with_scroll(c, cols, widths, height=9,
                                        editable=[1, 2],
                                        on_change=self._recalc_det_c)
        for p in DET_PREFS:
            self.t_det_c.add_row([p, "0", "0", "0"])
        self.t_det_c.add_total(["TOTAL", "0", "0", "0"])
        self.t_det_c.add_fuente("FUENTE: Brain", 4)


    def _recalc_det_c(self):
        data = [r for r in self.t_det_c.get_children()
                if "total" not in self.t_det_c.item(r,"tags")
                and "fuente" not in self.t_det_c.item(r,"tags")]
        ta, td = 0, 0
        for r in data:
            vals = list(self.t_det_c.item(r,"values"))
            a = safe_int(vals[1]); d = safe_int(vals[2])
            vals[3] = fmt(a + d)
            self.t_det_c.item(r, values=vals)
            ta += a; td += d
        for r in self.t_det_c.get_children():
            if "total" in self.t_det_c.item(r,"tags"):
                self.t_det_c.item(r, values=("TOTAL", fmt(ta), fmt(td), fmt(ta+td)))

    # =========================================================================
    # UTILIDAD GENÉRICA DE RECÁLCULO
    # =========================================================================
    def _recalc_simple(self, tree, total_label, col_indices):
        data = [r for r in tree.get_children()
                if "total" not in tree.item(r,"tags")
                and "fuente" not in tree.item(r,"tags")]
        tots = [sum(safe_int(tree.item(r,"values")[i]) for r in data)
                for i in col_indices]
        for r in tree.get_children():
            if "total" in tree.item(r,"tags"):
                vals = list(tree.item(r,"values"))
                vals[0] = total_label
                for j, ci in enumerate(col_indices):
                    vals[ci] = fmt(tots[j])
                tree.item(r, values=vals)

    # =========================================================================
    # LECTURA DE DATOS DESDE TREEVIEW (para exportación)
    # =========================================================================
