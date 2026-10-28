"""Vista IX — Concurrencias y Diligencias Relevantes (recuadros de 2 columnas)."""
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.fechas import fecha_corta
from comite_policial.core.recuadros import (
    TIPO_DEFECTO, TIPO_IMPORTADO, CLAVES_META, nuevo_recuadro, campos_def,
    grupos_izquierda, parrafos_derecha, resumen)
from comite_policial.ui.widgets import *
from comite_policial.ui.calendarios import CalendarPopup


CASILLA_SI, CASILLA_NO = "Sí", "No"


class RecuadrosMixin:
    """Lista, editor y vista previa de los recuadros de la sección IX."""

    # =========================================================================
    # CONSTRUCCIÓN
    # =========================================================================
    def _build_v_recuadros(self):
        self.recuadros: list[dict] = []
        self._rec_idx = None            # índice en edición (None = nuevo)
        self._rec_widgets = {}          # etiqueta -> (widget, clase)
        self._rec_defs_imp = []         # campos del último recuadro importado
        self._rec_fotos: list[str] = []  # fotos del recuadro (solo para el PowerPoint)
        self._rec_after = None

        v = self._make_view("recuadros")
        section_title(v, "IX", "Concurrencias y diligencias relevantes",
                      "una lámina por recuadro en el PowerPoint")

        # ── Lista ────────────────────────────────────────────────────────────
        _, c = card(v, "Recuadros registrados")
        hint_label(c, "Salen en este orden. «PPT» indica si el recuadro se exporta al "
                      "PowerPoint (clic en la columna para cambiarlo). Doble clic para editar.")
        cols = ["PPT", "N°", "UNIDAD", "TIPO", "FECHA", "LUGAR", "DETENIDOS"]
        widths = [50, 44, 300, 230, 90, 170, 100]
        self.t_rec = tree_with_scroll(c, cols, widths, height=8, min_rows=4,
                                      editable=[99])
        self.t_rec.bind("<Button-1>", self._rec_click_casilla, add="+")
        self.t_rec.bind("<Double-1>", self._rec_doble_clic, add="+")
        self.v_rec_pptx = tk.StringVar()
        tk.Label(c, textvariable=self.v_rec_pptx, bg=SURFACE, fg=INK_MID,
                 font=(FONT, SIZE_SMALL, "bold"), anchor="w").pack(fill="x")

        bar = tk.Frame(c, bg=SURFACE)
        bar.pack(fill="x", pady=(10, 0))
        for txt, cmd, tipo in [
            ("Nuevo", self._rec_nuevo, "primary"),
            ("Editar", self._rec_editar, "secondary"),
            ("Duplicar", self._rec_duplicar, "secondary"),
            ("Subir", lambda: self._rec_mover(-1), "secondary"),
            ("Bajar", lambda: self._rec_mover(1), "secondary"),
            ("Eliminar", self._rec_eliminar, "danger"),
        ]:
            Boton(bar, txt, cmd, tipo, chico=True).pack(side="left", padx=(0, 6))
        Boton(bar, "Ninguna al PPT", lambda: self._rec_marcar_todas(False), "ghost",
              chico=True).pack(side="right")
        Boton(bar, "Todas al PPT", lambda: self._rec_marcar_todas(True), "ghost",
              chico=True).pack(side="right", padx=(0, 6))

        # ── Editor ───────────────────────────────────────────────────────────
        self._rec_card_title = tk.StringVar(value="Nuevo recuadro")
        borde = tk.Frame(v, bg=LINE)
        borde.pack(fill="x", padx=20, pady=(0, 12))
        e = tk.Frame(borde, bg=SURFACE)
        e.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(e, textvariable=self._rec_card_title, bg=SURFACE, fg=NAVY,
                 font=(FONT, SIZE_BASE + 1, "bold"), anchor="w", padx=16, pady=11
                 ).pack(fill="x")
        tk.Frame(e, bg=LINE, height=1).pack(fill="x")
        body = tk.Frame(e, bg=SURFACE, padx=16, pady=14)
        body.pack(fill="x")
        body.grid_columnconfigure(1, weight=1)

        self.rv_tipo = tk.StringVar(value=TIPO_DEFECTO)
        self.rv_unidad = tk.StringVar()
        self.rv_grupo = tk.StringVar()
        self.rv_det = tk.StringVar()

        # Columna izquierda: encabezado del recuadro
        izq = tk.Frame(body, bg=SURFACE)
        izq.grid(row=0, column=0, sticky="n", padx=(0, 28))
        tk.Label(izq, text="Encabezado", bg=SURFACE, fg=INK, anchor="w",
                 font=(FONT, SIZE_BASE, "bold")).pack(fill="x", pady=(0, 8))

        self._rec_lbl(izq, "Tipo de procedimiento")
        cb = ttk.Combobox(izq, textvariable=self.rv_tipo, state="readonly",
                          values=list(TIPOS_RECUADRO) + [TIPO_IMPORTADO], width=40,
                          font=(FONT, SIZE_BASE))
        cb.pack(anchor="w", pady=(2, 10))
        cb.bind("<<ComboboxSelected>>", lambda e: self._rec_cambiar_tipo())

        self._rec_lbl(izq, "Unidad / Brigada")
        ttk.Combobox(izq, textvariable=self.rv_unidad, values=UNIDADES_RECUADRO,
                     width=40, font=(FONT, SIZE_BASE)).pack(anchor="w", pady=(2, 10))

        fila = tk.Frame(izq, bg=SURFACE)
        fila.pack(anchor="w", fill="x", pady=(0, 10))
        c1 = tk.Frame(fila, bg=SURFACE)
        c1.pack(side="left", padx=(0, 12))
        self._rec_lbl(c1, "Grupo (opcional)")
        ttk.Combobox(c1, textvariable=self.rv_grupo, values=GRUPOS_RECUADRO,
                     width=22, font=(FONT, SIZE_BASE)).pack(pady=(2, 0))
        c2 = tk.Frame(fila, bg=SURFACE)
        c2.pack(side="left")
        self._rec_lbl(c2, "N.º detenidos")
        tk.Entry(c2, textvariable=self.rv_det, width=8, justify="center"
                 ).pack(pady=(2, 0), ipady=4)

        self._rec_lbl(izq, "Título o delito (una línea por renglón)")
        self.rv_titulo = tk.Text(izq, width=44, height=4, wrap="word", padx=6, pady=4)
        self.rv_titulo.pack(anchor="w", pady=(2, 10))

        self._rec_lbl(izq, "Fotos para el PowerPoint (opcional)")
        fotos = tk.Frame(izq, bg=SURFACE)
        fotos.pack(anchor="w", fill="x", pady=(4, 0))
        Boton(fotos, "Agregar fotos", self._rec_fotos_agregar, "secondary",
              chico=True).pack(side="left")
        Boton(fotos, "Quitar", self._rec_fotos_quitar, "ghost",
              chico=True).pack(side="left", padx=6)
        self._rec_fotos_txt = tk.StringVar()
        tk.Label(fotos, textvariable=self._rec_fotos_txt, bg=SURFACE, fg=INK_MID,
                 font=(FONT, SIZE_SMALL), anchor="w").pack(side="left", padx=4)

        # Columna derecha: campos dinámicos según el tipo
        der = tk.Frame(body, bg=SURFACE)
        der.grid(row=0, column=1, sticky="nsew")
        tk.Label(der, text="Detalle", bg=SURFACE, fg=INK, anchor="w",
                 font=(FONT, SIZE_BASE, "bold")).pack(fill="x", pady=(0, 8))
        self._rec_campos_frame = tk.Frame(der, bg=SURFACE)
        self._rec_campos_frame.pack(fill="both", expand=True)

        eb = tk.Frame(e, bg=SURFACE)
        eb.pack(fill="x", padx=16, pady=(0, 16))
        Boton(eb, "Guardar recuadro", self._rec_guardar, "success").pack(side="left")
        Boton(eb, "Limpiar formulario", self._rec_nuevo, "secondary"
              ).pack(side="left", padx=8)

        # ── Vista previa ─────────────────────────────────────────────────────
        _, p = card(v, "Vista previa (así se verá en el Word)")
        marco = tk.Frame(p, bg="#000000", padx=2, pady=2)
        marco.pack(fill="x")
        txt = dict(bg="#FFFFFF", fg="#000000", relief="flat", bd=0, wrap="word",
                   padx=10, pady=10, cursor="arrow", highlightthickness=0)
        self._pv_izq = tk.Text(marco, width=26, font=("Arial", 9, "bold"), height=10, **txt)
        self._pv_izq.pack(side="left", fill="y", padx=(0, 2))
        self._pv_der = tk.Text(marco, font=("Arial", 9), height=10, **txt)
        self._pv_der.pack(side="left", fill="both", expand=True)
        self._pv_izq.tag_configure("c", justify="center")
        self._pv_der.tag_configure("b", font=("Arial", 9, "bold"))
        self._pv_der.tag_configure("j", justify="left")
        # al mostrarse la vista el ancho ya es real: se recalcula el alto
        self._pv_der.bind("<Map>", lambda e: self._rec_preview_later())

        for var in (self.rv_unidad, self.rv_grupo, self.rv_det):
            var.trace_add("write", lambda *a: self._rec_preview_later())
        self.rv_titulo.bind("<KeyRelease>", lambda e: self._rec_preview_later())

        self._rec_fotos_refrescar()
        self._rec_construir_campos({}, fecha_corta())
        self._rec_preview()
        self._rec_actualizar_contador()

    # =========================================================================
    # CAMPOS DINÁMICOS
    # =========================================================================
    @staticmethod
    def _rec_lbl(parent, texto):
        etiqueta(parent, texto).pack(anchor="w")

    def _rec_construir_campos(self, valores: dict, fecha_def=""):
        for w in self._rec_campos_frame.winfo_children():
            w.destroy()
        self._rec_widgets = {}
        cf = self._rec_campos_frame
        cf.grid_columnconfigure(1, weight=1)
        for i, (nombre, clase) in enumerate(self._rec_defs_actuales()):
            tk.Label(cf, text=nombre, bg=SURFACE, fg=INK, font=(FONT, SIZE_SMALL, "bold"),
                     anchor="ne", justify="right", width=20, wraplength=150
                     ).grid(row=i, column=0, sticky="ne", padx=(0, 10), pady=(7, 0))
            val = valores.get(nombre, fecha_def if nombre == "Fecha" else "")
            if clase == "linea":
                ent = tk.Entry(cf)
                ent.insert(0, val)
                ent.grid(row=i, column=1, sticky="ew", pady=3, ipady=4)
                ent.bind("<KeyRelease>", lambda e: self._rec_preview_later())
                if nombre == "Fecha":
                    Boton(cf, "Calendario", lambda en=ent: self._rec_elegir_fecha(en),
                          "secondary", chico=True).grid(row=i, column=2, padx=(8, 0))
                self._rec_widgets[nombre] = (ent, clase)
            else:
                txt = tk.Text(cf, height=3, wrap="word", padx=6, pady=4)
                txt.insert("1.0", val)
                txt.grid(row=i, column=1, sticky="ew", pady=3)
                txt.bind("<KeyRelease>", lambda e: self._rec_preview_later())
                self._rec_widgets[nombre] = (txt, clase)
                if clase == "lista":
                    tk.Label(cf, text="un ítem\npor línea", bg=SURFACE, fg=INK_SOFT,
                             font=(FONT, 8), justify="left"
                             ).grid(row=i, column=2, padx=(8, 0), sticky="w")

    def _rec_fotos_refrescar(self):
        n = len(self._rec_fotos)
        self._rec_fotos_txt.set("Sin fotos" if not n else
                                f"{n} foto" + ("" if n == 1 else "s"))

    def _rec_fotos_agregar(self):
        rutas = filedialog.askopenfilenames(
            title="Fotos del recuadro",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"),
                       ("Todos", "*.*")])
        self._rec_fotos.extend(r for r in rutas if r not in self._rec_fotos)
        self._rec_fotos_refrescar()

    def _rec_fotos_quitar(self):
        self._rec_fotos = []
        self._rec_fotos_refrescar()

    def _rec_defs_actuales(self) -> list:
        if self.rv_tipo.get() == TIPO_IMPORTADO:
            return [tuple(d) for d in self._rec_defs_imp]
        return TIPOS_RECUADRO[self.rv_tipo.get()]

    def _rec_valores_campos(self) -> dict:
        out = {}
        for etiqueta, (w, clase) in self._rec_widgets.items():
            out[etiqueta] = (w.get() if clase == "linea"
                             else w.get("1.0", "end-1c")).strip()
        return out

    def _rec_cambiar_tipo(self):
        """Cambia los campos conservando lo ya escrito en los que coinciden."""
        previos = self._rec_valores_campos()
        if self.rv_tipo.get() == TIPO_IMPORTADO and not self._rec_defs_imp:
            self.rv_tipo.set(TIPO_DEFECTO)
        self._rec_construir_campos(previos)
        self._rec_preview()

    def _rec_elegir_fecha(self, entry):
        var = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        inicial = var.get()
        pop = CalendarPopup(self, var, "Fecha del procedimiento")
        self.wait_window(pop)
        if var.get() != inicial:
            try:
                d = datetime.datetime.strptime(var.get(), "%d/%m/%Y").date()
                entry.delete(0, "end")
                entry.insert(0, fecha_corta(d))
                self._rec_preview()
            except ValueError:
                pass

    # =========================================================================
    # VISTA PREVIA
    # =========================================================================
    def _rec_preview_later(self):
        if self._rec_after:
            self.after_cancel(self._rec_after)
        self._rec_after = self.after(150, self._rec_preview)

    def _rec_preview(self):
        self._rec_after = None
        rec = self._rec_leer_form()
        izq, der = self._pv_izq, self._pv_der
        for t in (izq, der):
            t.configure(state="normal")
            t.delete("1.0", "end")
        # Izquierda: bloques separados por línea en blanco, centrados
        for i, bloque in enumerate(grupos_izquierda(rec)):
            if i:
                izq.insert("end", "\n\n")
            izq.insert("end", "\n".join(bloque), "c")
        # Derecha: párrafos con etiquetas en negrita
        for i, par in enumerate(parrafos_derecha(rec)):
            if i:
                der.insert("end", "\n")
            for texto, negrita in par:
                der.insert("end", texto, "b" if negrita else "")
        # Alto según el contenido, pero acotado: con la vista oculta el ancho es
        # 1 px y cada letra contaría como una línea (el widget crecía sin límite).
        if der.winfo_ismapped() and der.winfo_width() > 50:
            self.update_idletasks()
            alturas = []
            for t in (izq, der):
                n = t.count("1.0", "end", "displaylines")
                alturas.append(n[0] if n else 1)
            h = max(6, min(max(alturas) + 1, 24))
        else:
            h = 8
        for t in (izq, der):
            t.configure(height=h, state="disabled")

    # =========================================================================
    # CRUD
    # =========================================================================
    def _rec_leer_form(self) -> dict:
        rec = {
            "tipo": self.rv_tipo.get(),
            "unidad": self.rv_unidad.get().strip(),
            "grupo": self.rv_grupo.get().strip(),
            "detenidos": self.rv_det.get().strip(),
            "titulo": self.rv_titulo.get("1.0", "end-1c").strip(),
            "campos": self._rec_valores_campos(),
            "fotos": list(self._rec_fotos),
        }
        if rec["tipo"] == TIPO_IMPORTADO:
            rec["defs"] = [list(d) for d in self._rec_defs_imp]
        return rec

    def _rec_cargar_form(self, rec: dict):
        if rec.get("defs"):
            self._rec_defs_imp = [tuple(d) for d in rec["defs"]]
        self.rv_tipo.set(rec["tipo"])
        self.rv_unidad.set(rec["unidad"])
        self.rv_grupo.set(rec["grupo"])
        self.rv_det.set(rec["detenidos"])
        self.rv_titulo.delete("1.0", "end")
        self.rv_titulo.insert("1.0", rec["titulo"])
        self._rec_fotos = list(rec.get("fotos", []))
        self._rec_fotos_refrescar()
        self._rec_construir_campos(rec["campos"])
        self._rec_preview()

    def _rec_refrescar_lista(self, seleccionar=None):
        self.t_rec.clear()
        for i, rec in enumerate(self.recuadros, 1):
            self.t_rec.add_row([CASILLA_SI if rec.get("en_pptx", True) else CASILLA_NO,
                                i, *resumen(rec)])
        if seleccionar is not None and 0 <= seleccionar < len(self.recuadros):
            iid = self.t_rec.get_children()[seleccionar]
            self.t_rec.selection_set(iid)
            self.t_rec.see(iid)
        self._rec_actualizar_contador()

    # ── Selección de recuadros para el PowerPoint ────────────────────────────
    def _rec_actualizar_contador(self):
        n = sum(1 for r in self.recuadros if r.get("en_pptx", True))
        total = len(self.recuadros)
        self.v_rec_pptx.set("Aún no hay recuadros: completa el formulario y guarda el primero."
                            if not total else
                            f"Se exportarán {n} de {total} recuadros al PowerPoint.")

    def _rec_click_casilla(self, event):
        """Un clic en la columna PPT marca/desmarca el recuadro."""
        if (self.t_rec.identify_region(event.x, event.y) != "cell"
                or self.t_rec.identify_column(event.x) != "#1"):
            return None
        iid = self.t_rec.identify_row(event.y)
        if not iid:
            return None
        i = list(self.t_rec.get_children()).index(iid)
        rec = self.recuadros[i]
        rec["en_pptx"] = not rec.get("en_pptx", True)
        vals = self.t_rec.raw_values(iid)
        vals[0] = CASILLA_SI if rec["en_pptx"] else CASILLA_NO
        self.t_rec.item(iid, values=vals)
        self._rec_actualizar_contador()
        return "break"

    def _rec_doble_clic(self, event):
        if self.t_rec.identify_column(event.x) == "#1":
            return "break"
        self._rec_editar()

    def _rec_marcar_todas(self, valor: bool):
        for rec in self.recuadros:
            rec["en_pptx"] = valor
        self._rec_refrescar_lista(self._rec_seleccion())

    def _rec_seleccion(self):
        sel = self.t_rec.selection()
        if not sel:
            return None
        return list(self.t_rec.get_children()).index(sel[0])

    def _rec_nuevo(self):
        self._rec_idx = None
        self._rec_card_title.set("Nuevo recuadro")
        self.rv_tipo.set(TIPO_DEFECTO)
        self.rv_unidad.set(""); self.rv_grupo.set(""); self.rv_det.set("")
        self.rv_titulo.delete("1.0", "end")
        self._rec_fotos = []
        self._rec_fotos_refrescar()
        self._rec_construir_campos({}, fecha_corta())
        self._rec_preview()

    def _rec_guardar(self):
        rec = self._rec_leer_form()
        if not rec["unidad"]:
            messagebox.showwarning("Falta la unidad",
                                   "Indica la unidad o brigada del recuadro.")
            return
        if self._rec_idx is None:
            self.recuadros.append(rec)
            idx = len(self.recuadros) - 1
        else:
            previo = self.recuadros[self._rec_idx]
            for k in CLAVES_META + ("en_pptx",):    # conserva vínculo con el documento
                if k in previo:
                    rec[k] = previo[k]
            self.recuadros[self._rec_idx] = rec
            idx = self._rec_idx
        self._rec_refrescar_lista(idx)
        self._doc_cambio()
        self.v_status.set(f"Recuadro guardado ({len(self.recuadros)} en total)")
        self._rec_nuevo()

    def _rec_editar(self):
        i = self._rec_seleccion()
        if i is None:
            return
        self._rec_idx = i
        self._rec_card_title.set(f"Editando el recuadro N° {i + 1}")
        self._rec_cargar_form(self.recuadros[i])

    def _rec_duplicar(self):
        i = self._rec_seleccion()
        if i is None:
            return
        import copy
        nuevo = copy.deepcopy(self.recuadros[i])
        origen = nuevo.get("_orig") if nuevo.get("_orig") is not None else nuevo.get("_clone_of")
        nuevo["_orig"], nuevo["_snap"], nuevo["_clone_of"] = None, None, origen
        self.recuadros.insert(i + 1, nuevo)
        self._rec_refrescar_lista(i + 1)
        self._doc_cambio()

    def _rec_eliminar(self):
        i = self._rec_seleccion()
        if i is None:
            return
        if not messagebox.askyesno("Eliminar recuadro",
                                   f"¿Eliminar el recuadro N° {i + 1}?"):
            return
        del self.recuadros[i]
        if self._rec_idx == i:
            self._rec_nuevo()
        elif self._rec_idx is not None and self._rec_idx > i:
            self._rec_idx -= 1
        self._rec_refrescar_lista(min(i, len(self.recuadros) - 1))
        self._doc_cambio()

    def _rec_mover(self, d):
        i = self._rec_seleccion()
        if i is None:
            return
        j = i + d
        if not 0 <= j < len(self.recuadros):
            return
        self.recuadros[i], self.recuadros[j] = self.recuadros[j], self.recuadros[i]
        if self._rec_idx == i:
            self._rec_idx = j
        elif self._rec_idx == j:
            self._rec_idx = i
        self._rec_refrescar_lista(j)
        self._doc_cambio()
