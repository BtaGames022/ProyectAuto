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
        section_title(v, "IX", "CONCURRENCIAS Y DILIGENCIAS RELEVANTES",
                      "recuadros del informe")

        # ── Lista ────────────────────────────────────────────────────────────
        _, c = card(v, "Recuadros registrados (en el orden en que saldrán en el informe)")
        hint_label(c, "Selecciona uno y pulsa «Editar» (o doble clic) para modificarlo")
        cols = ["N°", "UNIDAD", "TIPO", "FECHA", "LUGAR", "DETENIDOS"]
        widths = [45, 330, 260, 100, 200, 110]
        self.t_rec = tree_with_scroll(c, cols, widths, height=5, editable=[99])
        self.t_rec.bind("<Double-1>", lambda e: self._rec_editar(), add="+")

        bar = tk.Frame(c, bg=CARD_BG)
        bar.pack(pady=8)
        for txt, cmd, col in [
            ("✚ Nuevo", self._rec_nuevo, ACCENT),
            ("✏ Editar", self._rec_editar, TEAL),
            ("⧉ Duplicar", self._rec_duplicar, PURPLE),
            ("▲ Subir", lambda: self._rec_mover(-1), TEXT_MID),
            ("▼ Bajar", lambda: self._rec_mover(1), TEXT_MID),
            ("🗑 Eliminar", self._rec_eliminar, DANGER),
        ]:
            action_btn(bar, txt, cmd, col, small=True).pack(side="left", padx=4)

        # ── Editor ───────────────────────────────────────────────────────────
        self._rec_card_title = tk.StringVar(value="Editor del recuadro — NUEVO")
        _, e = card(v, "Editor del recuadro")
        tk.Label(e, textvariable=self._rec_card_title, bg="#eff6ff", fg="#1d4ed8",
                 font=(FONT, 9, "bold"), anchor="w", padx=12, pady=5
                 ).pack(fill="x", padx=12, pady=(10, 0))
        body = tk.Frame(e, bg=CARD_BG)
        body.pack(fill="x", padx=12, pady=10)

        # Columna izquierda: encabezado del recuadro
        izq = tk.LabelFrame(body, text="  Columna izquierda (encabezado)  ",
                            bg=CARD_BG, fg=ACCENT2, font=(FONT, 9, "bold"),
                            padx=10, pady=8)
        izq.pack(side="left", fill="y", padx=(0, 12))

        self.rv_tipo = tk.StringVar(value=TIPO_DEFECTO)
        self.rv_unidad = tk.StringVar()
        self.rv_grupo = tk.StringVar()
        self.rv_det = tk.StringVar()

        self._rec_lbl(izq, "Tipo de procedimiento")
        cb = ttk.Combobox(izq, textvariable=self.rv_tipo, state="readonly",
                          values=list(TIPOS_RECUADRO) + [TIPO_IMPORTADO], width=42, font=(FONT, 9))
        cb.pack(anchor="w", pady=(0, 6))
        cb.bind("<<ComboboxSelected>>", lambda e: self._rec_cambiar_tipo())

        self._rec_lbl(izq, "Unidad / Brigada")
        ttk.Combobox(izq, textvariable=self.rv_unidad, values=UNIDADES_RECUADRO,
                     width=42, font=(FONT, 9)).pack(anchor="w", pady=(0, 6))

        fila = tk.Frame(izq, bg=CARD_BG)
        fila.pack(anchor="w", fill="x", pady=(0, 6))
        c1 = tk.Frame(fila, bg=CARD_BG); c1.pack(side="left", padx=(0, 10))
        self._rec_lbl(c1, "Grupo (opcional)")
        ttk.Combobox(c1, textvariable=self.rv_grupo, values=GRUPOS_RECUADRO,
                     width=18, font=(FONT, 9)).pack()
        c2 = tk.Frame(fila, bg=CARD_BG); c2.pack(side="left")
        self._rec_lbl(c2, "N.º detenidos")
        tk.Entry(c2, textvariable=self.rv_det, width=8, font=(FONT, 10),
                 justify="center", relief="solid", bd=1).pack()

        self._rec_lbl(izq, "Título / delito (una línea por renglón)")
        self.rv_titulo = tk.Text(izq, width=44, height=4, font=(FONT, 9),
                                 relief="solid", bd=1, wrap="word")
        self.rv_titulo.pack(anchor="w")

        self._rec_lbl(izq, "Fotos para el PowerPoint (opcional)")
        fotos = tk.Frame(izq, bg=CARD_BG)
        fotos.pack(anchor="w", fill="x")
        self._rec_fotos_txt = tk.StringVar()
        tk.Label(fotos, textvariable=self._rec_fotos_txt, bg=CARD_BG, fg=TEXT_MID,
                 font=(FONT, 9, "italic"), anchor="w", wraplength=260, justify="left"
                 ).pack(side="left", padx=(0, 8))
        action_btn(fotos, "➕ Agregar", self._rec_fotos_agregar, ACCENT,
                   small=True).pack(side="left", padx=2)
        action_btn(fotos, "Quitar", self._rec_fotos_quitar, TEXT_MID,
                   small=True).pack(side="left", padx=2)

        # Columna derecha: campos dinámicos según el tipo
        der = tk.LabelFrame(body, text="  Columna derecha (detalle)  ",
                            bg=CARD_BG, fg=ACCENT2, font=(FONT, 9, "bold"),
                            padx=10, pady=8)
        der.pack(side="left", fill="both", expand=True)
        self._rec_campos_frame = tk.Frame(der, bg=CARD_BG)
        self._rec_campos_frame.pack(fill="both", expand=True)

        # Botones del editor
        eb = tk.Frame(e, bg=CARD_BG)
        eb.pack(pady=(0, 10))
        action_btn(eb, "💾 Guardar recuadro", self._rec_guardar, SUCCESS).pack(side="left", padx=4)
        action_btn(eb, "↺ Limpiar formulario", self._rec_nuevo, TEXT_MID).pack(side="left", padx=4)

        # ── Vista previa ─────────────────────────────────────────────────────
        _, p = card(v, "Vista previa (así se verá en Word)", accent_color=SUCCESS)
        marco = tk.Frame(p, bg="#000000", padx=2, pady=2)
        marco.pack(fill="x", padx=14, pady=12)
        self._pv_izq = tk.Text(marco, width=26, bg="#ffffff", fg="#000000",
                               font=("Arial", 9, "bold"), relief="flat", bd=0,
                               wrap="word", padx=10, pady=10, cursor="arrow",
                               height=10)
        self._pv_izq.pack(side="left", fill="y", padx=(0, 2))
        self._pv_der = tk.Text(marco, bg="#ffffff", fg="#000000",
                               font=("Arial", 9), relief="flat", bd=0,
                               wrap="word", padx=10, pady=10, cursor="arrow",
                               height=10)
        self._pv_der.pack(side="left", fill="both", expand=True)
        self._pv_izq.tag_configure("c", justify="center")
        self._pv_der.tag_configure("b", font=("Arial", 9, "bold"))
        self._pv_der.tag_configure("j", justify="left")

        for var in (self.rv_unidad, self.rv_grupo, self.rv_det):
            var.trace_add("write", lambda *a: self._rec_preview_later())
        self.rv_titulo.bind("<KeyRelease>", lambda e: self._rec_preview_later())

        self._rec_fotos_refrescar()
        self._rec_construir_campos({}, fecha_corta())
        self._rec_preview()

    # =========================================================================
    # CAMPOS DINÁMICOS
    # =========================================================================
    @staticmethod
    def _rec_lbl(parent, texto):
        tk.Label(parent, text=texto, bg=CARD_BG, fg=TEXT_MID,
                 font=(FONT, 8, "bold")).pack(anchor="w")

    def _rec_construir_campos(self, valores: dict, fecha_def=""):
        for w in self._rec_campos_frame.winfo_children():
            w.destroy()
        self._rec_widgets = {}
        for etiqueta, clase in self._rec_defs_actuales():
            fila = tk.Frame(self._rec_campos_frame, bg=CARD_BG)
            fila.pack(fill="x", pady=2)
            tk.Label(fila, text=etiqueta + ":", bg=CARD_BG, fg=TEXT_DARK,
                     font=(FONT, 9, "bold"), width=19, anchor="nw"
                     ).pack(side="left", anchor="n")
            val = valores.get(etiqueta, fecha_def if etiqueta == "Fecha" else "")
            if clase == "linea":
                ent = tk.Entry(fila, font=(FONT, 10), relief="solid", bd=1)
                ent.insert(0, val)
                ent.pack(side="left", fill="x", expand=True, ipady=2)
                ent.bind("<KeyRelease>", lambda e: self._rec_preview_later())
                if etiqueta == "Fecha":
                    b = action_btn(fila, "🗓", lambda en=ent: self._rec_elegir_fecha(en),
                                   ACCENT, small=True)
                    b.pack(side="left", padx=(6, 0))
                self._rec_widgets[etiqueta] = (ent, clase)
            else:
                alto = 3 if clase == "parrafo" else 3
                txt = tk.Text(fila, font=(FONT, 9), height=alto, wrap="word",
                              relief="solid", bd=1)
                txt.insert("1.0", val)
                txt.pack(side="left", fill="x", expand=True)
                txt.bind("<KeyRelease>", lambda e: self._rec_preview_later())
                self._rec_widgets[etiqueta] = (txt, clase)
                if clase == "lista":
                    tk.Label(fila, text="un ítem\npor línea", bg=CARD_BG,
                             fg=TEXT_LIGHT, font=(FONT, 7)).pack(side="left", padx=4)

    def _rec_fotos_refrescar(self):
        n = len(self._rec_fotos)
        self._rec_fotos_txt.set("sin fotos" if not n else
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
        self.update_idletasks()
        for t in (izq, der):
            n = t.count("1.0", "end", "displaylines")
            lineas = (n[0] if n else 1)
            t.configure(height=max(lineas + 1, 6), state="disabled")
        # igualar alturas
        h = max(int(izq.cget("height")), int(der.cget("height")))
        izq.configure(height=h); der.configure(height=h)

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
            self.t_rec.add_row([i, *resumen(rec)])
        if seleccionar is not None and 0 <= seleccionar < len(self.recuadros):
            iid = self.t_rec.get_children()[seleccionar]
            self.t_rec.selection_set(iid)
            self.t_rec.see(iid)

    def _rec_seleccion(self):
        sel = self.t_rec.selection()
        if not sel:
            return None
        return list(self.t_rec.get_children()).index(sel[0])

    def _rec_nuevo(self):
        self._rec_idx = None
        self._rec_card_title.set("Editor del recuadro — NUEVO")
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
            for k in CLAVES_META:               # conserva el vínculo con el documento
                if k in previo:
                    rec[k] = previo[k]
            self.recuadros[self._rec_idx] = rec
            idx = self._rec_idx
        self._rec_refrescar_lista(idx)
        self._doc_cambio()
        self.v_status.set(f"✅  Recuadro guardado ({len(self.recuadros)} en total)")
        self._rec_nuevo()

    def _rec_editar(self):
        i = self._rec_seleccion()
        if i is None:
            return
        self._rec_idx = i
        self._rec_card_title.set(f"Editor del recuadro — EDITANDO N° {i + 1}")
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
