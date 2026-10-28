"""Widgets reutilizables: tablas editables, tarjetas, botones y títulos."""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.calculos import *

_ESPECIALES = ("total", "fuente", "header_row", "subtotal")


# ──────────────────────────────────────────────────────────────────────────────
# BOTONES
# ──────────────────────────────────────────────────────────────────────────────
# tipo → (fondo, texto, borde, fondo al pasar el mouse)
_BOTONES = {
    "primary":   (NAVY,      "#FFFFFF", NAVY,      NAVY_SOFT),
    "accent":    (GOLD,      NAVY,      GOLD,      "#E6AC00"),
    "success":   (OK,        "#FFFFFF", OK,        "#14633B"),
    "secondary": (SURFACE,   NAVY,      LINE_STRONG, SKY),
    "danger":    (SURFACE,   ERR,       "#E3B3AF", ERR_SOFT),
    "ghost":     (None,      NAVY,      None,      SKY),
    "side":      (NAVY_DEEP, "#FFFFFF", "#3A5192", NAVY_SOFT),
}


class Boton(tk.Frame):
    """Botón plano con borde de 1 px, foco visible y activación con teclado."""

    def __init__(self, parent, texto, comando, tipo="primary", chico=False,
                 ancho=None):
        bg, fg, borde, hover = _BOTONES[tipo]
        fondo_padre = parent.cget("bg") if bg is None else bg
        self._bg = fondo_padre
        self._hover = hover
        self._cmd = comando
        self._habilitado = True
        super().__init__(parent, bg=borde or fondo_padre, bd=0,
                         highlightthickness=2, highlightbackground=parent.cget("bg"),
                         highlightcolor=GOLD, takefocus=1)
        self._lbl = tk.Label(self, text=texto, bg=self._bg, fg=fg,
                             font=(FONT, SIZE_SMALL if chico else SIZE_BASE, "bold"),
                             padx=10 if chico else 16, pady=3 if chico else 7,
                             cursor="hand2", width=ancho or 0)
        self._lbl.pack(fill="both", expand=True, padx=1, pady=1)
        for w in (self, self._lbl):
            w.bind("<Button-1>", self._clic)
            w.bind("<Enter>", self._entra)
            w.bind("<Leave>", self._sale)
        self.bind("<Return>", self._clic)
        self.bind("<space>", self._clic)

    def _clic(self, _e=None):
        if self._habilitado:
            self.focus_set()
            self._cmd()
        return "break"

    def _entra(self, _e=None):
        if self._habilitado:
            self._lbl.configure(bg=self._hover)

    def _sale(self, _e=None):
        self._lbl.configure(bg=self._bg)

    def texto(self, nuevo):
        self._lbl.configure(text=nuevo)

    def habilitar(self, si=True):
        self._habilitado = si
        self._lbl.configure(cursor="hand2" if si else "arrow",
                            fg=_BOTONES_FG_OFF if not si else self._lbl.cget("fg"))


_BOTONES_FG_OFF = INK_SOFT

# colores heredados del código antiguo → tipo de botón
_TIPO_POR_COLOR = {SUCCESS: "success", DANGER: "danger", ACCENT: "primary",
                   TEAL: "secondary", WARNING: "secondary"}


def action_btn(parent, text, cmd, color=ACCENT, small=False):
    """Compatibilidad: acepta el color antiguo y elige el tipo de botón."""
    return Boton(parent, text, cmd, _TIPO_POR_COLOR.get(color, "secondary"), chico=small)


# ──────────────────────────────────────────────────────────────────────────────
# TABLA EDITABLE
# ──────────────────────────────────────────────────────────────────────────────
class AutoScrollbar(ttk.Scrollbar):
    """Barra que solo aparece cuando hay algo que desplazar."""

    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.grid_remove()
        else:
            self.grid()
        super().set(lo, hi)


class EditableTree(ttk.Treeview):
    """
    Tabla con el aspecto de las del informe: encabezado azul institucional,
    filas alternas, fila TOTAL azul con texto blanco y FUENTE en gris tenue.
    • Doble clic o Enter edita la celda (Tab / Enter pasan a la siguiente).
    • La altura se ajusta al número de filas (hasta `height`).
    """

    on_any_change = None  # callback global (sincronización con el documento)
    _estilo_listo = False

    def __init__(self, parent, columns, col_widths=None, height=8,
                 editable_cols=None, on_change=None, min_rows=1, **kw):
        self._estilo()
        super().__init__(parent, columns=columns, show="headings",
                         height=min_rows, style="PDI.Treeview", **kw)
        self._columns = columns
        self._editable_cols = editable_cols or list(range(1, len(columns)))
        self._on_change = on_change
        self._entry_popup = None
        self._hover_item = None
        self._sel_item = None
        self._max_rows = height
        self._min_rows = min_rows
        self._ajuste = None
        self.v_fuente = tk.StringVar()          # la fuente va bajo la tabla, no como fila
        self._pedido = [col_widths[i] if col_widths and i < len(col_widths) else 130
                        for i in range(len(columns))]
        self._ancho_total = 0

        for i, col in enumerate(columns):
            w = col_widths[i] if col_widths and i < len(col_widths) else 130
            anchor = "w" if i == 0 else "center"
            self.heading(col, text=col, anchor=anchor)
            self.column(col, anchor=anchor, width=w, minwidth=40, stretch=True)

        f = (FONT, SIZE_BASE)
        self.tag_configure("odd", background=ROW_ODD, foreground=INK, font=f)
        self.tag_configure("even", background=ROW_EVEN, foreground=INK, font=f)
        self.tag_configure("alt", background=ROW_EVEN)
        self.tag_configure("hover", background=ROW_HOVER_BG, foreground=INK, font=f)
        self.tag_configure("selrow", background=ROW_SEL_BG, foreground=INK,
                           font=(FONT, SIZE_BASE, "bold"))
        self.tag_configure("total", background=ROW_TOTAL_BG, foreground=ROW_TOTAL_FG,
                           font=(FONT, SIZE_BASE, "bold"))
        self.tag_configure("subtotal", background=ROW_SUBTOT_BG, foreground=ROW_SUBTOT_FG,
                           font=(FONT, SIZE_BASE, "bold"))
        self.tag_configure("fuente", background=ROW_FUENTE_BG, foreground=ROW_FUENTE_FG,
                           font=(FONT, SIZE_SMALL, "italic"))
        self.tag_configure("header_row", background=ROW_HEADER_BG, foreground=ROW_HEADER_FG,
                           font=(FONT, SIZE_BASE, "bold"))

        self.bind("<Configure>", self._repartir_columnas, add="+")
        self.bind("<Double-1>", self._on_double_click)
        self.bind("<Return>", self._on_double_click)
        self.bind("<Motion>", self._on_hover)
        self.bind("<Leave>", self._on_leave)
        self.bind("<<TreeviewSelect>>", self._on_select)

    # ── Estilo ttk (una sola vez) ─────────────────────────────────────────────
    @classmethod
    def _estilo(cls):
        if cls._estilo_listo:
            return
        cls._estilo_listo = True
        s = ttk.Style()
        # alto de fila proporcional a la letra (respeta la escala de pantalla)
        raiz = tk._default_root
        px = raiz.winfo_fpixels("1i") / 72 if raiz else 1.333
        alto = int(round(SIZE_BASE * max(px, 1.0) * 2.25))
        s.configure("PDI.Treeview", background=ROW_ODD, foreground=INK,
                    fieldbackground=ROW_ODD, rowheight=alto, borderwidth=0,
                    relief="flat", font=(FONT, SIZE_BASE))
        s.configure("PDI.Treeview.Heading", background=HDR_BG, foreground=HDR_FG,
                    font=HDR_FONT, relief="flat", borderwidth=1,
                    bordercolor=NAVY_SOFT, lightcolor=HDR_BG, darkcolor=HDR_BG,
                    padding=(10, 8))
        s.map("PDI.Treeview.Heading", background=[("active", NAVY_SOFT)])
        s.map("PDI.Treeview", background=[("selected", ROW_SEL_BG)],
              foreground=[("selected", ROW_SEL_FG)])
        s.layout("PDI.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    # ── Columnas proporcionales al ancho disponible ───────────────────────────
    def _repartir_columnas(self, event):
        """Reparte el ancho entre las columnas según el ancho pedido: así la tabla
        siempre cabe en la ventana y no aparece la barra horizontal."""
        ancho = event.width
        if ancho < 120 or ancho == self._ancho_total:
            return
        self._ancho_total = ancho
        total = sum(self._pedido)
        for col, pedido in zip(self._columns, self._pedido):
            self.column(col, width=max(60, int(ancho * pedido / total)))

    # ── Altura automática ─────────────────────────────────────────────────────
    def insert(self, *a, **k):
        r = super().insert(*a, **k)
        self._programar_ajuste()
        return r

    def delete(self, *items):
        super().delete(*items)
        self._programar_ajuste()

    def _programar_ajuste(self):
        if self._ajuste is None:
            self._ajuste = self.after_idle(self._ajustar_altura)

    def _ajustar_altura(self):
        self._ajuste = None
        try:
            n = len(self.get_children())
            self.configure(height=max(self._min_rows, min(n, self._max_rows)))
        except tk.TclError:
            pass

    # ── Selección y hover ─────────────────────────────────────────────────────
    def _on_select(self, event=None):
        """Los tags anulan la selección nativa: se marca a mano."""
        if self._sel_item:
            prev, self._sel_item = self._sel_item, None
            if self.exists(prev):
                self._restore_row_tag(prev)
        sel = self.selection()
        if sel:
            tags = self.item(sel[0], "tags")
            if not any(t in tags for t in _ESPECIALES):
                self._sel_item = sel[0]
                self.item(sel[0], tags=("selrow",))

    def _on_hover(self, event):
        item = self.identify_row(event.y)
        if item == self._hover_item:
            return
        if self._hover_item and self.exists(self._hover_item):
            self._restore_row_tag(self._hover_item)
        self._hover_item = None
        if item:
            tags = self.item(item, "tags")
            if not any(t in tags for t in _ESPECIALES + ("selrow",)):
                self._hover_item = item
                self.item(item, tags=("hover",))

    def _on_leave(self, event):
        if self._hover_item and self.exists(self._hover_item):
            self._restore_row_tag(self._hover_item)
        self._hover_item = None

    def _restore_row_tag(self, item):
        try:
            if item == self._sel_item:
                self.item(item, tags=("selrow",))
                return
            idx = list(self.get_children()).index(item)
            self.item(item, tags=("even" if idx % 2 else "odd",))
        except (ValueError, tk.TclError):
            pass

    # ── API pública ───────────────────────────────────────────────────────────
    def clear(self):
        self.delete(*self.get_children())
        self._hover_item = None
        self._sel_item = None
        self.v_fuente.set("")

    def raw_values(self, item) -> list:
        """Valores de la fila como texto exacto (sin que Tk convierta '01' en 1)."""
        raw = self.tk.call(self._w, "item", item, "-values")
        return [str(v) for v in self.tk.splitlist(raw)]

    def add_row(self, values, tag=""):
        idx = len(self.get_children())
        t = tag if tag else ("even" if idx % 2 else "odd")
        self.insert("", "end", values=values, tags=(t,))

    def add_total(self, values):
        self.insert("", "end", values=values, tags=("total",))

    def add_subtotal(self, values):
        self.insert("", "end", values=values, tags=("subtotal",))

    def add_fuente(self, texto, ncols=0):
        self.v_fuente.set(texto)

    # ── Edición en la celda ───────────────────────────────────────────────────
    def _on_double_click(self, event):
        if event.type == tk.EventType.KeyPress:
            fila = self.focus()
            if not fila or not self._editable_cols:
                return
            return self._abrir_si_editable(fila, self._editable_cols[0])
        if self.identify("region", event.x, event.y) != "cell":
            return
        fila = self.identify_row(event.y)
        col = int(self.identify_column(event.x).replace("#", "")) - 1
        if fila:
            self._abrir_si_editable(fila, col)

    def _abrir_si_editable(self, fila, col):
        if col not in self._editable_cols:
            return
        if any(t in self.item(fila, "tags") for t in ("total", "fuente", "header_row")):
            return
        self._abrir_editor(fila, col)

    def _abrir_editor(self, row_id, col_idx):
        if self._entry_popup:
            try:
                self._entry_popup.destroy()
            except tk.TclError:
                pass
        bbox = self.bbox(row_id, self._columns[col_idx])
        if not bbox:
            return
        x, y, w, h = bbox
        valor_actual = self.raw_values(row_id)[col_idx]

        marco = tk.Frame(self, bg=NAVY, padx=2, pady=2)
        marco.place(x=x - 2, y=y - 2, width=max(w, 90) + 4, height=h + 4)
        popup = tk.Entry(marco, font=(FONT_MONO, 11), justify="center",
                         bg=GOLD_SOFT, fg=INK, relief="flat", bd=0,
                         highlightthickness=0)
        popup.pack(fill="both", expand=True)
        popup.insert(0, str(valor_actual))
        popup.select_range(0, "end")
        popup.focus_set()
        self._entry_popup = marco
        listo = [False]

        def siguiente(dr, dc):
            filas = list(self.get_children())
            cols = sorted(self._editable_cols)
            try:
                r = filas.index(row_id)
            except ValueError:
                return None
            if dc:                                  # Tab: columna siguiente, luego fila
                pos = cols.index(col_idx) if col_idx in cols else -1
                if pos + 1 < len(cols):
                    return filas[r], cols[pos + 1]
                r2, c2 = r + 1, cols[0]
            else:                                   # Enter: fila de abajo
                r2, c2 = r + 1, col_idx
            while r2 < len(filas):
                if not any(t in self.item(filas[r2], "tags") for t in _ESPECIALES):
                    return filas[r2], c2
                r2 += 1
            return None

        def confirmar(e=None, mover=None):
            if listo[0]:
                return
            listo[0] = True
            nuevo = popup.get().strip()
            marco.destroy()
            self._entry_popup = None
            vals = self.raw_values(row_id)
            vals[col_idx] = nuevo
            self.item(row_id, values=vals)
            if self._on_change:
                self._on_change()
            if EditableTree.on_any_change:
                EditableTree.on_any_change()
            if mover:
                sig = siguiente(*mover)
                if sig:
                    self.after(10, lambda: self._abrir_editor(sig[0], sig[1]))

        def cancelar(e=None):
            listo[0] = True
            marco.destroy()
            self._entry_popup = None

        popup.bind("<Return>", lambda e: confirmar(mover=(1, 0)) or "break")
        popup.bind("<Tab>", lambda e: confirmar(mover=(0, 1)) or "break")
        popup.bind("<Escape>", cancelar)
        popup.bind("<FocusOut>", lambda e: confirmar())


# ──────────────────────────────────────────────────────────────────────────────
# ESTRUCTURA
# ──────────────────────────────────────────────────────────────────────────────
def card(parent, title=None, pady=6, accent_color=None):
    """Tarjeta blanca con borde fino. Devuelve (marco, contenido)."""
    borde = tk.Frame(parent, bg=LINE)
    borde.pack(fill="x", padx=20, pady=(0, 12))
    inner = tk.Frame(borde, bg=SURFACE)
    inner.pack(fill="both", expand=True, padx=1, pady=1)

    if title:
        cab = tk.Frame(inner, bg=SURFACE)
        cab.pack(fill="x")
        tk.Label(cab, text=title, bg=SURFACE, fg=NAVY,
                 font=(FONT, SIZE_BASE + 1, "bold"), anchor="w",
                 padx=16, pady=11).pack(side="left")
        tk.Frame(inner, bg=LINE, height=1).pack(fill="x")

    content = tk.Frame(inner, bg=SURFACE, padx=16, pady=12)
    content.pack(fill="both", expand=True)
    return inner, content


def section_title(parent, num, titulo, subtitulo=""):
    """Título de sección: numeral romano, nombre y contexto."""
    f = tk.Frame(parent, bg=BG)
    f.pack(fill="x", padx=20, pady=(22, 12))
    tk.Label(f, text=num, bg=SKY, fg=NAVY, font=(FONT, SIZE_BASE, "bold"),
             padx=9, pady=3).pack(side="left", padx=(0, 12))
    tk.Label(f, text=titulo, bg=BG, fg=NAVY,
             font=(FONT, SIZE_H1, "bold")).pack(side="left")
    if subtitulo:
        tk.Label(f, text=subtitulo, bg=BG, fg=INK_MID,
                 font=(FONT, SIZE_BASE)).pack(side="left", padx=(12, 0), pady=(4, 0))


def hint_label(parent, text):
    """Texto de ayuda breve bajo el título de una tarjeta."""
    tk.Label(parent, text=text, bg=SURFACE, fg=INK_MID, font=(FONT, SIZE_SMALL),
             anchor="w", justify="left", wraplength=900
             ).pack(fill="x", pady=(0, 8))


def etiqueta(parent, texto):
    """Rótulo de un campo de formulario."""
    return tk.Label(parent, text=texto, bg=parent.cget("bg"), fg=INK_MID,
                    font=(FONT, SIZE_SMALL, "bold"), anchor="w")


def tree_with_scroll(parent, cols, widths=None, height=8,
                     editable=None, on_change=None, accent=None,
                     min_rows=1) -> EditableTree:
    """Tabla editable con borde y barras que aparecen solo si hacen falta."""
    borde = tk.Frame(parent, bg=LINE)
    borde.pack(fill="x", pady=(0, 4))
    wrap = tk.Frame(borde, bg=SURFACE)
    wrap.pack(fill="both", expand=True, padx=1, pady=1)

    tree = EditableTree(wrap, cols, widths, height, editable_cols=editable,
                        on_change=on_change, min_rows=min_rows)
    vsb = AutoScrollbar(wrap, orient="vertical", command=tree.yview,
                        style="Thin.Vertical.TScrollbar")
    hsb = AutoScrollbar(wrap, orient="horizontal", command=tree.xview,
                        style="Thin.Horizontal.TScrollbar")
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    fuente = tk.Label(wrap, textvariable=tree.v_fuente, bg=SURFACE, fg=INK_SOFT,
                      font=(FONT, SIZE_SMALL, "italic"), anchor="w", padx=8, pady=5)
    fuente.grid(row=2, column=0, columnspan=2, sticky="ew")
    tree.v_fuente.trace_add("write", lambda *a: fuente.grid() if tree.v_fuente.get()
                            else fuente.grid_remove())
    fuente.grid_remove()
    wrap.grid_rowconfigure(0, weight=1)
    wrap.grid_columnconfigure(0, weight=1)
    return tree
