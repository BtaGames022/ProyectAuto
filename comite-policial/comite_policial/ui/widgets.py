"""Widgets reutilizables: EditableTree, tarjetas, botones y títulos."""
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.calculos import *

_style_done = False   # aplicar estilo global solo una vez

class EditableTree(ttk.Treeview):
    """
    Treeview profesional con:
    • Filas alternas bicolor
    • Fila TOTAL en azul PDI con texto blanco
    • Fila FUENTE en gris tenue itálica
    • Hover resaltado (Motion binding)
    • Edición inline con doble clic — Entry flotante sobre la celda
    • Columna 0 izquierda, resto centrado con fuente monoespaciada
    """

    _instance_count = 0   # id único de estilo por instancia
    on_any_change = None  # callback global (sincronización con el documento)

    def __init__(self, parent, columns, col_widths=None, height=8,
                 editable_cols=None, on_change=None, **kw):
        EditableTree._instance_count += 1
        self._style_name = f"ET{EditableTree._instance_count}.Treeview"

        super().__init__(parent, columns=columns, show="headings",
                         height=height, style=self._style_name, **kw)

        self._columns       = columns
        self._editable_cols = editable_cols or list(range(1, len(columns)))
        self._on_change     = on_change
        self._entry_popup   = None
        self._hover_item    = None

        self._apply_style()

        # Anchos de columna
        for i, col in enumerate(columns):
            w = col_widths[i] if col_widths and i < len(col_widths) else 130
            self.heading(col, text=col)
            anchor = "w" if i == 0 else "center"
            self.column(col, anchor=anchor, width=w, minwidth=40, stretch=True)

        # ── Tags de fila ──────────────────────────────────────────────────────
        self.tag_configure("odd",
            background=ROW_ODD, foreground=TEXT_DARK,
            font=(FONT, 10))
        self.tag_configure("even",
            background=ROW_EVEN, foreground=TEXT_DARK,
            font=(FONT, 10))
        self.tag_configure("hover",
            background=ROW_HOVER_BG, foreground=TEXT_DARK,
            font=(FONT, 10))
        self.tag_configure("total",
            background=ROW_TOTAL_BG, foreground=ROW_TOTAL_FG,
            font=(FONT, 10, "bold"))
        self.tag_configure("subtotal",
            background=ROW_SUBTOT_BG, foreground=ROW_SUBTOT_FG,
            font=(FONT, 10, "bold"))
        self.tag_configure("fuente",
            background=ROW_FUENTE_BG, foreground=ROW_FUENTE_FG,
            font=(FONT, 8, "italic"))
        self.tag_configure("header_row",
            background=ROW_HEADER_BG, foreground=ROW_HEADER_FG,
            font=(FONT, 10, "bold"))
        # alias de compatibilidad con código existente
        self.tag_configure("alt",  background=ROW_EVEN)

        # ── Bindings ──────────────────────────────────────────────────────────
        self.bind("<Double-1>",    self._on_double_click)
        self.bind("<Return>",      self._on_double_click)
        self.bind("<Motion>",      self._on_hover)
        self.bind("<Leave>",       self._on_leave)

    # ── Estilo ttk ────────────────────────────────────────────────────────────
    def _apply_style(self):
        s = ttk.Style()
        sn = self._style_name
        s.configure(sn,
                    background=ROW_ODD,
                    foreground=TEXT_DARK,
                    fieldbackground=ROW_ODD,
                    rowheight=36,
                    borderwidth=0,
                    relief="flat",
                    font=(FONT, 10))
        s.configure(f"{sn}.Heading",
                    background=HDR_BG,
                    foreground=HDR_FG,
                    font=HDR_FONT,
                    relief="flat",
                    borderwidth=1,
                    bordercolor="#1e4a7a",
                    lightcolor=HDR_BG,
                    darkcolor=HDR_BG,
                    padding=(10, 9))
        s.map(f"{sn}.Heading",
              background=[("active", "#1e4a7a")],
              relief=[("active", "flat")])
        s.map(sn,
              background=[("selected", ROW_SEL_BG)],
              foreground=[("selected", ROW_SEL_FG)])
        s.layout(sn, [
            ("Treeview.treearea", {"sticky": "nswe"})
        ])
        self.tag_configure("selrow",
            background=ROW_SEL_BG, foreground=ROW_SEL_FG,
            font=(FONT, 10, "bold"))
        self._sel_item = None
        self.bind("<<TreeviewSelect>>", self._on_select)

    def _on_select(self, event=None):
        """Marca visualmente la fila seleccionada (los tags anulan la selección nativa)."""
        if self._sel_item:
            prev, self._sel_item = self._sel_item, None
            if self.exists(prev):
                self._restore_row_tag(prev)
        sel = self.selection()
        if sel:
            tags = self.item(sel[0], "tags")
            if not any(t in tags for t in ("total", "fuente", "header_row", "subtotal")):
                self._sel_item = sel[0]
                self.item(sel[0], tags=("selrow",))
                return
        self._sel_item = None

    # ── Hover row ─────────────────────────────────────────────────────────────
    def _on_hover(self, event):
        item = self.identify_row(event.y)
        if item == self._hover_item:
            return
        # Restaurar fila anterior
        if self._hover_item:
            self._restore_row_tag(self._hover_item)
        # Resaltar nueva fila
        if item:
            tags = self.item(item, "tags")
            if not any(t in tags for t in ("total", "fuente", "header_row",
                                           "subtotal", "selrow")):
                self._hover_item = item
                self.item(item, tags=("hover",))
            else:
                self._hover_item = None
        else:
            self._hover_item = None

    def _on_leave(self, event):
        if self._hover_item:
            self._restore_row_tag(self._hover_item)
            self._hover_item = None

    def _restore_row_tag(self, item):
        """Restaura el tag odd/even según la posición del item."""
        try:
            if item == getattr(self, "_sel_item", None):
                self.item(item, tags=("selrow",))
                return
            children = self.get_children()
            idx = list(children).index(item)
            natural_tag = "even" if idx % 2 else "odd"
            self.item(item, tags=(natural_tag,))
        except (ValueError, tk.TclError):
            pass

    # ── API pública ───────────────────────────────────────────────────────────
    def clear(self):
        self.delete(*self.get_children())
        self._hover_item = None
        self._sel_item = None

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

    def add_fuente(self, texto, ncols):
        vals = [texto] + [""] * (ncols - 1)
        self.insert("", "end", values=vals, tags=("fuente",))

    # ── Edición inline ────────────────────────────────────────────────────────
    def _on_double_click(self, event):
        region = self.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id  = self.identify_row(event.y)
        col_id  = self.identify_column(event.x)
        col_idx = int(col_id.replace("#", "")) - 1
        if not row_id:
            return
        if col_idx not in self._editable_cols:
            return
        tags = self.item(row_id, "tags")
        if any(t in tags for t in ("total", "fuente", "header_row")):
            return
        self._abrir_editor(row_id, col_idx)

    def _abrir_editor(self, row_id, col_idx):
        if self._entry_popup:
            try: self._entry_popup.destroy()
            except: pass

        bbox = self.bbox(row_id, self._columns[col_idx])
        if not bbox:
            return
        x, y, w, h = bbox
        valor_actual = self.raw_values(row_id)[col_idx]

        # Marco azul alrededor de la celda editada
        frame = tk.Frame(self, bg=ACCENT, padx=2, pady=2)
        frame.place(x=x - 2, y=y - 2, width=max(w, 90) + 4, height=h + 4)

        popup = tk.Entry(frame, font=("Consolas", 11), justify="center",
                         bg="#eff6ff", fg=TEXT_DARK,
                         relief="flat", bd=0,
                         highlightthickness=0)
        popup.pack(fill="both", expand=True)
        popup.insert(0, str(valor_actual))
        popup.select_range(0, "end")
        popup.focus_set()
        self._entry_popup = frame

        done = [False]

        def _siguiente(dr, dc):
            """Devuelve (row, col) editable siguiente o None."""
            filas = list(self.get_children())
            cols = sorted(self._editable_cols)
            try:
                r = filas.index(row_id)
            except ValueError:
                return None
            if dc:  # Tab: siguiente columna editable, luego siguiente fila
                pos = cols.index(col_idx) if col_idx in cols else -1
                if pos + 1 < len(cols):
                    return filas[r], cols[pos + 1]
                r2, c2 = r + 1, cols[0]
            else:   # Enter: misma columna, fila de abajo
                r2, c2 = r + 1, col_idx
            while r2 < len(filas):
                tg = self.item(filas[r2], "tags")
                if not any(t in tg for t in ("total", "fuente", "header_row", "subtotal")):
                    return filas[r2], c2
                r2 += 1
            return None

        def confirmar(e=None, mover=None):
            if done[0]:
                return
            done[0] = True
            nuevo_val = popup.get().strip()
            frame.destroy()
            self._entry_popup = None
            vals = self.raw_values(row_id)
            vals[col_idx] = nuevo_val
            self.item(row_id, values=vals)
            if self._on_change:
                self._on_change()
            if EditableTree.on_any_change:
                EditableTree.on_any_change()
            if mover:
                sig = _siguiente(*mover)
                if sig:
                    self.after(10, lambda: self._abrir_editor(sig[0], sig[1]))

        def cancelar(e=None):
            done[0] = True
            frame.destroy()
            self._entry_popup = None

        popup.bind("<Return>",   lambda e: confirmar(mover=(1, 0)) or "break")
        popup.bind("<Tab>",      lambda e: confirmar(mover=(0, 1)) or "break")
        popup.bind("<Escape>",   cancelar)
        popup.bind("<FocusOut>", lambda e: confirmar())


# ──────────────────────────────────────────────────────────────────────────────
# COMPONENTES UI MEJORADOS
# ──────────────────────────────────────────────────────────────────────────────

def card(parent, title=None, pady=6, accent_color=None):
    """
    Tarjeta profesional con:
    • Sombra simulada (frame oscuro detrás + offset)
    • Borde izquierdo coloreado (4px)
    • Cabecera con gradiente simulado
    """
    accent = accent_color or CARD_BORDER_L

    # Sombra (frame ligeramente desplazado)
    shadow = tk.Frame(parent, bg=CARD_SHADOW)
    shadow.pack(fill="x", padx=16, pady=(pady, 0))

    # Marco con borde izquierdo coloreado
    outer = tk.Frame(shadow, bg=accent)
    outer.pack(fill="both", expand=True, padx=(0, 2), pady=(0, 2))

    inner = tk.Frame(outer, bg=CARD_BG)
    inner.pack(fill="both", expand=True, padx=(4, 0), pady=0)

    if title:
        # Cabecera con barra de acento + título
        hdr_wrap = tk.Frame(inner, bg=HDR_BG)
        hdr_wrap.pack(fill="x")

        # Franja de color en la cabecera
        tk.Frame(hdr_wrap, bg=accent, width=4).pack(side="left", fill="y")

        hdr = tk.Frame(hdr_wrap, bg=HDR_BG, pady=0)
        hdr.pack(side="left", fill="both", expand=True)

        tk.Label(hdr, text=title, bg=HDR_BG, fg=HDR_FG,
                 font=(FONT, 10, "bold"), padx=12, pady=9
                 ).pack(side="left")

        content = tk.Frame(inner, bg=CARD_BG, padx=0, pady=0)
        content.pack(fill="both", expand=True)
        return inner, content

    content = tk.Frame(inner, bg=CARD_BG, padx=8, pady=8)
    content.pack(fill="both", expand=True)
    return inner, content


def section_title(parent, num, titulo, subtitulo=""):
    """Título de sección con badge numerado y línea divisoria."""
    f = tk.Frame(parent, bg=LIGHT_BG)
    f.pack(fill="x", padx=14, pady=(16, 2))

    # Badge número
    badge = tk.Frame(f, bg=ACCENT)
    badge.pack(side="left")
    tk.Label(badge, text=f" {num} ", bg=ACCENT, fg=TEXT_WHITE,
             font=(FONT, 10, "bold"), pady=3).pack()

    # Título
    tk.Label(f, text=f"  {titulo}", bg=LIGHT_BG, fg=TEXT_DARK,
             font=(FONT, 13, "bold")).pack(side="left")

    # Subtítulo
    if subtitulo:
        tk.Label(f, text=f"  ·  {subtitulo}", bg=LIGHT_BG, fg=TEXT_MID,
                 font=(FONT, 10)).pack(side="left")

    # Línea divisoria
    line = tk.Frame(parent, bg=BORDER, height=1)
    line.pack(fill="x", padx=14, pady=(2, 0))


def action_btn(parent, text, cmd, color=ACCENT, small=False):
    """Botón con efecto hover suave."""
    pad_x = 10 if small else 16
    pad_y = 4  if small else 7
    fsize = 9  if small else 10

    # Versión oscura del color para hover
    def darken(hex_c):
        try:
            r = int(hex_c[1:3], 16)
            g = int(hex_c[3:5], 16)
            b = int(hex_c[5:7], 16)
            r = max(0, r - 30)
            g = max(0, g - 30)
            b = max(0, b - 30)
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return hex_c

    dark = darken(color)

    f = tk.Frame(parent, bg=color, cursor="hand2")
    lbl = tk.Label(f, text=text, bg=color, fg=TEXT_WHITE,
                   font=(FONT, fsize, "bold"),
                   padx=pad_x, pady=pad_y)
    lbl.pack()

    f.bind("<Button-1>",   lambda e: cmd())
    lbl.bind("<Button-1>", lambda e: cmd())
    f.bind("<Enter>",      lambda e: [f.configure(bg=dark), lbl.configure(bg=dark)])
    f.bind("<Leave>",      lambda e: [f.configure(bg=color), lbl.configure(bg=color)])
    lbl.bind("<Enter>",    lambda e: [f.configure(bg=dark), lbl.configure(bg=dark)])
    lbl.bind("<Leave>",    lambda e: [f.configure(bg=color), lbl.configure(bg=color)])
    return f


def tree_with_scroll(parent, cols, widths=None, height=8,
                     editable=None, on_change=None,
                     accent=None) -> EditableTree:
    """
    Contenedor completo de tabla con:
    • Barra de acento superior (3px)
    • Treeview con EditableTree
    • Scrollbars vertical y horizontal
    • Pie con leyenda de edición
    """
    accent = accent or ACCENT_BAR

    outer = tk.Frame(parent, bg=CARD_BG)
    outer.pack(fill="both", expand=True, padx=2, pady=2)

    # Barra de acento superior
    tk.Frame(outer, bg=accent, height=ACCENT_BAR_H).pack(fill="x")

    # Contenedor del Treeview + scrollbars
    wrap = tk.Frame(outer, bg=CARD_BG)
    wrap.pack(fill="both", expand=True)

    tree = EditableTree(wrap, cols, widths, height,
                        editable_cols=editable, on_change=on_change)

    # Scrollbars estilizadas
    style = ttk.Style()
    style.configure("Thin.Vertical.TScrollbar",
                    troughcolor=LIGHT_BG, background=TEXT_LIGHT,
                    arrowsize=12)
    style.configure("Thin.Horizontal.TScrollbar",
                    troughcolor=LIGHT_BG, background=TEXT_LIGHT,
                    arrowsize=12)

    vsb = ttk.Scrollbar(wrap, orient="vertical",   command=tree.yview,
                        style="Thin.Vertical.TScrollbar")
    hsb = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview,
                        style="Thin.Horizontal.TScrollbar")
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid( row=0, column=1, sticky="ns")
    hsb.grid( row=1, column=0, sticky="ew")
    wrap.grid_rowconfigure(0, weight=1)
    wrap.grid_columnconfigure(0, weight=1)

    return tree


def hint_label(parent, text):
    """Etiqueta de ayuda con icono."""
    f = tk.Frame(parent, bg=CARD_BG)
    f.pack(fill="x", padx=2, pady=(2, 0))
    tk.Frame(f, bg=ACCENT, width=3).pack(side="left", fill="y")
    tk.Label(f, text=f"  ✏  {text}",
             bg="#eff6ff", fg="#1d4ed8",
             font=(FONT, 8, "italic"),
             padx=8, pady=4).pack(side="left", fill="x", expand=True)
