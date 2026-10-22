"""Ventana principal: barra superior, menú lateral y estado."""
import datetime
import tkinter as tk
from tkinter import ttk
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.calculos import *
from comite_policial.config.settings import settings
from comite_policial.core.fechas import semana_str
from comite_policial.ui.widgets import *
from comite_policial.ui.calendarios import *
from comite_policial.ui.views import ViewsMixin
from comite_policial.ui.recuadros import RecuadrosMixin
from comite_policial.ui.views_extra import ExtraViewsMixin
from comite_policial.ui.documento import DocumentoMixin
from comite_policial.exporters.pptx_export import PptxExporterMixin


class App(ViewsMixin, RecuadrosMixin, ExtraViewsMixin, DocumentoMixin,
          PptxExporterMixin, tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(settings.app_title)
        self.minsize(960, 560)
        self._apply_initial_geometry()
        self.configure(bg=DARK)
        try:
            ttk.Style().theme_use("clam")
        except tk.TclError:
            pass

        hoy = datetime.date.today()
        self.v_semana = tk.StringVar(
            value=semana_str(hoy) if settings.semana_auto else settings.semana_default)
        self.v_fecha  = tk.StringVar(value=hoy.strftime("%d/%m/%Y"))
        self.v_turno  = tk.StringVar(value=settings.turno_responsable)
        self.v_status = tk.StringVar(value="Listo — doble clic sobre cualquier celda para editar")

        self._det_a_rows: list[dict] = []
        self.fotos_pptx: dict[str, str] = {}   # clave de marco -> ruta de la foto
        self._doc_init()
        EditableTree.on_any_change = self._doc_cambio
        self.v_semana.trace_add("write", lambda *a: self._doc_cambio())
        self.v_turno.trace_add("write", lambda *a: self._doc_cambio())
        self._build_ui()

    def _apply_initial_geometry(self):
        """Ajusta el tamaño inicial de la ventana según la resolución de la pantalla."""
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        # En pantallas de notebook (<= 1440px ancho o <= 860px alto), maximizar o ajustar al área útil
        if sw <= 1440 or sh <= 860:
            w = max(960, min(settings.win_width, sw - 24))
            h = max(560, min(settings.win_height, sh - 72))
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h - 40) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
            try:
                self.state("zoomed")
            except tk.TclError:
                pass
        else:
            w = min(settings.win_width, sw - 60)
            h = min(settings.win_height, sh - 80)
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h - 40) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        self._build_topbar()
        body = tk.Frame(self, bg=DARK)
        body.pack(fill="both", expand=True)
        self._build_sidebar(body)
        self._build_workspace(body)
        self._build_statusbar()

    # ── helper: campo + botón calendario ──────────────────────────────────────
    def _topbar_field(self, parent, label, var, icon, popup_fn, width=22):
        grp = tk.Frame(parent, bg=DARK); grp.pack(side="left", padx=8)
        tk.Label(grp, text=label, bg=DARK, fg=TEXT_LIGHT,
                 font=(FONT, 8)).pack(anchor="w")
        inner = tk.Frame(grp, bg="#1e293b",
                         highlightthickness=1, highlightbackground="#334155")
        inner.pack()
        tk.Label(inner, textvariable=var, bg="#1e293b", fg=TEXT_WHITE,
                 font=(FONT, 9), width=width, anchor="w", padx=6
                 ).pack(side="left", ipady=4)
        cal = tk.Label(inner, text=icon, bg=ACCENT, fg=TEXT_WHITE,
                       font=(FONT, 10), padx=8, cursor="hand2")
        cal.pack(side="right", ipady=2)
        cal.bind("<Button-1>", lambda e: popup_fn())
        cal.bind("<Enter>", lambda e: cal.configure(bg=ACCENT2))
        cal.bind("<Leave>", lambda e: cal.configure(bg=ACCENT))

    def _build_topbar(self):
        bar = tk.Frame(self, bg=DARK, height=62)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        # Logo + título
        left = tk.Frame(bar, bg=DARK); left.pack(side="left", padx=12, pady=8)
        tk.Label(left, text="⚓", bg=DARK, fg=ACCENT,
                 font=(FONT, 20)).pack(side="left", padx=(0, 10))
        ttls = tk.Frame(left, bg=DARK); ttls.pack(side="left")
        tk.Label(ttls, text="PDI — COMITÉ POLICIAL",
                 bg=DARK, fg=TEXT_WHITE, font=(FONT, 12, "bold")).pack(anchor="w")
        tk.Label(ttls, text="Región Policial Valparaíso",
                 bg=DARK, fg=TEXT_LIGHT, font=(FONT, 9)).pack(anchor="w")

        tk.Frame(bar, bg="#334155", width=1).pack(side="left", fill="y", padx=10, pady=10)

        # Turno
        gt = tk.Frame(bar, bg=DARK); gt.pack(side="left", pady=10)
        tk.Label(gt, text="Turno a/c:", bg=DARK, fg=TEXT_LIGHT,
                 font=(FONT, 8)).pack(anchor="w")
        tk.Entry(gt, textvariable=self.v_turno, width=28,
                 bg="#1e293b", fg=TEXT_WHITE, insertbackground=TEXT_WHITE,
                 relief="flat", font=(FONT, 9),
                 highlightthickness=1, highlightbackground="#334155"
                 ).pack(ipady=4)

        # Exportación a PowerPoint (extremo derecho)
        exp = tk.Frame(bar, bg=DARK); exp.pack(side="right", padx=(0, 12), pady=14)
        action_btn(exp, "📊  Exportar PowerPoint", self._exportar_pptx,
                   SUCCESS).pack(side="left", padx=(0, 6))
        action_btn(exp, "🖼  Fotos", self._fotos_dialogo, TEAL).pack(side="left")

        # Campos con calendario (lado derecho)
        right = tk.Frame(bar, bg=DARK); right.pack(side="right", padx=12, pady=10)

        self._topbar_field(
            right, "📅 Fecha del Informe", self.v_fecha, "🗓",
            lambda: CalendarPopup(self, self.v_fecha, "Seleccionar Fecha"),
            width=12
        )
        tk.Frame(right, bg="#334155", width=1).pack(side="left", fill="y", padx=8, pady=2)
        self._topbar_field(
            right, "📆 Semana Operativa", self.v_semana, "🗓",
            lambda: WeekRangePopup(self, self.v_semana),
            width=24
        )

    def _build_statusbar(self):
        bar = tk.Frame(self, bg="#1e293b", height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, textvariable=self.v_status, bg="#1e293b",
                 fg=TEXT_LIGHT, font=(FONT, 8), padx=12).pack(side="left", pady=4)

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=SIDEBAR_BG, width=220)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        self._sb = sb

        # Encabezado fijo superior
        logo_frame = tk.Frame(sb, bg=HEADER_BG)
        logo_frame.pack(side="top", fill="x")
        self._sb_logo_lbl = tk.Label(
            logo_frame, text="INFORME SEMANAL", bg=HEADER_BG,
            fg=TEXT_WHITE, font=(FONT, 10, "bold"), pady=10
        )
        self._sb_logo_lbl.pack()

        # Zona central adaptable y con scroll automático si la pantalla es baja
        mid = tk.Frame(sb, bg=SIDEBAR_BG)
        mid.pack(side="top", fill="both", expand=True)

        sb_canvas = tk.Canvas(mid, bg=SIDEBAR_BG, highlightthickness=0, bd=0, width=212)
        sb_vsb = ttk.Scrollbar(mid, orient="vertical", command=sb_canvas.yview)
        sb_canvas.configure(yscrollcommand=sb_vsb.set)
        sb_canvas.pack(side="left", fill="both", expand=True)

        sb_inner = tk.Frame(sb_canvas, bg=SIDEBAR_BG)
        sb_win = sb_canvas.create_window((0, 0), window=sb_inner, anchor="nw")

        self._sb_canvas = sb_canvas
        self._sb_inner = sb_inner
        self._sb_win = sb_win
        self._sb_vsb = sb_vsb
        self._sb_scroll_needed = False
        self._sb_compact = None

        # Grupos de menú
        groups = [
            ("OPERATIVO",   [
                ("I    Armas",              "armas",  "🔫"),
                ("II   Drogas",             "drogas", "💊"),
            ]),
            ("HOMICIDIOS",  [
                ("III‑A  Por Brigada BH",   "homi_a", "📍"),
                ("III‑B  Modalidad",        "homi_b", "📋"),
                ("III‑C  Resultado (%)",    "homi_c", "📊"),
                ("III‑D  Comparativo",      "homi_d", "📈"),
            ]),
            ("LOGÍSTICO",   [
                ("IV   Migraciones",        "migr",   "🛂"),
                ("IV‑B  Flujo Migratorio",   "flujo",  "🏔"),
                ("V    Vehículos",          "vehic",  "🚗"),
                ("VI   Delitos Sexuales",   "brisex", "🔴"),
                ("VII  Bandas",             "bandas", "🕵"),
            ]),
            ("DETENIDOS",   [
                ("VIII‑A  Por Delito",      "det_a",  "⛓"),
                ("VIII‑B  Por Prefectura",  "det_b",  "🏛"),
                ("VIII‑C  Acumulado Año",   "det_c",  "📅"),
            ]),
            ("RELEVANTES",  [
                ("IX   Concurrencias y Diligencias", "recuadros", "📝"),
                ("X    Diligencias de relevancia", "textox", "⭐"),
            ]),
        ]

        self._vistas: dict[str, tk.Frame] = {}
        self._nav_btns: dict[str, tuple[tk.Frame, tk.Label, tk.Label]] = {}
        self._grp_lbls: list[tk.Label] = []
        self._current_vista = None

        for grp_name, items in groups:
            gl = tk.Label(sb_inner, text=grp_name, bg=SIDEBAR_BG, fg="#4a7ab5",
                          font=(FONT, 7, "bold"), padx=10, anchor="w")
            gl.pack(fill="x", pady=(6, 2))
            self._grp_lbls.append(gl)
            for label, key, icon in items:
                row = tk.Frame(sb_inner, bg=SIDEBAR_BG, cursor="hand2")
                row.pack(fill="x", padx=6, pady=1)
                ic = tk.Label(row, text=icon, bg=SIDEBAR_BG, fg=TEXT_LIGHT,
                              font=(FONT, 9), width=2)
                ic.pack(side="left", padx=(4, 0))
                lbl = tk.Label(row, text=label, bg=SIDEBAR_BG, fg=TEXT_LIGHT,
                               font=(FONT, 9), anchor="w", pady=5, padx=4)
                lbl.pack(side="left", fill="x", expand=True)
                self._nav_btns[key] = (row, lbl, ic)
                for widget in (row, lbl, ic):
                    widget.bind("<Button-1>", lambda e, k=key: self._switch(k))
                    widget.bind("<Enter>",    lambda e, r=row, l=lbl, i=ic: (
                        r.configure(bg=SIDEBAR_HOV),
                        l.configure(bg=SIDEBAR_HOV, fg=TEXT_WHITE),
                        i.configure(bg=SIDEBAR_HOV)))
                    widget.bind("<Leave>",    lambda e, r=row, l=lbl, i=ic, k=key: (
                        r.configure(bg=SIDEBAR_BG if self._current_vista != k else ACCENT2),
                        l.configure(bg=SIDEBAR_BG if self._current_vista != k else ACCENT2,
                                    fg=TEXT_LIGHT if self._current_vista != k else TEXT_WHITE),
                        i.configure(bg=SIDEBAR_BG if self._current_vista != k else ACCENT2)))

        self._build_doc_panel(sb_inner)

        sb_canvas.bind("<Configure>", lambda e: self._update_sidebar_layout())
        sb_inner.bind("<Configure>", lambda e: self._update_sidebar_scroll())

    def _update_sidebar_layout(self):
        """Adapta la densidad vertical del menú según el alto disponible y actualiza el scroll."""
        c = self._sb_canvas
        alto_disp = c.winfo_height()
        if alto_disp <= 1:
            return
        # Modo compacto para pantallas de notebook (< 680 px disponibles en zona central)
        compact = alto_disp < 680
        if compact != self._sb_compact:
            self._sb_compact = compact
            nav_pady = 2 if compact else 5
            grp_pady = (4, 1) if compact else (7, 2)
            for gl in self._grp_lbls:
                gl.pack_configure(pady=grp_pady)
            for row, lbl, ic in self._nav_btns.values():
                lbl.configure(pady=nav_pady)
        self._update_sidebar_scroll()

    def _update_sidebar_scroll(self):
        """Activa el scroll del menú lateral solo cuando el alto de la ventana no alcanza."""
        c = self._sb_canvas
        c.update_idletasks()
        alto_vista = c.winfo_height()
        alto_contenido = self._sb_inner.winfo_reqheight()
        necesita = alto_contenido > alto_vista + 2
        if necesita:
            if not self._sb_scroll_needed:
                self._sb_vsb.pack(side="right", fill="y", before=c)
                self._sb_scroll_needed = True
            c.itemconfig(self._sb_win, width=max(1, c.winfo_width()))
            c.configure(scrollregion=(0, 0, c.winfo_width(), alto_contenido))
        else:
            if self._sb_scroll_needed:
                self._sb_vsb.pack_forget()
                self._sb_scroll_needed = False
            c.itemconfig(self._sb_win, width=max(1, c.winfo_width()))
            c.configure(scrollregion=(0, 0, c.winfo_width(), alto_vista))
            c.yview_moveto(0)

    def _build_workspace(self, parent):
        # Scroll canvas para el panel principal
        self._ws_outer = tk.Frame(parent, bg=LIGHT_BG)
        self._ws_outer.pack(side="right", fill="both", expand=True)

        canvas = tk.Canvas(self._ws_outer, bg=LIGHT_BG,
                           highlightthickness=0)
        vsb = ttk.Scrollbar(self._ws_outer, orient="vertical",
                            command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)

        self._ws = tk.Frame(canvas, bg=LIGHT_BG)
        self._ws_win = canvas.create_window((0, 0), window=self._ws, anchor="nw")
        self._vsb = vsb
        self._vsb_visible = False
        self._canvas = canvas

        def on_resize(e):
            canvas.itemconfig(self._ws_win, width=e.width)
            self._update_scroll()
        def on_frame(e):
            self._update_scroll()

        def on_wheel(e):
            # Si el cursor está sobre el menú lateral, desplazar el menú lateral
            px = self.winfo_pointerx() - self.winfo_rootx()
            if hasattr(self, "_sb") and 0 <= px <= self._sb.winfo_width():
                if getattr(self, "_sb_scroll_needed", False):
                    self._sb_canvas.yview_scroll(-1 * (e.delta // 120), "units")
                return
            if self._vsb_visible:
                canvas.yview_scroll(-1 * (e.delta // 120), "units")

        canvas.bind("<Configure>", on_resize)
        self._ws.bind("<Configure>", on_frame)
        canvas.bind_all("<MouseWheel>", on_wheel)

        # Construir todas las vistas
        self._build_v_armas()
        self._build_v_drogas()
        self._build_v_homi_a()
        self._build_v_homi_b()
        self._build_v_homi_c()
        self._build_v_homi_d()
        self._build_v_migr()
        self._build_v_vehic()
        self._build_v_brisex()
        self._build_v_bandas()
        self._build_v_det_a()
        self._build_v_det_b()
        self._build_v_det_c()
        self._build_v_recuadros()
        self._build_v_flujo()
        self._build_v_textox()

        self._switch("armas")

    def _update_scroll(self):
        """Muestra la barra de scroll solo si el contenido supera el área visible."""
        c = self._canvas
        c.update_idletasks()
        alto_vista = c.winfo_height()
        alto_contenido = self._ws.winfo_reqheight()
        necesita = alto_contenido > alto_vista + 1
        if necesita:
            c.configure(scrollregion=(0, 0, c.winfo_width(), alto_contenido))
            if not self._vsb_visible:
                self._vsb.pack(side="right", fill="y", before=c)
                self._vsb_visible = True
        else:
            c.configure(scrollregion=(0, 0, c.winfo_width(), alto_vista))
            c.yview_moveto(0)
            if self._vsb_visible:
                self._vsb.pack_forget()
                self._vsb_visible = False

    def _switch(self, nombre):
        for v in self._vistas.values():
            v.pack_forget()
        self._vistas[nombre].pack(fill="both", expand=True)
        self._canvas.yview_moveto(0)
        self.after_idle(self._update_scroll)
        self._current_vista = nombre
        # Actualizar estilo sidebar
        for k, (row, lbl, ic) in self._nav_btns.items():
            if k == nombre:
                row.configure(bg=ACCENT2)
                lbl.configure(bg=ACCENT2, fg=TEXT_WHITE)
                ic.configure(bg=ACCENT2)
            else:
                row.configure(bg=SIDEBAR_BG)
                lbl.configure(bg=SIDEBAR_BG, fg=TEXT_LIGHT)
                ic.configure(bg=SIDEBAR_BG)
        self.v_status.set("✏️  Doble clic sobre cualquier celda numérica para editar directamente")

    # =========================================================================
    # VISTA AUXILIAR: envuelve cada sección
    # =========================================================================
    def _make_view(self, key):
        f = tk.Frame(self._ws, bg=LIGHT_BG)
        self._vistas[key] = f
        return f



