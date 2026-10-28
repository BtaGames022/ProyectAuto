"""Ventana principal: menú lateral, barra de contexto, área de trabajo y estado."""
import datetime
import os
import tkinter as tk
from tkinter import messagebox
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.calculos import *
from comite_policial.config.settings import settings
from comite_policial.core.fechas import semana_str
from comite_policial.ui.estilo import aplicar_estilo
from comite_policial.ui.scroll import ScrollFrame, instalar_rueda
from comite_policial.ui.logo import LOGO_AZUL, LOGO_BLANCO
from comite_policial.ui.widgets import *
from comite_policial.ui.calendarios import *
from comite_policial.ui.views import ViewsMixin
from comite_policial.ui.recuadros import RecuadrosMixin
from comite_policial.ui.views_extra import ExtraViewsMixin
from comite_policial.ui.documento import DocumentoMixin
from comite_policial.exporters.pptx_export import PptxExporterMixin

# Menú: (grupo, [(numeral, nombre, clave de la vista)])
MENU = [
    ("OPERATIVO", [("I", "Armas", "armas"), ("II", "Drogas", "drogas")]),
    ("HOMICIDIOS", [("III-A", "Por brigada", "homi_a"),
                    ("III-B", "Modalidad", "homi_b"),
                    ("III-C", "Resultado (%)", "homi_c"),
                    ("III-D", "Comparativo", "homi_d")]),
    ("LOGÍSTICO", [("IV", "Migraciones", "migr"),
                   ("IV-B", "Flujo migratorio", "flujo"),
                   ("V", "Vehículos", "vehic"),
                   ("VI", "Delitos sexuales", "brisex"),
                   ("VII", "Bandas", "bandas")]),
    ("DETENIDOS", [("VIII-A", "Por delito", "det_a"),
                   ("VIII-B", "Por prefectura", "det_b"),
                   ("VIII-C", "Acumulado del año", "det_c")]),
    ("RELEVANTES", [("IX", "Concurrencias y diligencias", "recuadros"),
                    ("X", "Diligencias de relevancia", "textox")]),
]

# Árboles de cada vista (para saber si la sección ya tiene datos)
ARBOLES = {
    "armas": ["t_armas"], "drogas": ["t_drogas_g", "t_drogas_p"],
    "homi_a": ["t_homi_a"], "homi_b": ["t_homi_b"], "homi_c": ["t_homi_c"],
    "homi_d": ["t_homi_d"], "migr": ["t_migr"], "flujo": ["t_flujo"],
    "vehic": ["t_veh_e", "t_veh_t"], "brisex": ["t_brisex"], "bandas": ["t_bandas"],
    "det_a": ["t_det_a"], "det_b": ["t_det_b"], "det_c": ["t_det_c"],
}
VACIOS = {"", "0", "0,0", "0,00", "0%", "0,0%", "100%", "-"}

NAV_TXT, NAV_NUM = "#E4EAF7", "#9FB1DC"


class App(ViewsMixin, RecuadrosMixin, ExtraViewsMixin, DocumentoMixin,
          PptxExporterMixin, tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(settings.app_title)
        self.minsize(1100, 640)
        self.configure(bg=BG)
        aplicar_estilo(self)
        instalar_rueda(self)
        self._logos = (tk.PhotoImage(data=LOGO_AZUL), tk.PhotoImage(data=LOGO_BLANCO))
        try:
            self.iconphoto(True, self._logos[0])
        except tk.TclError:
            pass
        self._apply_initial_geometry()

        hoy = datetime.date.today()
        self.v_semana = tk.StringVar(
            value=semana_str(hoy) if settings.semana_auto else settings.semana_default)
        self.v_fecha = tk.StringVar(value=hoy.strftime("%d/%m/%Y"))
        self.v_turno = tk.StringVar(value=settings.turno_responsable)
        self.v_status = tk.StringVar(value="Listo")

        self._det_a_rows: list[dict] = []
        self.fotos_pptx: dict[str, str] = {}   # clave de marco -> ruta de la foto
        self._estado_after = None
        self._doc_init()
        EditableTree.on_any_change = self._doc_cambio
        self.v_semana.trace_add("write", lambda *a: self._doc_cambio())
        self.v_turno.trace_add("write", lambda *a: self._doc_cambio())
        self._build_ui()
        self.bind_all("<Control-e>", lambda e: self._exportar_pptx())
        self.bind_all("<Control-s>", lambda e: self._doc_guardar(auto=False))
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

    def report_callback_exception(self, exc, val, tb):
        """Un error dentro de un botón no debe pasar en silencio: se registra y se avisa."""
        import traceback
        detalle = "".join(traceback.format_exception(exc, val, tb))
        try:
            os.makedirs(settings.output_dir, exist_ok=True)
            with open(os.path.join(settings.output_dir, "errores.log"), "a",
                      encoding="utf-8") as f:
                f.write(f"--- {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n{detalle}\n")
        except OSError:
            pass
        self.v_status.set(f"Error: {val}")
        try:
            messagebox.showerror(
                "Error inesperado",
                f"{exc.__name__}: {val}\n\nLo ocurrido quedó registrado en "
                f"{os.path.join(settings.output_dir, 'errores.log')}.\n"
                "Tus datos siguen en la aplicación.")
        except tk.TclError:
            pass

    def _cerrar(self):
        """Antes de salir se escribe en el Word lo que estuviera pendiente."""
        try:
            if self._doc is not None and self._doc_after:
                self.after_cancel(self._doc_after)
                self._doc_guardar(auto=True)
        finally:
            self.destroy()

    def _apply_initial_geometry(self):
        """Ajusta el tamaño inicial según la pantalla (maximiza en notebooks)."""
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        if sw <= 1440 or sh <= 860:
            w = max(1000, min(settings.win_width, sw - 24))
            h = max(600, min(settings.win_height, sh - 72))
            self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h - 40) // 2)}")
            try:
                self.state("zoomed")
            except tk.TclError:
                pass
        else:
            w = min(settings.win_width, sw - 60)
            h = min(settings.win_height, sh - 80)
            self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h - 40) // 2)}")

    # ══════════════════════════════════════════════════════════════════════════
    # ESTRUCTURA
    # ══════════════════════════════════════════════════════════════════════════
    def _build_ui(self):
        self._build_statusbar()                      # primero: reserva el pie
        self._build_sidebar()
        main = tk.Frame(self, bg=BG)
        main.pack(side="left", fill="both", expand=True)
        self._build_topbar(main)
        self._build_workspace(main)

    # ── Barra superior ───────────────────────────────────────────────────────
    def _build_topbar(self, parent):
        bar = tk.Frame(parent, bg=SURFACE)
        bar.pack(fill="x")
        fila = tk.Frame(bar, bg=SURFACE)
        fila.pack(fill="x", padx=20, pady=12)
        tk.Label(fila, text="Informe semanal del comité", bg=SURFACE, fg=NAVY,
                 font=(FONT, SIZE_H2, "bold")).pack(side="left")
        Boton(fila, "Exportar PowerPoint", self._exportar_pptx, "accent"
              ).pack(side="right")
        Boton(fila, "Fotos", self._fotos_dialogo, "secondary"
              ).pack(side="right", padx=(0, 8))

        # Contexto del informe: siempre visible, define toda la exportación
        ctx = tk.Frame(parent, bg=SKY_SOFT)
        ctx.pack(fill="x")
        tk.Frame(ctx, bg=LINE, height=1).pack(fill="x")
        campos = tk.Frame(ctx, bg=SKY_SOFT)
        campos.pack(fill="x", padx=20, pady=10)
        self._campo_fecha(campos, "Semana operativa", self.v_semana, 22,
                          lambda: WeekRangePopup(self, self.v_semana))
        self._campo_fecha(campos, "Fecha del informe", self.v_fecha, 10,
                          lambda: CalendarPopup(self, self.v_fecha, "Fecha del informe"))
        g = tk.Frame(campos, bg=SKY_SOFT)
        g.pack(side="left", fill="x", expand=True)
        etiqueta(g, "Turno a/c").pack(anchor="w", pady=(0, 3))
        tk.Entry(g, textvariable=self.v_turno, width=12).pack(fill="x", ipady=5)
        tk.Frame(parent, bg=LINE, height=1).pack(fill="x")

    def _campo_fecha(self, parent, rotulo, var, ancho, abrir):
        g = tk.Frame(parent, bg=SKY_SOFT)
        g.pack(side="left", padx=(0, 22))
        etiqueta(g, rotulo).pack(anchor="w", pady=(0, 3))
        fila = tk.Frame(g, bg=SKY_SOFT)
        fila.pack()
        caja = tk.Frame(fila, bg=SURFACE, highlightthickness=1,
                        highlightbackground=LINE_STRONG)
        caja.pack(side="left")
        tk.Label(caja, textvariable=var, bg=SURFACE, fg=INK, width=ancho,
                 anchor="w", padx=8, pady=5, font=(FONT, SIZE_BASE, "bold")
                 ).pack()
        Boton(fila, "Cambiar", abrir, "secondary", chico=True).pack(side="left", padx=(6, 0))

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=SURFACE)
        bar.pack(side="bottom", fill="x")
        tk.Frame(bar, bg=LINE, height=1).pack(fill="x")
        tk.Label(bar, textvariable=self.v_status, bg=SURFACE, fg=INK_MID,
                 font=(FONT, SIZE_SMALL), padx=16, pady=5, anchor="w"
                 ).pack(side="left")
        tk.Label(bar, text="Ctrl+E exportar  ·  Ctrl+S guardar en el documento",
                 bg=SURFACE, fg=INK_SOFT, font=(FONT, SIZE_SMALL), padx=16
                 ).pack(side="right")

    # ── Menú lateral ─────────────────────────────────────────────────────────
    def _build_sidebar(self):
        escala = max(self.winfo_fpixels("1i") / 96, 1.0)   # pantallas de alta densidad
        sb = tk.Frame(self, bg=NAVY_DEEP, width=int(250 * escala))
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        self._sb = sb

        marca = tk.Frame(sb, bg=NAVY_DEEP)
        marca.pack(fill="x", padx=18, pady=(18, 12))
        tk.Label(marca, image=self._logos[1], bg=NAVY_DEEP).pack(anchor="w")
        tk.Label(marca, text="Comité Policial", bg=NAVY_DEEP, fg="#FFFFFF",
                 font=(FONT, SIZE_H2, "bold")).pack(anchor="w", pady=(8, 0))
        tk.Label(marca, text=settings.unidad, bg=NAVY_DEEP, fg=NAV_NUM,
                 font=(FONT, SIZE_SMALL)).pack(anchor="w")
        tk.Frame(sb, bg="#1D3470", height=1).pack(fill="x", padx=18)

        # Panel de documento fijo abajo; el menú ocupa el resto y se desplaza
        self._build_doc_panel(sb)
        nav = ScrollFrame(sb, bg=NAVY_DEEP)
        nav.pack(fill="both", expand=True, pady=(6, 0))
        interior = nav.interior

        self._nav = {}
        for grupo, items in MENU:
            tk.Label(interior, text=grupo, bg=NAVY_DEEP, fg=NAV_NUM,
                     font=(FONT, 8, "bold"), anchor="w", padx=18
                     ).pack(fill="x", pady=(12, 3))
            for num, nombre, clave in items:
                self._nav[clave] = self._nav_item(interior, num, nombre, clave)

    def _nav_item(self, parent, num, nombre, clave):
        fila = tk.Frame(parent, bg=NAVY_DEEP, cursor="hand2")
        fila.pack(fill="x", padx=8, pady=1)
        marca = tk.Frame(fila, bg=NAVY_DEEP, width=4)
        marca.pack(side="left", fill="y")
        n = tk.Label(fila, text=num, bg=NAVY_DEEP, fg=NAV_NUM, width=6, anchor="w",
                     font=(FONT, SIZE_SMALL, "bold"), pady=5, padx=8)
        n.pack(side="left")
        t = tk.Label(fila, text=nombre, bg=NAVY_DEEP, fg=NAV_TXT, anchor="w",
                     font=(FONT, SIZE_BASE))
        t.pack(side="left", fill="x", expand=True)
        punto = tk.Label(fila, text="", bg=NAVY_DEEP, fg="#6FD6A4", width=2,
                         font=(FONT, SIZE_SMALL))
        punto.pack(side="right", padx=(0, 6))
        fila._partes = (marca, n, t, punto)

        def colorear(fondo):
            for w in (fila, n, t, punto):
                w.configure(bg=fondo)

        def entra(_e):
            if self._current_vista != clave:
                colorear(NAVY_SOFT)

        def sale(_e):
            if self._current_vista != clave:
                colorear(NAVY_DEEP)

        for w in (fila, marca, n, t, punto):
            w.bind("<Button-1>", lambda e, k=clave: self._switch(k))
            w.bind("<Enter>", entra)
            w.bind("<Leave>", sale)
        fila._colorear = colorear
        return fila

    def _nav_pintar(self):
        for clave, fila in self._nav.items():
            marca, n, t, punto = fila._partes
            activo = clave == self._current_vista
            fila._colorear(NAVY_SOFT if activo else NAVY_DEEP)
            marca.configure(bg=GOLD if activo else NAVY_DEEP)
            t.configure(fg="#FFFFFF" if activo else NAV_TXT,
                        font=(FONT, SIZE_BASE, "bold" if activo else "normal"))
            n.configure(fg=GOLD if activo else NAV_NUM)

    # ── Estado de cada sección (punto verde = ya tiene datos) ────────────────
    def _tiene_datos(self, clave) -> bool:
        if clave == "recuadros":
            return bool(getattr(self, "recuadros", []))
        if clave == "textox":
            return bool(self.txt_x.get("1.0", "end-1c").strip())
        for attr in ARBOLES.get(clave, []):
            t = getattr(self, attr, None)
            for r in (t.get_children() if t is not None else ()):
                if {"total", "fuente"} & set(t.item(r, "tags")):
                    continue
                if any(v.strip() not in VACIOS for v in t.raw_values(r)[1:]):
                    return True
        return False

    def _nav_estado_later(self):
        if self._estado_after is None:
            self._estado_after = self.after(400, self._nav_estado)

    def _nav_estado(self):
        self._estado_after = None
        for clave, fila in self._nav.items():
            try:
                con = self._tiene_datos(clave)
            except (tk.TclError, AttributeError):
                con = False
            fila._partes[3].configure(text="●" if con else "")

    # ── Área de trabajo ──────────────────────────────────────────────────────
    def _build_workspace(self, parent):
        self._scroll = ScrollFrame(parent, bg=BG)
        self._scroll.pack(fill="both", expand=True)
        self._ws = self._scroll.interior
        self._canvas = self._scroll.canvas

        self._vistas: dict[str, tk.Frame] = {}
        self._current_vista = None
        for construir in (
                self._build_v_armas, self._build_v_drogas, self._build_v_homi_a,
                self._build_v_homi_b, self._build_v_homi_c, self._build_v_homi_d,
                self._build_v_migr, self._build_v_vehic, self._build_v_brisex,
                self._build_v_bandas, self._build_v_det_a, self._build_v_det_b,
                self._build_v_det_c, self._build_v_recuadros, self._build_v_flujo,
                self._build_v_textox):
            construir()
        self._switch("armas")
        self._nav_estado()

    def _switch(self, nombre):
        for v in self._vistas.values():
            v.pack_forget()
        self._vistas[nombre].pack(fill="x", anchor="n", pady=(0, 24))
        self._current_vista = nombre
        self._scroll.al_inicio()
        self._nav_pintar()
        self._nav_estado()
        self.v_status.set("Doble clic en una celda para editarla")

    def _make_view(self, key):
        f = tk.Frame(self._ws, bg=BG)
        self._vistas[key] = f
        return f
