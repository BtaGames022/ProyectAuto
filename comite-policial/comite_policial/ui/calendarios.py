"""Selectores de fecha y rango semanal (calendarios interactivos)."""
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from comite_policial.config.theme import *
from comite_policial.config.catalogos import *
from comite_policial.core.calculos import *
from comite_policial.ui.widgets import *

import calendar as _cal

MESES_ES    = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
               "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
MESES_ABREV = ["ENE","FEB","MAR","ABR","MAY","JUN",
               "JUL","AGO","SEP","OCT","NOV","DIC"]
DIAS_ES     = ["Lu","Ma","Mi","Ju","Vi","Sá","Do"]


class CalendarPopup(tk.Toplevel):
    """Popup de calendario para seleccionar una fecha simple."""

    def __init__(self, master, var: tk.StringVar, title="Seleccionar fecha"):
        super().__init__(master)
        self.var = var
        self.title(title)
        self.resizable(False, False)
        self.configure(bg=DARK)
        self.grab_set()

        try:
            p = var.get().split("/")
            self._sel = datetime.date(int(p[2]), int(p[1]), int(p[0]))
        except Exception:
            self._sel = datetime.date.today()
        self._vy = self._sel.year
        self._vm = self._sel.month

        self._build()
        self._center(master)

    def _center(self, master):
        self.update_idletasks()
        mx = master.winfo_rootx() + master.winfo_width()  // 2
        my = master.winfo_rooty() + master.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{mx - w//2}+{my - h//2}")

    def _nav_btn(self, parent, text, cmd):
        b = tk.Label(parent, text=text, bg=HEADER_BG, fg=TEXT_WHITE,
                     font=(FONT, 14, "bold"), padx=14, cursor="hand2")
        b.pack(side="left" if text == "◀" else "right")
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.configure(fg=ACCENT))
        b.bind("<Leave>", lambda e: b.configure(fg=TEXT_WHITE))

    def _build(self):
        hdr = tk.Frame(self, bg=HEADER_BG, pady=8)
        hdr.pack(fill="x")
        self._nav_btn(hdr, "◀", self._prev)
        self._lbl_m = tk.Label(hdr, bg=HEADER_BG, fg=TEXT_WHITE,
                                font=(FONT, 12, "bold"))
        self._lbl_m.pack(side="left", expand=True)
        self._nav_btn(hdr, "▶", self._next)

        df = tk.Frame(self, bg="#1e293b", pady=4); df.pack(fill="x")
        for i, d in enumerate(DIAS_ES):
            tk.Label(df, text=d, bg="#1e293b",
                     fg=DANGER if i == 6 else TEXT_LIGHT,
                     font=(FONT, 9, "bold"), width=4).pack(side="left", padx=2)

        self._gf = tk.Frame(self, bg=DARK, padx=6, pady=6); self._gf.pack()

        foot = tk.Frame(self, bg=DARK, pady=8); foot.pack(fill="x")
        for txt, cmd, col, sd in [("✓  Confirmar", self._ok, SUCCESS, "left"),
                                   ("✕  Cancelar", self.destroy, "#475569", "right")]:
            b = tk.Label(foot, text=txt, bg=col, fg=TEXT_WHITE,
                         font=(FONT, 10, "bold"), padx=14, pady=6, cursor="hand2")
            b.pack(side=sd, padx=8)
            b.bind("<Button-1>", lambda e, fn=cmd: fn())
        self._refresh()

    def _refresh(self):
        for w in self._gf.winfo_children(): w.destroy()
        self._lbl_m.configure(text=f"{MESES_ES[self._vm-1]}  {self._vy}")
        hoy = datetime.date.today()
        for semana in _cal.monthcalendar(self._vy, self._vm):
            row = tk.Frame(self._gf, bg=DARK); row.pack(fill="x", pady=1)
            for ci, dia in enumerate(semana):
                if dia == 0:
                    tk.Label(row, text="", bg=DARK, width=4).pack(side="left", padx=2)
                    continue
                f = datetime.date(self._vy, self._vm, dia)
                es_sel = (f == self._sel)
                es_hoy = (f == hoy)
                es_dom = (ci == 6)
                if es_sel:
                    bg, fg, bold = ACCENT, TEXT_WHITE, "bold"
                elif es_hoy:
                    bg, fg, bold = HEADER_BG, TEXT_WHITE, "bold"
                elif es_dom:
                    bg, fg, bold = DARK, "#f87171", "normal"
                else:
                    bg, fg, bold = DARK, TEXT_WHITE, "normal"
                lbl = tk.Label(row, text=str(dia), bg=bg, fg=fg,
                               font=(FONT, 10, bold), width=4, pady=5, cursor="hand2")
                lbl.pack(side="left", padx=2)
                lbl.bind("<Button-1>", lambda e, d=f: self._pick(d))
                lbl.bind("<Enter>",  lambda e, l=lbl: l.configure(bg=ACCENT2))
                lbl.bind("<Leave>",  lambda e, l=lbl, b=bg: l.configure(bg=b))

    def _pick(self, f): self._sel = f; self._refresh()
    def _prev(self):
        self._vm -= 1
        if self._vm == 0: self._vm, self._vy = 12, self._vy - 1
        self._refresh()
    def _next(self):
        self._vm += 1
        if self._vm == 13: self._vm, self._vy = 1, self._vy + 1
        self._refresh()
    def _ok(self):
        self.var.set(self._sel.strftime("%d/%m/%Y"))
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
# CALENDARIO DE RANGO — SELECTOR DE SEMANA OPERATIVA
# ──────────────────────────────────────────────────────────────────────────────
class WeekRangePopup(tk.Toplevel):
    """
    Selecciona el rango de días de la semana operativa.
    1er clic → día inicio · 2do clic → día fin
    Genera: "DD.MES AL DD.MES.YY"
    """

    def __init__(self, master, var: tk.StringVar):
        super().__init__(master)
        self.var = var
        self.title("Seleccionar Semana Operativa")
        self.resizable(False, False)
        self.configure(bg=DARK)
        self.grab_set()

        self._start: datetime.date | None = None
        self._end:   datetime.date | None = None
        self._step = 0
        hoy = datetime.date.today()
        self._vy, self._vm = hoy.year, hoy.month
        self._build()
        self._center(master)

    def _center(self, master):
        self.update_idletasks()
        mx = master.winfo_rootx() + master.winfo_width()  // 2
        my = master.winfo_rooty() + master.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{mx - w//2}+{my - h//2}")

    def _nav_btn(self, parent, text, cmd):
        b = tk.Label(parent, text=text, bg=HEADER_BG, fg=TEXT_WHITE,
                     font=(FONT, 14, "bold"), padx=14, cursor="hand2")
        b.pack(side="left" if text == "◀" else "right")
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.configure(fg=ACCENT))
        b.bind("<Leave>", lambda e: b.configure(fg=TEXT_WHITE))

    def _build(self):
        # Instrucción dinámica
        info = tk.Frame(self, bg="#1c2f4a", pady=10); info.pack(fill="x")
        self._lbl_i = tk.Label(info,
            text="① Haz clic en el DÍA DE INICIO de la semana",
            bg="#1c2f4a", fg="#fbbf24", font=(FONT, 10, "bold"))
        self._lbl_i.pack()

        # Visualización del rango
        self._lbl_r = tk.Label(self, text="  Semana: —  ",
                                bg="#0f2540", fg=TEXT_WHITE,
                                font=(FONT, 12, "bold"), pady=6)
        self._lbl_r.pack(fill="x")

        # Cabecera mes/año
        hdr = tk.Frame(self, bg=HEADER_BG, pady=8); hdr.pack(fill="x")
        self._nav_btn(hdr, "◀", self._prev)
        self._lbl_m = tk.Label(hdr, bg=HEADER_BG, fg=TEXT_WHITE,
                                font=(FONT, 12, "bold"))
        self._lbl_m.pack(side="left", expand=True)
        self._nav_btn(hdr, "▶", self._next)

        # Días de semana
        df = tk.Frame(self, bg="#1e293b", pady=4); df.pack(fill="x")
        for i, d in enumerate(DIAS_ES):
            tk.Label(df, text=d, bg="#1e293b",
                     fg=DANGER if i == 6 else TEXT_LIGHT,
                     font=(FONT, 9, "bold"), width=4).pack(side="left", padx=2)

        self._gf = tk.Frame(self, bg=DARK, padx=6, pady=6); self._gf.pack()

        # Botones
        foot = tk.Frame(self, bg=DARK, pady=10); foot.pack(fill="x")
        acciones = [
            ("↺  Reiniciar",  self._reset,     WARNING, "left"),
            ("✓  Confirmar",  self._ok,         SUCCESS, "left"),
            ("✕  Cancelar",   self.destroy,     "#475569", "right"),
        ]
        for txt, cmd, col, sd in acciones:
            b = tk.Label(foot, text=txt, bg=col, fg=TEXT_WHITE,
                         font=(FONT, 10, "bold"), padx=14, pady=6, cursor="hand2")
            b.pack(side=sd, padx=8)
            b.bind("<Button-1>", lambda e, fn=cmd: fn())

        self._refresh()

    def _rango_txt(self) -> str:
        if not self._start: return "  Semana: —  "
        s = self._start
        p1 = f"{s.day:02d}.{MESES_ABREV[s.month-1]}"
        if not self._end: return f"  Semana:  {p1}  AL  ...  "
        e = self._end
        p2 = f"{e.day:02d}.{MESES_ABREV[e.month-1]}.{str(e.year)[-3:]}"
        return f"  {p1}  AL  {p2}  "

    def _semana_str(self) -> str:
        if not self._start or not self._end: return ""
        s, e = self._start, self._end
        return (f"{s.day:02d}.{MESES_ABREV[s.month-1]} "
                f"AL {e.day:02d}.{MESES_ABREV[e.month-1]}.{str(e.year)[-3:]}")

    def _refresh(self):
        for w in self._gf.winfo_children(): w.destroy()
        self._lbl_m.configure(text=f"{MESES_ES[self._vm-1]}  {self._vy}")
        self._lbl_r.configure(text=self._rango_txt())
        hoy = datetime.date.today()

        for semana in _cal.monthcalendar(self._vy, self._vm):
            row = tk.Frame(self._gf, bg=DARK); row.pack(fill="x", pady=1)
            for ci, dia in enumerate(semana):
                if dia == 0:
                    tk.Label(row, text="", bg=DARK, width=4).pack(side="left", padx=2)
                    continue
                f = datetime.date(self._vy, self._vm, dia)
                en_rango  = (self._start and self._end and
                             self._start <= f <= self._end)
                es_inicio = (f == self._start)
                es_fin    = (f == self._end)
                es_hoy    = (f == hoy)
                es_dom    = (ci == 6)

                if es_inicio or es_fin:
                    bg, fg, bold = ACCENT, TEXT_WHITE, "bold"
                elif en_rango:
                    bg, fg, bold = "#1d4ed8", TEXT_WHITE, "normal"
                elif es_hoy:
                    bg, fg, bold = HEADER_BG, TEXT_WHITE, "bold"
                elif es_dom:
                    bg, fg, bold = DARK, "#f87171", "normal"
                else:
                    bg, fg, bold = DARK, TEXT_WHITE, "normal"

                lbl = tk.Label(row, text=str(dia), bg=bg, fg=fg,
                               font=(FONT, 10, bold), width=4, pady=5, cursor="hand2")
                lbl.pack(side="left", padx=2)
                lbl.bind("<Button-1>", lambda e, d=f: self._click(d))
                lbl.bind("<Enter>",  lambda e, l=lbl: l.configure(bg="#3b82f6"))
                lbl.bind("<Leave>",  lambda e, l=lbl, b=bg: l.configure(bg=b))

    def _click(self, f):
        if self._step == 0:
            self._start, self._end, self._step = f, None, 1
            self._lbl_i.configure(
                text="② Ahora haz clic en el DÍA DE FIN de la semana")
        else:
            if f < self._start: self._start, f = f, self._start
            self._end, self._step = f, 0
            self._lbl_i.configure(
                text="✅ Rango seleccionado — pulsa Confirmar")
        self._refresh()

    def _reset(self):
        self._start = self._end = None; self._step = 0
        self._lbl_i.configure(
            text="① Haz clic en el DÍA DE INICIO de la semana")
        self._refresh()

    def _prev(self):
        self._vm -= 1
        if self._vm == 0: self._vm, self._vy = 12, self._vy - 1
        self._refresh()

    def _next(self):
        self._vm += 1
        if self._vm == 13: self._vm, self._vy = 1, self._vy + 1
        self._refresh()

    def _ok(self):
        s = self._semana_str()
        if not s:
            messagebox.showwarning("Sin selección",
                                   "Selecciona el día de inicio y el día de fin.", parent=self)
            return
        self.var.set(s)
        self.destroy()

