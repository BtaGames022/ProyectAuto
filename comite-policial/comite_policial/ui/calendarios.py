"""Selectores de fecha y de semana operativa (ventanas con calendario)."""
from __future__ import annotations

import calendar as _cal
import datetime
import tkinter as tk
from tkinter import messagebox

from comite_policial.config.theme import *
from comite_policial.ui.widgets import Boton

MESES_ES    = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
               "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MESES_ABREV = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
               "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
DIAS_ES     = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]


class _Calendario(tk.Toplevel):
    """Base: cabecera con mes, días de la semana y cuadrícula de días."""

    def __init__(self, master, titulo):
        super().__init__(master)
        self.title(titulo)
        self.resizable(False, False)
        self.configure(bg=SURFACE)
        self.transient(master)
        hoy = datetime.date.today()
        self._vy, self._vm = hoy.year, hoy.month
        self._cuerpo = tk.Frame(self, bg=SURFACE)
        self._cuerpo.pack(fill="both", expand=True)

    def _armar_calendario(self, parent):
        cab = tk.Frame(parent, bg=SURFACE)
        cab.pack(fill="x", padx=16, pady=(14, 6))
        Boton(cab, "‹", self._prev, "ghost", chico=True).pack(side="left")
        self._lbl_m = tk.Label(cab, bg=SURFACE, fg=NAVY,
                               font=(FONT, SIZE_H2, "bold"))
        self._lbl_m.pack(side="left", expand=True)
        Boton(cab, "›", self._next, "ghost", chico=True).pack(side="right")

        dias = tk.Frame(parent, bg=SURFACE)
        dias.pack(padx=16)
        for i, d in enumerate(DIAS_ES):
            tk.Label(dias, text=d, bg=SURFACE, width=4, pady=4,
                     fg=ERR if i == 6 else INK_SOFT,
                     font=(FONT, SIZE_SMALL, "bold")).grid(row=0, column=i, padx=1)
        self._gf = tk.Frame(parent, bg=SURFACE)
        self._gf.pack(padx=16, pady=(0, 8))

    def _celda(self, fila, col, dia, bg, fg, negrita, comando):
        lbl = tk.Label(self._gf, text=str(dia), bg=bg, fg=fg, width=4, pady=6,
                       font=(FONT, SIZE_BASE, "bold" if negrita else "normal"),
                       cursor="hand2")
        lbl.grid(row=fila, column=col, padx=1, pady=1)
        lbl.bind("<Button-1>", lambda e: comando())
        lbl.bind("<Enter>", lambda e: lbl.configure(bg=SKY))
        lbl.bind("<Leave>", lambda e: lbl.configure(bg=bg))

    def _limpiar(self):
        for w in self._gf.winfo_children():
            w.destroy()
        self._lbl_m.configure(text=f"{MESES_ES[self._vm - 1]} {self._vy}")

    def _prev(self):
        self._vm -= 1
        if self._vm == 0:
            self._vm, self._vy = 12, self._vy - 1
        self._refresh()

    def _next(self):
        self._vm += 1
        if self._vm == 13:
            self._vm, self._vy = 1, self._vy + 1
        self._refresh()

    def _centrar(self, master):
        self.update_idletasks()
        mx = master.winfo_rootx() + master.winfo_width() // 2
        my = master.winfo_rooty() + master.winfo_height() // 2
        self.geometry(f"+{mx - self.winfo_width() // 2}+{my - self.winfo_height() // 2}")
        self.grab_set()

    @staticmethod
    def _estilo_dia(f, ci, hoy, activo=False, en_rango=False):
        if activo:
            return NAVY, "#FFFFFF", True
        if en_rango:
            return SKY, NAVY, True
        if f == hoy:
            return GOLD_SOFT, INK, True
        return SURFACE, (ERR if ci == 6 else INK), False


class CalendarPopup(_Calendario):
    """Selecciona una fecha (formato dd/mm/aaaa)."""

    def __init__(self, master, var: tk.StringVar, title="Seleccionar fecha"):
        super().__init__(master, title)
        self.var = var
        try:
            d, m, a = (int(x) for x in var.get().split("/"))
            self._sel = datetime.date(a, m, d)
        except Exception:
            self._sel = datetime.date.today()
        self._vy, self._vm = self._sel.year, self._sel.month

        self._armar_calendario(self._cuerpo)
        pie = tk.Frame(self._cuerpo, bg=SURFACE)
        pie.pack(fill="x", padx=16, pady=(0, 14))
        Boton(pie, "Confirmar", self._ok, "primary").pack(side="left")
        Boton(pie, "Cancelar", self.destroy, "secondary").pack(side="left", padx=8)
        Boton(pie, "Hoy", self._hoy, "ghost").pack(side="right")
        self._refresh()
        self._centrar(master)

    def _refresh(self):
        self._limpiar()
        hoy = datetime.date.today()
        for fi, semana in enumerate(_cal.monthcalendar(self._vy, self._vm)):
            for ci, dia in enumerate(semana):
                if not dia:
                    continue
                f = datetime.date(self._vy, self._vm, dia)
                bg, fg, neg = self._estilo_dia(f, ci, hoy, activo=(f == self._sel))
                self._celda(fi, ci, dia, bg, fg, neg, lambda d=f: self._pick(d))

    def _pick(self, f):
        self._sel = f
        self._refresh()

    def _hoy(self):
        self._sel = datetime.date.today()
        self._vy, self._vm = self._sel.year, self._sel.month
        self._refresh()

    def _ok(self):
        self.var.set(self._sel.strftime("%d/%m/%Y"))
        self.destroy()


class WeekRangePopup(_Calendario):
    """Selecciona la semana operativa: clic en el día de inicio y en el de fin.

    Atajos para la semana en curso y la anterior (lunes a domingo).
    Resultado: «DD.MES AL DD.MES.AAA».
    """

    def __init__(self, master, var: tk.StringVar):
        super().__init__(master, "Semana operativa")
        self.var = var
        self._start: datetime.date | None = None
        self._end: datetime.date | None = None
        self._step = 0

        self._lbl_i = tk.Label(self._cuerpo, bg=SKY_SOFT, fg=NAVY, anchor="w",
                               font=(FONT, SIZE_BASE, "bold"), padx=16, pady=9)
        self._lbl_i.pack(fill="x")
        self._lbl_r = tk.Label(self._cuerpo, bg=SURFACE, fg=INK, anchor="w",
                               font=(FONT, SIZE_H2, "bold"), padx=16, pady=8)
        self._lbl_r.pack(fill="x")
        rapidos = tk.Frame(self._cuerpo, bg=SURFACE)
        rapidos.pack(fill="x", padx=16)
        Boton(rapidos, "Esta semana", lambda: self._semana(0), "secondary",
              chico=True).pack(side="left")
        Boton(rapidos, "Semana pasada", lambda: self._semana(-1), "secondary",
              chico=True).pack(side="left", padx=6)

        self._armar_calendario(self._cuerpo)
        pie = tk.Frame(self._cuerpo, bg=SURFACE)
        pie.pack(fill="x", padx=16, pady=(0, 14))
        Boton(pie, "Confirmar", self._ok, "primary").pack(side="left")
        Boton(pie, "Cancelar", self.destroy, "secondary").pack(side="left", padx=8)
        Boton(pie, "Reiniciar", self._reset, "ghost").pack(side="right")
        self._reset()
        self._centrar(master)

    # ── Selección ────────────────────────────────────────────────────────────
    def _semana(self, desplazamiento):
        hoy = datetime.date.today() + datetime.timedelta(weeks=desplazamiento)
        lunes = hoy - datetime.timedelta(days=hoy.weekday())
        self._start, self._end = lunes, lunes + datetime.timedelta(days=6)
        self._step = 0
        self._vy, self._vm = lunes.year, lunes.month
        self._lbl_i.configure(text="Semana elegida. Pulsa Confirmar.")
        self._refresh()

    def _click(self, f):
        if self._step == 0:
            self._start, self._end, self._step = f, None, 1
            self._lbl_i.configure(text="Ahora haz clic en el último día de la semana.")
        else:
            if f < self._start:
                self._start, f = f, self._start
            self._end, self._step = f, 0
            self._lbl_i.configure(text="Rango elegido. Pulsa Confirmar.")
        self._refresh()

    def _reset(self):
        self._start = self._end = None
        self._step = 0
        self._lbl_i.configure(text="Haz clic en el primer día de la semana.")
        self._refresh()

    def _rango_txt(self) -> str:
        if not self._start:
            return "Sin selección"
        s = self._start
        p1 = f"{s.day:02d}.{MESES_ABREV[s.month - 1]}"
        if not self._end:
            return f"{p1} AL …"
        e = self._end
        return f"{p1} AL {e.day:02d}.{MESES_ABREV[e.month - 1]}.{str(e.year)[-3:]}"

    def _semana_str(self) -> str:
        return self._rango_txt() if self._start and self._end else ""

    def _refresh(self):
        self._limpiar()
        self._lbl_r.configure(text=self._rango_txt())
        hoy = datetime.date.today()
        for fi, semana in enumerate(_cal.monthcalendar(self._vy, self._vm)):
            for ci, dia in enumerate(semana):
                if not dia:
                    continue
                f = datetime.date(self._vy, self._vm, dia)
                borde = f in (self._start, self._end)
                en_rango = bool(self._start and self._end and self._start <= f <= self._end)
                bg, fg, neg = self._estilo_dia(f, ci, hoy, activo=borde, en_rango=en_rango)
                self._celda(fi, ci, dia, bg, fg, neg, lambda d=f: self._click(d))

    def _ok(self):
        s = self._semana_str()
        if not s:
            messagebox.showwarning("Sin selección",
                                   "Elige el primer y el último día de la semana.",
                                   parent=self)
            return
        self.var.set(s)
        self.destroy()
