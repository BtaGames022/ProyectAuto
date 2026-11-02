"""Editor del detalle del recuadro: fecha con calendario + un cuadro de texto como el del Word."""
import tkinter as tk

from comite_policial.config.catalogos import TIPOS_RECUADRO
from comite_policial.config.theme import *
from comite_policial.core import tipos_personalizados
from comite_policial.core.detalle import (
    marcas_fijo, marcas_libre, parse_fijo, parse_libre, render)
from comite_policial.core.recuadros import TIPO_IMPORTADO
from comite_policial.ui.widgets import Boton

AYUDA_FIJO = ("Escribe a continuación de cada rótulo. Los rótulos en negrita son del "
              "formato de este tipo y no se pueden borrar.")
AYUDA_LIBRE = ("Tipo libre: puedes escribir y modificar todo el cuadro. Una línea "
               "«Rótulo: texto» crea un campo en negrita.")


class _Campo:
    """Acceso a un campo del cuadro con la interfaz mínima de un Entry/Text
    (get, insert, delete). Sirve a pruebas y a código que edita un campo por nombre."""

    def __init__(self, app, etiqueta, clase):
        self._app, self._etq, self._clase = app, etiqueta, clase

    def winfo_class(self):
        return "Text"

    def get(self, *_a):
        return self._app._rec_det_leer()[1].get(self._etq, "")

    def insert(self, indice, texto):
        v = self.get()
        self._app._rec_det_campo(self._etq, self._clase,
                                 texto + v if str(indice) in ("0", "1.0") else v + texto)

    def delete(self, *_a):
        self._app._rec_det_campo(self._etq, self._clase, "")

    def __getattr__(self, nombre):
        return getattr(self._app._rec_texto, nombre)


class DetalleMixin:

    # ── Tipos ────────────────────────────────────────────────────────────────
    @staticmethod
    def _rec_es_libre(tipo: str) -> bool:
        return tipo == TIPO_IMPORTADO or tipo in tipos_personalizados.PERSONALIZADOS

    def _rec_defs_tipo(self, tipo: str) -> list:
        """Campos (etiqueta, clase) que define el tipo, sin mirar lo escrito."""
        if tipo == TIPO_IMPORTADO:
            return [tuple(d) for d in self._rec_defs_imp]
        return [tuple(d) for d in TIPOS_RECUADRO.get(tipo, [])]

    def _rec_defs_actuales(self) -> list:
        """Campos del formulario tal como están ahora (en modo libre, los del texto)."""
        if getattr(self, "_rec_det_libre", False) and hasattr(self, "_rec_texto"):
            defs = self._rec_det_leer()[0]
            fecha = [("Fecha", "linea")] if "Fecha" in self._rec_widgets else []
            return fecha + list(defs)
        return self._rec_defs_tipo(self._rec_tipo_valido())

    # ── Construcción ─────────────────────────────────────────────────────────
    def _rec_construir_campos(self, valores: dict, fecha_def="", defs=None):
        cf = self._rec_campos_frame
        for w in cf.winfo_children():
            w.destroy()
        self._rec_widgets = {}
        tipo = self._rec_tipo_valido()
        defs = [tuple(d) for d in (defs or self._rec_defs_tipo(tipo))]
        self._rec_det_libre = self._rec_es_libre(tipo)
        cf.grid_columnconfigure(1, weight=1)

        fila = 0
        if any(e == "Fecha" for e, _ in defs):
            tk.Label(cf, text="Fecha", bg=SURFACE, fg=INK, font=(FONT, SIZE_SMALL, "bold"),
                     anchor="ne", justify="right", width=20
                     ).grid(row=0, column=0, sticky="ne", padx=(0, 10), pady=(7, 0))
            ent = tk.Entry(cf)
            ent.insert(0, valores.get("Fecha", fecha_def))
            ent.grid(row=0, column=1, sticky="ew", pady=3, ipady=4)
            ent.bind("<KeyRelease>", lambda e: self._rec_preview_later())
            Boton(cf, "Calendario", lambda en=ent: self._rec_elegir_fecha(en),
                  "secondary", chico=True).grid(row=0, column=2, padx=(8, 0))
            self._rec_widgets["Fecha"] = (ent, "linea")
            fila = 1

        self._rec_det_defs = [d for d in defs if d[0] != "Fecha"]
        caja = tk.Text(cf, height=12, wrap="word", font=(FONT, SIZE_BASE), undo=True,
                       relief="solid", bd=1, padx=10, pady=8)
        caja.grid(row=fila, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        caja.tag_configure("etq", font=(FONT, SIZE_BASE, "bold"))
        caja.bind("<<Modified>>", self._rec_det_modificado)
        caja.bind("<Map>", self._rec_det_ajustar)
        self._rec_texto = caja
        self._rec_det_ayuda = tk.StringVar(value=AYUDA_LIBRE if self._rec_det_libre
                                           else AYUDA_FIJO)
        tk.Label(cf, textvariable=self._rec_det_ayuda, bg=SURFACE, fg=INK_SOFT,
                 font=(FONT, SIZE_SMALL), anchor="w", justify="left", wraplength=620
                 ).grid(row=fila + 1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        self._rec_det_poner(self._rec_det_defs, valores)
        self._rec_det_inicial = caja.get("1.0", "end-1c")
        self._rec_det_orig = (list(self._rec_det_defs), dict(valores))
        for e, c in self._rec_det_defs:
            self._rec_widgets[e] = (_Campo(self, e, c), c)

    # ── Texto del cuadro ─────────────────────────────────────────────────────
    def _rec_det_poner(self, defs, campos):
        """Escribe `campos` en el cuadro (carga programática, sin pasar por la guarda)."""
        c = self._rec_texto
        texto = render(defs, campos)
        self._rec_det_ok = texto
        c.delete("1.0", "end")
        c.insert("1.0", texto)
        c.edit_reset()
        c.edit_modified(False)
        self._rec_det_estilo()
        self._rec_det_ajustar()

    def _rec_det_modificado(self, _e=None):
        c = self._rec_texto
        try:
            if not c.edit_modified():
                return
            c.edit_modified(False)
            texto = c.get("1.0", "end-1c")
            if texto == self._rec_det_ok:
                return
            if not self._rec_det_libre and self.rv_tipo.get().strip() != self._rec_tipo_ok:
                self._rec_tipo_escrito()
            if (not self._rec_det_libre
                    and parse_fijo(texto, self._rec_det_defs) is None):
                # tocó un rótulo del formato: se deshace
                cursor = c.index("insert")
                c.delete("1.0", "end")
                c.insert("1.0", self._rec_det_ok)
                try:
                    c.mark_set("insert", cursor)
                except tk.TclError:
                    pass
                c.edit_modified(False)
                return
            self._rec_det_ok = texto
            self._rec_det_estilo()
            self._rec_det_ajustar()
            self._rec_preview_later()
        except tk.TclError:
            pass

    def _rec_det_estilo(self):
        c = self._rec_texto
        c.tag_remove("etq", "1.0", "end")
        texto = c.get("1.0", "end-1c")
        marcas = (marcas_libre(texto) if self._rec_det_libre
                  else marcas_fijo(texto, self._rec_det_defs)) or []
        for ini, fin in marcas:
            c.tag_add("etq", f"1.0+{ini}c", f"1.0+{fin}c")

    def _rec_det_ajustar(self, _e=None):
        c = self._rec_texto
        try:
            lineas = c.get("1.0", "end-1c").count("\n") + 1
            if c.winfo_ismapped() and c.winfo_width() > 50:
                c.update_idletasks()
                n = c.count("1.0", "end", "displaylines")
                if n:
                    lineas = max(lineas, n[0])
            c.configure(height=max(8, min(26, lineas + 1)))
        except tk.TclError:
            pass

    def _rec_det_modo_libre(self):
        """Pasa el cuadro a edición libre sin tocar su contenido (tipo recién escrito)."""
        self._rec_det_libre = True
        self._rec_det_ayuda.set(AYUDA_LIBRE)
        self._rec_det_estilo()

    # ── Lectura ──────────────────────────────────────────────────────────────
    def _rec_det_leer(self):
        """(defs, campos) del cuadro, sin la fecha."""
        if not self._rec_det_libre and self.rv_tipo.get().strip() != self._rec_tipo_ok:
            self._rec_tipo_escrito()
        texto = self._rec_texto.get("1.0", "end-1c")
        if self._rec_det_libre:
            if texto == self._rec_det_inicial and self._rec_det_orig:
                defs, campos = self._rec_det_orig
                return [d for d in defs if d[0] != "Fecha"], dict(campos)
            defs, campos = parse_libre(texto)
            return defs, campos
        return self._rec_det_defs, parse_fijo(texto, self._rec_det_defs) or {}

    def _rec_valores_campos(self) -> dict:
        out = {}
        if "Fecha" in self._rec_widgets:
            out["Fecha"] = self._rec_widgets["Fecha"][0].get().strip()
        out.update({k: v.strip() for k, v in self._rec_det_leer()[1].items()})
        return out

    def _rec_det_campo(self, etiqueta, clase, valor):
        """Cambia el valor de un campo y reescribe el cuadro."""
        defs, campos = self._rec_det_leer()
        defs = list(defs)
        if all(e != etiqueta for e, _ in defs):
            defs.append((etiqueta, clase))
        campos[etiqueta] = valor
        if not self._rec_det_libre:
            self._rec_det_defs = defs
        self._rec_det_poner(defs, campos)
        self._rec_preview_later()
