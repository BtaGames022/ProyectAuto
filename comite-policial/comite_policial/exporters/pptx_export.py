"""Exportación a Microsoft PowerPoint (.pptx) con el formato del comité."""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from comite_policial.config.settings import settings
from comite_policial.config.theme import *
from comite_policial.exporters.pptx_informe import SLOTS_FOTOS, exportar_informe
from comite_policial.ui.documento import TREES
from comite_policial.ui.widgets import action_btn

IMAGENES = [("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"), ("Todos", "*.*")]


class PptxExporterMixin:
    """Botón «Exportar PowerPoint»: reproduce el informe semanal del comité."""

    # ── Datos ────────────────────────────────────────────────────────────────
    def _tablas_export(self) -> dict:
        """Filas y total de cada tabla de la aplicación (incluye III-D)."""
        mapa = dict(TREES, homi_d="t_homi_d")
        tablas = {}
        for clave, attr in mapa.items():
            tree = getattr(self, attr, None)
            if tree is None:
                continue
            rows, total = [], None
            for r in tree.get_children():
                tags = tree.item(r, "tags")
                if "fuente" in tags:
                    continue
                vals = tree.raw_values(r)
                if "total" in tags:
                    total = vals
                else:
                    rows.append(vals)
            tablas[clave] = {"rows": rows, "total": total}
        return tablas

    def _estado_pptx(self) -> dict:
        return {"semana": self.v_semana.get().strip(),
                "tablas": self._tablas_export(),
                # solo los recuadros marcados con la casilla PPT (por defecto, todos)
                "recuadros": [r for r in getattr(self, "recuadros", [])
                              if r.get("en_pptx", True)]}

    def _gen_pptx(self, directorio: str, nombre: str = "") -> tuple:
        """Genera el .pptx en `directorio`; devuelve (ruta, avisos)."""
        estado = self._estado_pptx()
        nombre = nombre or f"COMITE POLICIAL {estado['semana']}.pptx"
        ruta = os.path.join(directorio, nombre)
        avisos = exportar_informe(
            estado, ruta, settings.plantilla_pptx,
            fotos=getattr(self, "fotos_pptx", {}),
            ano_ant=settings.ano_ant, ano_act=settings.ano_act,
            periodo=settings.periodo)
        return ruta, avisos

    # ── Botón ────────────────────────────────────────────────────────────────
    def _exportar_pptx(self):
        if not os.path.isfile(settings.plantilla_pptx):
            messagebox.showerror(
                "Falta la plantilla",
                f"No se encontró la plantilla del informe:\n{settings.plantilla_pptx}")
            return
        os.makedirs(settings.output_dir, exist_ok=True)
        inicial = f"COMITE POLICIAL {self.v_semana.get().strip()}.pptx"
        ruta = filedialog.asksaveasfilename(
            title="Exportar informe a PowerPoint", defaultextension=".pptx",
            initialdir=settings.output_dir, initialfile=inicial,
            filetypes=[("Presentación de PowerPoint", "*.pptx")])
        if not ruta:
            return
        self.v_status.set("⏳  Generando el PowerPoint…")
        self.update_idletasks()
        try:
            ruta, avisos = self._gen_pptx(os.path.dirname(ruta), os.path.basename(ruta))
        except PermissionError:
            self.v_status.set("⚠️  No se pudo escribir: el archivo está abierto en PowerPoint")
            messagebox.showwarning(
                "Archivo en uso",
                "No se pudo guardar: ese PowerPoint está abierto. Ciérralo y reintenta.")
            return
        except Exception as exc:
            self.v_status.set(f"❌  Error al exportar: {exc}")
            messagebox.showerror("Error al exportar", f"{type(exc).__name__}: {exc}")
            return
        self.v_status.set(f"✅  PowerPoint exportado: {os.path.basename(ruta)}")
        msg = f"Informe exportado en:\n{ruta}"
        if avisos:
            msg += "\n\nRevisar:\n• " + "\n• ".join(avisos)
        if messagebox.askyesno("Exportación lista", msg + "\n\n¿Abrirlo ahora?"):
            self._abrir_archivo(ruta)

    @staticmethod
    def _abrir_archivo(ruta: str):
        try:
            if hasattr(os, "startfile"):
                os.startfile(ruta)                       # Windows
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", ruta])
        except OSError:
            pass

    # ── Fotos de las láminas ─────────────────────────────────────────────────
    def _fotos_dialogo(self):
        if not hasattr(self, "fotos_pptx"):
            self.fotos_pptx = {}
        win = tk.Toplevel(self)
        win.title("Fotos del PowerPoint")
        win.configure(bg=LIGHT_BG)
        win.transient(self)
        tk.Label(win, text="Fotos de las láminas del PowerPoint", bg=LIGHT_BG,
                 fg=TEXT_DARK, font=(FONT, 11, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(win, text="Cada foto se recorta para llenar su marco, como en el informe.\n"
                 "Los marcos sin foto quedan vacíos. Las fotos de las diligencias "
                 "se agregan en cada recuadro (sección IX).",
                 bg=LIGHT_BG, fg=TEXT_MID, font=(FONT, 9), justify="left"
                 ).pack(anchor="w", padx=16, pady=(0, 8))

        marco = tk.Frame(win, bg=CARD_BG, padx=10, pady=8)
        marco.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        nombres = {}

        def refrescar(clave):
            ruta = self.fotos_pptx.get(clave)
            nombres[clave].set(os.path.basename(ruta) if ruta else "— sin foto —")

        def elegir(clave):
            ruta = filedialog.askopenfilename(parent=win, title="Elegir foto",
                                              filetypes=IMAGENES)
            if ruta:
                self.fotos_pptx[clave] = ruta
                refrescar(clave)

        def quitar(clave):
            self.fotos_pptx.pop(clave, None)
            refrescar(clave)

        for i, (clave, etiqueta, _lam, _geo) in enumerate(SLOTS_FOTOS):
            tk.Label(marco, text=etiqueta, bg=CARD_BG, fg=TEXT_DARK, font=(FONT, 9),
                     anchor="w", width=38).grid(row=i, column=0, sticky="w", pady=2)
            nombres[clave] = tk.StringVar()
            tk.Label(marco, textvariable=nombres[clave], bg=CARD_BG, fg=TEXT_MID,
                     font=(FONT, 9, "italic"), anchor="w", width=28
                     ).grid(row=i, column=1, sticky="w", padx=8)
            action_btn(marco, "Elegir…", lambda c=clave: elegir(c), ACCENT,
                       small=True).grid(row=i, column=2, padx=2)
            action_btn(marco, "Quitar", lambda c=clave: quitar(c), TEXT_MID,
                       small=True).grid(row=i, column=3, padx=2)
            refrescar(clave)
        action_btn(win, "Cerrar", win.destroy, SUCCESS).pack(pady=(0, 14))
