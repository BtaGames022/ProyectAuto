"""Exportación a Microsoft PowerPoint (.pptx) con el formato del comité."""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from comite_policial.config.settings import settings
from comite_policial.docio.semana import SEMANA_APP_RE
from comite_policial.config.theme import *
from comite_policial.exporters.pptx_informe import SLOTS_FOTOS, exportar_informe
from comite_policial.ui.documento import TREES
from comite_policial.ui.widgets import Boton, card, section_title

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
                "No se encontró la plantilla del informe:\n"
                f"{settings.plantilla_pptx}\n\nDebe estar en la carpeta «plantillas» "
                "junto a la aplicación.")
            return
        semana = self.v_semana.get().strip()
        if not SEMANA_APP_RE.fullmatch(semana):
            if not messagebox.askyesno(
                    "Semana operativa",
                    f"La semana «{semana or 'vacía'}» no tiene el formato "
                    "«31.AGO AL 06.SEP.026».\n\n¿Exportar igual?"):
                return
        try:
            os.makedirs(settings.output_dir, exist_ok=True)
        except OSError:
            pass
        ruta = filedialog.asksaveasfilename(
            title="Exportar informe a PowerPoint", defaultextension=".pptx",
            initialdir=settings.output_dir if os.path.isdir(settings.output_dir) else None,
            initialfile=f"COMITE POLICIAL {semana}.pptx".replace("/", "-"),
            filetypes=[("Presentación de PowerPoint", "*.pptx")])
        if not ruta:
            return
        self.v_status.set("Generando el PowerPoint…")
        self.update_idletasks()
        try:
            ruta, avisos = self._gen_pptx(os.path.dirname(ruta), os.path.basename(ruta))
        except PermissionError:
            self.v_status.set("No se pudo guardar el PowerPoint")
            messagebox.showwarning(
                "No se pudo guardar",
                "El archivo está abierto en PowerPoint o la carpeta no permite escribir.\n"
                "Ciérralo o elige otra ubicación y vuelve a exportar.")
            return
        except Exception as exc:
            self.v_status.set(f"Error al exportar: {exc}")
            messagebox.showerror(
                "No se pudo exportar",
                f"{type(exc).__name__}: {exc}\n\nNo se modificó ningún archivo.")
            return
        self.v_status.set(f"PowerPoint exportado: {os.path.basename(ruta)}")
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
        win.configure(bg=BG)
        win.transient(self)
        win.resizable(False, False)
        tk.Label(win, text="Fotos de las láminas", bg=BG, fg=NAVY,
                 font=(FONT, SIZE_H2, "bold")).pack(anchor="w", padx=20, pady=(18, 2))
        tk.Label(win, text="Cada foto se recorta para llenar su marco, como en el informe. "
                 "Un marco sin foto queda vacío.\nLas fotos de las diligencias se agregan "
                 "en cada recuadro (sección IX).",
                 bg=BG, fg=INK_MID, font=(FONT, SIZE_SMALL), justify="left"
                 ).pack(anchor="w", padx=20, pady=(0, 10))

        borde = tk.Frame(win, bg=LINE)
        borde.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        marco = tk.Frame(borde, bg=SURFACE, padx=14, pady=10)
        marco.pack(fill="both", expand=True, padx=1, pady=1)
        nombres = {}

        def refrescar(clave):
            ruta = self.fotos_pptx.get(clave)
            nombres[clave].set(os.path.basename(ruta) if ruta else "Sin foto")

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
            tk.Label(marco, text=etiqueta, bg=SURFACE, fg=INK, font=(FONT, SIZE_BASE),
                     anchor="w", width=40).grid(row=i, column=0, sticky="w", pady=3)
            nombres[clave] = tk.StringVar()
            tk.Label(marco, textvariable=nombres[clave], bg=SURFACE, fg=INK_MID,
                     font=(FONT, SIZE_SMALL), anchor="w", width=30
                     ).grid(row=i, column=1, sticky="w", padx=8)
            Boton(marco, "Elegir…", lambda c=clave: elegir(c), "secondary",
                  chico=True).grid(row=i, column=2, padx=2)
            Boton(marco, "Quitar", lambda c=clave: quitar(c), "ghost",
                  chico=True).grid(row=i, column=3, padx=2)
            refrescar(clave)
        Boton(win, "Cerrar", win.destroy, "primary").pack(anchor="e", padx=20, pady=(0, 18))
        win.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_rooty() + max(20, (self.winfo_height() - win.winfo_height()) // 3)
        win.geometry(f"+{x}+{y}")
