"""Exportación a Microsoft PowerPoint (.pptx) con el formato del comité."""
import os
import subprocess
import sys
from tkinter import filedialog, messagebox

from comite_policial.config.settings import settings
from comite_policial.docio.semana import SEMANA_APP_RE
from comite_policial.exporters.pptx_fotos import leer_marcos
from comite_policial.exporters.pptx_informe import exportar_informe
from comite_policial.ui.documento import TREES
from comite_policial.ui.fotos import DialogoFotos


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
        serv = getattr(self, "v_serv_fechas", None)
        nota = getattr(self, "txt_nota_sint", None)
        estado = {"semana": self.v_semana.get().strip(),
                  "servicio_fechas": serv.get().strip() if serv is not None else ""}
        if nota is not None:
            estado["nota_sintetica"] = nota.get("1.0", "end-1c")
        return {**estado,
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
        """Abre el diálogo de fotos; los marcos salen de la propia plantilla."""
        if not hasattr(self, "fotos_pptx"):
            self.fotos_pptx = {}
        if not os.path.isfile(settings.plantilla_pptx):
            messagebox.showerror("Falta la plantilla",
                                 f"No se encontró la plantilla del informe:\n{settings.plantilla_pptx}")
            return
        try:
            marcos = leer_marcos(settings.plantilla_pptx)
        except Exception as exc:
            messagebox.showerror("No se pudo leer la plantilla", f"{type(exc).__name__}: {exc}")
            return
        if not marcos:
            messagebox.showinfo(
                "La plantilla no tiene marcos de foto",
                "Es de un formato anterior. Regenerarla con "
                "tools/construir_plantilla_pptx.py para poder asignar fotos.")
            return
        return DialogoFotos(self, marcos, self.fotos_pptx)
