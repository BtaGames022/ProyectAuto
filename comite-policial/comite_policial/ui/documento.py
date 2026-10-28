"""Importación del informe .docx y sincronización de los cambios hacia el documento."""
import datetime
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from comite_policial.config.settings import settings
from comite_policial.config.theme import *
from comite_policial.docio.importador import importar
from comite_policial.docio.escritor import escribir
from comite_policial.ui.widgets import Boton

# clave del documento -> atributo (Treeview) de la aplicación
TREES = {
    "armas": "t_armas", "drogas_g": "t_drogas_g", "drogas_p": "t_drogas_p",
    "homi_a": "t_homi_a", "homi_b": "t_homi_b", "homi_c": "t_homi_c",
    "migr": "t_migr", "flujo": "t_flujo",
    "veh_e": "t_veh_e", "veh_t": "t_veh_t", "brisex": "t_brisex",
    "bandas": "t_bandas", "det_a": "t_det_a", "det_b": "t_det_b",
    "det_c": "t_det_c",
}


NOMBRES_TABLAS = {
    "armas": "I. Armas", "drogas_g": "II. Droga en gramos", "drogas_p": "II. Droga en plantas",
    "homi_a": "III-A. Homicidios por brigada", "homi_b": "III-B. Modalidad",
    "homi_c": "III-C. Resultado", "migr": "IV. Migraciones", "flujo": "IV-B. Flujo migratorio",
    "veh_e": "V. Vehículos por estado", "veh_t": "V. Vehículos por tipo",
    "brisex": "VI. Delitos sexuales", "bandas": "VII. Bandas",
    "det_a": "VIII-A. Detenidos por delito", "det_b": "VIII-B. Detenidos por prefectura",
    "det_c": "VIII-C. Acumulado del año",
}


class DocumentoMixin:
    """Importar un informe existente y mantenerlo actualizado con los cambios."""

    # ── Estado ───────────────────────────────────────────────────────────────
    def _doc_init(self):
        self._doc = None               # dict con path, bytes base y estado importado
        self._doc_after = None
        self._importando = False
        self.v_sync = tk.BooleanVar(value=True)
        self.v_docname = tk.StringVar(value="Sin documento importado")

    # ── Interfaz (panel fijo al pie del menú lateral) ────────────────────────
    def _build_doc_panel(self, sb):
        caja = tk.Frame(sb, bg=NAVY_DEEP)
        caja.pack(side="bottom", fill="x")
        tk.Frame(caja, bg="#1D3470", height=1).pack(fill="x", padx=18)
        tk.Label(caja, text="DOCUMENTO WORD", bg=NAVY_DEEP, fg="#9FB1DC",
                 font=(FONT, 8, "bold"), anchor="w", padx=18
                 ).pack(fill="x", pady=(10, 4))
        box = tk.Frame(caja, bg=NAVY_DEEP)
        box.pack(fill="x", padx=16, pady=(0, 14))
        Boton(box, "Importar informe (.docx)", self._doc_importar, "side",
              chico=True).pack(fill="x", pady=2)
        Boton(box, "Guardar en el documento",
              lambda: self._doc_guardar(auto=False), "side", chico=True
              ).pack(fill="x", pady=2)
        ttk.Checkbutton(box, text="Sincronizar automáticamente", variable=self.v_sync,
                        command=self._doc_cambio, style="TCheckbutton"
                        ).pack(fill="x", pady=(6, 0))
        tk.Label(box, textvariable=self.v_docname, bg=NAVY_DEEP, fg="#9FB1DC",
                 font=(FONT, 8, "italic"), anchor="w", wraplength=200,
                 justify="left").pack(fill="x", pady=(2, 0))

    # ── Importar ─────────────────────────────────────────────────────────────
    def _doc_importar(self):
        ruta = filedialog.askopenfilename(
            title="Seleccione el informe semanal (.docx)",
            filetypes=[("Documento Word", "*.docx")])
        if not ruta:
            return
        try:
            with open(ruta, "rb") as fh:
                base = fh.read()
            estado = importar(base)
        except PermissionError:
            messagebox.showerror("No se pudo importar",
                                 "No hay permiso para leer el archivo. "
                                 "Revisa que no esté protegido o en uso.")
            return
        except FileNotFoundError:
            messagebox.showerror("No se pudo importar", "El archivo ya no existe.")
            return
        except Exception as exc:
            messagebox.showerror(
                "No se pudo importar",
                "El archivo no es un informe .docx válido (o está dañado).\n\n"
                f"Detalle: {type(exc).__name__}: {exc}")
            return
        if not estado["tablas"] and not estado["recuadros"]:
            messagebox.showwarning(
                "Documento sin datos reconocibles",
                "No se encontró ninguna tabla ni recuadro del informe en ese archivo. "
                "Verifica que sea el informe semanal del comité.")
            return

        # Respaldo del original antes de tocarlo
        respaldo = ""
        try:
            carpeta = os.path.join(settings.output_dir, "respaldos")
            os.makedirs(carpeta, exist_ok=True)
            nombre = os.path.splitext(os.path.basename(ruta))[0]
            respaldo = os.path.join(
                carpeta, f"{nombre}_{datetime.datetime.now():%Y%m%d_%H%M%S}.docx")
            shutil.copy2(ruta, respaldo)
        except OSError:
            pass

        self._doc = {"path": ruta, "bytes": base, "import": estado}
        self._doc_aplicar(estado)
        self.v_docname.set(f"Documento: {os.path.basename(ruta)}")
        self._switch("armas")
        n_t = len(estado["tablas"]); n_r = len(estado["recuadros"])
        self.v_status.set(f"Documento importado: {n_t} tablas, {n_r} recuadros")
        faltan = [n for k, n in NOMBRES_TABLAS.items()
                  if k not in estado["tablas"] and k != "homi_d"]
        msg = (f"Se cargaron {n_t} tablas y {n_r} recuadros de concurrencias.\n\n"
               f"Los cambios que hagas en la aplicación se escribirán en:\n{ruta}\n\n")
        if respaldo:
            msg += f"Respaldo del original:\n{respaldo}\n\n"
        if faltan:
            msg += ("No se encontraron en el documento (quedan con sus valores "
                    "actuales y no se escriben en el Word):\n• " + "\n• ".join(faltan))
        (messagebox.showwarning if faltan else messagebox.showinfo)(
            "Documento importado", msg.strip())

    def _doc_aplicar(self, estado: dict):
        """Vuelca el estado leído del documento en las vistas de la aplicación."""
        self._importando = True
        try:
            if estado["semana"]:
                self.v_semana.set(estado["semana"])
            if estado["turno"]:
                self.v_turno.set(estado["turno"])
            for key, datos in estado["tablas"].items():
                tree = getattr(self, TREES.get(key, ""), None)
                if tree is None:
                    continue
                n = len(tree["columns"])
                pad = lambda r: (list(r) + [""] * n)[:n]
                tree.clear()
                for r in datos["rows"]:
                    tree.add_row(pad(r))
                if datos.get("total"):
                    tree.add_total(pad(datos["total"]))
                if datos.get("fuente"):
                    tree.add_fuente(datos["fuente"], n)
            self.recuadros = estado["recuadros"]
            self._rec_refrescar_lista()
            self._rec_nuevo()
            self.txt_x.delete("1.0", "end")
            self.txt_x.insert("1.0", estado["texto_x"])
        finally:
            self._importando = False

    # ── Estado actual de la app (para escribir) ──────────────────────────────
    def _doc_estado(self) -> dict:
        tablas = {}
        for key, attr in TREES.items():
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
            tablas[key] = {"rows": rows, "total": total}
        return {
            "semana": self.v_semana.get().strip(),
            "turno": self.v_turno.get().strip(),
            "tablas": tablas,
            "recuadros": self.recuadros,
            "texto_x": self.txt_x.get("1.0", "end-1c"),
        }

    # ── Sincronización ───────────────────────────────────────────────────────
    def _doc_cambio(self, *_):
        """Se llama ante cualquier cambio de datos; guarda con un breve retardo."""
        if hasattr(self, "_nav"):
            self._nav_estado_later()
        if self._doc is None or self._importando or not self.v_sync.get():
            return
        if self._doc_after:
            self.after_cancel(self._doc_after)
        self._doc_after = self.after(700, lambda: self._doc_guardar(auto=True))

    def _doc_guardar(self, auto=False):
        if self._doc is None:
            if not auto:
                messagebox.showinfo("Sin documento",
                                    "Primero importa un informe (.docx).")
            return
        self._doc_after = None
        d = self._doc
        try:
            escribir(d["bytes"], self._doc_estado(), d["import"], d["path"])
            self.v_status.set(
                f"Documento actualizado ({datetime.datetime.now():%H:%M:%S})")
        except PermissionError:
            msg = ("No se pudo escribir el documento: está abierto en Word. "
                   "Ciérralo y vuelve a guardar.")
            self.v_status.set(msg)
            if not auto:
                messagebox.showwarning("Documento en uso", msg)
        except Exception as exc:
            self.v_status.set(f"Error al guardar en el documento: {exc}")
            if not auto:
                messagebox.showerror("Error al guardar", f"{type(exc).__name__}: {exc}")
