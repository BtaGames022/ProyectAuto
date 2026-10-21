"""Importación del informe .docx y sincronización de los cambios hacia el documento."""
import datetime
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from comite_policial.config.settings import settings
from comite_policial.config.theme import *
from comite_policial.docio.importador import importar
from comite_policial.docio.escritor import escribir
from comite_policial.ui.widgets import action_btn

# clave del documento -> atributo (Treeview) de la aplicación
TREES = {
    "armas": "t_armas", "drogas_g": "t_drogas_g", "drogas_p": "t_drogas_p",
    "homi_a": "t_homi_a", "homi_b": "t_homi_b", "homi_c": "t_homi_c",
    "migr": "t_migr", "flujo": "t_flujo",
    "veh_e": "t_veh_e", "veh_t": "t_veh_t", "brisex": "t_brisex",
    "bandas": "t_bandas", "det_a": "t_det_a", "det_b": "t_det_b",
    "det_c": "t_det_c",
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

    # ── Interfaz (panel del menú lateral) ────────────────────────────────────
    def _build_doc_panel(self, sb):
        tk.Label(sb, text="DOCUMENTO", bg=SIDEBAR_BG, fg="#4a7ab5",
                 font=(FONT, 7, "bold"), padx=10, anchor="w").pack(fill="x", pady=(8, 2))
        box = tk.Frame(sb, bg=SIDEBAR_BG)
        box.pack(fill="x", padx=8)
        action_btn(box, "📂  Importar informe (.docx)", self._doc_importar,
                   ACCENT, small=True).pack(fill="x", pady=2)
        action_btn(box, "💾  Guardar en el documento",
                   lambda: self._doc_guardar(auto=False), TEAL, small=True
                   ).pack(fill="x", pady=2)
        tk.Checkbutton(box, text="Sincronizar automáticamente", variable=self.v_sync,
                       bg=SIDEBAR_BG, fg=TEXT_LIGHT, selectcolor=SIDEBAR_HOV,
                       activebackground=SIDEBAR_BG, activeforeground=TEXT_WHITE,
                       font=(FONT, 8), anchor="w",
                       command=self._doc_cambio).pack(fill="x", pady=(4, 0))
        tk.Label(box, textvariable=self.v_docname, bg=SIDEBAR_BG, fg="#7aa2d6",
                 font=(FONT, 8, "italic"), anchor="w", wraplength=210,
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
        except Exception as exc:
            messagebox.showerror("No se pudo importar", f"{type(exc).__name__}: {exc}")
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
        self.v_docname.set(f"📄 {os.path.basename(ruta)}")
        self._switch("armas")
        n_t = len(estado["tablas"]); n_r = len(estado["recuadros"])
        self.v_status.set(f"✅  Documento importado: {n_t} tablas, {n_r} recuadros")
        messagebox.showinfo(
            "Documento importado",
            f"Se cargaron {n_t} tablas y {n_r} recuadros de concurrencias.\n\n"
            f"Los cambios que hagas en la aplicación se escribirán en:\n{ruta}\n\n"
            + (f"Respaldo del original:\n{respaldo}" if respaldo else ""))

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
                f"✅  Documento actualizado ({datetime.datetime.now():%H:%M:%S})")
        except PermissionError:
            msg = ("No se pudo escribir el documento: está abierto en Word. "
                   "Ciérralo y vuelve a guardar.")
            self.v_status.set("⚠️  " + msg)
            if not auto:
                messagebox.showwarning("Documento en uso", msg)
        except Exception as exc:
            self.v_status.set(f"❌  Error al guardar en el documento: {exc}")
            if not auto:
                messagebox.showerror("Error al guardar", f"{type(exc).__name__}: {exc}")
