"""Diálogo «Fotos» del PowerPoint: una tarjeta por lámina con su miniatura.

Cada lámina del informe muestra sus marcos de foto numerados, tal como están en la
plantilla. Un clic en un marco (o en «Elegir…») asigna la foto, que se previsualiza
recortada como saldrá en la lámina. «Elegir varias…» reparte un lote de fotos en
orden. Las fotos asignadas viven en `app.fotos_pptx` (clave de marco -> ruta) y solo
durante la sesión.
"""
import base64
import io
import os
import tkinter as tk
from itertools import groupby
from tkinter import filedialog, messagebox

from PIL import Image, ImageOps

from comite_policial.config.theme import *
from comite_policial.exporters.pptx_informe import ALTO_LAMINA, ANCHO_LAMINA
from comite_policial.ui.scroll import ScrollFrame
from comite_policial.ui.widgets import Boton, card

IMAGENES = [("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"), ("Todos", "*.*")]
ANCHO_MINI = 400
ALTO_MINI = round(ANCHO_MINI * ALTO_LAMINA / ANCHO_LAMINA)


def miniatura_png(ruta: str, w: int, h: int) -> bytes:
    """La foto recortada para llenar un marco de w x h px (igual que en el informe)."""
    with Image.open(ruta) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((max(w, h) * 2, max(w, h) * 2))
        if im.width / im.height > w / h:
            ancho = round(im.height * w / h)
            x0 = (im.width - ancho) // 2
            caja = (x0, 0, x0 + ancho, im.height)
        else:
            alto = round(im.width * h / w)
            y0 = (im.height - alto) // 2
            caja = (0, y0, im.width, y0 + alto)
        buf = io.BytesIO()
        im.crop(caja).resize((w, h), Image.LANCZOS).save(buf, "PNG")
        return buf.getvalue()


class DialogoFotos(tk.Toplevel):
    def __init__(self, app, marcos, fotos: dict):
        super().__init__(app)
        self.app, self.marcos, self.fotos = app, marcos, fotos
        self.title("Fotos del PowerPoint")
        self.configure(bg=BG)
        self.transient(app)
        self.minsize(760, 480)
        self._imgs = {}              # clave -> PhotoImage (Tk no guarda la referencia)
        self._lienzos = {}           # n.º de lámina -> Canvas
        self._nombres = {}           # clave -> StringVar con el nombre del archivo
        self._por_lamina = {k: sorted(g, key=lambda m: m.indice)
                            for k, g in groupby(marcos, key=lambda m: m.lamina)}
        self.v_resumen = tk.StringVar()
        self._construir()
        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1040, sw - 60), min(760, sh - 100)
        x = max(0, min(app.winfo_rootx() + (app.winfo_width() - w) // 2, sw - w))
        self.geometry(f"{w}x{h}+{x}+{max(0, min(app.winfo_rooty() + 30, sh - h - 40))}")

    # ── Estructura ───────────────────────────────────────────────────────────
    def _construir(self):
        cab = tk.Frame(self, bg=BG)
        cab.pack(fill="x", padx=20, pady=(16, 6))
        tk.Label(cab, text="Fotos de las láminas", bg=BG, fg=NAVY,
                 font=(FONT, SIZE_H2, "bold")).pack(anchor="w")
        tk.Label(cab, text="Haz clic en un marco numerado (o en «Elegir…») para asignar su foto. "
                 "Cada foto se recorta para llenar el marco, como se verá en el informe; "
                 "un marco sin foto queda vacío.\nLas fotos de las diligencias se agregan en "
                 "cada recuadro (sección X). Las fotos no se guardan al cerrar la aplicación.",
                 bg=BG, fg=INK_MID, font=(FONT, SIZE_SMALL), justify="left"
                 ).pack(anchor="w", pady=(2, 0))

        barra = tk.Frame(self, bg=BG)
        barra.pack(side="bottom", fill="x", padx=20, pady=(6, 14))
        tk.Label(barra, textvariable=self.v_resumen, bg=BG, fg=INK_MID,
                 font=(FONT, SIZE_SMALL)).pack(side="left")
        Boton(barra, "Listo", self.destroy, "primary").pack(side="right")
        Boton(barra, "Quitar todas", self.quitar_todas, "danger").pack(side="right", padx=(0, 8))

        self._panel = ScrollFrame(self, bg=BG)
        self._panel.pack(fill="both", expand=True, pady=(6, 0))
        for lamina, grupo in self._por_lamina.items():
            self._tarjeta(self._panel.interior, lamina, grupo)
        self._refrescar_todo()

    def _tarjeta(self, padre, lamina, grupo):
        titulo = f"{grupo[0].seccion}   ·   lámina {lamina}"
        _, c = card(padre, titulo)
        izq = tk.Frame(c, bg=SURFACE)
        izq.pack(side="left", anchor="n")
        lienzo = tk.Canvas(izq, width=ANCHO_MINI, height=ALTO_MINI, bg=SKY_SOFT,
                           highlightthickness=1, highlightbackground=LINE_STRONG)
        lienzo.pack()
        self._lienzos[lamina] = lienzo

        der = tk.Frame(c, bg=SURFACE)
        der.pack(side="left", fill="both", expand=True, padx=(18, 0))
        for m in grupo:
            fila = tk.Frame(der, bg=SURFACE)
            fila.pack(fill="x", pady=2)
            tk.Label(fila, text=str(m.indice), bg=GOLD, fg=NAVY, width=2,
                     font=(FONT, SIZE_BASE, "bold")).pack(side="left")
            self._nombres[m.clave] = tk.StringVar()
            tk.Label(fila, textvariable=self._nombres[m.clave], bg=SURFACE, fg=INK,
                     font=(FONT, SIZE_SMALL), anchor="w", width=30
                     ).pack(side="left", padx=8)
            Boton(fila, "Elegir…", lambda mm=m: self.elegir(mm), "secondary", chico=True
                  ).pack(side="left", padx=2)
            Boton(fila, "Quitar", lambda k=m.clave: self.quitar(k), "ghost", chico=True
                  ).pack(side="left", padx=2)
        if len(grupo) > 1:
            Boton(der, "Elegir varias…", lambda g=grupo: self.elegir_varias(g), "secondary",
                  chico=True).pack(anchor="w", pady=(8, 0))

    # ── Dibujo ───────────────────────────────────────────────────────────────
    def _pintar(self, lamina):
        lienzo = self._lienzos[lamina]
        lienzo.delete("all")
        k = ANCHO_MINI / ANCHO_LAMINA
        for m in self._por_lamina[lamina]:
            x, y = round(m.x * k), round(m.y * k)
            w, h = max(round(m.w * k), 6), max(round(m.h * k), 6)
            tag = f"m{m.indice}"
            ruta = self.fotos.get(m.clave)
            img = None
            if ruta:
                try:
                    img = tk.PhotoImage(data=base64.b64encode(miniatura_png(ruta, w, h)))
                    self._imgs[m.clave] = img
                except Exception:
                    img = None
            if img is not None:
                lienzo.create_image(x, y, image=img, anchor="nw", tags=(tag,))
                lienzo.create_rectangle(x, y, x + w, y + h, outline=NAVY, width=1, tags=(tag,))
            else:
                lienzo.create_rectangle(x, y, x + w, y + h, outline=NAVY, fill=SKY, width=1,
                                        dash=(3, 2), tags=(tag,))
                if ruta:        # la foto asignada ya no se puede leer
                    lienzo.create_text(x + w / 2, y + h / 2, text="?", fill=ERR,
                                       font=(FONT, SIZE_BASE, "bold"), tags=(tag,))
            lienzo.create_rectangle(x, y, x + 14, y + 14, fill=GOLD, outline=NAVY, tags=(tag,))
            lienzo.create_text(x + 7, y + 7, text=str(m.indice), fill=NAVY,
                               font=(FONT, 8, "bold"), tags=(tag,))
            lienzo.tag_bind(tag, "<Button-1>", lambda e, mm=m: self.elegir(mm))
            lienzo.tag_bind(tag, "<Enter>", lambda e, c=lienzo: c.configure(cursor="hand2"))
            lienzo.tag_bind(tag, "<Leave>", lambda e, c=lienzo: c.configure(cursor=""))

    def _refrescar_todo(self):
        for m in self.marcos:
            ruta = self.fotos.get(m.clave)
            self._nombres[m.clave].set(os.path.basename(ruta) if ruta else "Sin foto")
        for lamina in self._por_lamina:
            self._pintar(lamina)
        n = sum(1 for m in self.marcos if self.fotos.get(m.clave))
        self.v_resumen.set(f"{n} de {len(self.marcos)} marcos con foto")

    # ── Acciones (también las usan las pruebas) ──────────────────────────────
    def asignar(self, clave: str, ruta: str):
        self.fotos[clave] = ruta
        self._refrescar_todo()

    def quitar(self, clave: str):
        self.fotos.pop(clave, None)
        self._refrescar_todo()

    def quitar_todas(self):
        for m in self.marcos:
            self.fotos.pop(m.clave, None)
        self._refrescar_todo()

    def elegir(self, marco):
        ruta = filedialog.askopenfilename(parent=self, filetypes=IMAGENES,
                                          title=f"Foto para {marco.etiqueta}")
        if ruta:
            self.asignar(marco.clave, ruta)

    def elegir_varias(self, grupo):
        """Reparte las fotos elegidas, en orden, desde el primer marco libre."""
        rutas = filedialog.askopenfilenames(parent=self, filetypes=IMAGENES,
                                            title=f"Fotos para {grupo[0].seccion} (en orden)")
        if not rutas:
            return
        libres = [m for m in grupo if not self.fotos.get(m.clave)] or list(grupo)
        for m, ruta in zip(libres, rutas):
            self.fotos[m.clave] = ruta
        self._refrescar_todo()
        if len(rutas) > len(libres):
            messagebox.showinfo(
                "Fotos de más", f"Esta lámina tiene {len(libres)} marcos libres: "
                f"quedaron sin usar {len(rutas) - len(libres)} foto(s).", parent=self)
