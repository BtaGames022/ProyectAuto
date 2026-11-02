"""Diálogo «Fotos» del PowerPoint: una lámina a la vez, con su miniatura.

Igual que el resto de la aplicación (menú lateral + una vista visible, ver
`App._switch`), las láminas no se apilan en una columna con barra de desplazamiento:
a la izquierda va la lista de láminas y a la derecha solo la elegida. Apilar las 8
tarjetas (casi 300 widgets nativos) en un panel desplazable dejaba restos duplicados
en pantalla al arrastrar la barra en algunos equipos Windows.

La lámina elegida muestra sus marcos numerados, tal como están en la plantilla. Un
clic en un marco (o en «Elegir…») asigna la foto, que se previsualiza recortada como
saldrá. «Elegir varias…» reparte un lote de fotos en orden. Las fotos asignadas viven
en `app.fotos_pptx` (clave de marco -> ruta) y solo durante la sesión.
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
from comite_policial.ui.widgets import Boton, card

IMAGENES = [("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"), ("Todos", "*.*")]
ANCHO_MINI = 410
ALTO_MINI = round(ANCHO_MINI * ALTO_LAMINA / ANCHO_LAMINA)
NAV_TXT, NAV_NUM = "#E4EAF7", "#9FB1DC"          # los mismos tonos del menú lateral de la app


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


def _corto(nombre: str, largo: int = 24) -> str:
    """Nombre de archivo que cabe en la fila (conserva el final, que lleva la extensión)."""
    return nombre if len(nombre) <= largo else nombre[:largo - 9].rstrip() + "…" + nombre[-8:]


class DialogoFotos(tk.Toplevel):
    def __init__(self, app, marcos, fotos: dict):
        super().__init__(app)
        self.app, self.marcos, self.fotos = app, marcos, fotos
        self.title("Fotos del PowerPoint")
        self.configure(bg=BG)
        self.transient(app)
        self.minsize(1000, 520)
        self._imgs = {}              # clave -> PhotoImage (Tk no guarda la referencia)
        self._lienzos = {}           # n.º de lámina -> Canvas de la miniatura
        self._nombres = {}           # clave -> StringVar con el nombre del archivo
        self._tarjetas = {}          # n.º de lámina -> tarjeta (solo una está visible)
        self._nav = {}               # n.º de lámina -> (fila, marca, rótulo, contador)
        self._actual = None
        self._por_lamina = {k: sorted(g, key=lambda m: m.indice)
                            for k, g in groupby(marcos, key=lambda m: m.lamina)}
        self.v_resumen = tk.StringVar()
        self._construir()
        self.mostrar(next(iter(self._por_lamina)))
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Up>", lambda e: self._mover(-1))
        self.bind("<Down>", lambda e: self._mover(1))
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(1080, sw - 60), min(640, sh - 100)
        x = max(0, min(app.winfo_rootx() + (app.winfo_width() - w) // 2, sw - w))
        self.geometry(f"{w}x{h}+{x}+{max(0, min(app.winfo_rooty() + 30, sh - h - 40))}")

    # ── Estructura ───────────────────────────────────────────────────────────
    def _construir(self):
        cab = tk.Frame(self, bg=BG)
        cab.pack(fill="x", padx=20, pady=(16, 6))
        tk.Label(cab, text="Fotos de las láminas", bg=BG, fg=NAVY,
                 font=(FONT, SIZE_H2, "bold")).pack(anchor="w")
        ayuda = tk.Label(cab, text="Elige una lámina a la izquierda. Haz clic en un marco numerado (o en "
                         "«Elegir…») para asignar su foto: se recorta para llenar el marco, como se verá "
                         "en el informe, y un marco sin foto queda vacío. Las fotos de las diligencias se "
                         "agregan en cada recuadro (sección X). Las fotos no se guardan al cerrar la "
                         "aplicación.", bg=BG, fg=INK_MID, font=(FONT, SIZE_SMALL), justify="left",
                         anchor="w")
        ayuda.pack(fill="x", pady=(2, 0))
        ayuda.bind("<Configure>", lambda e: ayuda.configure(wraplength=max(e.width - 4, 200)))

        barra = tk.Frame(self, bg=BG)
        barra.pack(side="bottom", fill="x", padx=20, pady=(6, 14))
        tk.Label(barra, textvariable=self.v_resumen, bg=BG, fg=INK_MID,
                 font=(FONT, SIZE_SMALL)).pack(side="left")
        Boton(barra, "Listo", self.destroy, "primary").pack(side="right")
        Boton(barra, "Quitar todas", self.quitar_todas, "danger").pack(side="right", padx=(0, 8))

        cuerpo = tk.Frame(self, bg=BG)
        cuerpo.pack(fill="both", expand=True, padx=20, pady=(8, 0))
        nav = tk.Frame(cuerpo, bg=NAVY_DEEP, width=250)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        self._detalle = tk.Frame(cuerpo, bg=BG)
        self._detalle.pack(side="left", fill="both", expand=True)

        tk.Label(nav, text="LÁMINAS", bg=NAVY_DEEP, fg=NAV_NUM, font=(FONT, 8, "bold"),
                 anchor="w", padx=18).pack(fill="x", pady=(14, 4))
        for lamina, grupo in self._por_lamina.items():
            self._nav[lamina] = self._item_nav(nav, lamina, grupo[0].seccion)
            self._tarjeta(self._detalle, lamina, grupo)
        self._refrescar_todo()

    def _item_nav(self, parent, lamina, seccion):
        fila = tk.Frame(parent, bg=NAVY_DEEP, cursor="hand2")
        fila.pack(fill="x", padx=8, pady=1)
        marca = tk.Frame(fila, bg=NAVY_DEEP, width=4)
        marca.pack(side="left", fill="y")
        rotulo = tk.Label(fila, text=seccion, bg=NAVY_DEEP, fg=NAV_TXT, anchor="w",
                          font=(FONT, SIZE_BASE), pady=7, padx=8)
        rotulo.pack(side="left", fill="x", expand=True)
        contador = tk.Label(fila, text="", bg=NAVY_DEEP, fg=NAV_NUM, width=5, anchor="e",
                            font=(FONT, SIZE_SMALL, "bold"), padx=6)
        contador.pack(side="right")
        for w in (fila, marca, rotulo, contador):
            w.bind("<Button-1>", lambda e, n=lamina: self.mostrar(n))
            w.bind("<Enter>", lambda e, n=lamina: self._pintar_nav(n, hover=True))
            w.bind("<Leave>", lambda e, n=lamina: self._pintar_nav(n))
        return fila, marca, rotulo, contador

    def _pintar_nav(self, lamina, hover=False):
        fila, marca, rotulo, contador = self._nav[lamina]
        activo = lamina == self._actual
        fondo = NAVY_SOFT if (activo or hover) else NAVY_DEEP
        for w in (fila, rotulo, contador):
            w.configure(bg=fondo)
        marca.configure(bg=GOLD if activo else fondo)
        rotulo.configure(fg="#FFFFFF" if activo else NAV_TXT,
                         font=(FONT, SIZE_BASE, "bold" if activo else "normal"))

    def _tarjeta(self, padre, lamina, grupo):
        inner, c = card(padre, f"{grupo[0].seccion}   ·   lámina {lamina}")
        borde = inner.master
        borde.pack_forget()                      # solo se muestra la lámina elegida
        self._tarjetas[lamina] = borde
        izq = tk.Frame(c, bg=SURFACE)
        izq.pack(side="left", anchor="n")
        lienzo = tk.Canvas(izq, width=ANCHO_MINI, height=ALTO_MINI, bg=SKY_SOFT,
                           highlightthickness=1, highlightbackground=LINE_STRONG)
        lienzo.pack()
        self._lienzos[lamina] = lienzo

        filas = tk.Frame(c, bg=SURFACE)
        filas.pack(side="left", fill="both", expand=True, padx=(18, 0))
        for m in grupo:
            fila = tk.Frame(filas, bg=SURFACE)
            fila.pack(fill="x", pady=2)
            tk.Label(fila, text=str(m.indice), bg=GOLD, fg=NAVY, width=2,
                     font=(FONT, SIZE_BASE, "bold")).pack(side="left")
            self._nombres[m.clave] = tk.StringVar()
            tk.Label(fila, textvariable=self._nombres[m.clave], bg=SURFACE, fg=INK,
                     font=(FONT, SIZE_SMALL), anchor="w", width=22
                     ).pack(side="left", padx=8)
            Boton(fila, "Elegir…", lambda mm=m: self.elegir(mm), "secondary", chico=True
                  ).pack(side="left", padx=2)
            Boton(fila, "Quitar", lambda k=m.clave: self.quitar(k), "ghost", chico=True
                  ).pack(side="left", padx=2)
        if len(grupo) > 1:
            Boton(filas, "Elegir varias…", lambda g=grupo: self.elegir_varias(g), "secondary",
                  chico=True).pack(anchor="w", pady=(8, 0))

    # ── Navegación ───────────────────────────────────────────────────────────
    def mostrar(self, lamina):
        """Muestra una sola lámina (las demás se ocultan, como las vistas de la app)."""
        if lamina not in self._tarjetas:
            return
        anterior, self._actual = self._actual, lamina
        if anterior is not None and anterior != lamina:
            self._tarjetas[anterior].pack_forget()
            self._pintar_nav(anterior)
        self._tarjetas[lamina].pack(fill="x", anchor="n", pady=(0, 12))
        self._pintar_nav(lamina)

    def _mover(self, paso):
        laminas = list(self._por_lamina)
        i = laminas.index(self._actual) + paso
        if 0 <= i < len(laminas):
            self.mostrar(laminas[i])

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
                self._imgs.pop(m.clave, None)
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
            self._nombres[m.clave].set(_corto(os.path.basename(ruta)) if ruta else "Sin foto")
        for lamina, grupo in self._por_lamina.items():
            self._pintar(lamina)
            con = sum(1 for m in grupo if self.fotos.get(m.clave))
            contador = self._nav[lamina][3]
            contador.configure(text=f"{con}/{len(grupo)}",
                               fg="#6FD6A4" if con else NAV_NUM)
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
