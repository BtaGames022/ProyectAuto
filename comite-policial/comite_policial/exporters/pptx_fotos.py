"""Fotos del PowerPoint: marcos de la plantilla, colocación y reparto en diligencias.

La plantilla (`tools/construir_plantilla_pptx.py`) deja un marco invisible
«FOTO:<clave>» donde va cada foto de las láminas fijas; su texto alternativo es
«sección|n|total». Aquí se leen esos marcos (para el botón «Fotos» de la app) y se
reemplaza cada uno por la foto elegida, en su misma capa, recortada para llenarlo.
Los marcos sin foto se eliminan: el archivo exportado no lleva formas ocultas.
"""
from __future__ import annotations

import io
import math
from dataclasses import dataclass

from PIL import Image, ImageOps
from pptx import Presentation

PREFIJO = "FOTO:"

# Región (EMU) de las fotos de una lámina de diligencia: columna derecha del texto
REGION_FOTOS = (8550000, 870000, 3400000, 5580000)
SEPARACION = 90000


@dataclass(frozen=True)
class Marco:
    clave: str        # «droga_2»: es la clave de la foto en el diccionario de la app
    seccion: str      # «II. Droga»
    indice: int       # n.º de foto dentro de la lámina (orden de lectura)
    total: int        # fotos de esa lámina
    lamina: int       # n.º de lámina en la plantilla (desde 1)
    x: int
    y: int
    w: int
    h: int

    @property
    def etiqueta(self) -> str:
        return f"{self.seccion} — foto {self.indice} de {self.total}"


def _marcos_de(slide):
    return [sh for sh in slide.shapes if sh.name.startswith(PREFIJO)]


def leer_marcos(plantilla) -> list:
    """Marcos de foto de una plantilla (ruta o `Presentation`), en orden de lámina."""
    prs = Presentation(plantilla) if isinstance(plantilla, (str, bytes)) or hasattr(plantilla, "__fspath__") \
        else plantilla
    out = []
    for n, sl in enumerate(prs.slides, 1):
        for sh in _marcos_de(sl):
            partes = (sh._element.xpath("./p:nvSpPr/p:cNvPr/@descr") or [""])[0].split("|")
            try:
                seccion, i, total = partes[0], int(partes[1]), int(partes[2])
            except (IndexError, ValueError):
                seccion, i, total = f"Lámina {n}", 1, 1
            out.append(Marco(sh.name[len(PREFIJO):], seccion, i, total, n,
                             sh.left, sh.top, sh.width, sh.height))
    return out


# ── Imágenes ─────────────────────────────────────────────────────────────────
def imagen_flujo(ruta: str):
    """Abre la foto respetando su rotación, la achica y devuelve (bytes, ancho, alto)."""
    with Image.open(ruta) as im:
        im = ImageOps.exif_transpose(im)
        im.thumbnail((1920, 1920))
        buf = io.BytesIO()
        if im.mode in ("RGBA", "LA", "P"):
            im.convert("RGBA").save(buf, "PNG")
        else:
            im.convert("RGB").save(buf, "JPEG", quality=88)
        buf.seek(0)
        return buf, im.size[0], im.size[1]


def motivo(exc) -> str:
    """Causa legible de un fallo al leer una foto."""
    if isinstance(exc, FileNotFoundError):
        return "el archivo ya no existe"
    if isinstance(exc, PermissionError):
        return "sin permiso para leerlo"
    if exc.__class__.__name__ in ("UnidentifiedImageError", "DecompressionBombError"):
        return "no es una imagen válida o es demasiado grande"
    return f"{exc.__class__.__name__}: {exc}"


def poner_foto(slide, ruta: str, x, y, w, h, modo="cubrir", antes_de=None):
    """Inserta la foto en el marco (x, y, w, h). `modo`: «cubrir» llena el marco
    recortando; «ajustar» la muestra entera y centrada. Con `antes_de` (un elemento
    XML de la lámina) la foto ocupa esa capa; si no, queda encima de todo."""
    buf, iw, ih = imagen_flujo(ruta)
    if modo == "cubrir":
        pic = slide.shapes.add_picture(buf, x, y, w, h)
        r_marco, r_img = w / h, iw / ih
        if r_img > r_marco:
            corte = (1 - r_marco / r_img) / 2
            pic.crop_left = pic.crop_right = corte
        else:
            corte = (1 - r_img / r_marco) / 2
            pic.crop_top = pic.crop_bottom = corte
    else:
        esc = min(w / iw, h / ih)
        w2, h2 = int(iw * esc), int(ih * esc)
        pic = slide.shapes.add_picture(buf, x + (w - w2) // 2, y + (h - h2) // 2, w2, h2)
    if antes_de is not None:
        antes_de.addprevious(pic._element)


def colocar_fotos(prs, fotos: dict, avisos: list):
    """Reemplaza cada marco por su foto (`fotos`: clave -> ruta) y elimina los vacíos."""
    usadas = set()
    for sl in prs.slides:
        for marco in _marcos_de(sl):
            clave = marco.name[len(PREFIJO):]
            ruta = fotos.get(clave)
            if ruta:
                usadas.add(clave)
                try:
                    poner_foto(sl, ruta, marco.left, marco.top, marco.width, marco.height,
                               modo="cubrir", antes_de=marco._element)
                except Exception as exc:
                    descr = (marco._element.xpath("./p:nvSpPr/p:cNvPr/@descr") or [clave])[0].split("|")[0]
                    avisos.append(f"No se pudo insertar la foto «{descr} ({clave})»: {motivo(exc)}")
            marco._element.getparent().remove(marco._element)
    for clave in sorted(set(fotos) - usadas):
        if fotos[clave]:
            avisos.append(f"La foto «{clave}» no tiene marco en la plantilla y no se insertó.")


def celdas_fotos(n: int, region=REGION_FOTOS):
    """Reparte la región de fotos en `n` celdas (x, y, w, h)."""
    x0, y0, W, H = region
    if n <= 0:
        return []
    if n == 1:
        filas = [[1]]
    elif n == 2:
        filas = [[1], [1]]
    elif n == 3:
        filas = [[1], [1, 1]]
    else:
        filas = [[1, 1] for _ in range(math.ceil(n / 2))]
        if n % 2:
            filas[-1] = [1]
    alto = (H - SEPARACION * (len(filas) - 1)) / len(filas)
    celdas, y = [], y0
    for fila in filas:
        ancho = (W - SEPARACION * (len(fila) - 1)) / len(fila)
        for k in range(len(fila)):
            celdas.append((int(x0 + k * (ancho + SEPARACION)), int(y), int(ancho), int(alto)))
        y += alto + SEPARACION
    return celdas[:n]
