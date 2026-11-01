"""Tipos de recuadro de la sección X que se incorporan desde los informes importados.

Cuando un informe trae un recuadro con campos que no coinciden con ningún tipo
conocido, ese formato se guarda en `datos/tipos_recuadro.json` y queda disponible
en el desplegable «Tipo de procedimiento» (también en próximas sesiones). El
recuadro se importa igual, con o sin tipo en el catálogo.
"""
import json
import os

from comite_policial.config.catalogos import TIPOS_RECUADRO
from comite_policial.config.settings import BASE_DIR
from comite_policial.core.recuadros import TIPO_IMPORTADO
from comite_policial.docio.xml import norm

# COMITE_TIPOS_JSON permite redirigir el archivo (las pruebas no tocan el real)
RUTA = (os.environ.get("COMITE_TIPOS_JSON")
        or os.path.join(str(BASE_DIR), "datos", "tipos_recuadro.json"))
MAX_NOMBRE = 60

# nombres de los tipos agregados por el usuario (para distinguirlos de los de fábrica)
PERSONALIZADOS: set = set()


def _firma(defs) -> tuple:
    return tuple(norm(str(d[0])) for d in defs)


def _firmas_conocidas() -> set:
    return {_firma(defs) for defs in TIPOS_RECUADRO.values()}


def _guardar(ruta: str):
    ruta = ruta or RUTA
    datos = {n: [list(d) for d in TIPOS_RECUADRO[n]] for n in PERSONALIZADOS
             if n in TIPOS_RECUADRO}
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        with open(ruta, "w", encoding="utf-8") as fh:
            json.dump(datos, fh, ensure_ascii=False, indent=2)
    except OSError:
        pass                      # carpeta de solo lectura: el tipo vale en esta sesión


def cargar(ruta: str = None) -> list:
    """Suma al catálogo los tipos guardados. Devuelve los nombres agregados."""
    ruta = ruta or RUTA
    try:
        with open(ruta, encoding="utf-8") as fh:
            datos = json.load(fh)
    except (OSError, ValueError):
        return []
    agregados = []
    for nombre, defs in datos.items():
        try:
            defs = [(str(e), str(c)) for e, c in defs]
        except (TypeError, ValueError):
            continue
        if nombre in TIPOS_RECUADRO:          # ya cargado en esta sesión
            PERSONALIZADOS.add(nombre)
            continue
        if not defs:
            continue
        TIPOS_RECUADRO[nombre] = defs
        PERSONALIZADOS.add(nombre)
        agregados.append(nombre)
    return agregados


def _nombre(rec: dict) -> str:
    titulo = " ".join(ln.strip() for ln in rec.get("titulo", "").splitlines() if ln.strip())
    base = " ".join(titulo.lower().split()) or "formato importado"
    base = base[0].upper() + base[1:]
    if len(base) > MAX_NOMBRE:
        base = base[:MAX_NOMBRE - 1].rstrip() + "…"
    nombre, k = base, 2
    while nombre in TIPOS_RECUADRO or nombre == TIPO_IMPORTADO:
        nombre, k = f"{base} ({k})", k + 1
    return nombre


def registrar(recuadros: list, ruta: str = None) -> list:
    """Agrega al catálogo los formatos de `recuadros` que aún no existen.

    Devuelve los nombres de los tipos nuevos (lista vacía si no hubo)."""
    conocidas = _firmas_conocidas()
    nuevos = []
    for rec in recuadros:
        defs = rec.get("defs") or []
        if not defs:
            continue
        firma = _firma(defs)
        if firma in conocidas:
            continue
        nombre = _nombre(rec)
        TIPOS_RECUADRO[nombre] = [(str(e), str(c)) for e, c in defs]
        PERSONALIZADOS.add(nombre)
        conocidas.add(firma)
        nuevos.append(nombre)
    if nuevos:
        _guardar(ruta or RUTA)
    return nuevos
