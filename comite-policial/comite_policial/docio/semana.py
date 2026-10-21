"""Detección y sustitución de la semana operativa dentro de los textos del documento."""
import re

# 31.AGO AL 06.SEP.026  |  31.AGO al 06.SEP.026  |  31 AGO AL 06.SEP.026  |  31.AGO.026 AL 06.SEP.026
SEMANA_RE = re.compile(
    r"(?<!\d)(?!01\.ENE)(\d{1,2})([. ])([A-ZÁÉÍÓÚ]{3})(?:\.(\d{3}))?(\s+)(AL|al)(\s+)"
    r"(\d{1,2})\.([A-ZÁÉÍÓÚ]{3})\.(\d{3})")
ACUM_RE = re.compile(r"(01\.ENE\s+(?:AL|al)\s+)(\d{1,2}\.[A-ZÁÉÍÓÚ]{3}\.\d{3})")
# Formato estándar de la app: "DD.MES AL DD.MES.YYY"
SEMANA_APP_RE = re.compile(r"(\d{1,2})\.([A-ZÁÉÍÓÚ]{3})\s+AL\s+(\d{1,2})\.([A-ZÁÉÍÓÚ]{3})\.(\d{3})",
                           re.IGNORECASE)


def semana_normalizada(texto: str) -> str:
    """Extrae la semana del texto del documento y la lleva al formato de la app."""
    m = SEMANA_RE.search(texto)
    if not m:
        return ""
    d1, _, m1, _, _, _, _, d2, m2, y = m.groups()
    return f"{int(d1):02d}.{m1} AL {int(d2):02d}.{m2}.{y}"


def _partes(semana_app: str):
    m = SEMANA_APP_RE.fullmatch(semana_app.strip())
    if not m:
        return None
    d1, m1, d2, m2, y = m.groups()
    return int(d1), m1.upper(), int(d2), m2.upper(), y


def funcion_semana(semana_app: str):
    """Devuelve la función de reemplazo para SEMANA_RE (None si el formato es inválido)."""
    p = _partes(semana_app)
    if not p:
        return None
    d1, m1, d2, m2, y = p

    def fn(m):
        _, sep, _, ycon, ws1, al, ws2, _, _, _ = m.groups()
        anio1 = f".{y}" if ycon else ""
        return f"{d1:02d}{sep}{m1}{anio1}{ws1}{al}{ws2}{d2:02d}.{m2}.{y}"
    return fn


def funcion_acum(semana_app: str):
    p = _partes(semana_app)
    if not p:
        return None
    _, _, d2, m2, y = p
    return lambda m: f"{m.group(1)}{d2:02d}.{m2}.{y}"
