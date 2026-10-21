"""
Modelo de los recuadros de la sección IX (Concurrencias y Diligencias Relevantes).

Un recuadro es una tabla de 1 fila x 2 columnas:
  • Izquierda: unidad, grupo, N.º de detenidos y título del procedimiento (centrado, negrita).
  • Derecha: campos rotulados (Fecha, Lugar, Detenido, Diligencia, Resultados…).
Este módulo produce una representación neutra que usan la vista previa, los
exportadores y la sincronización con el documento original.
"""
from comite_policial.config.catalogos import TIPOS_RECUADRO

TIPO_DEFECTO = next(iter(TIPOS_RECUADRO))
TIPO_IMPORTADO = "Importado del documento (campos originales)"

# Claves de control para la sincronización con el .docx original
CLAVES_META = ("_orig", "_snap", "_clone_of")


def nuevo_recuadro(tipo: str = TIPO_DEFECTO, fecha: str = "") -> dict:
    return {"tipo": tipo, "unidad": "", "grupo": "", "detenidos": "",
            "titulo": "", "campos": {"Fecha": fecha}}


def campos_def(rec: dict) -> list:
    """Campos (etiqueta, clase) del recuadro: propios si fue importado."""
    if rec.get("tipo") == TIPO_IMPORTADO and rec.get("defs"):
        return [tuple(d) for d in rec["defs"]]
    return TIPOS_RECUADRO.get(rec.get("tipo"), [])


def contenido(rec: dict) -> dict:
    """Parte editable del recuadro (para detectar cambios frente al original)."""
    return {
        "izq": [rec.get("unidad", ""), rec.get("grupo", ""),
                str(rec.get("detenidos", "")), rec.get("titulo", "")],
        "der": [(e, c, str(rec.get("campos", {}).get(e, "")))
                for e, c in campos_def(rec)],
    }


def texto_detenidos(valor) -> str:
    """'1' -> '01 DETENIDO'; '2' -> '02 DETENIDOS'; vacío o 0 -> ''."""
    try:
        n = int(str(valor).strip())
    except ValueError:
        return ""
    if n <= 0:
        return ""
    return f"{n:02d} DETENIDO" + ("" if n == 1 else "S")


def grupos_izquierda(rec: dict) -> list:
    """Bloques de líneas de la columna izquierda (el renderer deja una línea
    en blanco entre bloques): unidad, grupo, detenidos, título."""
    bloques = [
        [rec.get("unidad", "").strip().upper()],
        [rec.get("grupo", "").strip().upper()],
        [texto_detenidos(rec.get("detenidos", ""))],
        [ln.strip().upper() for ln in rec.get("titulo", "").splitlines()],
    ]
    return [[ln for ln in b if ln] for b in bloques if any(ln for ln in b)]


def lineas_planas_izquierda(rec: dict) -> list:
    """Columna izquierda como lista plana de líneas ('' = línea en blanco)."""
    out = []
    for i, bloque in enumerate(grupos_izquierda(rec)):
        if i:
            out.append("")
        out.extend(bloque)
    return out


def parrafos_derecha(rec: dict) -> list:
    """Párrafos de la columna derecha.
    Cada párrafo es una lista de (texto, negrita); [] = línea en blanco."""
    campos = rec.get("campos", {})
    out, prev_linea = [], False
    for etiqueta, clase in campos_def(rec):
        valor = str(campos.get(etiqueta, "")).strip()
        if not valor:
            continue
        if out and not (clase == "linea" and prev_linea):
            out.append([])
        lineas = [ln.rstrip() for ln in valor.splitlines() if ln.strip()]
        if clase == "lista":
            out.append([(f"{etiqueta}:", True)])
            out.extend([(ln, False)] for ln in lineas)
        else:
            out.append([(f"{etiqueta}: ", True), (lineas[0], False)])
            out.extend([(ln, False)] for ln in lineas[1:])
        prev_linea = clase == "linea"
    return out


def resumen(rec: dict) -> tuple:
    """Datos para la tabla de la lista: (unidad, tipo, fecha, lugar, detenidos)."""
    c = rec.get("campos", {})
    tipo = rec.get("tipo", "")
    if tipo == TIPO_IMPORTADO:
        tipo = "Importado"
    return (rec.get("unidad", ""), tipo, c.get("Fecha", ""),
            c.get("Lugar", ""), texto_detenidos(rec.get("detenidos", "")))
