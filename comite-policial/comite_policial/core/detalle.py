"""Texto del recuadro (columna derecha) como un solo cuadro, igual que en el Word.

Dos formas de editarlo:
  • fijo  (tipos de fábrica): los rótulos («Lugar:», «Hechos:» …) quedan en negrita y
    no se pueden borrar ni mover; solo se escribe a continuación de cada uno.
  • libre (tipos escritos a mano, los incorporados desde un informe y los
    «importados»): todo el cuadro es editable; una línea «Rótulo: texto» define un
    campo en negrita.
La fecha va aparte (campo con botón de calendario) en ambos casos.
"""
import re

LINEA = ("FECHA", "LUGAR")
_ETQ = re.compile(r"^[ \t]*([A-ZÁÉÍÓÚÑ][^:\n]{1,59}?)[ \t]*:[ \t]*(.*)$")


def es_etiqueta(texto: str) -> bool:
    t = texto.strip()
    return 0 < len(t) <= 60 and len(t.split()) <= 7


_SUFIJO_DUP = re.compile(r"\s+\(\d+\)$")


def rotulo_visible(etq: str) -> str:
    return _SUFIJO_DUP.sub("", str(etq))


def _bloque(etq: str, clase: str, valor) -> str:
    rot = rotulo_visible(etq)
    valor = str(valor or "").strip("\n")
    if clase == "lista":
        return f"{rot}:" + (f"\n{valor}" if valor else "")
    return f"{rot}: {valor}"


def render(defs, campos: dict) -> str:
    """Texto del cuadro: un bloque «Rótulo: valor» por campo, separados por una línea en blanco."""
    return "\n\n".join(_bloque(e, c, campos.get(e, "")) for e, c in defs)


# ── Modo fijo ────────────────────────────────────────────────────────────────
def marcas_fijo(texto: str, defs):
    """[(inicio, fin)] de cada rótulo esperado (en orden, al comienzo de línea)
    o None si falta alguno o hay texto antes del primero."""
    pos, marcas = 0, []
    for e, _ in defs:
        m = re.compile("^" + re.escape(rotulo_visible(e)) + ":", re.M).search(texto, pos)
        if not m:
            return None
        marcas.append((m.start(), m.end()))
        pos = m.end()
    if marcas and texto[:marcas[0][0]].strip():
        return None
    return marcas


def parse_fijo(texto: str, defs):
    """{etiqueta: valor} o None si la estructura de rótulos no está íntegra."""
    marcas = marcas_fijo(texto, defs)
    if marcas is None:
        return None
    campos = {}
    for i, (e, _) in enumerate(defs):
        fin = marcas[i + 1][0] if i + 1 < len(marcas) else len(texto)
        campos[e] = texto[marcas[i][1]:fin].strip()
    return campos


# ── Modo libre ───────────────────────────────────────────────────────────────
def marcas_libre(texto: str):
    """[(inicio, fin)] de los rótulos de las líneas «Rótulo: …»."""
    out, off = [], 0
    for ln in texto.split("\n"):
        m = _ETQ.match(ln)
        if m and es_etiqueta(m.group(1)):
            ini = off + m.start(1)
            out.append((ini, off + ln.find(":", m.end(1)) + 1))
        off += len(ln) + 1
    return out


def parse_libre(texto: str):
    """(defs, campos) a partir de un texto libre."""
    defs, campos, actual, previo = [], {}, None, []
    for ln in texto.split("\n"):
        m = _ETQ.match(ln)
        if m and es_etiqueta(m.group(1)):
            base = " ".join(m.group(1).split())
            etq, k = base, 2
            while etq in campos:
                etq, k = f"{base} ({k})", k + 1
            resto = m.group(2).strip()
            clase = ("linea" if base.upper() in LINEA
                     else ("parrafo" if resto else "lista"))
            defs.append((etq, clase))
            campos[etq] = resto
            actual = etq
        elif actual is None:
            if ln.strip():
                previo.append(ln.strip())
        else:
            campos[actual] = campos[actual] + "\n" + ln
    campos = {k: v.strip() for k, v in campos.items()}
    if previo:                                    # texto sin rótulo al comienzo
        nombre = "Detalle"
        while nombre in campos:
            nombre += "*"
        defs.insert(0, (nombre, "parrafo"))
        campos[nombre] = "\n".join(previo)
    return defs, campos
