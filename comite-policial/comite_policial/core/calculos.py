"""Motor de cálculo analítico: variaciones, porcentajes y formato numérico."""

def variacion(a: float, b: float) -> str:
    if a == 0 and b == 0: return "0%"
    if a == 0: return "+100%"
    v = (b - a) / a * 100
    return f"{'+'if v>0 else ''}{v:.1f}%"

def porcentaje(p: float, t: float) -> str:
    if t == 0: return "0,0%"
    return f"{p/t*100:.1f}%".replace(".", ",")

def fmt(v) -> str:
    try:
        f = float(v)
        if f == int(f): return f"{int(f):,}".replace(",", ".")
        return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(v)

def safe_float(s):
    try: return float(str(s).replace(",", ".").replace(".", "", s.count(".")-1) if s.count(".")>1 else str(s).replace(",", "."))
    except: return 0.0

def safe_int(s):
    try: return int(str(s).replace(".", ""))
    except: return 0
