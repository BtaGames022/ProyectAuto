"""Utilidades de fecha: detección automática de la semana operativa (lunes a domingo)."""
import datetime

MESES_ABREV = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
               "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]


def rango_semana(hoy: datetime.date | None = None):
    """Devuelve (lunes, domingo) de la semana que contiene a `hoy`."""
    hoy = hoy or datetime.date.today()
    lunes = hoy - datetime.timedelta(days=hoy.weekday())
    return lunes, lunes + datetime.timedelta(days=6)


def semana_str(hoy: datetime.date | None = None) -> str:
    """Formato del informe: '12.OCT AL 18.OCT.026'."""
    s, e = rango_semana(hoy)
    return (f"{s.day:02d}.{MESES_ABREV[s.month-1]} "
            f"AL {e.day:02d}.{MESES_ABREV[e.month-1]}.{str(e.year)[-3:]}")


def fecha_corta(d: datetime.date | None = None) -> str:
    """Formato de fecha del informe: '31.AGO.026'."""
    d = d or datetime.date.today()
    return f"{d.day:02d}.{MESES_ABREV[d.month-1]}.{str(d.year)[-3:]}"
