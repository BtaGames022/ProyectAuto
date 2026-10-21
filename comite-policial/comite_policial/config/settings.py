"""
Configuración centralizada de la aplicación.

Lee variables desde el archivo `.env` (raíz del proyecto) y/o el entorno del
sistema. Implementación propia (sin dependencias) para mantener el proyecto
100 % offline y portable.
"""
import os
import sys
from dataclasses import dataclass
from pathlib import Path


def _base_dir() -> Path:
    """Raíz del proyecto (o carpeta del .exe cuando se compila con PyInstaller)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


BASE_DIR = _base_dir()


def _load_env(path: Path) -> dict:
    data = {}
    if not path.is_file():
        return data
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        val = val.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
            val = val[1:-1]
        data[key.strip()] = val
    return data


_ENV = _load_env(BASE_DIR / ".env")


def _get(key: str, default: str) -> str:
    return os.environ.get(key, _ENV.get(key, default))


def _get_int(key: str, default: int) -> int:
    try:
        return int(_get(key, str(default)))
    except ValueError:
        return default


def _path(key: str, default: str) -> str:
    p = Path(_get(key, default))
    return str(p if p.is_absolute() else BASE_DIR / p)


@dataclass(frozen=True)
class Settings:
    app_title: str
    unidad: str
    ano_ant: str
    ano_act: str
    periodo: str
    semana_auto: bool
    semana_default: str
    turno_responsable: str
    win_width: int
    win_height: int
    output_dir: str
    plantilla_docx: str
    assets_dir: str


settings = Settings(
    app_title=_get("APP_TITLE", "Sistema de Automatización de Informes — PDI Valparaíso"),
    unidad=_get("UNIDAD", "Región Policial Valparaíso"),
    ano_ant=_get("ANO_ANTERIOR", "2025"),
    ano_act=_get("ANO_ACTUAL", "2026"),
    periodo=_get("PERIODO", "ENE - AGO"),
    semana_auto=_get("SEMANA_AUTO", "true").lower() in ("1", "true", "si", "sí", "yes"),
    semana_default=_get("SEMANA_DEFAULT", "31.AGO AL 06.SEP.026"),
    turno_responsable=_get("TURNO_RESPONSABLE", "PROF (A/C) Nombre APELLIDO"),
    win_width=_get_int("WINDOW_WIDTH", 1400),
    win_height=_get_int("WINDOW_HEIGHT", 860),
    output_dir=_path("OUTPUT_DIR", "output"),
    plantilla_docx=_path("PLANTILLA_DOCX", "plantillas/plantilla_base.docx"),
    assets_dir=_path("ASSETS_DIR", "assets"),
)

