# Sistema de Automatización de Informes Policiales

Aplicación local, offline y portable (Python + Tkinter) que consolida el informe
operativo semanal y lo exporta a **Word (.docx)** y **PowerPoint (.pptx)**.

## Estructura

```
comite-policial/
├── run.py                     # Punto de entrada
├── .env / .env.example        # Configuración (el .env no se versiona)
├── requirements.txt
├── build.bat                  # Compilación a .exe (PyInstaller)
├── comite_policial/
│   ├── config/
│   │   ├── settings.py        # Lectura de .env
│   │   ├── theme.py           # Colores y tipografía
│   │   └── catalogos.py       # Catálogos de tablas (armas, drogas, unidades…)
│   ├── core/
│   │   └── calculos.py        # Variaciones, porcentajes, formato numérico
│   ├── ui/
│   │   ├── app.py             # Ventana principal
│   │   ├── views.py           # Vistas de las secciones I–VIII-C
│   │   ├── widgets.py         # EditableTree, tarjetas, botones
│   │   └── calendarios.py     # Selector de fecha y rango semanal
│   └── exporters/
│       ├── word.py            # Exportación .docx
│       └── pptx_export.py     # Exportación .pptx
├── plantillas/                # plantilla_base.docx (opcional)
├── assets/                    # Logos / imágenes
├── output/                    # Informes generados
├── docs/
└── tests/
```

## Uso

```
pip install -r requirements.txt
copy .env.example .env
python run.py
```

## Configuración (.env)

| Variable | Descripción |
|---|---|
| `APP_TITLE` | Título de la ventana |
| `ANO_ANTERIOR` / `ANO_ACTUAL` / `PERIODO` | Comparación interanual |
| `SEMANA_DEFAULT` / `TURNO_RESPONSABLE` | Valores iniciales de la barra superior |
| `WINDOW_WIDTH` / `WINDOW_HEIGHT` | Tamaño de ventana |
| `OUTPUT_DIR` / `PLANTILLA_DOCX` / `ASSETS_DIR` | Rutas (relativas a la raíz o absolutas) |
