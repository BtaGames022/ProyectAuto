# Sistema de Automatización de Informes — Comité Policial PDI Valparaíso

Aplicación de escritorio **local, offline y portable** (Python + Tkinter) para
actualizar el **Informe Semanal del Comité Policial** de la Región Policial
Valparaíso. El flujo oficial es:

1. **Importar** el `.docx` oficial (`INFORME REPOL VPO RRF.docx`): la app carga
   semana, turno, las 15 tablas, los recuadros de la Sección IX y el texto de la X.
2. **Editar** en la interfaz (tablas tipo hoja de cálculo, recálculo automático
   de totales, diferencias y Δ%).
3. **Sincronizar** sobre el `.docx` original (automática o manual desde el panel
   `DOCUMENTO`). Se edita el XML directamente, así que se conserva el formato
   (estilos, colores `#002060`, Calibri, celdas combinadas, logos, pies de página).
   Antes de modificar el archivo se guarda una copia en `output/respaldos/`.
4. **Exportar a PowerPoint** el informe del comité (botón en la barra superior).

## Requisitos y uso

Python 3.10+ (probado con 3.14).

```
pip install -r requirements.txt
copy .env.example .env
python run.py
```

Dependencias de ejecución: `python-docx`, `python-pptx`, `lxml`, `Pillow`
(`tkinter` viene con Python). Para compilar además se usa `pyinstaller`
(`requirements-dev.txt`).

## Estructura

```
comite-policial/
├── run.py                    # Punto de entrada
├── ComitePolicial.spec       # Configuración de PyInstaller
├── build.bat                 # Compila dist\ComitePolicial.exe
├── .env / .env.example       # Configuración (el .env no se versiona)
├── requirements.txt          # Dependencias de ejecución
├── requirements-dev.txt      # + pyinstaller
├── comite_policial/
│   ├── config/               # settings (.env propio), theme, catalogos
│   ├── core/                 # calculos, fechas (semana lunes-domingo), recuadros
│   ├── docio/                # lectura/escritura directa sobre el XML del .docx
│   │   ├── importador.py     #   .docx -> estado de la app
│   │   ├── escritor.py       #   estado de la app -> .docx (preserva formato)
│   │   ├── mapa.py, xml.py, semana.py
│   ├── exporters/
│   │   ├── pptx_export.py    #   botones Exportar PowerPoint y Fotos
│   │   ├── pptx_informe.py   #   motor que rellena la plantilla .pptx
│   │   ├── recuadros_export.py  # dibujo de recuadros IX (lo usa docio/escritor)
│   │   └── word.py           #   exportador Word desde cero (legado, sin uso en la UI)
│   └── ui/                   # app, views, views_extra, recuadros, documento,
│                             # widgets, calendarios
├── plantillas/plantilla_comite.pptx
├── tools/construir_plantilla_pptx.py
├── tests/
├── output/                   # respaldos y exportaciones (no se versiona)
└── assets/
```

## Interfaz

Diseño institucional PDI: azul marino `#002060`, celeste de las tablas del informe,
amarillo `#FFC000` solo como acento (sección activa y botón de exportar) y Arial,
la tipografía de los documentos del comité.

- **Menú lateral:** numeral romano de cada sección; un punto verde indica que ya tiene datos.
- **Barra de contexto:** semana operativa, fecha del informe y turno, siempre visibles.
- **Atajos:** `Ctrl+E` exporta a PowerPoint y `Ctrl+S` guarda en el Word importado.
- **Rueda del mouse:** `ui/scroll.py` decide quién se desplaza (tabla, texto o panel); un
  desplegable nunca cambia de valor al pasar la rueda (era la causa de que el editor de
  recuadros «se repitiera» al hacer scroll). Prueba: `tests/scroll_test.py`.
- **Colores y tamaños:** todo sale de `config/theme.py`; los widgets en `ui/widgets.py`
  (`Boton`, `card`, `section_title`, `EditableTree`) y los valores globales en `ui/estilo.py`.

## Exportar a PowerPoint

El botón **Exportar PowerPoint** abre `plantillas/plantilla_comite.pptx` (el
informe real sin las fotos de la semana) y solo reescribe textos y celdas, de
modo que fondos, fuentes y posiciones son los del original. Cubre: portada,
I Armas, II Droga, III Homicidios (A-D), IV Migraciones, IV-B Flujo,
V Vehículos, VII Bandas, VIII Detenidos (B y C), una lámina por cada recuadro de
la Sección IX y el cierre PDI. Las secciones que no figuran en el formato del
comité (VI, VIII-A y X) no se exportan al PowerPoint.

- **Fotos:** el botón **Fotos** asigna una imagen a cada marco de las láminas
  fijas (se recorta para llenarlo). Las fotos de cada diligencia se agregan en el
  recuadro (sección IX). Las fotos solo viven durante la sesión: no se guardan.
- **Cambiar el diseño:** si el comité cambia el formato, regenerar la plantilla con
  `python tools/construir_plantilla_pptx.py "<informe nuevo>.pptx"`. El exportador
  busca las tablas por nombre de objeto, conviene mantenerlos.

## Compilar el .exe

```
build.bat
```

Genera `dist\ComitePolicial.exe` (`--onefile --windowed`, definido en
`ComitePolicial.spec`) y copia `.env` y `plantillas\plantilla_comite.pptx` junto
al ejecutable. Edita `dist\.env` para configurar.

## Pruebas

```
python tests/run_all.py
```

Corre seis pruebas (no dependen del directorio actual ni de un informe real):

- `roundtrip_docx` e `integracion_docx`: importar un Word, editarlo y reescribirlo sin
  perder formato, y el flujo App y .docx con el guardado automático.
- `smoke_recuadros`: sección IX y exportación a PowerPoint.
- `pptx_export_test`: exportador sin interfaz.
- `scroll_test`: la rueda del mouse no cambia desplegables ni reconstruye el formulario.
- `ui_completa_test`: recorre todas las vistas y botones, importa un Word válido, uno
  dañado y uno ajeno, exporta con y sin fotos y comprueba el cierre ordenado.

Si falta `tests/data/informe_ejemplo.docx` (el informe real no se versiona; `*.docx` está
en `.gitignore`) se genera uno sintético con `tests/make_informe_docx.py`, que tiene la
misma estructura que el real.

## Configuración (.env)

| Variable | Descripción |
|---|---|
| `APP_TITLE` | Título de la ventana |
| `ANO_ANTERIOR` / `ANO_ACTUAL` / `PERIODO` | Comparación interanual |
| `SEMANA_AUTO` | `true` detecta la semana actual (lunes-domingo) |
| `SEMANA_DEFAULT` / `TURNO_RESPONSABLE` | Valores iniciales de la barra superior |
| `WINDOW_WIDTH` / `WINDOW_HEIGHT` | Tamaño de ventana |
| `OUTPUT_DIR` / `PLANTILLA_PPTX` / `ASSETS_DIR` | Rutas (relativas a la raíz o absolutas) |
