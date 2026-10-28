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

## Pruebas automatizadas

Sirven para comprobar, en un minuto y sin abrir la aplicación a mano, que nada se rompió
después de un cambio: importar y guardar el Word, exportar el PowerPoint, el scroll y todos
los botones de la interfaz.

### Cómo ejecutarlas

1. Instalar las dependencias (una sola vez): `pip install -r requirements-dev.txt`.
2. Cerrar la aplicación si está abierta y, desde la carpeta `comite-policial`, correr:

```
python tests/run_all.py
```

Se abren y cierran ventanas de la aplicación durante la prueba: es normal, no hay que
tocarlas. Necesita una pantalla (no corre por SSH ni en un servidor sin escritorio) y
Python 3.10 o superior. Tarda cerca de un minuto.

Resultado esperado, una línea por prueba:

```
OK     roundtrip_docx.py
OK     integracion_docx.py
OK     smoke_recuadros.py
OK     pptx_export_test.py
OK     scroll_test.py
OK     ui_completa_test.py
```

Si alguna dice `FALLA`, debajo aparece el detalle (la línea que no se cumplió). Un
código de salida distinto de 0 indica fallo, así que sirve también en un flujo automático.

Para ejecutar una sola prueba, por ejemplo la del scroll:

```
python tests/scroll_test.py
```

Cada una imprime `OK` al final si todo salió bien.

### Qué comprueba cada una

| Prueba | Qué verifica |
|---|---|
| `roundtrip_docx` | Importar un Word, modificarlo y reescribirlo sin perder el formato; sin cambios el XML queda idéntico. |
| `integracion_docx` | Flujo completo App y Word: importar, editar una celda o un recuadro y comprobar que el guardado automático llega al archivo. |
| `smoke_recuadros` | Sección IX (crear, duplicar, casilla PPT) y exportación a Word y PowerPoint. |
| `pptx_export_test` | El exportador de PowerPoint sin interfaz: sin restos de la semana de la plantilla, cantidad de láminas, textos y cierre. |
| `scroll_test` | Regresión del error «se repite la vista»: la rueda del mouse no cambia desplegables ni reconstruye el formulario, y la vista previa no crece sin límite. |
| `ui_completa_test` | Abre todas las vistas; edita celdas con el teclado; agrega y elimina filas; calendarios; recuadros; fotos; exporta el PowerPoint (con una foto perdida y el archivo bloqueado); importa un Word válido, uno dañado y uno ajeno; comprueba que al cerrar se guarda lo pendiente. |

### Informe de ejemplo

Las pruebas de Word necesitan un informe. El real no se versiona (`*.docx` está en
`.gitignore`), por eso, si falta `tests/data/informe_ejemplo.docx`, se genera uno
sintético con la misma estructura (`tests/make_informe_docx.py`). Para crearlo a mano:

```
python tests/make_informe_docx.py
```

Si se quiere probar con el informe real, basta copiarlo con ese nombre a `tests/data/`.
Las pruebas trabajan siempre sobre copias: el original no se modifica.

### Agregar una prueba nueva

Crear `tests/mi_prueba.py` que termine con `print("OK")` si todo pasó (o falle con un
`assert`) y agregar su nombre a la lista `PRUEBAS` de `tests/run_all.py`. Para ejercitar
la interfaz, copiar el arranque de `tests/ui_completa_test.py`: reemplaza
`report_callback_exception` para que un error dentro de un botón haga fallar la prueba en
lugar de pasar en silencio.

## Configuración (.env)

| Variable | Descripción |
|---|---|
| `APP_TITLE` | Título de la ventana |
| `ANO_ANTERIOR` / `ANO_ACTUAL` / `PERIODO` | Comparación interanual |
| `SEMANA_AUTO` | `true` detecta la semana actual (lunes-domingo) |
| `SEMANA_DEFAULT` / `TURNO_RESPONSABLE` | Valores iniciales de la barra superior |
| `WINDOW_WIDTH` / `WINDOW_HEIGHT` | Tamaño de ventana |
| `OUTPUT_DIR` / `PLANTILLA_PPTX` / `ASSETS_DIR` | Rutas (relativas a la raíz o absolutas) |
