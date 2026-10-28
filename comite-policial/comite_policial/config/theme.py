"""Sistema visual de la aplicación (identidad PDI).

Colores tomados del informe oficial del comité (PowerPoint): azul marino
institucional, celeste de las tablas y amarillo de acento. Tipografía Arial,
la misma de los documentos. Los nombres antiguos se conservan como alias para
no romper el resto del código.
"""

# ── Marca ────────────────────────────────────────────────────────────────────
NAVY        = "#002060"   # azul institucional (títulos, encabezados de tabla)
NAVY_DEEP   = "#001540"   # menú lateral
NAVY_SOFT   = "#0B3080"   # hover sobre azul
SKY         = "#DAE9F8"   # celeste de las tablas del informe
SKY_SOFT    = "#EEF4FB"   # filas alternas
GOLD        = "#FFC000"   # amarillo de acento (activo, foco)
GOLD_SOFT   = "#FFF2CC"   # selección

# ── Neutros (con un leve tinte azul) ─────────────────────────────────────────
BG          = "#EDF1F7"   # fondo del área de trabajo
SURFACE     = "#FFFFFF"   # tarjetas y campos
LINE        = "#D3DBE8"   # bordes
LINE_STRONG = "#AAB7CC"   # bordes de campos
INK         = "#0F1B33"   # texto principal
INK_MID     = "#46566F"   # texto secundario
INK_SOFT    = "#7A8AA3"   # ayudas, fuentes

# ── Estados ──────────────────────────────────────────────────────────────────
OK          = "#1B7A4B"
OK_SOFT     = "#E3F4EA"
WARN        = "#B26A00"
WARN_SOFT   = "#FFF1D6"
ERR         = "#B3261E"
ERR_SOFT    = "#FBE5E3"

# ── Tipografía ───────────────────────────────────────────────────────────────
FONT        = "Arial"
FONT_MONO   = "Consolas"
SIZE_BASE   = 10
SIZE_SMALL  = 9
SIZE_H1     = 16
SIZE_H2     = 12

# ── Alias de compatibilidad con el código existente ──────────────────────────
DARK        = NAVY_DEEP
SIDEBAR_BG  = NAVY_DEEP
SIDEBAR_HOV = NAVY_SOFT
HEADER_BG   = NAVY
ACCENT      = NAVY
ACCENT2     = NAVY_SOFT
SUCCESS     = OK
DANGER      = ERR
WARNING     = WARN
PURPLE      = INK_MID
TEAL        = INK_MID
LIGHT_BG    = BG
CARD_BG     = SURFACE
BORDER      = LINE
TEXT_DARK   = INK
TEXT_MID    = INK_MID
TEXT_LIGHT  = INK_SOFT
TEXT_WHITE  = "#FFFFFF"
ROW_ALT     = SKY_SOFT
ROW_TOTAL   = SKY
ROW_FUENTE  = BG

# Tablas (réplica de las del informe PowerPoint)
ROW_ODD       = SURFACE
ROW_EVEN      = SKY_SOFT
ROW_HOVER_BG  = SKY
ROW_SEL_BG    = GOLD_SOFT
ROW_SEL_FG    = INK
ROW_TOTAL_BG  = NAVY
ROW_TOTAL_FG  = "#FFFFFF"
ROW_SUBTOT_BG = SKY
ROW_SUBTOT_FG = NAVY
ROW_FUENTE_BG = SURFACE
ROW_FUENTE_FG = INK_SOFT
ROW_HEADER_BG = SKY
ROW_HEADER_FG = NAVY
HDR_BG        = NAVY
HDR_FG        = "#FFFFFF"
HDR_FONT      = (FONT, SIZE_BASE, "bold")
ACCENT_BAR    = NAVY
ACCENT_BAR_H  = 0
CARD_SHADOW   = LINE
CARD_BORDER_L = LINE
