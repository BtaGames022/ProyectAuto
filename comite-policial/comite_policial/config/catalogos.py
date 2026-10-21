"""Catálogos estructurales de las tablas del informe semanal."""
from comite_policial.config.settings import settings

ANO_ANT = settings.ano_ant
ANO_ACT = settings.ano_act
PERIODO = settings.periodo

TIPOS_ARMAS    = ["Armas de Fuego", "Aire Comprimido", "Armas Hechizas",
                  "Armas Fogueo", "Armas de Fantasía"]
TIPOS_DROGAS_G = ["Cannabis Sativa (g.)", "Clorhidrato Cocaína (g.)",
                  "Cocaína Base (g.)", "Sintéticas (g.) (*)"]
TIPOS_DROGAS_P = ["Cannabis Sativa Plantas (unidad)"]
BH_PREF        = ["BH Valparaíso", "BH Los Andes", "BH San Antonio"]
MODALIDADES    = ["Arma de Fuego", "Arma Cortante",
                  "Elemento Contundente", "Otras Modalidades"]
RESULTADOS     = ["Con Detenido PDI", "Esclarecido", "Identificado", "Investigación"]
HOMI_COMP_CATS = ["Homicidios", "Femicidios", "Parricidios", "Infanticidios"]
MIGR_UNITS     = ["POLINT VPO", "POLINT VMA", "POLINT LAD",
                  "POLINT SAN", "SEPI QTO", "SEPI IPA"]
VEH_ESTADO     = ["Incautado", "Recuperado"]
VEH_TIPOS      = ["Camión", "Motocicleta", "Tractocamión",
                  "Veh. Furgón", "Maquinaria", "Automóvil (Suv)"]
BRISEX_UNITS   = ["Brisex VPO", "Brisex SAN", "Brisex LAD"]
BANDAS_PREFS   = ["PROVPO", "PREVMA", "PROSAN", "PROLAD"]
DET_PREFS      = ["PREVMA", "PROLAD", "PROVPO", "PROSAN", "PRECOCEN", "JENACIBЕР"]


# ──────────────────────────────────────────────────────────────────────────────
# IX. CONCURRENCIAS Y DILIGENCIAS RELEVANTES  (recuadros de 2 columnas)
# Cada tipo define los campos de la columna derecha, en orden:
#   (etiqueta, clase)   clase: "linea" = Etiqueta: valor (campos cortos, sin salto entre sí)
#                              "parrafo" = Etiqueta: texto corrido
#                              "lista" = Etiqueta: y debajo un ítem por línea
# ──────────────────────────────────────────────────────────────────────────────
_F, _L = ("Fecha", "linea"), ("Lugar", "linea")

TIPOS_RECUADRO = {
    "Detención / Infracción Ley 20.000 (MT-0)": [
        _F, _L, ("Detenidos", "parrafo"), ("Diligencias", "parrafo"),
        ("Droga incautada", "lista"), ("Especies incautadas", "lista"),
        ("Dinero incautado", "lista"), ("Avalúo droga incautada", "lista"),
        ("Resultados", "parrafo")],
    "Concurrencia por muerte / suicidio": [
        _F, _L, ("Víctima", "parrafo"), ("Hechos", "parrafo"),
        ("Diligencias", "parrafo")],
    "Homicidio / Lesiones graves": [
        _F, _L, ("Detenido", "parrafo"), ("Víctima", "parrafo"),
        ("Hechos", "parrafo"), ("Diligencias", "parrafo"),
        ("Resultados", "parrafo")],
    "Delitos sexuales": [
        _F, _L, ("Víctima", "parrafo"), ("Detenido", "parrafo"),
        ("Hechos", "parrafo"), ("Diligencias", "parrafo"),
        ("Resultados", "parrafo")],
    "Denuncia / Amenaza": [
        _F, _L, ("Hechos", "parrafo"), ("Diligencia", "parrafo"),
        ("Resultados", "parrafo")],
    "Fiscalización migratoria": [
        _F, _L, ("Diligencias", "parrafo"), ("Resultados", "lista")],
    "Hallazgo / Recuperación de vehículo": [
        _F, _L, ("Diligencia", "parrafo"), ("Especies recuperadas", "lista"),
        ("Avalúo de las especies recuperadas", "lista"),
        ("Resultados", "parrafo")],
    "Operativo con incautación de especies": [
        _F, _L, ("Detenidos", "parrafo"), ("Diligencias", "parrafo"),
        ("Especies incautadas", "lista"),
        ("Avalúos de las especies incautadas", "lista"),
        ("Resultados", "parrafo")],
    "Otro (formato libre)": [
        _F, _L, ("Hechos", "parrafo"), ("Diligencias", "parrafo"),
        ("Resultados", "parrafo")],
}

UNIDADES_RECUADRO = [
    "Brigada de Investigación Criminal Villa Alemana",
    "Brigada de Investigación Criminal Valparaíso",
    "Brigada de Investigación Criminal Viña del Mar",
    "Brigada de Investigación Criminal Quillota",
    "Brigada de Investigación Criminal La Calera",
    "Brigada de Investigación Criminal Quintero",
    "Brigada de Investigación Criminal Los Andes",
    "Brigada de Investigación Criminal San Felipe",
    "Brigada de Homicidios Valparaíso",
    "Brigada de Homicidios San Antonio",
    "Brigada de Robos Valparaíso",
    "Brigada Investigadora de Delitos Sexuales San Antonio",
    "Brigada Investigadora de Delitos Sexuales Valparaíso",
    "Brigada Investigadora de Delitos contra el Medio Ambiente y Patrimonio Cultural Valparaíso",
    "Brigada Investigadora de Delitos en Recintos Portuarios Valparaíso",
    "Departamento de Migraciones y Policía Internacional Valparaíso",
    "Departamento de Migraciones y Policía Internacional Quintero",
    "Departamento de Migraciones Isla de Pascua",
    "Sección Migraciones y Policía Internacional San Antonio",
]

GRUPOS_RECUADRO = ["", "Grupo MT-0"]
