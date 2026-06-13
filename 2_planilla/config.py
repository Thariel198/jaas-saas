from pathlib import Path

# ── Rutas ──────────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
INPUTS_DIR  = BASE_DIR / "inputs"
OUTPUTS_DIR = BASE_DIR / "outputs"

# Inputs con fecha (se resuelven en runtime con el YYYY-MM detectado)
def lecturas_path(mes: str) -> Path:
    return INPUTS_DIR / "lecturas" / f"lecturas_planilla_{mes}.xlsx"

def deuda_path(mes: str) -> Path:
    return INPUTS_DIR / "deuda_anterior" / f"arrastre_deuda_{mes}.xlsx"

def corte_path(mes: str) -> Path:
    return INPUTS_DIR / "corte" / f"arrastre_corte_{mes}.xlsx"

# Inputs estáticos (el tesorero los actualiza cada mes)
CONVENIOS_PATH        = INPUTS_DIR / "convenios"        / "convenios.xlsx"
MULTAS_PATH           = INPUTS_DIR / "multas"           / "multas.xlsx"
ACUERDOS_PATH         = INPUTS_DIR / "acuerdos_asamblea" / "acuerdos_asamblea.xlsx"

def output_path(mes: str) -> Path:
    return OUTPUTS_DIR / f"planilla_{mes}.xlsx"

OUTPUT_SHEET = "Planilla"

# ── Reglas de negocio ──────────────────────────────────────────────────────
TARIFA_M3  = 1   # S/ por m³
TARIFA_MIN = 5   # consumo mínimo en soles
MANT_FIJO  = 3   # mantenimiento fijo universal

# ── Columnas requeridas por input ──────────────────────────────────────────
COLS_LECTURAS  = ["MZ", "LT", "NOMBRE", "MES_ANO", "MARC_ANT", "MARC_ACT", "M3"]
COLS_DEUDA     = ["MZ", "LT", "monto"]
COLS_CORTE     = ["MZ", "LT", "monto"]
COLS_CONVENIOS = ["MZ", "LT", "cuota_mes"]
COLS_MULTAS    = ["MZ", "LT", "monto_mes"]
COLS_ACUERDOS  = ["MZ", "LT", "monto_mes"]

# ── Columnas del output — orden exacto ────────────────────────────────────
OUTPUT_COLS = [
    "MZ", "LT", "NOMBRE", "MES_ANO",           # ¿Quién es?
    "MARC_ANT", "MARC_ACT", "M3",              # Lectura
    "MES_ACTUAL", "MANTENIMIENTO",             # Cobro — consumo
    "MES_ANTERIOR", "CORTE_RECONEXION",        # Cobro — arrastres
    "CONVENIO", "MULTA", "ACUERDOS_ASAMBLEA",  # Cobro — seguimiento
    "BLANCO", "DEVOLUCION",                    # Descuentos (4_pagos escribe negativos)
    "TOTAL_A_PAGAR",                           # Fórmula Excel
    "MONTO_YAPE", "MONTO_EFECTIVO",            # Pago (vacío → 4_pagos)
    "ESTADO", "FECHA_PAGO",
]

# ── Formato Excel — colores por sección ───────────────────────────────────
# Cada sección: header_bg, header_fg, data_bg, data_fg
SEC_QUIEN = {
    "header_bg": "EBF5FB", "header_fg": "1A5276",
    "data_bg":   "F4FAFF", "data_fg":   "1A5276",
}
SEC_LECTURA = {
    "header_bg": "E6F1FB", "header_fg": "0C447C",
    "data_bg":   "F0F8FF", "data_fg":   "111111",
}
SEC_CARGOS = {
    "header_bg": "E9F7EF", "header_fg": "1E5C3A",
    "data_bg":   "F4FBF7", "data_fg":   "1E5C3A",
}
SEC_DESCUENTOS = {
    "header_bg": "EDE9FE", "header_fg": "4C1D95",
    "data_bg":   "F5F3FF", "data_fg":   "6D28D9",
}
SEC_TOTAL = {
    "header_bg": "1E8449", "header_fg": "FFFFFF",
    "data_bg":   "D5F5E3", "data_fg":   "1E5C3A",
}
SEC_PAGO = {
    "header_bg": "F3E8FF", "header_fg": "5B21B6",
    "data_bg":   "FAF5FF", "data_fg":   "8B5CF6",
}

# Mapa columna → sección
COL_SECTION = {
    "MZ":       SEC_QUIEN, "LT":       SEC_QUIEN,
    "NOMBRE":   SEC_QUIEN, "MES_ANO":  SEC_QUIEN,
    "MARC_ANT": SEC_LECTURA, "MARC_ACT": SEC_LECTURA, "M3": SEC_LECTURA,
    "MES_ACTUAL":        SEC_CARGOS, "MANTENIMIENTO":    SEC_CARGOS,
    "MES_ANTERIOR":      SEC_CARGOS, "CORTE_RECONEXION": SEC_CARGOS,
    "CONVENIO":          SEC_CARGOS, "MULTA":            SEC_CARGOS,
    "ACUERDOS_ASAMBLEA": SEC_CARGOS,
    "BLANCO":     SEC_DESCUENTOS, "DEVOLUCION": SEC_DESCUENTOS,
    "TOTAL_A_PAGAR": SEC_TOTAL,
    "MONTO_YAPE":    SEC_PAGO, "MONTO_EFECTIVO": SEC_PAGO,
    "ESTADO":        SEC_PAGO, "FECHA_PAGO":     SEC_PAGO,
}

# Formato de número por columna
FMT_ENTERO = "0"
FMT_MONTO  = '"S/"\\ #,##0.00'
FMT_FECHA  = "DD/MM/YYYY"

COL_FORMAT = {
    "MZ": "@", "LT": "@", "NOMBRE": "@", "MES_ANO": "@",
    "MARC_ANT": FMT_ENTERO, "MARC_ACT": FMT_ENTERO, "M3": FMT_ENTERO,
    "MES_ACTUAL":        FMT_MONTO, "MANTENIMIENTO":    FMT_MONTO,
    "MES_ANTERIOR":      FMT_MONTO, "CORTE_RECONEXION": FMT_MONTO,
    "CONVENIO":          FMT_MONTO, "MULTA":            FMT_MONTO,
    "ACUERDOS_ASAMBLEA": FMT_MONTO,
    "BLANCO":     FMT_MONTO, "DEVOLUCION":     FMT_MONTO,
    "TOTAL_A_PAGAR":  FMT_MONTO,
    "MONTO_YAPE":     FMT_MONTO, "MONTO_EFECTIVO": FMT_MONTO,
    "ESTADO":         "@",        "FECHA_PAGO":     FMT_FECHA,
}

# Ancho de columna en unidades Excel
COL_WIDTH = {
    "MZ": 5, "LT": 6, "NOMBRE": 28, "MES_ANO": 10,
    "MARC_ANT": 10, "MARC_ACT": 10, "M3": 6,
    "MES_ACTUAL": 13, "MANTENIMIENTO": 13,
    "MES_ANTERIOR": 17, "CORTE_RECONEXION": 17,
    "CONVENIO": 13, "MULTA": 13, "ACUERDOS_ASAMBLEA": 19,
    "BLANCO": 13, "DEVOLUCION": 13,
    "TOTAL_A_PAGAR": 14,
    "MONTO_YAPE": 14, "MONTO_EFECTIVO": 14,
    "ESTADO": 11, "FECHA_PAGO": 12,
}

# Alturas de fila (puntos)
ROW_HEIGHT_SECTIONS = 18   # fila 1: encabezados de sección
ROW_HEIGHT_COLS     = 22   # fila 2: nombres de columna

FREEZE_PANES = "A3"
