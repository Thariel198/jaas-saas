"""Constantes y paths del módulo 1_lecturas.

Modificar si la junta decide cambiar umbrales o si se renombra alguna ruta.
Las anomalías que dependen de estas constantes están documentadas en
docs/arquitectura_1_lecturas.html sección "Reglas de validación".
"""
from pathlib import Path

# ── Umbrales de anomalías ──────────────────────────────────────────────────────
M3_MINIMO   = 5.0   # consumo declarado por el operario cuando aplica el mínimo de cobro
M3_EXCESIVO = 50.0  # consumo calculado > esto dispara EXCESIVO
SALTO_MESES = 4     # ciclos previos usados para calcular el promedio histórico
SALTO_FACTOR = 3.0  # factor sobre el promedio que dispara SALTO_HISTORICO

# Si MARC_ACT < MARC_ANT (sin obs=M que legitime):
#   ant − act ≤ UMBRAL_INVERSION  →  MEDIDOR_INVERTIDO (el medidor está al revés)
#   ant − act >  UMBRAL_INVERSION →  POSIBLE_CAMBIO_MEDIDOR (lo cambiaron sin reportar)
UMBRAL_INVERSION = M3_EXCESIVO

# ── Códigos válidos de obs_operario ────────────────────────────────────────────
# M = medidor cambiado (legitima MEDIDOR_INVERTIDO / POSIBLE_CAMBIO_MEDIDOR → MEDIDOR_CAMBIADO)
# F = fuga visible reportada (informativo siempre)
# P = predio cerrado o inaccesible (legitima SIN_LECTURA)
OBS_CODIGOS = {"M", "F", "P"}

# ── Paths del módulo ──────────────────────────────────────────────────────────
ROOT = Path(__file__).parent

INPUTS_DIR  = ROOT / "inputs"
OUTPUTS_DIR = ROOT / "outputs"
BACKUPS_DIR = INPUTS_DIR / "backups"
SYNC_DIR    = OUTPUTS_DIR / "sync"

# Inputs (uno lo llena el operario, el otro lo lee/actualiza main.py + sync)
REGISTRO_MES_PATH        = INPUTS_DIR / "registro_operario_mes.xlsx"
REGISTRO_ACUMULADO_PATH  = INPUTS_DIR / "registro_operario_acumulado.xlsx"

# Sync — input externo (0_padron) y outputs propios
PADRON_RECONCILIADO_PATH = ROOT.parent / "0_padron" / "02_matching" / "outputs" / "padron_reconciliado.xlsx"
REPORTE_SYNC_PATH        = SYNC_DIR / "reporte_sincronizacion.xlsx"
TRAZABILIDAD_SYNC_PATH   = SYNC_DIR / "trazabilidad_sincronizacion.xlsx"

# Outputs por ciclo (YYYY-MM se reemplaza en runtime)
def correcciones_path(mes_ano: str) -> Path:
    return OUTPUTS_DIR / f"correcciones_{mes_ano}.xlsx"

def trazabilidad_path(mes_ano: str) -> Path:
    return OUTPUTS_DIR / f"trazabilidad_{mes_ano}.xlsx"

def lecturas_planilla_path(mes_ano: str) -> Path:
    return OUTPUTS_DIR / f"lecturas_planilla_{mes_ano}.xlsx"

def orden_verificacion_path(mes_ano: str) -> Path:
    return OUTPUTS_DIR / f"orden_verificacion_{mes_ano}.pdf"

LOG_PATH = OUTPUTS_DIR / "run.log"

# ── Schema del acumulado ──────────────────────────────────────────────────────
# Columnas fijas a la izquierda del bloque mensual.
# La 4ta columna SIN_SERVICIO se introdujo el 2026-06-20 (decisión 9).
# Al leer, si fila1[col4] es YYYY-MM se asume schema legacy (sin SIN_SERVICIO).
COLS_FIJAS = ("MZ", "LT", "NOMBRE", "SIN_SERVICIO")
N_COLS_FIJAS = len(COLS_FIJAS)
