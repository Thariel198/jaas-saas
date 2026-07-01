"""6b_corte_multas/config.py — constantes compartidas entre los 3 scripts.

Módulo espejo de 6_corte. Penaliza la deuda de MULTA + ACUERDOS_ASAMBLEA
(no el saldo de consumo de agua). Diferencias clave con 6_corte:

  - Elegibilidad: DEUDA_MULTA > 0, calculada por cascada de prioridad
    (agua → corte → multa → acuerdos → convenio). CONVENIO se excluye:
    es deuda grande de pago opcional, nunca dispara corte.
  - La penalidad de S/40 se escribe en CORTE_RECONEXION — el mismo cargo de
    corte físico y reconexión que usa 6_corte. El módulo trata exactamente de
    eso: cortar y reconectar. Al ir a CORTE_RECONEXION (col 13, dentro de
    SUM(J:R)=TOTAL_A_PAGAR), 5_cobranza la propaga sola a SALDO en ciclo 2.
  - registro_cortes vive en shared/ (estado compartido con 6_corte).
  - Prioridad: quien ya va a corte por AGUA (6_corte) se excluye de aquí —
    un solo corte físico, un solo cargo CORTE_RECONEXION.

Cada script (generar_lista_multas, aplicar_penalidad_multas,
seguimiento_multas) importa de acá.
"""
from pathlib import Path

# ── PATHS ────────────────────────────────────────────────────────────────────
ROOT        = Path(__file__).parent
INPUTS_DIR  = ROOT / "inputs"
OUTPUTS_DIR = ROOT / "outputs"
BACKUP_DIR  = ROOT / "backup"

# Inputs externos — sólo lectura
COBRANZA_DIR          = ROOT.parent / "5_cobranza" / "outputs"
PLANILLA_COBRADO_PATH = COBRANZA_DIR / "planilla_cobrado.xlsx"

# Pagos en efectivo — para enriquecer pagaron_penalidad_multas con trazabilidad
PAGOS_EFECTIVO_PATH = ROOT.parent / "4_pagos" / "efectivo" / "outputs" / "pagos_efectivo.xlsx"

# Pagos por Yape — para enriquecer pagaron_penalidad_multas con trazabilidad
PAGOS_YAPE_TEPAGO_PATH = (
    ROOT.parent / "4_pagos" / "yape" / "motor_matching" / "outputs" / "pagos_yape_tepago.xlsx"
)

# Reclamos (4b_reclamos/outputs) — generar_lista_multas.py cruza por (MZ, LT)
RECLAMOS_DIR = ROOT.parent / "4b_reclamos" / "outputs"

def reclamos_path(mes_ano: str) -> Path:
    """reclamos_YYYY-MM.xlsx — lista cruda de reclamos del mes."""
    return RECLAMOS_DIR / f"reclamos_{mes_ano}.xlsx"

def resolucion_reclamos_path(mes_ano: str) -> Path:
    """resolucion_reclamos_YYYY-MM.xlsx — decisiones del supervisor."""
    return RECLAMOS_DIR / f"resolucion_reclamos_{mes_ano}.xlsx"

# shared/ — recursos cross-módulo
SHARED_DIR       = ROOT.parent / "shared"
PLANILLA_MES_DIR = SHARED_DIR / "planilla_mes"

# Reportes procesados de Yape — definen el ancla del periodo de cobro.
SHARED_PROCESADO_DIR = SHARED_DIR / "reporte_acumulado_procesado"

# registro_cortes — estado persistente COMPARTIDO entre 6_corte y 6b_corte_multas.
# generar_lista_multas lo lee para excluir predios ya CORTADO/EXONERADO.
# seguimiento_multas hace append de los nuevos CORTADO por multa.
REGISTRO_CORTES_PATH = SHARED_DIR / "registro_cortes.xlsx"

# lista_corte de 6_corte — para excluir a quien ya va a corte por AGUA este ciclo
# (prioridad: agua sobre multa, un solo corte físico).
LISTA_CORTE_AGUA_PATH = ROOT.parent / "6_corte" / "outputs" / "lista_corte.xlsx"

# pagos en efectivo — fuente del tag "compromiso" para generar_compromisos.py
PAGOS_EFECTIVO_OUT_PATH = ROOT.parent / "4_pagos" / "efectivo" / "outputs" / "pagos_efectivo.xlsx"

# compromisos.xlsx — fuente de verdad de los socios que firmaron compromiso de pago
# de su multa. Estado persistente append-idempotente. generar_lista_multas lo lee:
# compromiso VIGENTE → EJECUTAR_CORTE=NO. Es la ÚNICA exoneración del corte por multa
# (reclamos y EXONERADO/CORTADO de registro_cortes son del mundo agua, no aplican aquí).
COMPROMISOS_PATH = ROOT / "compromisos.xlsx"
TAG_COMPROMISO   = "compromi"      # cubre "compromiso" y "compromizo"
TAG_EXONERACION  = "exonera"       # COMPROMISO_EXONERACION si el comentario lo incluye
COMPROMISO_FECHA_LIMITE = "2026-07-05"   # default editable de FECHA_LIMITE

# Outputs propios
LISTA_MULTAS_PATH     = OUTPUTS_DIR / "lista_multas.xlsx"
AUDIT_PENALIDAD_PATH  = OUTPUTS_DIR / "audit_penalidad_multas.xlsx"
PAGARON_PATH          = OUTPUTS_DIR / "pagaron_penalidad_multas.xlsx"
CORTE_FISICO_PATH     = OUTPUTS_DIR / "corte_fisico_multas.xlsx"

def arrastre_multa_path(mes_ano: str) -> Path:
    """arrastre_multa_YYYY-MM.xlsx — se consume en 2_planilla del mes siguiente."""
    return OUTPUTS_DIR / f"arrastre_multa_{mes_ano}.xlsx"

# ── REGLAS DE NEGOCIO ────────────────────────────────────────────────────────
# TODO: Junio 2026 — secretaria aprobó S/40 directo desde Día 0 (sin ventana en S/20),
# igual que su hermano 6_corte. Valor normal es PENALIDAD=20, PENALIDAD_FINAL=40
# (escalada tras corte físico). Volver a 20/40 el próximo ciclo.
PENALIDAD       = 40.0   # S/ — temporal: S/40 directo (normal: S/20)
PENALIDAD_FINAL = 40.0   # S/ — sin escalada este mes (normal: S/40 solo para cortados)
TOL             = 0.005  # tolerancia de redondeo en comparaciones de S/

# Columna de la planilla donde se carga la penalidad. Es la MISMA que 6_corte:
# el cargo de corte/reconexión. 6b y 6_corte nunca tocan el mismo predio el
# mismo ciclo (registro_cortes + exclusión por lista_corte de agua lo garantizan).
COL_PENALIDAD = "CORTE_RECONEXION"
