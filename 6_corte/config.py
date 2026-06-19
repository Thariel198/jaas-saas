"""6_corte/config.py — constantes compartidas entre los 3 scripts del módulo.

Un solo lugar para paths, reglas de negocio y tolerancia. Cada script
(generar_lista, aplicar_penalidad, seguimiento) importa de acá.
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

# Pagos en efectivo — para enriquecer lista_corte con MESA + COBRADOR
PAGOS_EFECTIVO_PATH = ROOT.parent / "4_pagos" / "efectivo" / "outputs" / "pagos_efectivo.xlsx"

# Reclamos (4b_reclamos/outputs) — generar_lista.py cruza por (MZ, LT)
RECLAMOS_DIR = ROOT.parent / "4b_reclamos" / "outputs"

def reclamos_path(mes_ano: str) -> Path:
    """reclamos_YYYY-MM.xlsx — lista cruda de reclamos del mes."""
    return RECLAMOS_DIR / f"reclamos_{mes_ano}.xlsx"

def resolucion_reclamos_path(mes_ano: str) -> Path:
    """resolucion_reclamos_YYYY-MM.xlsx — decisiones del supervisor (FUNDADO/RECHAZADO/EN_REVISION)."""
    return RECLAMOS_DIR / f"resolucion_reclamos_{mes_ano}.xlsx"

# shared/planilla_mes — único archivo externo que escribimos (aplicar_penalidad)
SHARED_DIR       = ROOT.parent / "shared"
PLANILLA_MES_DIR = SHARED_DIR / "planilla_mes"

# Inputs propios — estado persistente
REGISTRO_CORTES_PATH = INPUTS_DIR / "registro_cortes.xlsx"

# Outputs propios
LISTA_CORTE_PATH      = OUTPUTS_DIR / "lista_corte.xlsx"
AUDIT_PENALIDAD_PATH  = OUTPUTS_DIR / "audit_penalidad.xlsx"
PAGARON_PATH          = OUTPUTS_DIR / "pagaron_penalidad.xlsx"
CORTE_FISICO_PATH     = OUTPUTS_DIR / "corte_fisico.xlsx"

def arrastre_corte_path(mes_ano: str) -> Path:
    """arrastre_corte_YYYY-MM.xlsx — se consume en 2_planilla del mes siguiente."""
    return OUTPUTS_DIR / f"arrastre_corte_{mes_ano}.xlsx"

# ── REGLAS DE NEGOCIO ────────────────────────────────────────────────────────
PENALIDAD        = 20.0   # S/ — penalidad inicial al generar lista_corte
PENALIDAD_FINAL  = 40.0   # S/ — escalada para los efectivamente cortados (20 + 20 reconexión)
MES_ANTERIOR_MIN = 8      # ≥ 8 confirma elegibilidad para corte
TOL              = 0.005  # tolerancia de redondeo en comparaciones de S/
