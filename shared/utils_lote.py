"""
shared/utils_lote.py — Primitivos para correcciones de lote

Fuente: 5_cobranza/inputs/correcciones_lote.xlsx
        (gestionado por 5_cobranza, consumido por cualquier módulo que
         necesite remapear MZ+LT antes de buscar en planilla o DATA_boletas)
"""
from pathlib import Path

import pandas as pd

_BASE = Path(__file__).parent.parent
CORR_LOTE_PATH = _BASE / "5_cobranza" / "inputs" / "correcciones_lote.xlsx"


def _norm(v: str) -> str:
    return str(v).strip().upper().replace(" ", "")


def leer_correcciones_lote() -> dict:
    """
    Lee correcciones_lote.xlsx de 5_cobranza.
    Retorna {(mz_orig, lt_orig): (mz_dest, lt_dest)}.
    Retorna {} si el archivo no existe.
    """
    if not CORR_LOTE_PATH.exists():
        return {}
    df = pd.read_excel(CORR_LOTE_PATH, header=0, dtype=str)
    df.columns = [str(c).strip().upper() for c in df.columns]
    corr = {}
    for _, row in df.iterrows():
        mo = _norm(row.get("MZ_ORIGEN", ""))
        lo = _norm(row.get("LT_ORIGEN", ""))
        md = _norm(row.get("MZ_DESTINO", ""))
        ld = _norm(row.get("LT_DESTINO", ""))
        if mo and lo and md and ld:
            corr[(mo, lo)] = (md, ld)
    return corr
