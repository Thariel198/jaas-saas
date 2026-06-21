"""
shared/utils_sort_mz_lt.py — Orden canónico de MZ y LT para todo el pipeline.

Se usa en 0_padron (padron_reconciliado.xlsx) y 1_lecturas
(registro_operario_acumulado.xlsx) para que ambos archivos compartan
exactamente el mismo orden de filas.

MZ:  etapa 1 (A-Z una letra) primero, etapa 2 (A1-H1 compuestas) después.
LT:  numérico + sufijo como desempate, soporta 3, 3A, 3B, 11A, 11-B, 14 B...
"""
import re


def mz_orden(mz) -> tuple:
    """
    Etapa 1 (una sola letra): A, B, C … Z  → (0, letra, 0)
    Etapa 2 (letra + número): A1, B1 … H1  → (1, letra, número)
    """
    mz = str(mz).strip().upper()
    m = re.match(r"^([A-Z]+)(\d+)$", mz)
    if m:
        return (1, m.group(1), int(m.group(2)))
    return (0, mz, 0)


def lt_orden(lt) -> tuple:
    """
    Ordena lotes numéricamente con sufijo como desempate.
    Soporta: "3", "3A", "3B", "11A", "11-B", "14 B", "16C"
    → (número, sufijo)  ej. (3, ""), (3, "A"), (11, "B")
    """
    lt = str(lt).strip().upper()
    m = re.match(r"^(\d+)\s*-?\s*([A-Z]*)$", lt)
    if m:
        return (int(m.group(1)), m.group(2))
    return (0, lt)


def clave_orden(mz, lt) -> tuple:
    """Clave compuesta para sorted() — orden canónico (mz, lt) del pipeline."""
    mzo = mz_orden(mz)
    lto = lt_orden(lt)
    return (mzo[0], mzo[1], mzo[2], lto[0], lto[1])
