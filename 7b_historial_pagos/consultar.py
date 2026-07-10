"""
7b_historial_pagos/consultar.py — Lectores read-only (tools del agente)

API PÚBLICA:
    analizar_reclamo(mz, lt) -> DataFrame   todos los pagos de un predio, todos los
                                             meses/canales cargados en el store
    reporte_pagos(mz, lt)    -> str         vista de texto mes a mes, para explicar
                                             al usuario

Nunca escribe. Lee todos los YYYY-MM_historial.xlsx del store — seguro para que un
agente lo invoque sin riesgo de mutar el ledger.
"""

from pathlib import Path

import pandas as pd

import historial_repo

STORE_DIR = historial_repo.STORE_DIR


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip().upper().replace(" ", "")


def _meses_disponibles() -> list[str]:
    if not STORE_DIR.exists():
        return []
    meses = []
    for p in STORE_DIR.glob("*_historial.xlsx"):
        meses.append(p.name.replace("_historial.xlsx", ""))
    return sorted(meses)


def analizar_reclamo(mz: str, lt: str) -> pd.DataFrame:
    """Todos los eventos (pago + correccion) de un predio, en todos los meses
    cargados al store, ordenados por FECHA."""
    mz_n, lt_n = _norm(mz), _norm(lt)
    partes = []
    for mes in _meses_disponibles():
        df = historial_repo.leer_eventos(mes)
        if df.empty:
            continue
        sub = df[(df["MZ"].map(_norm) == mz_n) & (df["LT"].map(_norm) == lt_n)]
        if not sub.empty:
            partes.append(sub)
    if not partes:
        return pd.DataFrame(columns=historial_repo.COLUMNAS)
    return pd.concat(partes, ignore_index=True)


def reporte_pagos(mz: str, lt: str) -> str:
    """Vista de texto mes a mes — lo que se le explica al vecino."""
    df = analizar_reclamo(mz, lt)
    if df.empty:
        return f"Predio {mz}-{lt}: sin pagos registrados en el histórico."

    lineas = [f"Predio {mz}-{lt} — {len(df)} evento(s) en el histórico:"]
    for mes in sorted(df["MES_CICLO"].unique()):
        sub = df[df["MES_CICLO"] == mes]
        total = sub["MONTO"].astype(float).sum()
        lineas.append(f"\n  {mes} — total S/ {total:.2f}")
        for _, r in sub.sort_values("FECHA").iterrows():
            tag = "" if r["FUENTE"] == "pago" else f" [{r['FUENTE']}: {r['MOTIVO']}]"
            lineas.append(f"    {r['FECHA']:20} {r['CANAL']:9} S/ {float(r['MONTO']):>8.2f}  "
                           f"{r['REFERENCIA']}{tag}")
    return "\n".join(lineas)


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Uso: python consultar.py <MZ> <LT>")
        sys.exit(1)
    print(reporte_pagos(sys.argv[1], sys.argv[2]))
