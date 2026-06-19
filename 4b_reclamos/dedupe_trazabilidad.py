"""
dedupe_trazabilidad.py — One-off cleanup de duplicados en trazabilidad_reclamos.xlsx

Causa: main.py 2da corrida appendeaba sin chequear si la fila ya estaba.
Esto se arregló en _append_trazabilidad (fix idempotencia). Este script limpia
los duplicados acumulados antes del fix.

Clave de deduplicación: (MESA, MZ, LT, FECHA_COBRO).
Keep='last' — preserva la versión más reciente (con datos sincronizados).
Backup automático antes de tocar.

ELIMINAR ESTE ARCHIVO después de correrlo — es one-off, no debe persistir.
"""
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
from openpyxl import Workbook

from main import (
    BASE_DIR,
    _COLS_TRAZAB, _SECCIONES_TRAZAB,
    _ruta_trazab,
    _write_headers, _write_fila,
)


def main() -> None:
    p = _ruta_trazab()
    if not p.exists():
        print(f"No existe: {p}")
        return

    df = pd.read_excel(p, sheet_name="Trazabilidad", header=1, dtype=str).fillna("")
    antes = len(df)
    print(f"Antes: {antes} filas")

    backup_dir = BASE_DIR / "backup" / "trazabilidad"
    backup_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup = backup_dir / f"trazabilidad_reclamos_predupe_{ts}.xlsx"
    shutil.copy2(p, backup)
    print(f"Backup: {backup.name}")

    df = df.drop_duplicates(
        subset=["MESA", "MZ", "LT", "FECHA_COBRO"], keep="last"
    ).reset_index(drop=True)
    despues = len(df)
    print(f"Después: {despues} filas  (eliminados: {antes - despues})")

    if despues == antes:
        print("No había duplicados — no se reescribe nada")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Trazabilidad"
    _write_headers(ws, _COLS_TRAZAB, _SECCIONES_TRAZAB)
    for ri, (_, row) in enumerate(df.iterrows(), start=3):
        _write_fila(ws, ri, row.to_dict(), _COLS_TRAZAB)
    wb.save(p)
    print(f"Guardado limpio: {p.name}")


if __name__ == "__main__":
    main()
