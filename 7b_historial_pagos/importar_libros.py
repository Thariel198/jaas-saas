"""
7b_historial_pagos/importar_libros.py — Productor: libros Excel viejos (Drive)

USO:
    python importar_libros.py --libro "ruta/al/libro.xlsx" --mes 2026-05

LOGICA:
    Lee un libro Excel del formato legacy (hojas Efectivo + Reporte) y registra
    cada pago vía historial_repo.registrar_pago(). No escribe el store directamente.

    Hoja "Efectivo" (bloque 1: Mz/Lt/Monto/Llave/Comentario/Fecha):
        CANAL=efectivo. MZ/LT siempre presentes (el cobrador está físicamente ahí)
        → ESTADO=identificado siempre. Sin COBRADOR en el libro legacy —
        REFERENCIA = "(sin cobrador — libro histórico)".

    Hoja "Reporte" (yape crudo):
        Solo TIPO="TE PAGÓ" (ingreso real a la JASS) — se excluye "PAGASTE"
        (egreso, plata que sale de la JASS, no es pago de un usuario).
        mz="blanco" (literal) → ESTADO=blanco, PREDIO vacío.
        REFERENCIA = Origen (quién envió). FECHA con hora:min:seg — es lo que,
        junto con REFERENCIA, identifica el pago yape de forma única.

    --mes es OBLIGATORIO y no se infiere del nombre del archivo: los títulos de
    los libros viejos usan rangos de fecha que no calzan un mes calendario
    (ej. "mayo-planilla 2026-03-11 A 2026-04-10"). El mes lo decide un humano.

Contrato visual: docs/formato_evento_pago.html
"""

import argparse
import logging
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent / "shared"))
import data_boletas_repo  # noqa: E402

import historial_repo  # noqa: E402

log = logging.getLogger(__name__)


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip().upper().replace(" ", "")


def _limpio(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in ("nan", "None", "NaT") else s


def _importar_efectivo(path: Path, mes: str, lookup: dict) -> dict:
    """Hoja Efectivo — solo el bloque 1 (Mz/Lt/Monto/Llave/Comentario/Fecha).
    El bloque 2 (Mz.1/Lt.1/...) es un espejo de fórmula del mismo libro Excel —
    usarlo también duplicaría cada pago."""
    df = pd.read_excel(path, sheet_name="Efectivo", header=0, dtype=str).fillna("")
    n_ok = n_skip = n_err = 0
    for i, row in df.iterrows():
        mz, lt = _limpio(row.get("Mz", "")), _limpio(row.get("Lt", ""))
        monto = _limpio(row.get("Monto", ""))
        if not mz or not lt or not monto:
            continue
        try:
            monto_f = float(monto)
        except ValueError:
            n_err += 1
            log.warning(f"Efectivo fila {i+2}: monto no numérico {monto!r} — saltada")
            continue

        fecha_raw = row.get("Fecha", "")
        try:
            fecha = pd.to_datetime(fecha_raw, dayfirst=True).strftime("%d/%m/%Y")
        except Exception:
            fecha = _limpio(fecha_raw)

        nombre = lookup.get((_norm(mz), _norm(lt)), {}).get("NOMBRES", "")
        origen_archivo = f"{path.name}#Efectivo-r{i+2}"

        r = historial_repo.registrar_pago(
            mz, lt, nombre, "efectivo", monto_f, fecha, mes,
            "(sin cobrador - libro historico)",
            origen_archivo=origen_archivo, estado="identificado",
        )
        n_skip += 1 if r["skipped"] else 0
        n_ok += 0 if r["skipped"] else 1

    return {"ok": n_ok, "skip": n_skip, "err": n_err, "total_filas": len(df)}


def _importar_reporte_yape(path: Path, mes: str, lookup: dict) -> dict:
    df = pd.read_excel(path, sheet_name="Reporte", header=0, dtype=str).fillna("")
    tcol = next(c for c in df.columns if "Tipo de Transacc" in c)
    fcol = next(c for c in df.columns if "Fecha de operaci" in c)

    mask_ingreso = df[tcol].str.contains("TE PAG", case=False, na=False)
    te = df[mask_ingreso]
    n_ok = n_skip = n_err = n_blanco = 0
    for i, row in te.iterrows():
        monto = _limpio(row.get("Monto", ""))
        if not monto:
            continue
        try:
            monto_f = float(monto)
        except ValueError:
            n_err += 1
            log.warning(f"Reporte fila {i+2}: monto no numérico {monto!r} — saltada")
            continue

        fecha_raw = row.get(fcol, "")
        try:
            fecha = pd.to_datetime(fecha_raw, dayfirst=True).strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            fecha = _limpio(fecha_raw)

        mz_raw = _limpio(row.get("mz", ""))
        lt_raw = _limpio(row.get("lt", ""))
        origen = _limpio(row.get("Origen", "")) or "(origen desconocido)"
        origen_archivo = f"{path.name}#Reporte-r{i+2}"

        if mz_raw.lower() == "blanco" or not mz_raw or not lt_raw:
            if mz_raw and mz_raw.lower() != "blanco":
                log.warning(f"Reporte fila {i+2}: MZ/LT no identifican un predio "
                            f"(mz={mz_raw!r} lt={lt_raw!r}) — tratado como blanco")
            n_blanco += 1
            r = historial_repo.registrar_pago(
                "", "", "", "yape", monto_f, fecha, mes, origen,
                origen_archivo=origen_archivo, estado="blanco",
            )
        else:
            nombre = lookup.get((_norm(mz_raw), _norm(lt_raw)), {}).get("NOMBRES", "")
            r = historial_repo.registrar_pago(
                mz_raw, lt_raw, nombre, "yape", monto_f, fecha, mes, origen,
                origen_archivo=origen_archivo, estado="identificado",
            )
        n_skip += 1 if r["skipped"] else 0
        n_ok += 0 if r["skipped"] else 1

    return {"ok": n_ok, "skip": n_skip, "err": n_err, "blanco": n_blanco,
            "total_te_pago": len(te), "total_filas": len(df)}


def main(libro: str, mes: str) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
        force=True,
    )
    path = Path(libro)
    if not path.exists():
        raise FileNotFoundError(f"Libro no encontrado: {path}")

    log.info(f"=== importar_libros.py — {path.name} -> mes_ciclo={mes} ===")
    lookup = data_boletas_repo.get_predio_lookup()

    r_ef = _importar_efectivo(path, mes, lookup)
    log.info(f"Efectivo: {r_ef['ok']} registrados, {r_ef['skip']} ya existían, "
              f"{r_ef['err']} con error (de {r_ef['total_filas']} filas)")

    r_ya = _importar_reporte_yape(path, mes, lookup)
    log.info(f"Reporte (yape, solo TE PAGÓ): {r_ya['ok']} registrados, {r_ya['skip']} ya existían, "
              f"{r_ya['err']} con error, {r_ya['blanco']} en blanco "
              f"(de {r_ya['total_te_pago']} TE PAGÓ / {r_ya['total_filas']} filas totales)")

    log.info("=== completado ===")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--libro", required=True, help="Ruta al libro Excel legacy")
    ap.add_argument("--mes", required=True, help="Mes ciclo YYYY-MM al que se atribuye (no se infiere del nombre del archivo)")
    args = ap.parse_args()
    main(args.libro, args.mes)
