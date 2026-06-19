"""
aplicar_correcciones.py — Aplica VALOR_A_CORREGIR a DATA_boletas vía repo

USO:
    python aplicar_correcciones.py --mes 2026-06

LOGICA:
    1. Lee outputs/resolucion_reclamos_{mes}.xlsx
    2. Filtra ESTADO=RESUELTO + CAMPO + VALOR_A_CORREGIR no vacíos
    3. Por cada fila → shared/data_boletas_repo.apply_correction(...)
    4. Resumen en consola: aplicados / skipped (idempotente) / errores

EL SCRIPT NO ESCRIBE ARCHIVOS DIRECTAMENTE — toda la persistencia es vía el repo:
    - update    → 3_boletas/inputs/DATA_boletas.xlsx
    - append    → shared/data_boletas_audit.xlsx
    - snapshot  → 3_boletas/backup/DATA_boletas/

Contratos visuales: docs/diagrama_4b_reclamos.html (etapa 4)
                    shared/docs/diagrama_repo_pattern.html
"""

import argparse
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "shared"))

import pandas as pd

from main import OUTPUTS_DIR, _clean
from data_boletas_repo import apply_correction

log = logging.getLogger(__name__)


def _ruta_resolucion(mes: str) -> Path:
    return OUTPUTS_DIR / f"resolucion_reclamos_{mes}.xlsx"


def _leer_resolucion(mes: str) -> pd.DataFrame:
    p = _ruta_resolucion(mes)
    if not p.exists():
        log.warning(f"resolucion_reclamos_{mes}.xlsx no encontrado: {p}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(p, sheet_name="Correcciones", header=1, dtype=str)
        return df.fillna("")
    except Exception as e:
        log.error(f"Error leyendo {p.name}: {e}")
        return pd.DataFrame()


def _audit_ref(row: dict) -> str:
    """Convención 4b_reclamos: MESA|FECHA_COBRO."""
    mesa  = _clean(row.get("MESA", ""))
    fecha = _clean(row.get("FECHA_COBRO", ""))
    return f"{mesa}|{fecha}"


def main(mes: str) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
        force=True,
    )
    log.info(f"=== 4b_reclamos/aplicar_correcciones.py — mes {mes} ===")

    df = _leer_resolucion(mes)
    if df.empty:
        log.info("Sin filas en resolucion — nada que aplicar")
        return

    resueltos = df[
        (df["ESTADO"].str.strip().str.upper() == "RESUELTO") &
        (df["CAMPO"].str.strip() != "") &
        (df["VALOR_A_CORREGIR"].str.strip() != "")
    ]
    log.info(f"Candidatos (RESUELTO + CAMPO + VALOR_A_CORREGIR): {len(resueltos)}")

    if resueltos.empty:
        log.info("Ningún reclamo cumple el filtro — nada que aplicar")
        return

    aplicados = 0
    skipped   = 0
    errores   = 0

    for _, row in resueltos.iterrows():
        mz        = _clean(row.get("MZ", ""))
        lt        = _clean(row.get("LT", ""))
        campo     = _clean(row.get("CAMPO", ""))
        valor     = _clean(row.get("VALOR_A_CORREGIR", ""))
        motivo    = _clean(row.get("RESOLUCION", "")) or _clean(row.get("RECLAMO", ""))
        audit_ref = _audit_ref(row.to_dict())

        try:
            result = apply_correction(
                mz=mz, lt=lt, campo=campo, valor=valor,
                source="4b_reclamos",
                audit_ref=audit_ref,
                motivo=motivo,
            )
            if result["skipped"]:
                skipped += 1
                log.info(f"  [SKIP] ya aplicado - MZ={mz} LT={lt} CAMPO={campo}")
            else:
                aplicados += 1
                log.info(
                    f"  [OK]   aplicado - MZ={mz} LT={lt} CAMPO={campo}: "
                    f"{result['valor_antes']} -> {result['valor_despues']}"
                )
        except Exception as e:
            errores += 1
            log.error(f"  [ERR]  fallo - MZ={mz} LT={lt} CAMPO={campo}: {e}")

    log.info("")
    log.info(f"Resumen: aplicados={aplicados}  skipped={skipped}  errores={errores}")
    log.info("=== completado ===")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Aplica correcciones acordadas a DATA_boletas vía repo")
    parser.add_argument("--mes", required=True, help="Mes a procesar (YYYY-MM)")
    args = parser.parse_args()
    main(args.mes)
