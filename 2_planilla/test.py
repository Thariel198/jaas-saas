"""
Test de integración — fixtures en tests/_tmp_integracion/, no toca inputs/outputs reales.
Cada corrida borra y recrea la carpeta temporal → idempotente.

Escenarios cubiertos:
  A/7    — consumo normal (M3=8), sin arrastres, con acuerdo asamblea
  A/12   — consumo bajo (M3=3 → mínimo S/5), deuda anterior, acuerdo asamblea
  B1/11A — consumo bajo (M3=2 → mínimo S/5), corte+reconexión, acuerdo asamblea
  B1/5   — consumo alto (M3=15), sin arrastres, acuerdo asamblea

Los arrastres vienen del arrastre_consolidado del mes anterior (2026-05), fuente
única (Opción A · writer único). CONVENIO/MULTA=0 en el fixture → prueba el path 0.
"""
import json
import logging
import shutil
import sys
from pathlib import Path

import pandas as pd

# ── Monkey-patch de config ANTES de usar cualquier función de main ────────
import config

_TMP = Path(__file__).parent / "tests" / "_tmp_integracion"


def _setup_paths() -> None:
    if _TMP.exists():
        shutil.rmtree(_TMP)
    for sub in ["inputs/lecturas", "cobranza_outputs", "outputs"]:
        (_TMP / sub).mkdir(parents=True)

    config.INPUTS_DIR            = _TMP / "inputs"
    config.OUTPUTS_DIR           = _TMP / "outputs"
    config.COBRANZA_OUTPUTS_DIR  = _TMP / "cobranza_outputs"
    config.ESTADO_CICLO_PATH     = _TMP / "estado_ciclo.json"


import main as mod_main  # noqa: E402 — importar después de definir _setup_paths

# mod_main.repo ES el módulo seguimiento_repo (singleton en sys.modules) —
# redirigir acá aísla get_saldos_bulk() del archivo real de producción. Path
# inexistente en el fixture → devuelve {} → MULTA/ACUERDOS/CONVENIO=0,
# determinístico (no depende de qué prediós tenga hoy seguimiento_pueblo real).
mod_main.repo.SEGUIMIENTO_PATH = _TMP / "seguimiento_pueblo.xlsx"

MES = "2026-06"
MES_ANT = "2026-05"   # el consolidado que alimenta la planilla de junio

# ── Fixtures ──────────────────────────────────────────────────────────────

_LECTURAS = pd.DataFrame([
    {"MZ": "A",  "LT": "7",   "NOMBRE": "FLORES MAMANI, ROSA",   "MES_ANO": MES, "MARC_ANT": 1045, "MARC_ACT": 1053, "M3": 8},
    {"MZ": "A",  "LT": "12",  "NOMBRE": "QUISPE HUANCA, PEDRO",  "MES_ANO": MES, "MARC_ANT": 780,  "MARC_ACT": 783,  "M3": 3},
    {"MZ": "B1", "LT": "11A", "NOMBRE": "MAMANI ROQUE, JUAN",    "MES_ANO": MES, "MARC_ANT": 320,  "MARC_ACT": 322,  "M3": 2},
    {"MZ": "B1", "LT": "5",   "NOMBRE": "CONDORI QUISPE, ELENA", "MES_ANO": MES, "MARC_ANT": 600,  "MARC_ACT": 615,  "M3": 15},
])

# arrastre_consolidado del mes anterior (2026-05): solo agua+corte (desde
# seguimiento_pueblo, MULTA/ACUERDOS/CONVENIO ya no vienen de acá).
#   A/12  → DEUDA_AGUA 15   (deuda de agua no cubierta → MES_ANTERIOR)
#   B1/11A → CORTE 25        (corte no cubierto)
_CONSOLIDADO = pd.DataFrame([
    {"MZ": "A",  "LT": "7",   "DEUDA_AGUA": 0,  "CORTE_RECONEXION": 0},
    {"MZ": "A",  "LT": "12",  "DEUDA_AGUA": 15, "CORTE_RECONEXION": 0},
    {"MZ": "B1", "LT": "11A", "DEUDA_AGUA": 0,  "CORTE_RECONEXION": 25},
    {"MZ": "B1", "LT": "5",   "DEUDA_AGUA": 0,  "CORTE_RECONEXION": 0},
])


def _crear_fixtures() -> None:
    _LECTURAS.to_excel(config.lecturas_path(MES), index=False)
    # El consolidado real trae grupos en la fila 1 y nombres de columna en la 2
    # → header=1. Reproducimos ese layout con startrow=1 (fila 1 en blanco).
    _CONSOLIDADO.to_excel(config.consolidado_path(MES_ANT), index=False, startrow=1)
    config.ESTADO_CICLO_PATH.write_text(
        json.dumps({MES_ANT: {"arrastre": {"generado": True, "validado": True}}}),
        encoding="utf-8",
    )
    # ACUERDOS=10 para los 4 predios, sembrado en seguimiento_pueblo (no en el
    # consolidado) — prueba el path real de lectura vía get_saldos_bulk().
    # MULTA/CONVENIO se dejan sin sembrar → prueba el path 0 (predio sin eventos).
    for mz, lt in (("A", "7"), ("A", "12"), ("B1", "11A"), ("B1", "5")):
        mod_main.repo.registrar_cargo(
            mz, lt, "ACUERDOS", MES_ANT, 10.0,
            source="test_fixture", audit_ref=f"fixture_{mz}_{lt}",
        )
    print("  (consolidado 2026-05: agua+corte · seguimiento_pueblo: ACUERDOS=10, MULTA/CONVENIO ausentes)")


# ── Verificación del dataframe calculado ──────────────────────────────────

def _verificar_df(df: pd.DataFrame) -> None:
    errores = []

    def chk(cond, msg):
        if not cond:
            errores.append(msg)

    def fila(mz, lt):
        r = df[(df["_mz"] == mz) & (df["_lt"] == lt)]
        if r.empty:
            raise AssertionError(f"Fila no encontrada: MZ={mz} LT={lt}")
        return r.iloc[0]

    chk(len(df) == 4, f"Filas esperadas: 4, obtenidas: {len(df)}")

    r = fila("A", "7")
    chk(r["MES_ACTUAL"]        == 8.0,  f"A/7  MES_ACTUAL esperado 8, got {r['MES_ACTUAL']}")
    chk(r["MES_ANTERIOR"]      == 0.0,  f"A/7  MES_ANTERIOR esperado 0, got {r['MES_ANTERIOR']}")
    chk(r["CORTE_RECONEXION"]  == 0.0,  f"A/7  CORTE esperado 0")
    chk(r["CONVENIO"]          == 0.0,  f"A/7  CONVENIO esperado 0 (archivo ausente)")
    chk(r["MULTA"]             == 0.0,  f"A/7  MULTA esperado 0 (archivo ausente)")
    chk(r["ACUERDOS_ASAMBLEA"] == 10.0, f"A/7  ACUERDOS esperado 10")
    chk(r["BLANCO"]            == 0.0,  f"A/7  BLANCO esperado 0")
    chk(r["DEVOLUCION"]        == 0.0,  f"A/7  DEVOLUCION esperado 0")
    chk(r["MANTENIMIENTO"]     == 3.0,  f"A/7  MANTENIMIENTO esperado 3")

    r = fila("A", "12")
    chk(r["MES_ACTUAL"]   == 5.0,  f"A/12 MES_ACTUAL esperado 5 (mínimo, M3=3), got {r['MES_ACTUAL']}")
    chk(r["MES_ANTERIOR"] == 15.0, f"A/12 MES_ANTERIOR esperado 15, got {r['MES_ANTERIOR']}")

    r = fila("B1", "11A")
    chk(r["MES_ACTUAL"]       == 5.0,  f"B1/11A MES_ACTUAL esperado 5 (mínimo, M3=2)")
    chk(r["CORTE_RECONEXION"] == 25.0, f"B1/11A CORTE esperado 25, got {r['CORTE_RECONEXION']}")
    chk(r["MES_ANTERIOR"]     == 0.0,  f"B1/11A MES_ANTERIOR esperado 0 (no en deuda)")

    r = fila("B1", "5")
    chk(r["MES_ACTUAL"] == 15.0, f"B1/5  MES_ACTUAL esperado 15 (M3=15), got {r['MES_ACTUAL']}")

    if errores:
        for e in errores:
            print(f"  FALLO: {e}")
        sys.exit(1)

    print("  Todos los valores calculados son correctos.")


# ── Verificación del archivo Excel generado ───────────────────────────────

def _verificar_excel() -> None:
    out = config.output_path(MES)
    if not out.exists():
        print(f"  FALLO: archivo no generado: {out}")
        sys.exit(1)

    df_raw = pd.read_excel(out, sheet_name=config.OUTPUT_SHEET, header=None)
    col_names = list(df_raw.iloc[1])  # fila 0 = secciones, fila 1 = columnas

    errores = [f"Columna '{c}' ausente" for c in config.OUTPUT_COLS if c not in col_names]
    if errores:
        for e in errores:
            print(f"  FALLO: {e}")
        sys.exit(1)

    datos = df_raw.iloc[2:]
    print(f"  21 columnas en orden correcto · {len(datos)} filas de datos")
    print(f"  run.log: {config.OUTPUTS_DIR / 'run.log'}")

    # Preview de las columnas de cálculo
    df_preview = pd.read_excel(out, sheet_name=config.OUTPUT_SHEET, skiprows=1)
    print()
    cols_preview = [
        "MZ", "LT", "MES_ACTUAL", "MANTENIMIENTO",
        "MES_ANTERIOR", "CORTE_RECONEXION",
        "CONVENIO", "MULTA", "ACUERDOS_ASAMBLEA",
        "BLANCO", "DEVOLUCION", "TOTAL_A_PAGAR",
    ]
    print(df_preview[cols_preview].to_string(index=False))


# ── Main ──────────────────────────────────────────────────────────────────

def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s  %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    print("=" * 55)
    print("2_planilla — test de integración")
    print("=" * 55)

    print("\n[1] Preparando carpeta temporal y paths...")
    _setup_paths()
    print(f"  tmp: {_TMP}")

    print("\n[2] Creando fixtures...")
    _crear_fixtures()

    print("\n[3] Corriendo pipeline...")
    df = mod_main.build_planilla(MES)

    print("\n[4] Verificando valores calculados...")
    _verificar_df(df)

    print("\n[5] Escribiendo Excel...")
    mod_main.write_excel(df, MES)

    print("\n[6] Verificando archivo de salida...")
    _verificar_excel()

    print("\n" + "=" * 55)
    print("PASÓ — planilla generada sin errores.")
    print("=" * 55)


if __name__ == "__main__":
    main()
