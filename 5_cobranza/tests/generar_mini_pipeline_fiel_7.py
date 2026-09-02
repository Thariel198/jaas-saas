"""Run the real 5_cobranza flow in a temporary ledger and filter seven lots."""
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

THIS = Path(__file__).resolve()
ROOT = THIS.parents[2]
sys.path.insert(0, str(THIS.parent.parent))
import main as mod  # noqa: E402

REFERENCIAS = ROOT / "docs" / "aprendizaje" / "solucion de proble" / "referencias"
OUT = Path(r"C:\Users\wilde\AppData\Local\Temp\opencode\mini_pipeline_fiel_7_20260817")
INPUTS = OUT / "inputs"
OUTPUTS = REFERENCIAS
LIVE_LEDGER = ROOT / "shared" / "seguimiento_pueblo.xlsx"
TARGETS = {("I", "9"), ("L", "5"), ("P", "12"), ("P", "3"),
           ("Q", "5"), ("S", "2"), ("W", "5")}


def _prepare_output():
    if OUT.exists():
        backup = OUT.parent / f"{OUT.name}_backup_{datetime.now():%Y%m%d_%H%M%S}"
        shutil.copytree(OUT, backup)
        shutil.rmtree(OUT)
        print(f"BACKUP_PREVIO={backup}")
    INPUTS.mkdir(parents=True)
    OUTPUTS.mkdir(parents=True, exist_ok=True)


def main():
    _prepare_output()
    original_abonos = mod.ABONOS_REZAGADOS_PATH
    original_manifest = mod.ABONOS_MANIFEST_PATH
    original_ledger = mod.repo.SEGUIMIENTO_PATH
    try:
        if not LIVE_LEDGER.exists():
            raise FileNotFoundError(f"No existe el ledger vivo: {LIVE_LEDGER}")
        if Path(original_ledger).resolve() != LIVE_LEDGER.resolve():
            raise RuntimeError(
                f"El repositorio no apunta al ledger vivo: {original_ledger}"
            )
        abonos_copy = INPUTS / "abonos_rezagados.xlsx"
        manifest_copy = INPUTS / "abonos_rezagados_manifest_2026-08.json"
        shutil.copy2(original_abonos, abonos_copy)
        shutil.copy2(original_manifest, manifest_copy)
        mod.ABONOS_REZAGADOS_PATH = abonos_copy
        mod.ABONOS_MANIFEST_PATH = manifest_copy

        plan_path = mod._validar_inputs()
        _, mes_ano = mod._cargar_planilla(plan_path)
        mod._cargar_abonos_rezagados(mes_ano)
        sys.path.insert(0, str(ROOT / "4b_reclamos" / "herramienta"))
        import reporte_historico as rh  # noqa: E402

        eventos = mod.repo._leer_eventos()
        historicos = rh._cargar_historicos()
        mapa_raw = rh._cargar_mapa_raw()
        dfp = pd.read_excel(
            rh.comun._planilla_cobrado_path(mes_ano),
            sheet_name="planilla_cobrado", header=1,
        )
        source_df = pd.read_excel(abonos_copy, sheet_name="Abonos_Raw", header=1)
        nombres = mod.repo._lookup_nombres()
        redirects = rh._cargar_redirects()
        report_rows = []
        for mz, lt in sorted(TARGETS):
            tabla = rh.tabla_predio(
                mz, lt, historicos, eventos, dfp, mapa_raw,
                nombres.get((mz, lt), ""), incluir_abonos_rezagados=True,
                deuda_conceptos_desde_ledger=True,
            )
            tabla = rh.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
            actual = tabla[tabla["MES"].astype(str).str.strip() == mes_ano].iloc[-1]
            abono = source_df[
                (source_df["MZ"].astype(str).str.strip() == mz) &
                (source_df["LT"].astype(str).str.replace(".0", "", regex=False).str.strip() == lt) &
                (source_df["MES_ANO_APLICA"].astype(str).str[:7] == mes_ano)
            ]["MONTO"].sum()
            saldo = round(float(actual["DEUDA_TOTAL"] - actual["PAGO_TOTAL"]), 2)
            report_rows.append({
                "MZ": mz, "LT": lt, "MES_ANO": mes_ano,
                "TOTAL_DEUDA": round(float(actual["DEUDA_TOTAL"]), 2),
                "ABONO_REZAGADO": round(float(abono), 2),
                "TOTAL_PAGADO": round(float(actual["PAGO_TOTAL"]), 2),
                "SALDO": saldo,
                "ESTADO": "EXCESO" if saldo < 0 else ("PENDIENTE" if saldo == 0 else "PARCIAL"),
            })
        report = pd.DataFrame(report_rows)
        nuevos = pd.DataFrame(columns=["MZ", "LT", "MES", "CONCEPTO", "PAGO"])
        report.to_excel(OUTPUTS / "mini_resultado_pipeline_fiel.xlsx", index=False)
        nuevos.to_excel(OUTPUTS / "mini_ledger_predicho.xlsx", index=False)
        rh.generar_deuda_ledger(
            "2026-07",
            REFERENCIAS / "mini_pipeline_fiel_7_20260817.pdf",
            predios=TARGETS,
            incluir_abonos_rezagados=True,
            mini_abonos=source_df,
        )
        print(f"OK mini fiel: {len(report)} lotes · ledger vivo leído={LIVE_LEDGER}")
        print(f"OUTPUT={OUTPUTS}")
    finally:
        mod.ABONOS_REZAGADOS_PATH = original_abonos
        mod.ABONOS_MANIFEST_PATH = original_manifest
        mod.repo.SEGUIMIENTO_PATH = original_ledger


if __name__ == "__main__":
    main()
