"""Guard tests for the approved 18 delayed-payment rows."""
import sys
import json
import tempfile
from pathlib import Path

import pandas as pd

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))
import main as mod  # noqa: E402


def main():
    assert mod._fecha_str("2026-06-03") == "03/06/2026"
    assert mod._fecha_hora_str("2026-06-03") == "03/06/2026"

    df = pd.read_excel(mod.ABONOS_REZAGADOS_PATH, header=1)
    df.columns = mod._norm_cols(df)

    manifest = json.loads(mod.ABONOS_MANIFEST_PATH.read_text(encoding="utf-8"))
    expected = {
        mod._abono_manifest_key(row)
        for row in manifest
        if row["ESTADO"] == "CONFIRMADO" and row["MES_ANO_APLICA"] == "2026-08"
    }
    active = mod._validar_abonos_manifest(df, "2026-08")

    usuarios, mes_ano = mod._cargar_planilla(mod._validar_inputs())
    abonos = mod._cargar_abonos_rezagados(mes_ano)
    resultado, _ = mod._calcular(
        usuarios, [], [], {}, {}, {}, {}, 1, set(), abonos_rezagados=abonos
    )
    f1_4 = next(row for row in resultado if row["key"] == "F1-4")
    assert f1_4["abono_rezagado"] == 101.0
    assert f1_4["total_pagado"] == 101.0
    assert f1_4["saldo"] == 16.0
    aplicaciones, pendientes = mod._aplicaciones_por_fuente(f1_4, mes_ano)
    aplicado = aplicaciones["abonos_rezagados"]
    assert sum(aplicado.values()) == 101.0
    assert aplicado["AGUA"] == 51.0
    assert aplicado["MULTA"] == 50.0
    assert pendientes["AGUA_ACT"] == 13.0
    assert pendientes["MANT_ACT"] == 3.0
    assert sum(pendientes.values()) == f1_4["saldo"]
    l_5 = next(row for row in resultado if row["key"] == "L-5")
    assert l_5["total_a_pagar"] == 176.0
    assert l_5["abono_rezagado"] == 126.0
    assert l_5["saldo"] == 50.0

    manifest = df.iloc[0:0].copy()
    extra = {"MZ": "F1", "LT": "4", "MONTO": 10, "MES_CICLO": "2026-07", "MES_ANO_APLICA": "2026-08"}
    bad = pd.concat([df, pd.DataFrame([extra])], ignore_index=True)
    try:
        mod._validar_abonos_manifest(bad, "2026-08")
    except RuntimeError:
        pass
    else:
        raise AssertionError("una fila adicional debe bloquear la corrida")

    original_path = mod.repo.SEGUIMIENTO_PATH
    with tempfile.TemporaryDirectory() as tmp:
        mod.repo.SEGUIMIENTO_PATH = Path(tmp) / "seguimiento_pueblo.xlsx"
        mod.repo.registrar_cargo("Z", "1", "CONVENIO", "2026-08", 50.0,
                                 source="test", audit_ref="cargo_Z_1")
        row = {
            "mz": "Z", "lt": "1", "mes_actual": 0.0, "mantenimiento": 0.0,
            "mes_anterior": 100.0, "blanco_final": 0.0, "devolucion": 0.0,
            "corte_reconexion": 0.0, "convenio": 50.0, "acuerdos_asamblea": 0.0,
            "multa": 0.0, "total_pagado": 150.0,
            "total_pagado_normal": 100.0, "abono_rezagado": 50.0,
        }
        objetivos = mod._objetivos_ledger([row], "2026-08")
        mod.repo.reconciliar_objetivos_batch("2026-08", "a" * 64, objetivos)
        ledger = pd.read_excel(mod.repo.SEGUIMIENTO_PATH, header=1)
        abono = ledger[(ledger["SOURCE"] == "abonos_rezagados") & (ledger["CONCEPTO"] == "CONVENIO")]
        assert len(abono) == 1 and float(abono.iloc[0]["PAGO"]) == 50.0
    mod.repo.SEGUIMIENTO_PATH = original_path

    print(f"OK manifest: {len(active)} activos, fuente y manifest coinciden, extra bloqueado")


if __name__ == "__main__":
    main()
