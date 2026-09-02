"""Run the real cascade against every delayed-payment source row in isolation."""
import sys
import tempfile
from pathlib import Path

import pandas as pd

THIS = Path(__file__).resolve()
ROOT = THIS.parents[2]
sys.path.insert(0, str(THIS.parent.parent))
import main as mod  # noqa: E402


def _month(value):
    text = str(value or "").strip()[:7]
    return text if text in {"2026-07", "2026-08"} else "2026-08"


def main():
    source = pd.read_excel(mod.ABONOS_REZAGADOS_PATH, header=1)
    source.columns = mod._norm_cols(source)
    assert len(source) == 42, len(source)

    reports = []
    for mes_ano, planilla in (
        ("2026-07", ROOT / "shared" / "planilla_mes" / "planilla_2026-07.xlsx"),
        ("2026-08", ROOT / "shared" / "planilla_mes" / "planilla_2026-08.xlsx"),
    ):
        usuarios, loaded_month = mod._cargar_planilla(planilla)
        assert loaded_month == mes_ano, (loaded_month, mes_ano)
        subset = source[source["MES_ANO_APLICA"].map(_month) == mes_ano]
        abonos = {}
        for _, row in subset.iterrows():
            key = (mod._norm_mz(row["MZ"]), mod._norm_lt(row["LT"]))
            amount = mod._float(row["MONTO"])
            closed, current = abonos.get(key, (0.0, 0.0))
            if str(row.get("MES_CICLO", "")).strip()[:7] == mes_ano:
                current += amount
            else:
                closed += amount
            abonos[key] = (closed, current)

        resultado, _ = mod._calcular(
            usuarios, [], [], {}, {}, {}, {}, 1, set(),
            abonos_rezagados=abonos,
        )
        by_key = {row["key"]: row for row in resultado}
        for key, (closed, current) in abonos.items():
            row = by_key.get(f"{key[0]}-{key[1]}")
            assert row is not None, key
            applied = round(closed + current, 2)
            pueblo = mod._descomponer_pago(row, applied)
            reports.append({
                "MZ": key[0],
                "LT": key[1],
                "MES_ANO_APLICA": mes_ano,
                "ABONO_FUENTE": applied,
                "TOTAL_DEUDA": row["total_a_pagar"],
                "TOTAL_PAGADO": row["total_pagado"],
                "SALDO": row["saldo"],
                "CONVENIO": pueblo["CONVENIO"],
                "ACUERDOS": pueblo["ACUERDOS"],
                "MULTA": pueblo["MULTA"],
            })

    output = Path(tempfile.mkdtemp(prefix="mini_ledger_abonos_")) / "mini_ledger.xlsx"
    pd.DataFrame(reports).sort_values(["MES_ANO_APLICA", "MZ", "LT"]).to_excel(
        output, index=False
    )
    print(f"OK mini-ledger: {len(source)} filas fuente, {len(reports)} predio/mes")
    print(f"OUTPUT={output}")


if __name__ == "__main__":
    main()
