"""Compare the seven list-cutoff lots against secretary expectations."""
import sys
from pathlib import Path

import pandas as pd

THIS = Path(__file__).resolve()
ROOT = THIS.parents[2]
MINI_OUT = Path(r"C:\Users\wilde\AppData\Local\Temp\opencode\mini_corrida_lista_corte_7_20260816")
NOTAS = ROOT / "4b_reclamos" / "pendientes_secretaria" / "notas_2026-07.xlsx"

TARGET_KEYS = {"I-9", "L-5", "P-12", "P-3", "Q-5", "S-2", "W-5"}

# These are business expectations extracted from the secretary's notes.
EXPECTED = {
    "I-9": {"estado": "FUERA_DE_CORTE", "abono_minimo": 50},
    "Q-5": {"estado": "FUERA_DE_CORTE", "abono_minimo": 69},
    "L-5": {"estado": "PENDIENTE_CONFIRMACION"},
    "S-2": {"estado": "PENDIENTE_CONFIRMACION"},
    "W-5": {"estado": "PENDIENTE_CONFIRMACION"},
    "P-3": {"estado": "SIN_NOTA"},
    "P-12": {"estado": "SIN_NOTA"},
}


def _key(row):
    mz = str(row["MZ"]).strip().upper()
    lt = str(row["LT"]).strip().replace(".0", "")
    return f"{mz}-{lt}"


def main():
    resultado = pd.read_excel(MINI_OUT / "outputs" / "mini_resultado_cascada.xlsx")
    resultado["KEY"] = resultado.apply(_key, axis=1)
    observado = resultado.set_index("KEY")

    notas = pd.read_excel(NOTAS, sheet_name="notas")
    notas["KEY"] = notas.apply(_key, axis=1)
    notas_por_lote = notas[notas["KEY"].isin(TARGET_KEYS)].groupby("KEY").size().to_dict()

    assert set(observado.index) == TARGET_KEYS, set(observado.index) ^ TARGET_KEYS
    assert set(EXPECTED) == TARGET_KEYS

    errores = []
    for key in sorted(TARGET_KEYS):
        expected = EXPECTED[key]
        row = observado.loc[key]
        actual_abono = float(row["ABONO_REZAGADO"])
        print(
            f"{key}: esperado={expected['estado']} "
            f"abono_observado=S/{actual_abono:.0f} "
            f"saldo_observado=S/{float(row['SALDO']):.0f} "
            f"notas={notas_por_lote.get(key, 0)}"
        )
        if "abono_minimo" in expected and actual_abono < expected["abono_minimo"]:
            errores.append(
                f"{key}: abono observado S/{actual_abono:.0f}, "
                f"se esperaba al menos S/{expected['abono_minimo']}"
            )

    if errores:
        raise AssertionError("\n".join(errores))
    print("OK: expectativas de secretaria compatibles con el mini-pipeline")


if __name__ == "__main__":
    main()
