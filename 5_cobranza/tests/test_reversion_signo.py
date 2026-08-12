"""tests/test_reversion_signo.py — la SECUENCIA completa de la reversión de un pago.

El test que faltaba. `test_reconciliacion_pueblo.py` verifica un solo paso a la vez;
el bug del signo solo se ve recorriendo la secuencia entera, porque el error vive en
el par (lo que se ESCRIBE, lo que se LEE) y cada mitad por separado parece coherente:

    corrida → el insumo encoge → re-corrida → el pago reaparece

Qué es "el insumo encoge": una re-corrida (típicamente `--force`) donde la cascada ya
no le asigna a ese concepto la plata que le había asignado antes — porque se corrigió
el crudo de pagos, o porque un humano sacó esa deuda de la planilla.

La regla, en una línea: revertir un PAGO es DEVOLVER la deuda, así que el saldo tiene
que SUBIR. La columna AJUSTE del ledger está en unidades de DEUDA (`_registrar`:
`saldo += monto`), mientras `delta` está en unidades de CRÉDITO — cruzar la frontera
sin traducir es lo que dejaba el saldo 2× el monto por debajo de la verdad
(D-16 ACUERDOS y D1-6 MULTA en julio 2026, ver LEER_ANTES.md).

Uso:
    python tests/test_reversion_signo.py
"""
import sys
from pathlib import Path

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))   # 5_cobranza/ → para importar main

import main as mod  # noqa: E402

TEST_ROOT = THIS.parent / "_tmp_reversion_signo"
MES_TEST  = "2099-12"
CARGO     = 200.0
PAGO      = 75.0

errores = []


def check(cond, msg):
    print(("OK  " if cond else "FAIL") + " " + msg)
    if not cond:
        errores.append(msg)


def _reset_repo():
    TEST_ROOT.mkdir(parents=True, exist_ok=True)
    path = TEST_ROOT / "seguimiento_pueblo.xlsx"
    if path.exists():
        path.unlink()
    mod.repo.SEGUIMIENTO_PATH = path
    return path


def _usuario(total_pagado):
    """Solo MULTA tiene componente, así el reparto de la cascada es inequívoco."""
    return {
        "mz": "A", "lt": "1",
        "mes_actual": 0.0, "mantenimiento": 0.0, "mes_anterior": 0.0,
        "blanco_final": 0.0, "devolucion": 0.0, "corte_reconexion": 0.0,
        "multa": CARGO, "acuerdos_asamblea": 0.0, "convenio": 0.0,
        "total_pagado": total_pagado,
    }


def _eventos(path):
    import pandas as pd
    df = pd.read_excel(path, header=1, dtype={"MZ": str, "LT": str, "TIPO_EVENTO": str})
    return df[df["TIPO_EVENTO"] == "AJUSTE"], len(df)


def main():
    path = _reset_repo()
    mod.repo.registrar_cargo("A", "1", "MULTA", MES_TEST, CARGO,
                             source="test", audit_ref="cargo_A_1_MULTA")
    saldo = lambda: mod.repo.get_saldo("A", "1", "MULTA", MES_TEST)
    check(saldo() == CARGO, f"paso 0: el CARGO deja saldo {CARGO} (obtuve {saldo()})")

    # ── paso 1: corrida normal, el vecino pagó 75 de los 200 ──────────────────
    mod._reconciliar_pagos_pueblo([_usuario(PAGO)], MES_TEST)
    check(saldo() == CARGO - PAGO,
          f"paso 1: tras acreditar {PAGO}, saldo {CARGO - PAGO} (obtuve {saldo()})")

    # ── paso 2: el insumo encoge — ese pago ya no le corresponde a MULTA ──────
    # Revertir el pago DEVUELVE la deuda: el saldo tiene que volver al cargo entero.
    mod._reconciliar_pagos_pueblo([_usuario(0.0)], MES_TEST)
    ajustes, _ = _eventos(path)
    check(len(ajustes) == 1, f"paso 2: exactamente 1 evento AJUSTE (obtuve {len(ajustes)})")
    if len(ajustes):
        check(float(ajustes.iloc[0]["AJUSTE"]) == PAGO,
              f"paso 2: el AJUSTE se escribe en unidades de DEUDA, +{PAGO} "
              f"(obtuve {ajustes.iloc[0]['AJUSTE']})")
    check(saldo() == CARGO,
          f"paso 2: la deuda vuelve entera, saldo {CARGO} (obtuve {saldo()})")

    # ── paso 3: re-corrida idéntica → idempotencia, no se escribe nada ────────
    _, n_antes = _eventos(path)
    mod._reconciliar_pagos_pueblo([_usuario(0.0)], MES_TEST)
    ajustes3, n_despues = _eventos(path)
    check(n_despues == n_antes,
          f"paso 3: re-corrida sin cambios no escribe (eventos {n_antes} -> {n_despues})")
    check(len(ajustes3) == 1, f"paso 3: sigue habiendo 1 solo AJUSTE (obtuve {len(ajustes3)})")
    check(saldo() == CARGO, f"paso 3: saldo estable en {CARGO} (obtuve {saldo()})")

    # ── paso 4: el pago reaparece (se corrigió el crudo) → vuelve a acreditarse ─
    mod._reconciliar_pagos_pueblo([_usuario(PAGO)], MES_TEST)
    check(saldo() == CARGO - PAGO,
          f"paso 4: el pago vuelve a acreditarse, saldo {CARGO - PAGO} (obtuve {saldo()})")
    check(mod.repo.pago_registrado("A", "1", "MULTA", MES_TEST, source="5_cobranza") == PAGO * 2,
          "paso 4: los dos PAGO quedan en el ledger (append-only: nada se borra)")

    # ── el detector de producción: el saldo nunca queda POR DEBAJO de la verdad ─
    # Con el bug, paso 2 dejaba saldo 50 (2 × 75 por debajo de 200) y encima
    # POSITIVO, o sea invisible para el chequeo de "0 saldos negativos".
    check(saldo() >= 0, f"el saldo nunca queda negativo en esta secuencia (obtuve {saldo()})")

    if path.exists():
        path.unlink()

    print()
    if errores:
        print(f"FALLARON {len(errores)}:")
        for e in errores:
            print(" -", e)
        sys.exit(1)
    else:
        print("TODOS LOS CHECKS PASARON")


if __name__ == "__main__":
    main()
