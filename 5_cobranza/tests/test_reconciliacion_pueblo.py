"""tests/test_reconciliacion_pueblo.py — test sintético de _reconciliar_pagos_pueblo

Verifica el patrón de reconciliación por delta (metodología 3.6g) contra
seguimiento_repo: SET_DEBE (recalculado por 5_cobranza) − SET_TIENE
(Σ PAGO ya registrado) = delta a anotar. Casos forzados:

  1. Primer pago parcial      → registrar_pago(delta completo)
  2. Re-corrida sin cambios   → delta=0, no escribe nada
  3. Pago nuevo (incremento)  → registrar_pago(solo el delta incremental)
  4. Corrección a la baja     → registrar_ajuste(delta negativo)

Uso:
    python tests/test_reconciliacion_pueblo.py
"""
import sys
from pathlib import Path

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))   # 5_cobranza/ → para importar main

import main as mod  # noqa: E402

TEST_ROOT = THIS.parent / "_tmp_reconciliacion"
MES_TEST  = "2099-11"

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


def _usuario(mz, lt, multa, total_pagado):
    """Usuario sintético — solo MULTA tiene componente (agua/corte/acuerdos/convenio=0),
    así el total_pagado completo se puede rastrear sin ambigüedad de reparto."""
    return {
        "mz": mz, "lt": lt,
        "mes_actual": 0.0, "mantenimiento": 0.0, "mes_anterior": 0.0,
        "blanco_final": 0.0, "devolucion": 0.0, "corte_reconexion": 0.0,
        "multa": multa, "acuerdos_asamblea": 0.0, "convenio": 0.0,
        "total_pagado": total_pagado,
    }


def main():
    path = _reset_repo()

    # 1) Primer pago parcial: debe 50 de MULTA, paga 30
    r1 = _usuario("A", "1", multa=50.0, total_pagado=30.0)
    mod._reconciliar_pagos_pueblo([r1], MES_TEST)
    ya = mod.repo.pago_registrado("A", "1", "MULTA", MES_TEST)
    check(ya == 30.0, f"caso 1: pago_registrado=30 (obtuve {ya})")

    # 2) Re-corrida idéntica → delta=0, no debe cambiar nada
    mod._reconciliar_pagos_pueblo([r1], MES_TEST)
    ya2 = mod.repo.pago_registrado("A", "1", "MULTA", MES_TEST)
    check(ya2 == 30.0, f"caso 2: re-corrida sin cambios sigue en 30 (obtuve {ya2})")

    # 3) Pago nuevo: ahora pagó 45 en total (15 más) → delta incremental = 15
    r3 = _usuario("A", "1", multa=50.0, total_pagado=45.0)
    mod._reconciliar_pagos_pueblo([r3], MES_TEST)
    ya3 = mod.repo.pago_registrado("A", "1", "MULTA", MES_TEST)
    check(ya3 == 45.0, f"caso 3: tras pago incremental, pago_registrado=45 (obtuve {ya3})")

    # 4) Corrección a la baja: se detecta que en realidad solo pagó 20 (corrección
    # de un pago mal cargado) → debe generar AJUSTE de -25, no tocar el PAGO=45
    r4 = _usuario("A", "1", multa=50.0, total_pagado=20.0)
    mod._reconciliar_pagos_pueblo([r4], MES_TEST)
    ya4 = mod.repo.pago_registrado("A", "1", "MULTA", MES_TEST)
    check(ya4 == 45.0, f"caso 4: pago_registrado (solo PAGO) sigue en 45, el ajuste no lo toca (obtuve {ya4})")

    import pandas as pd
    df = pd.read_excel(path, header=1, dtype={"MZ": str, "LT": str, "CONCEPTO": str, "TIPO_EVENTO": str})
    ajustes = df[(df["TIPO_EVENTO"] == "AJUSTE") & (df["MZ"] == "A") & (df["LT"] == "1")]
    check(len(ajustes) == 1, f"caso 4: exactamente 1 evento AJUSTE (obtuve {len(ajustes)})")
    if len(ajustes):
        check(float(ajustes.iloc[0]["AJUSTE"]) == -25.0,
              f"caso 4: AJUSTE = -25.0 (obtuve {ajustes.iloc[0]['AJUSTE']})")

    # get_saldo tras todo: MULTA=50 nunca se registró como CARGO en este test
    # (foco es la reconciliación de pagos) → saldo = 0 - 45(pago) - 25(ajuste, ya negativo)
    saldo = mod.repo.get_saldo("A", "1", "MULTA", MES_TEST)
    check(saldo == -70.0, f"get_saldo refleja pago+ajuste acumulados: -70 (obtuve {saldo})")

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
