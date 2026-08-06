"""tests/test_reconciliacion_pueblo.py — test sintético de _reconciliar_pagos_pueblo

Verifica el patrón de reconciliación por delta (metodología 3.6g) contra
seguimiento_repo: SET_DEBE (recalculado por 5_cobranza) − SET_TIENE
(Σ PAGO ya registrado) = delta a anotar. Casos forzados:

  1. Primer pago parcial      → registrar_pago(delta completo)
  2. Re-corrida sin cambios   → delta=0, no escribe nada
  3. Pago nuevo (incremento)  → registrar_pago(solo el delta incremental)
  4. Corrección a la baja     → registrar_ajuste(delta negativo)
  5. Concepto ya saldado por una declaración manual que igual recibe plata del
     ciclo → se acredita y el saldo queda negativo A PROPÓSITO (es la señal de
     exceso); 5_cobranza solo avisa por log

El CARGO se siembra de entrada (06/08/2026): antes el test reconciliaba pagos
contra un concepto sin CARGO, cosa imposible en producción — la deuda que ve
5_cobranza viene de la planilla, y la planilla la lee del propio ledger.

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
    # El cargo tiene que existir: 5_cobranza reconcilia contra deuda real, y el
    # tope de _reconciliar_pagos_pueblo no deja acreditar más de lo que se debe.
    mod.repo.registrar_cargo("A", "1", "MULTA", MES_TEST, 50.0,
                             source="test", audit_ref="cargo_A_1_MULTA")

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

    # get_saldo tras todo: 50(cargo) - 45(pagos) - 25(ajuste, ya negativo) = -20
    saldo = mod.repo.get_saldo("A", "1", "MULTA", MES_TEST)
    check(saldo == -20.0, f"get_saldo refleja cargo+pago+ajuste acumulados: -20 (obtuve {saldo})")

    # 5) Concepto ya saldado por una declaración manual y que además recibe plata
    # del ciclo. Decisión 06/08/2026: se acredita igual y el saldo queda NEGATIVO
    # a propósito — es la señal de que ese predio pagó de más. Taparlo con un tope
    # silencioso borraría un exceso que puede corresponder devolver o dejar a favor.
    # 5_cobranza solo avisa por log; queda pendiente cruzarlo contra planilla_cobrado
    # para separar el exceso real del negativo por defecto del sistema.
    mod.repo.registrar_cargo("A", "2", "MULTA", MES_TEST, 50.0,
                             source="test", audit_ref="cargo_A_2_MULTA")
    mod.repo.registrar_pago("A", "2", "MULTA", MES_TEST, 50.0, source="manual",
                            audit_ref="declaracion_A_2", clase="DECLARACION_SECRETARIA")
    check(mod.repo.get_saldo("A", "2", "MULTA", MES_TEST) == 0.0,
          "caso 5: la declaración deja el saldo en 0")

    mod._reconciliar_pagos_pueblo([_usuario("A", "2", multa=50.0, total_pagado=50.0)], MES_TEST)
    saldo5 = mod.repo.get_saldo("A", "2", "MULTA", MES_TEST)
    check(saldo5 == -50.0, f"caso 5: el exceso queda visible como saldo -50 (obtuve {saldo5})")
    propio = mod.repo.pago_registrado("A", "2", "MULTA", MES_TEST, source="5_cobranza")
    check(propio == 50.0, f"caso 5: 5_cobranza sí acredita su parte (obtuve {propio})")
    ajeno = mod.repo.pago_registrado("A", "2", "MULTA", MES_TEST, source="manual")
    check(ajeno == 50.0, f"caso 5: la declaración manual sigue intacta (obtuve {ajeno})")

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
