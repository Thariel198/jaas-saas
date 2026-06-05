import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from main import _float, _norm_lt, _estado, _norm_fecha, _calcular_ajuste


# ── _float ────────────────────────────────────────────────────────────────────
def test_float_normal():
    assert _float("45.5") == 45.5
    assert _float(100)    == 100.0

def test_float_coma():
    assert _float("45,5") == 45.5

def test_float_invalido():
    assert _float("")    == 0.0
    assert _float(None)  == 0.0
    assert _float("abc") == 0.0


# ── _norm_lt ──────────────────────────────────────────────────────────────────
def test_norm_lt_decimal():
    assert _norm_lt("1.0")  == "1"
    assert _norm_lt("3.0")  == "3"

def test_norm_lt_alfanumerico():
    assert _norm_lt("11A") == "11A"

def test_norm_lt_vacio():
    assert _norm_lt("")    == ""
    assert _norm_lt(None)  == ""
    assert _norm_lt("nan") == ""


# ── _estado ───────────────────────────────────────────────────────────────────
def test_estado_cancelado():
    assert _estado(0.0,   50.0) == "CANCELADO"
    assert _estado(0.003, 50.0) == "CANCELADO"   # dentro de tolerancia
    assert _estado(-0.003, 50.0) == "CANCELADO"

def test_estado_exceso():
    assert _estado(-10.0, 65.0) == "EXCESO"
    assert _estado(-0.5,  55.0) == "EXCESO"

def test_estado_parcial():
    assert _estado(20.0, 30.0) == "PARCIAL"    # debe algo, pagó algo
    assert _estado(0.1,   5.0) == "PARCIAL"

def test_estado_pendiente():
    assert _estado(50.0, 0.0)   == "PENDIENTE"  # debe todo, no pagó nada
    assert _estado(30.0, 0.003) == "PENDIENTE"  # total_pagado bajo tolerancia


# ── _norm_fecha ───────────────────────────────────────────────────────────────
def test_norm_fecha_none():
    assert _norm_fecha(None) == ""

def test_norm_fecha_datetime_obj():
    dt = datetime(2025, 5, 1, 14, 30)
    assert _norm_fecha(dt) == "2025-05-01 14:30"

def test_norm_fecha_str_dmy():
    assert _norm_fecha("01/05/2025 14:30:00") == "2025-05-01 14:30"
    assert _norm_fecha("01/05/2025 14:30")    == "2025-05-01 14:30"

def test_norm_fecha_str_ymd():
    assert _norm_fecha("2025-05-01 14:30:00") == "2025-05-01 14:30"
    assert _norm_fecha("2025-05-01 14:30")    == "2025-05-01 14:30"

def test_norm_fecha_str_microsegundos():
    # Excel a veces exporta con microsegundos — el split(".") los descarta
    assert _norm_fecha("2025-05-01 14:30:00.123456") == "2025-05-01 14:30"

def test_norm_fecha_str_desconocido():
    # Formato no reconocido → devuelve el string tal cual (sin microsegundos)
    assert _norm_fecha("texto-invalido") == "texto-invalido"


# ── _calcular_ajuste ──────────────────────────────────────────────────────────
def test_ajuste_sumar_pago():
    # Agregar S/20 al lote nuevo (signo +1) — pagó algo, queda saldo → PARCIAL
    pago, topag, saldo, estado = _calcular_ajuste(0.0, 0.0, 50.0, 20.0, +1)
    assert pago   == 20.0
    assert topag  == 20.0
    assert saldo  == 30.0
    assert estado == "PARCIAL"

def test_ajuste_restar_pago():
    # Retirar S/20 del lote anterior — queda S/10 pagado, saldo S/40 → PARCIAL
    pago, topag, saldo, estado = _calcular_ajuste(30.0, 0.0, 50.0, 20.0, -1)
    assert pago   == 10.0
    assert topag  == 10.0
    assert saldo  == 40.0
    assert estado == "PARCIAL"

def test_ajuste_cancelado():
    # Sumar exactamente el total adeudado → CANCELADO
    pago, topag, saldo, estado = _calcular_ajuste(0.0, 0.0, 50.0, 50.0, +1)
    assert pago   == 50.0
    assert topag  == 50.0
    assert abs(saldo) <= 0.005
    assert estado == "CANCELADO"

def test_ajuste_exceso():
    # Sumar más del total → EXCESO, saldo negativo
    pago, topag, saldo, estado = _calcular_ajuste(0.0, 0.0, 50.0, 65.0, +1)
    assert pago   == 65.0
    assert topag  == 65.0
    assert saldo  == -15.0
    assert estado == "EXCESO"

def test_ajuste_no_puede_quedar_negativo():
    # Restar más de lo que hay → max(0, ...) clampea en 0
    pago, topag, saldo, estado = _calcular_ajuste(10.0, 0.0, 50.0, 20.0, -1)
    assert pago   == 0.0
    assert topag  == 0.0
    assert saldo  == 50.0
    assert estado == "PENDIENTE"

def test_ajuste_con_otro_pago():
    # Hay S/5 de efectivo · se suman S/15 de yape → topag = 20, saldo = 30
    pago, topag, saldo, estado = _calcular_ajuste(0.0, 5.0, 50.0, 15.0, +1)
    assert pago   == 15.0
    assert topag  == 20.0
    assert saldo  == 30.0
    assert estado == "PARCIAL"

def test_ajuste_parcial():
    # Sumar S/30 de S/50 → PARCIAL
    pago, topag, saldo, estado = _calcular_ajuste(0.0, 0.0, 50.0, 30.0, +1)
    assert pago   == 30.0
    assert topag  == 30.0
    assert saldo  == 20.0
    assert estado == "PARCIAL"

def test_ajuste_restar_con_otro():
    # Lote anterior: yape=25, efectivo=10 · restar yape 25 → topag=10, saldo=40, PARCIAL (efectivo queda)
    pago, topag, saldo, estado = _calcular_ajuste(25.0, 10.0, 50.0, 25.0, -1)
    assert pago   == 0.0
    assert topag  == 10.0
    assert saldo  == 40.0
    assert estado == "PARCIAL"


if __name__ == "__main__":
    tests = [
        test_float_normal, test_float_coma, test_float_invalido,
        test_norm_lt_decimal, test_norm_lt_alfanumerico, test_norm_lt_vacio,
        test_estado_cancelado, test_estado_exceso, test_estado_parcial, test_estado_pendiente,
        test_norm_fecha_none, test_norm_fecha_datetime_obj,
        test_norm_fecha_str_dmy, test_norm_fecha_str_ymd,
        test_norm_fecha_str_microsegundos, test_norm_fecha_str_desconocido,
        test_ajuste_sumar_pago, test_ajuste_restar_pago, test_ajuste_cancelado,
        test_ajuste_exceso, test_ajuste_no_puede_quedar_negativo,
        test_ajuste_con_otro_pago, test_ajuste_parcial, test_ajuste_restar_con_otro,
    ]
    passed = 0
    for t in tests:
        try:
            t()
            print(f"  OK  {t.__name__}")
            passed += 1
        except AssertionError as e:
            print(f"  FAIL  {t.__name__} -- {e}")
        except Exception as e:
            print(f"  ERROR  {t.__name__} -- {type(e).__name__}: {e}")
    print(f"\n{passed}/{len(tests)} tests pasaron")
