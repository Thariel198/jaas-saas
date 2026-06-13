import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from main import _float, _norm_mz, _norm_tipo, _diferencias_por_mz, TOLERANCIA


# ── _float ───────────────────────────────────────────────
def test_float_normal():
    assert _float("45.5")  == 45.5
    assert _float("100")   == 100.0
    assert _float(33)      == 33.0

def test_float_coma():
    assert _float("45,5")  == 45.5

def test_float_invalido():
    assert _float("")      == 0.0
    assert _float(None)    == 0.0
    assert _float("abc")   == 0.0


# ── _norm_mz ──────────────────────────────────────────────
def test_norm_mz():
    assert _norm_mz("a")    == "A"
    assert _norm_mz(" B1 ") == "B1"
    assert _norm_mz("G1")   == "G1"


# ── _norm_tipo ────────────────────────────────────────────
def test_norm_tipo_tepago():
    assert _norm_tipo("TE PAGÓ")   == "TE PAGO"
    assert _norm_tipo("TE PAGO")   == "TE PAGO"
    assert _norm_tipo("te pagó")   == "TE PAGO"

def test_norm_tipo_pagaste():
    assert _norm_tipo("PAGASTE")   == "PAGASTE"
    assert _norm_tipo("pagaste")   == "PAGASTE"

def test_norm_tipo_diferencia():
    assert _norm_tipo("TE PAGÓ") != _norm_tipo("PAGASTE")


# ── _diferencias_por_mz ───────────────────────────────────
def test_dif_todo_ok():
    rep = {"A": 100.0, "B": 200.0}
    pla = {"A": 100.0, "B": 200.0}
    res = _diferencias_por_mz(rep, pla)
    assert len(res) == 2
    assert all(r["ok"] for r in res)
    assert all(r["diferencia"] == 0.0 for r in res)

def test_dif_alerta():
    rep = {"A": 100.0, "B": 200.0}
    pla = {"A": 104.0, "B": 200.0}
    res = {r["mz"]: r for r in _diferencias_por_mz(rep, pla)}
    assert not res["A"]["ok"]
    assert res["A"]["diferencia"] == 4.0      # planilla - reporte
    assert res["B"]["ok"]

def test_dif_negativa():
    rep = {"C": 410.0}
    pla = {"C": 385.0}
    res = _diferencias_por_mz(rep, pla)
    assert res[0]["diferencia"] == -25.0
    assert not res[0]["ok"]

def test_dif_mz_solo_en_reporte():
    # MZ existe en reporte pero no en planilla → planilla asume 0
    rep = {"Z": 50.0}
    pla = {}
    res = _diferencias_por_mz(rep, pla)
    assert res[0]["planilla"]   == 0.0
    assert res[0]["diferencia"] == -50.0
    assert not res[0]["ok"]

def test_dif_mz_solo_en_planilla():
    # MZ existe en planilla pero no en reporte → reporte asume 0
    rep = {}
    pla = {"X": 30.0}
    res = _diferencias_por_mz(rep, pla)
    assert res[0]["reporte"]    == 0.0
    assert res[0]["diferencia"] == 30.0
    assert not res[0]["ok"]

def test_dif_tolerancia():
    # Diferencia menor a TOLERANCIA → OK
    rep = {"A": 100.0}
    pla = {"A": 100.004}
    res = _diferencias_por_mz(rep, pla)
    assert res[0]["ok"]

def test_dif_ordenado_por_mz():
    rep = {"C": 10.0, "A": 20.0, "B": 30.0}
    pla = {"C": 10.0, "A": 20.0, "B": 30.0}
    mzs = [r["mz"] for r in _diferencias_por_mz(rep, pla)]
    assert mzs == sorted(mzs)

def test_dif_vacio():
    assert _diferencias_por_mz({}, {}) == []


if __name__ == "__main__":
    tests = [
        test_float_normal, test_float_coma, test_float_invalido,
        test_norm_mz,
        test_norm_tipo_tepago, test_norm_tipo_pagaste, test_norm_tipo_diferencia,
        test_dif_todo_ok, test_dif_alerta, test_dif_negativa,
        test_dif_mz_solo_en_reporte, test_dif_mz_solo_en_planilla,
        test_dif_tolerancia, test_dif_ordenado_por_mz, test_dif_vacio,
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
