import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from main import _float, _norm_lt, _key, _estado_penalidad, _clasificar, PENALIDAD, RECONEXION, TOLERANCIA


# ── helpers ───────────────────────────────────────────────────────────────────
def _u(mz="A", lt="1", saldo=30.0, penalidad=20.0):
    return {
        "mz": mz, "lt": lt, "key": f"{mz}-{lt}", "nombre": "TEST",
        "saldo": saldo, "penalidad": penalidad,
        "total_a_pagar": round(saldo + penalidad, 2),
    }

def _cob(key, yape=0.0, efectivo=0.0, corte_reconexion=0.0,
         saldo=0.0, estado="PENDIENTE"):
    return {key: {"yape": yape, "efectivo": efectivo,
                  "corte_reconexion": corte_reconexion,
                  "saldo": saldo, "estado": estado}}


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
def test_norm_lt():
    assert _norm_lt("1.0") == "1"
    assert _norm_lt("11A") == "11A"
    assert _norm_lt("")    == ""
    assert _norm_lt(None)  == ""


# ── _estado_penalidad ─────────────────────────────────────────────────────────
def test_estado_cancelado():
    assert _estado_penalidad(0.0)        == "CANCELADO"
    assert _estado_penalidad(0.001)      == "CANCELADO"   # dentro de tolerancia
    assert _estado_penalidad(-0.001)     == "CANCELADO"

def test_estado_exceso():
    assert _estado_penalidad(-10.0)      == "EXCESO"
    assert _estado_penalidad(-0.5)       == "EXCESO"

def test_estado_parcial():
    assert _estado_penalidad(10.0)       == "PARCIAL"
    assert _estado_penalidad(0.1)        == "PARCIAL"


# ── _clasificar ───────────────────────────────────────────────────────────────
def test_clasificar_no_pago():
    # No pagó nada → corte físico
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {}, {})
    assert len(p) == 0
    assert len(c) == 1
    assert c[0]["total_reconexion"] == 70.0   # 30 + 40
    assert c[0]["pago_nuevo"]       == 0.0

def test_clasificar_pago_parcial_sin_penalidad():
    # Pagó S/15 — no alcanza los S/20 de penalidad → corte físico
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {"A-1": 15.0}, {})
    assert len(c) == 1
    assert c[0]["pago_nuevo"] == 15.0

def test_clasificar_pago_exacto_penalidad():
    # Pagó exactamente S/20 → salvado · todavía debe el saldo → PARCIAL
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {"A-1": 20.0}, {})
    assert len(p) == 1
    assert len(c) == 0
    assert p[0]["estado"]      == "PARCIAL"
    assert p[0]["saldo_final"] == 30.0
    assert p[0]["pago_nuevo"]  == 20.0

def test_clasificar_pago_total():
    # Pagó saldo + penalidad completos → CANCELADO
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {"A-1": 50.0}, {})   # 30 + 20
    assert p[0]["estado"]      == "CANCELADO"
    assert abs(p[0]["saldo_final"]) <= TOLERANCIA

def test_clasificar_pago_exceso():
    # Pagó más del total → EXCESO · saldo_final negativo
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {"A-1": 65.0}, {})   # 30 + 20 + 15 extra
    assert p[0]["estado"]      == "EXCESO"
    assert p[0]["saldo_final"] == -15.0

def test_clasificar_delta_correcto():
    # Ya había pagado S/10 en cobranza original · ahora total = S/25 → delta = S/15 → no alcanza
    u = _u(saldo=30.0)
    cob = _cob("A-1", yape=10.0)
    p, c = _clasificar([u], cob, {"A-1": 25.0}, {})   # delta = 15
    assert len(c) == 1
    assert c[0]["pago_nuevo"] == 15.0

def test_clasificar_delta_suficiente():
    # Ya había pagado S/5 · ahora total = S/25 → delta = S/20 = exactamente PENALIDAD → PARCIAL
    u = _u(saldo=30.0)
    cob = _cob("A-1", yape=5.0)
    p, c = _clasificar([u], cob, {"A-1": 25.0}, {})   # delta = 20
    assert len(p) == 1
    assert p[0]["pago_nuevo"] == 20.0
    assert p[0]["estado"]     == "PARCIAL"

def test_clasificar_pago_efectivo_suma():
    # Pago nuevo vía efectivo (S/20) → salvado
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {}, {"A-1": 20.0})
    assert len(p) == 1
    assert p[0]["pago_nuevo"] == 20.0

def test_clasificar_pago_mixto():
    # S/10 Yape + S/15 efectivo = S/25 delta → salvado · saldo_final = 25
    u = _u(saldo=30.0)
    cob = _cob("A-1")
    p, c = _clasificar([u], cob, {"A-1": 10.0}, {"A-1": 15.0})
    assert len(p) == 1
    assert p[0]["pago_nuevo"]  == 25.0
    assert p[0]["saldo_final"] == 25.0   # 30 - (25 - 20)

def test_clasificar_multiple_usuarios():
    # Varios usuarios: uno paga, otro no
    usuarios = [_u("A", "1", saldo=30.0), _u("B", "2", saldo=50.0)]
    cob = {**_cob("A-1"), **_cob("B-2")}
    p, c = _clasificar(usuarios, cob, {"A-1": 20.0}, {})
    assert len(p) == 1
    assert len(c) == 1
    assert p[0]["mz"] == "A"
    assert c[0]["mz"] == "B"

def test_clasificar_sin_penalidad_en_orig():
    # Pago delta no puede ser negativo (si alguien "devolvió" dinero)
    u = _u(saldo=30.0)
    cob = _cob("A-1", yape=50.0)   # original ya tenía más
    p, c = _clasificar([u], cob, {"A-1": 30.0}, {})   # "nuevo total" menor → delta = 0
    assert len(c) == 1
    assert c[0]["pago_nuevo"] == 0.0

def test_clasificar_total_reconexion():
    u = _u(saldo=45.0)
    cob = _cob("A-1")
    _, c = _clasificar([u], cob, {}, {})
    assert c[0]["reconexion"]       == RECONEXION          # S/40
    assert c[0]["total_reconexion"] == 85.0                # 45 + 40


if __name__ == "__main__":
    tests = [
        test_float_normal, test_float_coma, test_float_invalido,
        test_norm_lt,
        test_estado_cancelado, test_estado_exceso, test_estado_parcial,
        test_clasificar_no_pago, test_clasificar_pago_parcial_sin_penalidad,
        test_clasificar_pago_exacto_penalidad, test_clasificar_pago_total,
        test_clasificar_pago_exceso, test_clasificar_delta_correcto,
        test_clasificar_delta_suficiente, test_clasificar_pago_efectivo_suma,
        test_clasificar_pago_mixto, test_clasificar_multiple_usuarios,
        test_clasificar_sin_penalidad_en_orig, test_clasificar_total_reconexion,
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
