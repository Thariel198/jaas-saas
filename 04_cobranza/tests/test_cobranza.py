import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from main import _key, _norm_lt, _float, _estado, _calcular, _priorizar_pago


# ── _norm_lt ──────────────────────────────────────────────
def test_norm_lt_entero():
    assert _norm_lt("1")    == "1"
    assert _norm_lt("1.0")  == "1"
    assert _norm_lt("11")   == "11"

def test_norm_lt_alfanumerico():
    assert _norm_lt("11A")  == "11A"
    assert _norm_lt("A1")   == "A1"

def test_norm_lt_vacio():
    assert _norm_lt("")     == ""
    assert _norm_lt("nan")  == ""
    assert _norm_lt(None)   == ""


# ── _key ─────────────────────────────────────────────────
def test_key_simple():
    assert _key("A", "1")   == "A-1"
    assert _key("G1", "7")  == "G1-7"

def test_key_normaliza_mz():
    assert _key("a", "3")   == "A-3"
    assert _key(" B ", "2") == "B-2"

def test_key_normaliza_lt():
    assert _key("A", "1.0") == "A-1"
    assert _key("K", "4")   == "K-4"

def test_key_vacio():
    assert _key("", "1")    == ""
    assert _key("A", "")    == ""


# ── _estado ──────────────────────────────────────────────
def test_estado_cancelado():
    assert _estado(0.0,   50.0) == "CANCELADO"
    assert _estado(0.001, 50.0) == "CANCELADO"

def test_estado_exceso():
    assert _estado(-10.0, 60.0) == "EXCESO"
    assert _estado(-0.5,  50.5) == "EXCESO"

def test_estado_parcial():
    assert _estado(20.0, 30.0) == "PARCIAL"

def test_estado_pendiente():
    assert _estado(50.0,  0.0) == "PENDIENTE"


# ── _float ───────────────────────────────────────────────
def test_float_normal():
    assert _float("45.5") == 45.5
    assert _float("100")  == 100.0
    assert _float(33)     == 33.0

def test_float_coma():
    assert _float("45,5") == 45.5

def test_float_invalido():
    assert _float("")     == 0.0
    assert _float(None)   == 0.0
    assert _float("abc")  == 0.0


# ── _calcular ────────────────────────────────────────────
def _u(mz="A", lt="1", marc_ant=0, marc_act=0, arrastre=0,
       convenio=0, mant=3, corte_reconexion=0,
       reunion_faena=0, techado=0, devolucion=0, ajuste=0):
    m3 = marc_act - marc_ant
    return {
        "mz": mz, "lt": lt, "key": f"{mz}-{lt}", "nombre": "TEST",
        "marc_ant": marc_ant, "marc_act": marc_act, "m3": m3,
        "total_mes": round(m3 * 1.0, 2),  # COSTO_M3 = 1.0
        "arrastre": arrastre, "convenio": convenio, "mant": mant,
        "corte_reconexion": corte_reconexion,
        "reunion_faena": reunion_faena, "techado": techado,
        "devolucion": devolucion, "ajuste": ajuste,
    }

def test_calcular_cancelado_yape():
    # marc_ant=2, marc_act=11 → m3=9, total_mes=9, mant=3 → total_deuda=12
    res, _ = _calcular([_u(marc_ant=2, marc_act=11)], {"A-1": 12.0}, {}, {})
    r = res[0]
    assert r["total_deuda"] == 12.0
    assert r["yape"]        == 12.0
    assert r["saldo"]       == 0.0
    assert r["estado"]      == "CANCELADO"

def test_calcular_deuda_total_correcta():
    # marc_ant=2, marc_act=11 → m3=9, total_mes=9, mant=3 → total=12
    res, _ = _calcular([_u(marc_ant=2, marc_act=11, mant=3)], {}, {}, {})
    assert res[0]["total_deuda"] == 12.0

def test_calcular_blanco_reduce_deuda():
    # total_mes=17, mant=3, blanco=8 → total=12
    res, usados = _calcular([_u("E", "7", marc_ant=540, marc_act=557, mant=3)],
                            {"E-7": 12.0}, {}, {"E-7": 8.0})
    assert res[0]["blanco"]      == 8.0
    assert res[0]["total_deuda"] == 12.0   # 17 + 3 - 8
    assert res[0]["saldo"]       == 0.0
    assert res[0]["estado"]      == "CANCELADO"
    assert "E-7" in usados

def test_calcular_blanco_no_doble_aplica():
    # Si el key no está en blancos, blanco=0
    res, usados = _calcular([_u("A", "99")], {}, {}, {})
    assert res[0]["blanco"] == 0.0
    assert "A-99" not in usados

def test_calcular_exceso():
    res, _ = _calcular([_u(marc_ant=180, marc_act=192, mant=3)],
                       {"A-1": 20.0}, {}, {})
    # total=12+3=15, pagado=20 → saldo=-5
    assert res[0]["saldo"]  == -5.0
    assert res[0]["estado"] == "EXCESO"

def test_calcular_pendiente():
    res, _ = _calcular([_u(mant=3, arrastre=50)], {}, {}, {})
    assert res[0]["total_deuda"] == 53.0   # 0 consumo + 50 arrastre + 3 mant
    assert res[0]["estado"]      == "PENDIENTE"

def test_calcular_parcial():
    res, _ = _calcular([_u(mant=3, arrastre=30)], {"A-1": 10.0}, {}, {})
    assert res[0]["estado"] == "PARCIAL"
    assert res[0]["saldo"]  == 23.0

def test_calcular_corte_arrastre_exacto_8():
    # arrastre=8 → debe ir a corte (es el mínimo que confirma que no pagó nada)
    res, _ = _calcular([_u(arrastre=8, mant=0)], {}, {}, {})
    r = res[0]
    assert r["arrastre"] >= 8.0 - 0.005
    assert r["saldo"]    >  0.005

def test_calcular_arrastre_4_no_corte():
    # arrastre=4 → puede ser pago parcial del mes anterior, NO va a corte
    res, _ = _calcular([_u(arrastre=4, mant=0)], {}, {}, {})
    r = res[0]
    assert r["arrastre"] < 8.0  # no cumple condición corte


# ── _priorizar_pago ──────────────────────────────────────────────────────────
def _calc1(**kw):
    """Helper: crea un usuario, corre _calcular y devuelve el primer resultado."""
    u = _u(**kw)
    res, _ = _calcular([u], {u["key"]: kw.pop("pagado_yape", 0.0)},
                       {u["key"]: kw.pop("pagado_efec", 0.0)}, {})
    return res[0]

def _prio(**kw):
    pagado_yape = kw.pop("pagado_yape", 0.0)
    pagado_efec = kw.pop("pagado_efec", 0.0)
    u = _u(**kw)
    res, _ = _calcular([u], {u["key"]: pagado_yape}, {u["key"]: pagado_efec}, {})
    return res[0]


def test_prio_cancelado_todo_cero():
    r = _prio(marc_ant=5, marc_act=10, mant=3, pagado_yape=8.0)
    assert r["arrastre_nvo"] == 0.0
    assert r["devolucion_nvo"] == 0.0

def test_prio_pendiente_arrastre_incluye_mant():
    # total_mes=5, mant=3 — nadie pagó → arrastre_nvo = 5+0+3 = 8
    r = _prio(marc_ant=5, marc_act=10, mant=3)
    assert r["arrastre_nvo"] == 8.0
    assert r["corte_reconexion_nvo"] == 0.0

def test_prio_convenio_queda():
    # total=8+20=28, pagado=8 → convenio=20 queda (menor prioridad)
    r = _prio(marc_ant=5, marc_act=10, mant=3, convenio=20, pagado_yape=8.0)
    assert r["arrastre_nvo"] == 0.0
    assert r["convenio_nvo"] == 20.0

def test_prio_ejemplo_usuario():
    # arrastre=12, total_mes=5, mant=3, corte=20, convenio=20 → total=60, pagado=40, saldo=20
    # saldo 20 cae en convenio (menor prioridad que corte)
    r = _prio(marc_ant=5, marc_act=10, arrastre=12, mant=3,
              corte_reconexion=20, convenio=20, pagado_yape=40.0)
    assert r["arrastre_nvo"] == 0.0
    assert r["corte_reconexion_nvo"] == 0.0
    assert r["convenio_nvo"] == 20.0

def test_prio_saldo_mayor_que_convenio():
    # total_mes=5, mant=3, corte=20, convenio=15 → total=43, pagado=28, saldo=15
    # saldo 15 = convenio(15) completamente → corte queda pagado
    r = _prio(marc_ant=5, marc_act=10, mant=3, corte_reconexion=20, convenio=15, pagado_yape=28.0)
    assert r["corte_reconexion_nvo"] == 0.0
    assert r["convenio_nvo"] == 15.0

def test_prio_saldo_desborda_convenio():
    # total_mes=5, mant=3, corte=20, convenio=15 → total=43, pagado=18, saldo=25
    # saldo 25: convenio(15) + corte(10)
    r = _prio(marc_ant=5, marc_act=10, mant=3, corte_reconexion=20, convenio=15, pagado_yape=18.0)
    assert r["convenio_nvo"] == 15.0
    assert r["corte_reconexion_nvo"] == 10.0
    assert r["arrastre_nvo"] == 0.0

def test_prio_exceso_devolucion():
    # total=8, pagado=15 → exceso=7 → devolucion_nvo=7
    r = _prio(marc_ant=5, marc_act=10, mant=3, pagado_yape=15.0)
    assert r["devolucion_nvo"] == 7.0
    assert r["arrastre_nvo"] == 0.0

def test_prio_techado_y_reunion_al_final():
    # mant=3, reunion=10, techado=8 → total=21, pagado=0
    # saldo 21: techado(8) + reunion(10) + mant→arrastre(3)
    r = _prio(mant=3, reunion_faena=10, techado=8)
    assert r["techado_nvo"] == 8.0
    assert r["reunion_faena_nvo"] == 10.0
    assert r["arrastre_nvo"] == 3.0

def test_prio_arrastre_preexistente():
    # arrastre=20 (del mes anterior), mant=3, pagado=10 → saldo=13
    # saldo 13: convenio/corte/etc = 0, mant→arrastre(3) + arrastre(10) = 13
    r = _prio(arrastre=20, mant=3, pagado_yape=10.0)
    assert r["arrastre_nvo"] == 13.0  # 10 arrastre restante + 3 mant restante


if __name__ == "__main__":
    tests = [
        test_norm_lt_entero, test_norm_lt_alfanumerico, test_norm_lt_vacio,
        test_key_simple, test_key_normaliza_mz, test_key_normaliza_lt, test_key_vacio,
        test_estado_cancelado, test_estado_exceso, test_estado_parcial, test_estado_pendiente,
        test_float_normal, test_float_coma, test_float_invalido,
        test_calcular_cancelado_yape, test_calcular_deuda_total_correcta,
        test_calcular_blanco_reduce_deuda, test_calcular_blanco_no_doble_aplica,
        test_calcular_exceso, test_calcular_pendiente, test_calcular_parcial,
        test_calcular_corte_arrastre_exacto_8, test_calcular_arrastre_4_no_corte,
        test_prio_cancelado_todo_cero, test_prio_pendiente_arrastre_incluye_mant,
        test_prio_convenio_queda, test_prio_ejemplo_usuario,
        test_prio_saldo_mayor_que_convenio, test_prio_saldo_desborda_convenio,
        test_prio_exceso_devolucion, test_prio_techado_y_reunion_al_final,
        test_prio_arrastre_preexistente,
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
