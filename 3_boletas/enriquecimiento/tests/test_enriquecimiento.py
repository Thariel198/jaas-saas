import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from main import _float, _norm_lt, _str, _estado_boleta, _fecha_pago_str, _enriquecer

COSTO_M3 = 1.0


# ── _float ────────────────────────────────────────────────────────────────────
def test_float_normal():
    assert _float("45.5") == 45.5
    assert _float(100)    == 100.0

def test_float_coma():
    assert _float("45,5") == 45.5

def test_float_invalido():
    assert _float("")    == 0.0
    assert _float(None)  == 0.0


# ── _norm_lt ──────────────────────────────────────────────────────────────────
def test_norm_lt_decimal():
    assert _norm_lt("3.0") == "3"

def test_norm_lt_vacio():
    assert _norm_lt("")   == ""
    assert _norm_lt(None) == ""


# ── _estado_boleta ────────────────────────────────────────────────────────────
def test_estado_al_dia():
    assert _estado_boleta(0, 0, 0, 0, 0) == "USTED ESTÁ AL DÍA"

def test_estado_no_al_dia_por_arrastre():
    assert _estado_boleta(10, 0, 0, 0, 0) == "NO ESTÁ AL DÍA"

def test_estado_no_al_dia_por_corte():
    assert _estado_boleta(0, 20, 0, 0, 0) == "NO ESTÁ AL DÍA"

def test_estado_no_al_dia_por_convenio():
    assert _estado_boleta(0, 0, 15, 0, 0) == "NO ESTÁ AL DÍA"

def test_estado_solo_consumo_es_al_dia():
    # consumo del mes no es deuda acumulada → al día
    assert _estado_boleta(0, 0, 0, 0, 0) == "USTED ESTÁ AL DÍA"


# ── _fecha_pago_str ───────────────────────────────────────────────────────────
def test_fecha_pago_str():
    cfg = {"FECHA_PAGO": "02/05", "HORA_PAGO": "4-6 pm",
           "LUGAR_PAGO": "LOCAL DEL PUEBLO", "TELEFONO": "948 227 636"}
    s = _fecha_pago_str(cfg)
    assert "02/05" in s
    assert "4-6 pm" in s
    assert "LOCAL DEL PUEBLO" in s
    assert "948 227 636" in s


# ── _enriquecer ───────────────────────────────────────────────────────────────
def _usuario(mz="A", lt="1", marc_ant=0, marc_act=0,
             arrastre=0, convenio=0, mant=3, corte_reconex=0,
             reunion_faena=0, techado=0, devolucion=0, ajuste=0):
    m3 = marc_act - marc_ant
    return {
        "mz": mz, "lt": lt, "nombre": "TEST",
        "marc_ant": marc_ant, "marc_act": marc_act,
        "m3": m3, "total_mes": round(m3 * COSTO_M3, 2),
        "arrastre": arrastre, "convenio": convenio, "mant": mant,
        "corte_reconex": corte_reconex, "reunion_faena": reunion_faena,
        "techado": techado, "devolucion": devolucion, "ajuste": ajuste,
    }

def _cfg():
    return {
        "PERIODO": "11/03/2026 al 10/04/2026",
        "FECHA_VENCIMIENTO": "2026-05-02",
        "FECHA_EMISION": "2026-04-26",
        "LECTURA_ANT_FECHA": "2026-03-10",
        "LECTURA_ACT_FECHA": "2026-04-10",
        "FECHA_PAGO": "02/05",
        "HORA_PAGO": "4-6 pm",
        "LUGAR_PAGO": "LOCAL DEL PUEBLO",
        "TELEFONO": "948 227 636",
        "NUMERO_RECIBO_INICIO": 16900,
    }

def test_enriquecer_numero_recibo():
    usuarios = [_usuario("A", "1"), _usuario("A", "2"), _usuario("B", "3")]
    recs = _enriquecer(usuarios, _cfg())
    assert recs[0]["NUMERO DE RECIBO"] == 16900
    assert recs[1]["NUMERO DE RECIBO"] == 16901
    assert recs[2]["NUMERO DE RECIBO"] == 16902

def test_enriquecer_total_correcto():
    # marc_ant=5, marc_act=13 → m3=8, total_mes=8, mant=3 → total=11
    u = _usuario(marc_ant=5, marc_act=13, mant=3)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["Total"] == 11.0
    assert recs[0]["Importe a pagar"] == 11.0

def test_enriquecer_total_con_arrastre():
    # total_mes=8, mant=3, arrastre=20 → total=31
    u = _usuario(marc_ant=5, marc_act=13, mant=3, arrastre=20)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["Total"] == 31.0

def test_enriquecer_total_con_devolucion():
    # total_mes=8, mant=3, devolucion=5 → total=6
    u = _usuario(marc_ant=5, marc_act=13, mant=3, devolucion=5)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["Total"] == 6.0

def test_enriquecer_estado_al_dia():
    u = _usuario(marc_ant=5, marc_act=13, mant=3)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["Estado"] == "USTED ESTÁ AL DÍA"

def test_enriquecer_estado_no_al_dia():
    u = _usuario(marc_ant=5, marc_act=13, mant=3, arrastre=15)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["Estado"] == "NO ESTÁ AL DÍA"

def test_enriquecer_campos_config():
    u = _usuario(marc_ant=5, marc_act=13, mant=3)
    recs = _enriquecer([u], _cfg())
    r = recs[0]
    assert r["PERIODO"] == "11/03/2026 al 10/04/2026"
    assert r["LECTURA ANTERIOR"] == "2026-03-10"
    assert r["LECTURA ACTUAL"]   == "2026-04-10"
    assert r["FECHA_PAGO"] == "02/05"
    assert r["HORA_PAGO"]  == "4-6 pm"

def test_enriquecer_marcacion():
    u = _usuario(marc_ant=150, marc_act=163, mant=3)
    recs = _enriquecer([u], _cfg())
    r = recs[0]
    assert r["Marcación anterior"] == 150
    assert r["Marcacion altual"]   == 163
    assert r["M3"]                 == 13
    assert r["Total mes actual"]   == 13.0

def test_enriquecer_mes_anterior_es_arrastre():
    u = _usuario(marc_ant=5, marc_act=13, mant=3, arrastre=45)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["MES ANTERIOR"] == 45.0

def test_enriquecer_total_nunca_negativo():
    # devolucion mayor que todo → total = 0
    u = _usuario(marc_ant=5, marc_act=13, mant=3, devolucion=100)
    recs = _enriquecer([u], _cfg())
    assert recs[0]["Total"] == 0.0


if __name__ == "__main__":
    tests = [
        test_float_normal, test_float_coma, test_float_invalido,
        test_norm_lt_decimal, test_norm_lt_vacio,
        test_estado_al_dia, test_estado_no_al_dia_por_arrastre,
        test_estado_no_al_dia_por_corte, test_estado_no_al_dia_por_convenio,
        test_estado_solo_consumo_es_al_dia,
        test_fecha_pago_str,
        test_enriquecer_numero_recibo, test_enriquecer_total_correcto,
        test_enriquecer_total_con_arrastre, test_enriquecer_total_con_devolucion,
        test_enriquecer_estado_al_dia, test_enriquecer_estado_no_al_dia,
        test_enriquecer_campos_config, test_enriquecer_marcacion,
        test_enriquecer_mes_anterior_es_arrastre,
        test_enriquecer_total_nunca_negativo,
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
