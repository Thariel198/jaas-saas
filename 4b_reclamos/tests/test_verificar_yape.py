"""
tests/test_verificar_yape.py — unitarios de la verificacion del yape contra el banco.

REGLA DE ESTE ARCHIVO: no toca datos reales.
    · no escribe NINGUN archivo (ni en inputs/, ni en outputs/, ni en backup/)
    · no lee ningun .xlsx del repo: todo entra por monkeypatch con DataFrames
      construidos aca
El 12/08/2026 correr `pytest 4_pagos/efectivo/tests` sobrescribio mesa_1 y
mesa_2 del ciclo en curso con sus fixtures (59 y 106 filas de cobro reales, que
se recuperaron con git). Esa suite escribe sobre el inputs/ real. Esta no.

Los bloques marcados CONTRAFACTUAL reproducen un falso positivo real que la
herramienta emitio sobre datos de junio/julio/agosto 2026 y verifican que la
regla lo rechaza.

Uso: py -m pytest tests/test_verificar_yape.py -q
"""

import sys
from pathlib import Path

import pandas as pd

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent / "herramienta"))

import verificar_yape as vy  # noqa: E402


# ── Fixtures sinteticas ──────────────────────────────────────────────────────

def banco(filas):
    """DataFrame con la forma de _reporte_banco(). filas = [(origen, monto, mensaje, 'dd/mm/YYYY HH:MM')]"""
    df = pd.DataFrame([{"TIPO": "TE PAGÓ", "ORIGEN": o, "MONTO": m, "MENSAJE": g, "FECHA": f}
                       for o, m, g, f in filas])
    df["FECHA"] = pd.to_datetime(df["FECHA"], format="%d/%m/%Y %H:%M")
    return df


def sin_ventana(monkeypatch):
    """Desactiva el acote por ciclo para probar el resto de la logica aislada."""
    monkeypatch.setattr(vy, "ventana_del_ciclo", lambda m: (None, None))


# ── La verificacion basica ───────────────────────────────────────────────────

def test_encontrado_cuando_el_mensaje_nombra_el_lote(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Juan B*", 13.0, "mz F1 lt 8", "02/08/2026 10:00")]))
    r = vy.verificar_una(13.0, "02/08/2026", "F1", "8")
    assert r["estado"] == "ENCONTRADO"


def test_no_existe_cuando_no_hay_nada_que_calce(monkeypatch):
    # Caso real F1-8: Wagner anoto S/13 el 02/08 y no hay ninguna transaccion de
    # S/13 en la cuenta de la JASS en esos dias.
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Otro", 8.0, "", "02/08/2026 10:00")]))
    assert vy.verificar_una(13.0, "02/08/2026", "F1", "8")["estado"] == "NO_EXISTE"


def test_posible_cuando_solo_coincide_monto_y_fecha(monkeypatch):
    # Coincide el monto pero nada confirma que sea de este predio: los montos se
    # repiten muchisimo (36 transacciones de S/8 en el reporte real).
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Otro", 8.0, "", "02/08/2026 10:00")]))
    r = vy.verificar_una(8.0, "02/08/2026", "V", "8")
    assert r["estado"] == "POSIBLE"
    assert "nada confirma" in r["detalle"]


def test_sin_reporte_no_concluye_nada(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco", lambda: pd.DataFrame())
    assert vy.verificar_una(13.0, "02/08/2026", "F1", "8")["estado"] == "SIN_REPORTE"


def test_sin_fecha_no_concluye_nada(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Otro", 8.0, "", "02/08/2026 10:00")]))
    assert vy.verificar_una(13.0, "", "F1", "8")["estado"] == "SIN_FECHA"


def test_fuera_de_rango_no_es_lo_mismo_que_no_existe(monkeypatch):
    # No es lo mismo "no esta" que "no lo puedo saber": si el reporte no cubre
    # esa fecha, afirmar que el yape no entro seria inventar.
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Otro", 8.0, "", "02/08/2026 10:00")]))
    assert vy.verificar_una(13.0, "15/01/2026", "F1", "8")["estado"] == "FUERA_DE_RANGO"


def test_el_rango_se_mide_contra_la_ventana_no_contra_la_fecha(monkeypatch):
    # Un pago del 02/08 con el reporte terminando el 01/08 igual entra en la
    # ventana de +-3 dias y hay que buscarlo, no descartarlo.
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Juan B*", 13.0, "", "01/08/2026 10:00")]))
    assert vy.verificar_una(13.0, "02/08/2026", "F1", "8")["estado"] == "POSIBLE"


# ── CONTRAFACTUAL 1 — el mensaje nombra OTRO lote ───────────────────────────
# S/36 "Roman Lozano Mz H lote 21" se reportaba como el yape de H1-15, y S/30
# "mz D1 lt 1" como el de P-12. Coincidia solo el monto.

def test_descarta_el_pago_cuyo_mensaje_nombra_otro_lote(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("X", 36.0, "Roman Lozano Minaya Mz H lote 21",
                                        "05/07/2026 10:00")]))
    r = vy.verificar_una(36.0, "05/07/2026", "H1", "15")
    assert r["estado"] == "NO_EXISTE" and "nombra otro" in r["detalle"]


def test_no_descarta_cuando_el_mensaje_nombra_el_lote_propio(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("X", 36.0, "Mz H1 lt 15", "05/07/2026 10:00")]))
    assert vy.verificar_una(36.0, "05/07/2026", "H1", "15")["estado"] == "ENCONTRADO"


def test_lotes_en_mensaje_junta_las_dos_fuentes():
    # motor_matching aporta 31 patrones afinados ("M x L 11" -> X-11); los
    # propios aportan la forma con comas y el barrido de TODOS los lotes.
    assert ("X", "11") in vy._lotes_en_mensaje("Maria Rosa Jimenez Roca   M x L 11")
    assert ("B1", "3") in vy._lotes_en_mensaje("MZ,B1,Lt.3-Johan Rodriguez Flores")
    assert vy._lotes_en_mensaje("K-3, K-4.") >= {("K", "3"), ("K", "4")}
    assert vy._lotes_en_mensaje("pago de agua") == set()


def test_multilote_no_se_toma_como_lote_ajeno():
    # "K-3, K-4." nombra los dos: para K-4 NO es un lote ajeno.
    assert vy._nombra_otro_lote("K-3, K-4.", "K", "4") is False


# ── CONTRAFACTUAL 2 — el mensaje nombra a OTRA PERSONA ──────────────────────
# S/36 "pago de servicios de agua usuario Alejandro Melgarejo" matcheaba como el
# yape de H1-15, que es de Patricia Tarazona.

def test_descarta_el_pago_que_nombra_a_otro_usuario(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Jose Mel*", 36.0,
                                        "pago de servicios de agua usuario Alejandro Melgarejo",
                                        "05/07/2026 10:00")]))
    r = vy.verificar_una(36.0, "05/07/2026", "H1", "15", nombre="PATRICIA TARAZONA CARBAJAL")
    assert r["estado"] == "NO_EXISTE"


def test_no_descarta_si_el_mensaje_nombra_al_titular(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco",
                        lambda: banco([("Patricia T*", 36.0, "usuario Patricia Tarazona",
                                        "05/07/2026 10:00")]))
    r = vy.verificar_una(36.0, "05/07/2026", "H1", "15", nombre="PATRICIA TARAZONA CARBAJAL")
    assert r["estado"] == "POSIBLE"


# ── CONTRAFACTUAL 3 — la ventana del ciclo (ancla de corte) ─────────────────
# El reporte del banco abarca ~3 meses. Sin acotar, un yape de JUNIO que nombre
# el lote se devolvia como prueba de un pago de JULIO. Caso real: S/94 del 11/06
# con mensaje "c1-7".

_BANCO_2_CICLOS = [("Vecino", 94.0, "c1-7", "11/06/2026 18:37"),      # ciclo junio
                   ("Otro", 8.0, "", "05/07/2026 09:00")]             # ciclo julio


def test_sin_ventana_un_yape_de_junio_se_acredita_como_de_julio(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco", lambda: banco(_BANCO_2_CICLOS))
    sin_ventana(monkeypatch)
    assert vy.verificar_una(94.0, "05/07/2026", "C1", "7")["estado"] == "ENCONTRADO"


def test_con_ventana_el_yape_de_junio_queda_afuera(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco", lambda: banco(_BANCO_2_CICLOS))
    monkeypatch.setattr(vy, "ventana_del_ciclo",
                        lambda m: (pd.Timestamp("2026-06-17 20:32"),
                                   pd.Timestamp("2026-07-20 22:48")))
    r = vy.verificar_una(94.0, "05/07/2026", "C1", "7", mes="2026-07")
    assert r["estado"] == "NO_EXISTE"
    assert "ventana del ciclo" in r["detalle"]


def test_la_ventana_se_declara_en_el_veredicto(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco", lambda: banco(_BANCO_2_CICLOS))
    monkeypatch.setattr(vy, "ventana_del_ciclo",
                        lambda m: (pd.Timestamp("2026-06-17 20:32"),
                                   pd.Timestamp("2026-07-20 22:48")))
    r = vy.verificar_una(8.0, "05/07/2026", "V", "8", mes="2026-07")
    assert "ventana del ciclo 2026-07" in r["detalle"]


def test_ciclo_sin_transacciones_en_su_ventana_no_concluye(monkeypatch):
    monkeypatch.setattr(vy, "_reporte_banco", lambda: banco(_BANCO_2_CICLOS))
    monkeypatch.setattr(vy, "ventana_del_ciclo",
                        lambda m: (pd.Timestamp("2026-09-01"), pd.Timestamp("2026-09-30")))
    assert vy.verificar_una(8.0, "05/09/2026", "V", "8", mes="2026-09")["estado"] == "SIN_REPORTE"


# ── El nombre del cobrador se tipea a mano y sale mal ───────────────────────

def test_alias_de_cobrador_unifica_a_la_misma_persona():
    # En julio hay 3 filas que dicen "Yreald Romero" por "Yerald Romero". Sin
    # unificar, el resumen parte a la misma persona en dos totales inservibles.
    assert vy._cobrador_canon("Yreald Romero") == "Yerald Romero"
    assert vy._cobrador_canon("Wilder Trujillo Rosales") == "Wilder Trujillo"
    assert vy._cobrador_canon("Maximo Encarnacion") == "Maximo Encarnacion"


# ── filas_yape: se toma MONTO_YAPE, no MONTO ────────────────────────────────

def _mesa(filas):
    cols = ["MZ", "LT", "MONTO", "EFECTIVO", "YAPE", "COBRADOR", "MESA", "HOJA",
            "FECHA", "COMENTARIO", "CATEGORIA"]
    return pd.DataFrame(filas, columns=cols)


def test_filas_yape_toma_las_de_monto_yape_aunque_monto_sea_cero(monkeypatch):
    # Caso real W-5 de julio: MONTO=0 (el cobrador la marco como reclamo sin
    # cobro) con MONTO_YAPE=15. Filtrar por MONTO la escondia.
    monkeypatch.setattr(vy, "pagos_de_mesas", lambda mes: _mesa([
        ["W", "5", 0.0, 0.0, 15.0, "Wagner Trujillo", "mesa_4", "registro_1",
         "04/07/2026", "Cancelo reviza campo", "reclamo"],
        ["W", "5", 20.0, 20.0, 0.0, "Wagner Trujillo", "mesa_4", "registro_1",
         "05/07/2026", "", ""]]))
    r = vy.filas_yape("W", "5", "2026-07")
    assert len(r) == 1 and r[0]["monto"] == 15.0


def test_filas_yape_vacio_si_el_predio_no_tiene_yape(monkeypatch):
    monkeypatch.setattr(vy, "pagos_de_mesas", lambda mes: _mesa([
        ["A", "1", 8.0, 8.0, 0.0, "Wilder Trujillo", "mesa_1", "registro_1",
         "01/08/2026", "", ""]]))
    assert vy.filas_yape("A", "1", "2026-08") == []


# ── Normalizacion ───────────────────────────────────────────────────────────

def test_norm_saca_el_punto_cero_de_los_lotes_numericos():
    assert vy._norm("4.0") == "4"
    assert vy._clave("k", "3.0") == "K-3"


def test_numf_tolera_texto_vacio_y_nan():
    assert vy._numf("") == 0.0 and vy._numf(None) == 0.0
    assert vy._numf("16.0") == 16.0 and vy._numf("1,234.50") == 1234.5
