"""
tests/test_buscar_pago.py — unitarios de las reglas de buscar_pago.py.

Los 5 primeros bloques son CONTRAFACTUALES: cada uno reproduce un falso
positivo real que la 1a version del script emitio sobre los 29 reclamos
mes_anterior del ciclo 2026-08, y verifica que la regla lo rechaza. Sin estos
tests, cualquier refactor puede reintroducir el falso positivo sin que nadie se
entere -- que es exactamente el error que la herramienta existe para evitar
("un falso OK es peor que un no-se").

Uso: py -m pytest tests/test_buscar_pago.py -q
"""

import sys
from pathlib import Path

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent / "herramienta"))

import buscar_pago as bp  # noqa: E402


# ── Contrafactual 1 — el monto debe cubrir el CONCEPTO disputado ─────────────
# Caso real Q-9: se propuso un pago de S/3 (= su "mant") para una disputa de
# S/20 de mes anterior. Y Z-17: un pago de S/34 (= "consumo+mant") para una
# disputa de S/36. Ninguno de los dos subconjuntos incluye el concepto que el
# vecino reclama, asi que encontrarlos no explica nada.

CARGOS_Q9 = {"consumo": 12.0, "mant": 3.0, "anterior": 20.0,
             "corte": 0.0, "convenio": 0.0, "multa": 0.0, "cuota": 0.0}


def test_montos_solo_incluyen_combinaciones_con_el_concepto():
    montos = bp._montos_que_cubren(CARGOS_Q9, bp.CARGO_BOLETA)
    assert 20.0 in montos                       # anterior solo
    assert 32.0 in montos                       # anterior + consumo
    assert 35.0 in montos                       # anterior + consumo + mant
    assert 3.0 not in montos, "S/3 es 'mant' solo -- el falso positivo de Q-9"
    assert 15.0 not in montos, "S/15 es consumo+mant -- no toca mes anterior"


def test_montos_vacio_si_no_debe_el_concepto():
    cargos = dict(CARGOS_Q9, anterior=0.0)
    assert bp._montos_que_cubren(cargos, "anterior") == {}


def test_todas_las_etiquetas_nombran_el_concepto():
    for etiqueta in bp._montos_que_cubren(CARGOS_Q9, "anterior").values():
        assert "anterior" in etiqueta


# ── Contrafactual 2 — capa 4: el candidato debe estar IMPAGO ─────────────────
# Caso real O-28/O-21: se propusieron C-28, O-25, O-23 y O-24 como candidatos
# de tipeo, pero el monto del pago cuadraba EXACTO con la boleta propia de esos
# lotes -- era su propio pago, no una confusion.

BOLETAS = {
    "C-28": {"nombre": "vecino c28", "total": 8.0,  "cargos": {}},
    "W-2":  {"nombre": "vecino w2",  "total": 58.0, "cargos": {}},
}


def test_capa4_descarta_si_el_otro_lote_debia_ese_monto_ese_mes(monkeypatch):
    # O-2 debia consumo 8 + mant 3 en julio: un pago de S/11 en julio es
    # plausiblemente SUYO, no un tipeo de O-28. Se mide contra la planilla del
    # mes del pago, no contra la boleta vigente (que es de otro mes).
    monkeypatch.setattr(bp, "_cargos_planilla",
                        lambda mz, lt, mes: {"consumo": 8.0, "mant": 3.0, "anterior": 0.0,
                                             "corte": 0.0, "convenio": 0.0,
                                             "multa": 30.0, "cuota": 0.0})
    assert bp._explica_su_propio_pago("O-2", 11.0, "2026-07", BOLETAS) is True


def test_capa4_no_descarta_si_no_cuadra_con_nada_suyo(monkeypatch):
    monkeypatch.setattr(bp, "_cargos_planilla",
                        lambda mz, lt, mes: {"consumo": 12.0, "mant": 3.0, "anterior": 0.0,
                                             "corte": 0.0, "convenio": 0.0,
                                             "multa": 0.0, "cuota": 0.0})
    assert bp._explica_su_propio_pago("O-25", 11.0, "2026-07", BOLETAS) is False, \
        "12+3 no da 11 en ninguna combinacion -- el candidato sigue vivo"


def test_capa4_sin_planilla_del_mes_cae_al_total_de_la_boleta(monkeypatch):
    monkeypatch.setattr(bp, "_cargos_planilla", lambda mz, lt, mes: None)
    assert bp._explica_su_propio_pago("C-28", 8.0, "2025-11", BOLETAS) is True
    assert bp._explica_su_propio_pago("W-2", 8.0, "2025-11", BOLETAS) is False
    assert bp._explica_su_propio_pago("X-99", 8.0, "2025-11", BOLETAS) is False


# ── Contrafactual 3 — el multi-lote se detecta por el MENSAJE ────────────────
# Caso real K-3/K-4 (2026-08-10): un pago de S/34 cubria 2 lotes y el sistema
# acredito uno solo. No es un tipeo: K-3 y K-4 no son confundibles (3 y 4 no
# estan en la tabla de digitos), la evidencia es que el mensaje los nombra.

def test_mensaje_reconoce_el_lote_en_varios_formatos():
    for msg in ("K-3, K-4.", "mz K lt 4", "MZK LT4", "pago mz k-4 gracias"):
        assert bp._menciona_lote(msg, "K", "4") is True, msg


def test_mensaje_no_reconoce_un_lote_distinto():
    assert bp._menciona_lote("K-3, K-4.", "K", "5") is False
    assert bp._menciona_lote("", "K", "4") is False


def test_multilote_no_es_tipeo():
    assert bp.vl.confundible("K-3", "K-4") is None, \
        "si esto matchea, el multi-lote se reporta como error de tipeo"


# ── Contrafactual 4 — ventana temporal ──────────────────────────────────────
# La regla del negocio: el reclamo normal mira 1 mes atras. Un candidato de 2+
# meses solo vale si el vecino arrastro la deuda todos los meses intermedios.

def test_ventana_acepta_el_mes_inmediato_anterior():
    ok, _ = bp._plausible("A", "1", "2026-07", "2026-08")
    assert ok is True


def test_ventana_acepta_el_mismo_mes():
    ok, _ = bp._plausible("A", "1", "2026-08", "2026-08")
    assert ok is True


def test_ventana_rechaza_candidato_posterior_al_reclamo():
    ok, motivo = bp._plausible("A", "1", "2026-09", "2026-08")
    assert ok is False and "posterior" in motivo


def test_ventana_rechaza_sin_mes():
    assert bp._plausible("A", "1", "", "2026-08")[0] is False


def test_ventana_lejana_exige_probar_el_arrastre(monkeypatch):
    # Sin planilla del mes intermedio no se puede probar el arrastre -> se
    # rechaza. Preferir un no-se a proponer un candidato de 4 meses atras.
    monkeypatch.setattr(bp, "_debia", lambda *a: None)
    ok, motivo = bp._plausible("A", "1", "2026-04", "2026-08")
    assert ok is False and "arrastre" in motivo


def test_ventana_lejana_acepta_si_arrastro_todos_los_meses(monkeypatch):
    monkeypatch.setattr(bp, "_debia", lambda *a: 20.0)      # siempre debio
    ok, motivo = bp._plausible("A", "1", "2026-04", "2026-08")
    assert ok is True and "arrastr" in motivo


def test_ventana_lejana_rechaza_si_un_mes_quedo_en_cero(monkeypatch):
    # En 2026-06 no debia nada -> la deuda se cerro, el candidato viejo no puede
    # ser lo que reclama.
    monkeypatch.setattr(bp, "_debia",
                        lambda mz, lt, mes, col: 0.0 if mes == "2026-06" else 20.0)
    ok, motivo = bp._plausible("A", "1", "2026-04", "2026-08")
    assert ok is False and "no arrastró" in motivo


# ── Contrafactual 5 — se propone SOLO si queda 1 candidato ───────────────────
# Regla heredada de verificar_lotes.py. Caso real O-28: 3 candidatos, ninguno
# elegible -- proponer uno con esa evidencia es inventar con cara de certeza.

def test_un_solo_candidato_se_propone():
    r = bp._resolver_candidatos("CANDIDATO_TIPEO",
                                [{"candidato": "W-2", "monto": 8.0, "detalle": "x"}])
    assert r["candidato"] == "W-2" and r["monto"] == 8.0


def test_varios_candidatos_no_se_elige_ninguno():
    cands = [{"candidato": f"O-{n}", "monto": 8.0, "detalle": f"d{n}"} for n in (23, 24, 25)]
    r = bp._resolver_candidatos("CANDIDATO_TIPEO", cands)
    assert r["candidato"] == "3 candidatos"
    assert r["monto"] is None, "sin monto: no se afirma cual es"


def test_sin_candidatos_no_hay_veredicto():
    assert bp._resolver_candidatos("CANDIDATO_TIPEO", []) is None


# ── Aritmética de meses ─────────────────────────────────────────────────────

def test_mes_menos_cruza_el_ano():
    assert bp._mes_menos("2026-01", 1) == "2025-12"
    assert bp._mes_menos("2026-08", 1) == "2026-07"
    assert bp._mes_menos("2026-03", 14) == "2025-01"


def test_dist_meses_signo_y_cruce_de_ano():
    assert bp._dist_meses("2026-08", "2026-07") == 1
    assert bp._dist_meses("2026-01", "2025-12") == 1
    assert bp._dist_meses("2026-07", "2026-08") == -1
    assert bp._dist_meses("2026-08", "2026-08") == 0


# ── Contrafactual 6 — consumo+mant primero NO es una anomalia ───────────────
# La 1a version llamaba "PAGO_FUE_A_OTRO_CARGO / la plata existe, decidir si se
# re-imputa" a un pago que fue a consumo+mantenimiento. Eso es la cascada
# CORRECTA (P1: agua del mes primero). La anomalia real es que se cobre multa,
# acuerdos o convenio dejando el arrastre sin pagar.

_COLS_TABLA = ["MES", "CONSUMO", "MANT", "MES_ANT", "CORTE",
               "CONVENIO", "MULTA", "ACUERDOS", "TOTAL"]
CARGOS_U2 = {"consumo": 5.0, "mant": 3.0, "anterior": 8.0,
             "corte": 0.0, "convenio": 0.0, "multa": 0.0, "cuota": 0.0}


def _tabla(**vals):
    import pandas as pd
    fila = {c: 0.0 for c in _COLS_TABLA}
    fila["MES"] = "2026-08"
    fila.update(vals)
    return pd.DataFrame([fila], columns=_COLS_TABLA)


def test_pago_a_consumo_y_mant_no_es_anomalia(monkeypatch):
    monkeypatch.setattr(bp, "_debia", lambda *a: 8.0)          # si debia el arrastre
    r = bp._clasificar_mes("U", "2", "2026-08",
                           _tabla(CONSUMO=5.0, MANT=3.0, TOTAL=8.0), CARGOS_U2)
    assert r["veredicto"] == "PAGO_SOLO_EL_MES"
    assert "cascada correcta" in r["detalle"]
    assert "re-imputar" not in r["detalle"].replace("no re-imputar", "")


def test_pago_a_multa_dejando_el_arrastre_si_es_anomalia(monkeypatch):
    monkeypatch.setattr(bp, "_debia", lambda *a: 8.0)
    r = bp._clasificar_mes("U", "2", "2026-08",
                           _tabla(CONSUMO=5.0, MANT=3.0, MULTA=30.0, TOTAL=38.0), CARGOS_U2)
    assert r["veredicto"] == "CASCADA_FUERA_DE_ORDEN"
    assert "MULTA" in r["detalle"]


def test_monto_ambiguo_se_declara_en_vez_de_afirmar(monkeypatch):
    # U-2: consumo+mant = 5+3 = 8 y el cargo de mes anterior tambien es 8. El
    # monto NO permite decidir que quiso pagar -- la 1a version afirmaba
    # "es EXACTO el cargo que reclama, asi que pago justo eso".
    monkeypatch.setattr(bp, "_debia", lambda *a: 8.0)
    r = bp._clasificar_mes("U", "2", "2026-08",
                           _tabla(CONSUMO=5.0, MANT=3.0, TOTAL=8.0), CARGOS_U2)
    assert "no dice cuál de los dos quiso pagar" in r["detalle"]


def test_monto_no_ambiguo_no_agrega_la_nota(monkeypatch):
    monkeypatch.setattr(bp, "_debia", lambda *a: 16.0)
    cargos = dict(CARGOS_U2, anterior=16.0)                    # 5+3=8 != 16
    r = bp._clasificar_mes("J", "8", "2026-08",
                           _tabla(CONSUMO=8.0, MANT=3.0, TOTAL=11.0), cargos)
    assert "no dice cuál de los dos" not in r["detalle"]


def test_si_el_arrastre_recibio_algo_es_pago_parcial(monkeypatch):
    monkeypatch.setattr(bp, "_debia", lambda *a: 8.0)
    r = bp._clasificar_mes("U", "2", "2026-08",
                           _tabla(CONSUMO=5.0, MANT=3.0, MES_ANT=4.0, TOTAL=12.0), CARGOS_U2)
    assert r["veredicto"] == "PAGO_PARCIAL"


def test_sin_deuda_del_arrastre_ese_mes_no_hay_veredicto(monkeypatch):
    # Si no debia mes anterior, que reciba 0 es correcto -- no es hallazgo.
    monkeypatch.setattr(bp, "_debia", lambda *a: 0.0)
    assert bp._clasificar_mes("U", "2", "2026-08",
                              _tabla(CONSUMO=5.0, MANT=3.0, TOTAL=8.0), CARGOS_U2) is None


def test_sin_pago_ese_mes_no_hay_veredicto():
    assert bp._clasificar_mes("U", "2", "2026-08", _tabla(TOTAL=0.0), CARGOS_U2) is None


def test_historico_nunca_pago_un_arrastre():
    import pandas as pd
    vacio = pd.DataFrame([{c: 0.0 for c in _COLS_TABLA}])
    etiqueta, texto = bp._historico_mes_anterior(vacio)
    assert etiqueta == "nunca"                       # valor exacto: la columna se filtra por el
    assert "nunca se le aplicó nada" in texto


def test_historico_viene_pagando_arrastres():
    # El grupo que importa: 6 de los 14 de 2026-08 vienen pagando arrastres y aun
    # asi les sigue apareciendo -- ahi hay que buscar el problema real, no en los
    # 8 que nunca pagaron uno.
    import pandas as pd
    con = pd.DataFrame([dict({c: 0.0 for c in _COLS_TABLA}, MES="2026-06", MES_ANT=12.0),
                        dict({c: 0.0 for c in _COLS_TABLA}, MES="2026-07", MES_ANT=8.0)])
    etiqueta, texto = bp._historico_mes_anterior(con)
    assert etiqueta == "S/20.00 en 2 meses"
    assert "20.00" in texto and "2026-06" in texto and "2026-07" in texto


# ── Contrafactual 7 — la FUENTE (mesa_N) vs lo consolidado ──────────────────
# La cadena es  mesa_N.xlsx -> pagos_efectivo.xlsx -> planilla_cobrado.xlsx  y
# el pago se puede perder en cualquiera de los dos saltos. Buscar solo en el
# consolidado no encuentra nunca el pago que murio en el primer salto.

def _mesas(**por_mes):
    import pandas as pd
    cols = ["MZ", "LT", "MONTO", "EFECTIVO", "YAPE", "COBRADOR", "MESA", "HOJA",
            "FECHA", "COMENTARIO"]
    return {m: pd.DataFrame(filas, columns=cols) for m, filas in por_mes.items()}


def _tabla_sin_pago(mes):
    import pandas as pd
    return pd.DataFrame([dict({c: 0.0 for c in _COLS_TABLA}, MES=mes)])


def test_pago_en_la_mesa_que_no_llego_al_historial_se_reporta():
    # Caso real F1-8: Wagner anoto S/13 en mesa_4 el 02/08 y el historial de
    # agosto dice que no pago.
    mesas = _mesas(**{"2026-08": [["F1", "8", 13.0, 0.0, 13.0, "Wagner Trujillo",
                                   "mesa_4", "registro_1", "02/08/2026", "Yape"]]})
    r = bp._en_fuente_sin_consolidar("F1", "8", _tabla_sin_pago("2026-08"), mesas)
    assert len(r) == 1 and r[0]["monto"] == 13.0


def test_yape_anotado_en_la_hoja_de_efectivo_se_marca():
    # Cae entre dos modulos: 4_pagos/efectivo solo procesa MONTO_EFECTIVO y
    # motor_matching lee el banco, no las mesas. Nadie lo levanta.
    mesas = _mesas(**{"2026-08": [["F1", "8", 13.0, 0.0, 13.0, "Wagner Trujillo",
                                   "mesa_4", "registro_1", "02/08/2026", ""]]})
    r = bp._en_fuente_sin_consolidar("F1", "8", _tabla_sin_pago("2026-08"), mesas)
    assert "YAPE en la hoja de efectivo" in r[0]["detalle"]


def test_no_se_reporta_lo_que_un_precursor_ya_explica():
    # Caso real I-9: S/86 anotados por Wagner en mesa_6 de junio, que NO figuran
    # en junio porque abonos_rezagados los aplico en julio. Ya tiene dueño.
    mesas = _mesas(**{"2026-06": [["I", "9", 86.0, 86.0, 0.0, "Wagner Trujillo",
                                   "mesa_6", "registro_1", "05/06/2026", ""]]})
    assert bp._en_fuente_sin_consolidar("I", "9", _tabla_sin_pago("2026-06"), mesas,
                                        ya_explicados={86.0}) == []


def test_el_cobrador_que_nombra_el_vecino_va_primero():
    mesas = _mesas(**{"2026-07": [["O", "28", 8.0, 8.0, 0.0, "Wilder Trujillo",
                                   "mesa_1", "registro_1", "04/07/2026", ""],
                                  ["O", "28", 8.0, 8.0, 0.0, "Maximo Encarnacion",
                                   "mesa_3", "registro_1", "04/07/2026", ""]]})
    r = bp._en_fuente_sin_consolidar("O", "28", _tabla_sin_pago("2026-07"), mesas,
                                     cobrador_dicho="Maximo Encarnacion")
    assert r[0]["coincide_cobrador"] is True
    assert len(r) == 2, "se ordena por el cobrador nombrado, pero no se descarta el resto"


def test_cobrador_mencionado_detecta_el_nombre_en_el_texto():
    cobs = {"Maximo Encarnacion", "Wagner Trujillo", "Yerald Romero"}
    assert bp._cobrador_mencionado("Pago mes anterior con Maximo", cobs) == "Maximo Encarnacion"
    assert bp._cobrador_mencionado("le pago a Wagner Trujillo 15", cobs) == "Wagner Trujillo"
    assert bp._cobrador_mencionado("Pago mes anterior", cobs) == ""


# ── Contrafactual 8 — el yape se verifica contra el banco ───────────────────
# El reporte del banco es la UNICA prueba de que un yape entro a la cuenta de la
# JASS. Si el cobrador anota "yape" y ahi no hay nada que calce, el pago no
# entro -- pudo haber ido a un telefono personal.

def _banco(filas):
    import pandas as pd
    df = pd.DataFrame(filas, columns=["TIPO", "ORIGEN", "MONTO", "MENSAJE", "FECHA"])
    df["FECHA"] = pd.to_datetime(df["FECHA"], format="%d/%m/%Y")
    return df


def test_yape_que_no_esta_en_el_banco_se_reporta(monkeypatch):
    # Caso real F1-8: Wagner anoto S/13 el 02/08 y no hay ninguna transaccion
    # de S/13 en la cuenta de la JASS en esos dias.
    monkeypatch.setattr(bp, "_reporte_banco",
                        lambda: _banco([["TE PAGÓ", "Otro", 8.0, "", "02/08/2026"]]))
    r = bp._yape_en_banco(13.0, "02/08/2026", "F1", "8")
    assert r["estado"] == "NO_EXISTE"


def test_yape_que_si_esta_en_el_banco_se_encuentra(monkeypatch):
    monkeypatch.setattr(bp, "_reporte_banco",
                        lambda: _banco([["TE PAGÓ", "Juan B*", 13.0, "", "01/08/2026"]]))
    r = bp._yape_en_banco(13.0, "02/08/2026", "F1", "8")
    assert r["estado"] == "ENCONTRADO"


def test_yape_se_encuentra_por_mensaje_aunque_no_calce_la_fecha(monkeypatch):
    monkeypatch.setattr(bp, "_reporte_banco",
                        lambda: _banco([["TE PAGÓ", "X", 99.0, "mz F1 lt 8", "01/01/2026"]]))
    r = bp._yape_en_banco(13.0, "02/08/2026", "F1", "8")
    assert r["estado"] == "ENCONTRADO" and "nombra el lote" in r["detalle"]


def test_fecha_fuera_del_reporte_no_concluye_que_no_existe(monkeypatch):
    # No es lo mismo "no esta" que "no lo puedo saber": si el reporte no cubre
    # esa fecha, afirmar que el yape no entro seria inventar.
    monkeypatch.setattr(bp, "_reporte_banco",
                        lambda: _banco([["TE PAGÓ", "X", 8.0, "", "02/08/2026"]]))
    r = bp._yape_en_banco(13.0, "15/01/2026", "F1", "8")
    assert r["estado"] == "FUERA_DE_RANGO"


def test_sin_reporte_del_banco_no_se_concluye_nada(monkeypatch):
    import pandas as pd
    monkeypatch.setattr(bp, "_reporte_banco", lambda: pd.DataFrame())
    assert bp._yape_en_banco(13.0, "02/08/2026", "F1", "8")["estado"] == "SIN_REPORTE"


def test_solo_se_toman_las_filas_con_monto_yape():
    mesas = _mesas(**{"2026-08": [
        ["F1", "8", 13.0, 0.0, 13.0, "Wagner Trujillo", "mesa_4", "registro_1", "02/08/2026", ""],
        ["F1", "8", 20.0, 20.0, 0.0, "Wagner Trujillo", "mesa_4", "registro_1", "03/08/2026", ""]]})
    r = bp._filas_yape_de_mesas("F1", "8", mesas)
    assert len(r) == 1 and r[0]["monto"] == 13.0


# ── Alcance: solo mes_anterior ──────────────────────────────────────────────

def test_la_herramienta_es_solo_de_mes_anterior():
    # convenio y cuota NO se auditan aca: su causa es el orden de la cascada
    # (multa antes de convenio/acuerdos), que es otro trabajo. Si alguien
    # extiende el alcance, tiene que decidir esto a proposito, no por accidente.
    assert bp.TIPO == "mes_anterior"
    assert (bp.CARGO_BOLETA, bp.COL_TABLA, bp.COL_PLANILLA) == \
        ("anterior", "MES_ANT", "MES_ANTERIOR")


# ── Normalización ───────────────────────────────────────────────────────────

def test_norm_saca_el_punto_cero_de_los_lotes_numericos():
    # pagos_yape_tepago guarda LOTE como float: 4.0 debe comparar igual que "4"
    assert bp._norm("4.0") == "4"
    assert bp._clave("k", "3.0") == "K-3"


def test_numf_tolera_texto_vacio_y_nan():
    assert bp._numf("") == 0.0
    assert bp._numf(None) == 0.0
    assert bp._numf("16.0") == 16.0
    assert bp._numf("1,234.50") == 1234.5


# ── El tipo que audita sigue siendo un TIPO_RECLAMO valido del modulo ───────

def test_el_tipo_existe_en_el_modulo():
    import main as reclamos
    assert bp.TIPO in reclamos.TIPOS_RECLAMO_VALIDOS
