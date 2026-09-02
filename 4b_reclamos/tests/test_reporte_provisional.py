import sys
from pathlib import Path

import fitz
import pandas as pd
import pytest


ROOT = Path(__file__).parents[2]
sys.path.insert(0, str(ROOT / "4b_reclamos" / "herramienta"))
import reporte_historico as reporte  # noqa: E402


def _eventos():
    return pd.DataFrame([
        {"MZ": "A", "LT": "4", "CONCEPTO": "CONVENIO", "MES": "2026-06",
         "TIPO_EVENTO": "CARGO", "CARGO": 75.0, "PAGO": None, "AJUSTE": None,
         "SALDO": 75.0, "TIMESTAMP": "2026-06-01 08:00:00", "CLASE": "GENESIS",
         "SOURCE": "siembra", "AUDIT_REF": "cargo_convenio", "MOTIVO": ""},
        {"MZ": "A", "LT": "4", "CONCEPTO": "MULTA", "MES": "2026-06",
         "TIPO_EVENTO": "CARGO", "CARGO": 30.0, "PAGO": None, "AJUSTE": None,
         "SALDO": 30.0, "TIMESTAMP": "2026-06-01 08:00:01", "CLASE": "GENESIS",
         "SOURCE": "siembra", "AUDIT_REF": "cargo_multa", "MOTIVO": ""},
        {"MZ": "A", "LT": "4", "CONCEPTO": "MULTA", "MES": "2026-07",
         "TIPO_EVENTO": "PAGO", "CARGO": None, "PAGO": 30.0, "AJUSTE": None,
         "SALDO": 0.0, "TIMESTAMP": "2026-07-01 08:00:00", "CLASE": "COBRANZA",
         "SOURCE": "5_cobranza", "AUDIT_REF": "pago_multa", "MOTIVO": ""},
        {"MZ": "A", "LT": "4", "CONCEPTO": "MULTA", "MES": "2026-08",
         "TIPO_EVENTO": "AJUSTE", "CARGO": None, "PAGO": None, "AJUSTE": 30.0,
         "SALDO": 30.0, "TIMESTAMP": "2026-08-24 14:00:00", "CLASE": "REASIGNACION",
         "SOURCE": "reimputacion_ca1", "AUDIT_REF": "reasig", "MOTIVO": "Prioridad CA1"},
        {"MZ": "A", "LT": "4", "CONCEPTO": "CONVENIO", "MES": "2026-08",
         "TIPO_EVENTO": "AJUSTE", "CARGO": None, "PAGO": None, "AJUSTE": -30.0,
         "SALDO": 45.0, "TIMESTAMP": "2026-08-24 14:00:01", "CLASE": "REASIGNACION",
         "SOURCE": "reimputacion_ca1", "AUDIT_REF": "reasig", "MOTIVO": "Prioridad CA1"},
    ])


def test_tabla_sale_solo_de_eventos_activos():
    tabla = reporte.tabla_predio_ledger("A", "4", _eventos(), "2026-08")

    assert tabla["MES"].tolist() == ["2026-06", "2026-07", "2026-08"]
    assert tabla["COBERTURA"].tolist() == ["PARCIAL", "PARCIAL", "COMPLETA"]
    agosto = tabla.iloc[-1]
    assert agosto["AJUSTE_MULTA"] == 30.0
    assert agosto["AJUSTE_CONVENIO"] == -30.0
    assert agosto["SALDO_MULTA"] == 30.0
    assert agosto["SALDO_CONVENIO"] == 45.0
    assert agosto["SALDO_TOTAL"] == 75.0


def test_ajuste_no_se_presenta_como_pago():
    tabla = reporte.tabla_predio_ledger("A", "4", _eventos(), "2026-08")
    agosto = tabla.iloc[-1]
    assert agosto["PAGO_MULTA"] == 0.0
    assert agosto["PAGO_CONVENIO"] == 0.0
    assert agosto["AJUSTE_TOTAL"] == 0.0


def test_correccion_tecnica_de_siembra_no_se_muestra_como_ajuste():
    eventos = pd.concat([_eventos(), pd.DataFrame([{
        "MZ": "A", "LT": "4", "CONCEPTO": "CONVENIO", "MES": "2026-06",
        "TIPO_EVENTO": "AJUSTE", "CARGO": None, "PAGO": None, "AJUSTE": -10.0,
        "SALDO": 65.0, "TIMESTAMP": "2026-06-02 08:00:00", "CLASE": "CORRECCION_SISTEMA",
        "SOURCE": "correccion_genesis_formula", "AUDIT_REF": "fix_siembra", "MOTIVO": "",
    }])], ignore_index=True)

    tabla = reporte.tabla_predio_ledger("A", "4", eventos, "2026-08")
    ajustes = reporte._ajustes_predio("A", "4", eventos)

    assert tabla.loc[tabla["MES"] == "2026-06", "AJUSTE_CONVENIO"].iloc[0] == 0.0
    assert "CORRECCION_SISTEMA" not in ajustes["CLASE"].tolist()


def test_foto_boleta_muestra_deuda_y_aplicacion_completa(monkeypatch):
    deuda = {
        "CONSUMO": 5.0, "MANT": 3.0, "MES_ANT": 23.0, "CORTE": 0.0,
        "CONVENIO": 75.0, "MULTA": 30.0, "ACUERDOS": 75.0,
    }
    pago = {
        "CONSUMO": 5.0, "MANT": 3.0, "MES_ANT": 23.0, "CORTE": 0.0,
        "CONVENIO": 0.0, "MULTA": 30.0, "ACUERDOS": 75.0,
    }
    monkeypatch.setattr(reporte, "_foto_boleta", lambda *_args: (deuda, pago))
    monkeypatch.setattr(reporte, "_filas_historicas", lambda *_args: pd.DataFrame())

    tabla = reporte.tabla_predio_reporte("A", "4", _eventos(), "2026-08")
    julio = tabla[tabla["MES"] == "2026-07"].iloc[0]

    assert julio["DEUDA_TOTAL"] == 211.0
    assert julio["PAGO_TOTAL"] == 136.0
    assert julio["PAGO_CONSUMO"] == 5.0
    assert julio["PAGO_MANT"] == 3.0
    assert julio["PAGO_MES_ANT"] == 23.0
    assert julio["PAGO_MULTA"] == 30.0
    assert julio["PAGO_ACUERDOS"] == 75.0
    assert julio["SALDO_CONVENIO"] == 75.0


def test_foto_boleta_usa_coordenada_historica_reasignada(monkeypatch):
    boletas = pd.DataFrame([{
        "MZ": "F1", "LT": "7", "TOTAL MES ACTUAL": 5.0, "MANTENIMIENTO": 3.0,
        "MES ANTERIOR": 0.0, "CORTE Y RECONEXION": 0.0, "CONVENIO": 0.0,
        "MULTA (FAENA + REUNION)": 8.0, "CUOTA DIRECTA": 0.0,
    }])
    monkeypatch.setattr(reporte, "_coordenada_ciclo", lambda *_args: ("F1", "7"))
    monkeypatch.setattr(reporte, "_cargar_data_boletas", lambda *_args: boletas)
    monkeypatch.setattr(reporte, "_pago_confirmado_ciclo", lambda *_args: 0.0)

    deuda, pago = reporte._foto_boleta("F1", "6", "2026-07")

    assert deuda["CONSUMO"] == 5.0
    assert deuda["MULTA"] == 8.0
    assert sum(pago.values()) == 0.0


def test_saldo_de_planilla_cobrado_no_inventa_un_pago(monkeypatch):
    boletas = pd.DataFrame([{
        "MZ": "I", "LT": "9", "TOTAL MES ACTUAL": 5.0, "MANTENIMIENTO": 3.0,
        "MES ANTERIOR": 11.0, "CORTE Y RECONEXION": 0.0, "CONVENIO": 0.0,
        "MULTA (FAENA + REUNION)": 50.0, "CUOTA DIRECTA": 75.0,
    }])
    monkeypatch.setattr(reporte, "_coordenada_ciclo", lambda *_args: ("I", "9"))
    monkeypatch.setattr(reporte, "_cargar_data_boletas", lambda *_args: boletas)
    monkeypatch.setattr(reporte, "_pago_confirmado_ciclo", lambda *_args: 0.0)

    deuda, pago = reporte._foto_boleta("I", "9", "2026-07")

    assert sum(deuda.values()) == 144.0
    assert sum(pago.values()) == 0.0


def test_coordenada_ciclo_elige_el_lote_cuyo_nombre_coincide(monkeypatch):
    boletas = pd.DataFrame([
        {"MZ": "G", "LT": "17", "NOMBRES": "CARLOS ALBERTO SIGUENAS SAENZ"},
        {"MZ": "G", "LT": "16C", "NOMBRES": "ELOY SIGUENAS UGARTE"},
    ])
    monkeypatch.setattr(reporte, "_coordenada_historica", lambda *_args: ("G", "16C"))
    monkeypatch.setattr(reporte, "_nombres_actuales", lambda: {("G", "17"): "CARLOS ALBERTO SIGUENAS SAENZ"})
    monkeypatch.setattr(reporte, "_cargar_data_boletas", lambda *_args: boletas)

    assert reporte._coordenada_ciclo("G", "17", "2026-07") == ("G", "17")


def test_referencias_junio_julio_usan_coordenada_historica(monkeypatch):
    monkeypatch.setattr(reporte, "_coordenada_ciclo", lambda *_args: ("F1", "7"))
    monkeypatch.setattr(reporte.comun, "referencias_pago", lambda mz, lt, **_kwargs: (
        [{"MES": "2026-07", "MEDIO": "YAPE", "MONTO": 68.0},
         {"MES": "2026-08", "MEDIO": "EFECTIVO", "MONTO": 8.0}]
        if (mz, lt) == ("F1", "6")
        else [{"MES": "2026-07", "MEDIO": "ABONO REZ.", "MONTO": 30.0}]
    ))
    monkeypatch.setattr(reporte, "_aportes_tanque", lambda *_args: [])
    refs = reporte._referencias("F1", "6", "2026-08")

    assert [(r["MES"], r["MONTO"]) for r in refs] == [("2026-07", 30.0), ("2026-08", 8.0)]
    assert refs[0]["ESTADO_LEDGER"] == "NO ASENTADO EN LEDGER"


def test_fila_historica_reutiliza_regla_estable_de_deuda_y_pago():
    fuente = {
        "DEUDA_CONSUMO": 8.0, "DEUDA_MANT": 3.0, "DEUDA_MES_ANT": 10.0,
        "CONSUMO": 8.0, "MANT": 3.0, "MES_ANT": 10.0,
    }

    fila = reporte._fila_historica_reporte("2026-05", fuente)

    assert fila["DEUDA_TOTAL"] == 21.0
    assert fila["PAGO_TOTAL"] == 21.0
    assert fila["COBERTURA"] == "HISTORICO"


def test_no_admite_mes_no_comprometido(monkeypatch):
    monkeypatch.setattr(reporte.repo_estado, "ultimo_ledger_comprometido", lambda ruta: "2026-08")
    with pytest.raises(ValueError, match="no esta comprometido"):
        reporte._mes_comprometido("2026-09")


def test_referencias_solo_hasta_cierre(monkeypatch):
    monkeypatch.setattr(reporte.comun, "referencias_pago", lambda *_args, **_kwargs: [
        {"MES": "2026-05", "MEDIO": "YAPE", "MONTO": 5.0},
        {"MES": "2026-06", "MEDIO": "YAPE", "MONTO": 10.0},
        {"MES": "2026-08", "MEDIO": "EFECTIVO", "MONTO": 20.0},
        {"MES": "2026-09", "MEDIO": "YAPE", "MONTO": 30.0},
    ])
    monkeypatch.setattr(reporte, "_aportes_tanque", lambda *_args: [])
    refs = reporte._referencias("A", "4", "2026-08")
    assert [r["MES"] for r in refs] == ["2026-05", "2026-06", "2026-08"]
    assert {r["ESTADO_LEDGER"] for r in refs} == {"PAGO REGISTRADO"}


def test_abono_rezagado_indica_mes_aplicado_y_que_no_esta_en_ledger(monkeypatch):
    monkeypatch.setattr(reporte.comun, "referencias_pago", lambda *_args, **_kwargs: [{
        "MES": "2026-06", "MES_APLICA": "2026-07", "MEDIO": "ABONO REZ.",
        "FECHA_HORA": "05/06/2026 · aplicado en 2026-07", "MONTO": 58.0,
    }])
    monkeypatch.setattr(reporte, "_aportes_tanque", lambda *_args: [])
    refs = reporte._referencias("L", "4", "2026-08")

    assert refs[0]["MES"] == "2026-07"
    assert refs[0]["ESTADO_LEDGER"] == "NO ASENTADO EN LEDGER"


def test_referencia_yape_separa_aporte_tanque_con_fecha_completa(monkeypatch):
    monkeypatch.setattr(reporte.comun, "referencias_pago", lambda *_args, **_kwargs: [
        {"MES": "2026-07", "MEDIO": "YAPE", "FECHA_HORA": "fechas agregadas", "MONTO": 236.0},
    ])
    monkeypatch.setattr(reporte, "_aportes_tanque", lambda *_args: [
        {"MES": "2026-07", "MONTO": 100.0, "FECHA_REAL": "07/07/2026"},
    ])
    monkeypatch.setattr(reporte, "_yapes_crudos", lambda *_args: [
        {"MONTO": 136.0, "FECHA_HORA": "07/07/2026 07:20:05"},
        {"MONTO": 100.0, "FECHA_HORA": "07/07/2026 07:20:55"},
    ])

    refs = reporte._referencias("A", "4", "2026-08")

    assert refs == [
        {"MES": "2026-07", "MEDIO": "YAPE", "FECHA_HORA": "07/07/2026 07:20:05",
         "MONTO": 136.0, "ESTADO_LEDGER": "PAGO REGISTRADO"},
        {"MES": "2026-07", "MEDIO": "APORTE TANQUE", "FECHA_HORA": "07/07/2026 07:20:55",
         "ESTADO_LEDGER": "NO REDUCE DEUDA", "MONTO": 100.0},
    ]


def test_pdf_conserva_formato_y_muestra_reasignacion(tmp_path):
    tabla = reporte.tabla_predio_ledger("A", "4", _eventos(), "2026-08")
    tabla.loc[tabla["MES"] < "2026-08", "COBERTURA"] = "DATA_BOLETAS"
    ajustes = reporte._ajustes_predio("A", "4", _eventos())
    doc = fitz.open()
    reporte._dibujar_pagina_ledger(doc, "A", "4", "PRUEBA", tabla, ajustes)
    salida = tmp_path / "reporte.pdf"
    doc.save(salida)
    doc.close()

    texto = "".join(page.get_text() for page in fitz.open(salida))
    assert "historial mensual de deuda y pagos" in texto
    assert "REASIGNACION" in texto
    assert "Junio/julio: DATA_boletas" in texto
    assert texto.splitlines().count("PAGO SIM.") == 1
    assert texto.splitlines().count("SALDO") == 1
    assert texto.splitlines().count("AJUSTE") == 1
    assert "NO ASENTADO" not in texto


def test_pagina_referencias_tiene_cabecera_gris_distinta_del_predio(tmp_path):
    doc = fitz.open()
    reporte.comun._dibujar_pagina_referencias(doc, "I", "9", "PRUEBA", [])
    salida = tmp_path / "referencias.pdf"
    doc.save(salida)
    doc.close()

    with fitz.open(salida) as pdf:
        pagina = pdf[0]
        assert "Referencias del predio I-9" in pagina.get_text()
        assert pagina.get_drawings()[0]["fill"] == pytest.approx(reporte.comun._REFERENCIAS_BG)
        assert pagina.get_drawings()[0]["fill"] != pytest.approx(reporte.comun._AZUL)
