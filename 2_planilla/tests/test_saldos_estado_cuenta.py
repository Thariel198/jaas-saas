import importlib.util
import json
import sys
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).parents[2]
sys.path.insert(0, str(ROOT / "2_planilla"))
SPEC = importlib.util.spec_from_file_location("planilla_main", ROOT / "2_planilla" / "main.py")
planilla = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(planilla)
VALIDADOR_SPEC = importlib.util.spec_from_file_location(
    "planilla_validador", ROOT / "2_planilla" / "validar_arrastres.py"
)
validador = importlib.util.module_from_spec(VALIDADOR_SPEC)
VALIDADOR_SPEC.loader.exec_module(validador)


def test_septiembre_lee_saldos_comprometidos_de_agosto(tmp_path, monkeypatch):
    ledger = tmp_path / "estado_cuenta.xlsx"
    estado = tmp_path / "estado_ciclo.json"
    monkeypatch.setattr(planilla.repo, "SEGUIMIENTO_PATH", ledger)
    monkeypatch.setattr(planilla.config, "ESTADO_CICLO_PATH", estado)

    planilla.repo.registrar_cargo("A", "1", "AGUA", "2026-08", 20,
                                  source="test", audit_ref="agua")
    planilla.repo.registrar_cargo("A", "1", "MANTENIMIENTO", "2026-08", 3,
                                  source="test", audit_ref="mant")
    planilla.repo.registrar_cargo("A", "1", "CORTE_RECONEXION", "2026-08", 40,
                                  source="test", audit_ref="corte")
    estado.write_text(json.dumps({"2026-08": {"ledger": {
        "comprometido": True, "snapshot_hash": "a" * 64,
    }}}), encoding="utf-8")

    saldos = planilla._load_saldos_cuenta("2026-09")
    fila = saldos.iloc[0]
    assert fila["DEUDA_AGUA"] == 23.0
    assert fila["CORTE_RECONEXION"] == 40.0


def test_septiembre_bloquea_si_agosto_no_esta_comprometido(tmp_path, monkeypatch):
    estado = tmp_path / "estado_ciclo.json"
    estado.write_text("{}", encoding="utf-8")
    monkeypatch.setattr(planilla.config, "ESTADO_CICLO_PATH", estado)

    try:
        planilla._load_saldos_cuenta("2026-09")
    except ValueError as exc:
        assert "no está comprometido" in str(exc)
    else:
        raise AssertionError("debía bloquear una planilla construida sin cierre oficial")


def test_septiembre_une_lecturas_directas_y_toda_la_deuda_viva(tmp_path, monkeypatch):
    ledger = tmp_path / "seguimiento_pueblo.xlsx"
    estado = tmp_path / "estado_ciclo.json"
    lecturas_dir = tmp_path / "lecturas"
    lecturas_dir.mkdir()
    monkeypatch.setattr(planilla.repo, "SEGUIMIENTO_PATH", ledger)
    monkeypatch.setattr(planilla.config, "ESTADO_CICLO_PATH", estado)
    monkeypatch.setattr(planilla.config, "LECTURAS_DIR", lecturas_dir)

    pd.DataFrame([{
        "MZ": "A", "LT": "1", "NOMBRE": "USUARIO PRUEBA", "MES_ANO": "2026-09",
        "MARC_ANT": 100, "MARC_ACT": 108, "M3": 8,
    }]).to_excel(lecturas_dir / "lecturas_planilla_2026-09.xlsx", index=False, startrow=1)

    montos = {
        "AGUA": 20, "MANTENIMIENTO": 3, "CORTE_RECONEXION": 40,
        "CONVENIO": 50, "MULTA": 30, "ACUERDOS": 75,
    }
    for concepto, monto in montos.items():
        planilla.repo.registrar_cargo(
            "A", "1", concepto, "2026-08", monto,
            source="test", audit_ref=f"cargo_{concepto}",
        )
    estado.write_text(json.dumps({"2026-08": {"ledger": {
        "comprometido": True, "snapshot_hash": "a" * 64,
    }}}), encoding="utf-8")

    fila = planilla.build_planilla("2026-09").iloc[0]

    assert fila["MES_ACTUAL"] == 8.0
    assert fila["MANTENIMIENTO"] == 3.0
    assert fila["MES_ANTERIOR"] == 23.0
    assert fila["CORTE_RECONEXION"] == 40.0
    assert fila["CONVENIO"] == 50.0
    assert fila["MULTA"] == 30.0
    assert fila["ACUERDOS_ASAMBLEA"] == 75.0


def test_validador_septiembre_contrasta_toda_la_deuda_viva(monkeypatch):
    planilla_df = pd.DataFrame([{
        "MZ": "A", "LT": "1", "_mz": "A", "_lt": "1", "NOMBRE": "USUARIO PRUEBA",
        "MARC_ANT": 100.0, "MARC_ACT": 108.0, "M3": 8.0,
        "MES_ANTERIOR": 23.0, "CORTE_RECONEXION": 40.0,
        "CONVENIO": 50.0, "MULTA": 30.0, "ACUERDOS_ASAMBLEA": 75.0,
    }, {
        "MZ": "B", "LT": "2", "_mz": "B", "_lt": "2", "NOMBRE": "SIN LECTURA",
        "MARC_ANT": 0.0, "MARC_ACT": 0.0, "M3": 0.0,
        "MES_ANTERIOR": 10.0, "CORTE_RECONEXION": 0.0,
        "CONVENIO": 0.0, "MULTA": 0.0, "ACUERDOS_ASAMBLEA": 0.0,
    }])
    lecturas_df = planilla_df.iloc[:1][["MZ", "LT", "_mz", "_lt", "NOMBRE", "MARC_ANT", "MARC_ACT", "M3"]]
    deuda_df = pd.DataFrame([{
        "MZ": "A", "LT": "1", "_mz": "A", "_lt": "1",
        "DEUDA_AGUA": 23.0, "CORTE_RECONEXION": 40.0,
    }, {
        "MZ": "B", "LT": "2", "_mz": "B", "_lt": "2",
        "DEUDA_AGUA": 10.0, "CORTE_RECONEXION": 0.0,
    }])
    saldos = {"CONVENIO": 50.0, "MULTA": 30.0, "ACUERDOS": 75.0}

    monkeypatch.setattr(validador, "_cargar_planilla", lambda mes: planilla_df)
    monkeypatch.setattr(validador, "_cargar_lecturas_origen", lambda mes: lecturas_df)
    monkeypatch.setattr(validador, "_load_saldos_cuenta", lambda mes: deuda_df)
    monkeypatch.setattr(
        validador.repo, "get_saldos_bulk",
        lambda concepto, mes: {("A", "1"): saldos.get(concepto, 0.0)},
    )

    resumen, discrepancias = validador.comparar("2026-09")
    fuentes = {fila["CONCEPTO"]: fila["FUENTE"] for fila in resumen}

    assert discrepancias == []
    assert fuentes["MES_ANTERIOR"] == "seguimiento_repo.AGUA+MANTENIMIENTO"
    assert fuentes["CORTE_RECONEXION"] == "seguimiento_repo.CORTE_RECONEXION"
    cobertura = next(fila for fila in resumen if fila["CONCEPTO"] == "(MZ, LT)")
    assert cobertura["COINCIDEN"] == 2
    assert cobertura["DISCREPANCIAS"] == 0
    assert cobertura["NOTA"] == "1 sin lectura agregados por deuda viva"
