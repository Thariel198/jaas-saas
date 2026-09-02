import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "shared"))

import seguimiento_repo as repo


def test_anulacion_logica_oculta_par_compensado_sin_borrarlo(tmp_path, monkeypatch):
    monkeypatch.setattr(repo, "SEGUIMIENTO_PATH", tmp_path / "seguimiento.xlsx")
    monkeypatch.setattr(repo, "ANULACIONES_PATH", tmp_path / "anulaciones.json")

    repo.registrar_cargo("A", "1", "AGUA", "2026-08", 100,
                         source="2_planilla", audit_ref="cargo")
    repo.registrar_pago("A", "1", "AGUA", "2026-08", 30,
                        source="5_cobranza", audit_ref="pago_provisional")
    repo.registrar_ajuste("A", "1", "AGUA", "2026-08", 30,
                          source="5_cobranza", audit_ref="ajuste_compensatorio",
                          motivo="snapshot corrupto")
    repo.registrar_pago("A", "1", "AGUA", "2026-08", 30,
                        source="5_cobranza", audit_ref="pago_recuperado")

    repo.ANULACIONES_PATH.write_text(json.dumps({
        "schema": 1,
        "anulaciones": [{
            "estado": "ACTIVA",
            "eventos": [
                {"audit_ref": "pago_provisional"},
                {"audit_ref": "ajuste_compensatorio"},
            ],
        }],
    }), encoding="utf-8")

    assert len(repo._leer_eventos(incluir_anulados=True)) == 4
    assert len(repo._leer_eventos()) == 2
    assert repo._ya_registrado("5_cobranza", "pago_provisional", "A", "1", "AGUA")

    estado = repo.estado_cuenta("A", "1", "AGUA")
    agosto = estado[estado["MES"] == "2026-08"].iloc[0]
    assert agosto["PAGO"] == 30
    assert agosto["AJUSTE"] == 0
    assert agosto["SALDO"] == 70


def test_anulacion_recalcula_saldo_de_eventos_posteriores(tmp_path, monkeypatch):
    monkeypatch.setattr(repo, "SEGUIMIENTO_PATH", tmp_path / "seguimiento.xlsx")
    monkeypatch.setattr(repo, "ANULACIONES_PATH", tmp_path / "anulaciones.json")

    repo.registrar_cargo("A", "1", "AGUA", "2026-08", 100,
                         source="2_planilla", audit_ref="cargo")
    repo.registrar_pago("A", "1", "AGUA", "2026-08", 30,
                        source="abonos_rezagados", audit_ref="abono")
    repo.registrar_pago("A", "1", "AGUA", "2026-08", 20,
                        source="5_cobranza", audit_ref="pago_posterior")

    repo.ANULACIONES_PATH.write_text(json.dumps({
        "schema": 1,
        "anulaciones": [{"estado": "ACTIVA", "eventos": [{"audit_ref": "abono"}]}],
    }), encoding="utf-8")

    assert repo.get_saldo("A", "1", "AGUA", "2026-08") == 80
    assert len(repo._leer_eventos(incluir_anulados=True)) == 3
