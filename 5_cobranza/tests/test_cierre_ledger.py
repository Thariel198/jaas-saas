"""Contrato del commit único: proyección pura + writer batch idempotente."""
import sys
import tempfile
from pathlib import Path

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))
import main as cobranza  # noqa: E402

sys.path.insert(0, str(THIS.parents[2] / "shared"))
import seguimiento_repo as repo  # noqa: E402
import utils_estado_ciclo as estado  # noqa: E402


def _usuario(total):
    return {
        "mz": "A", "lt": "1", "mes_actual": 0.0, "mantenimiento": 0.0,
        "mes_anterior": 0.0, "blanco_final": 0.0, "devolucion": 0.0,
        "corte_reconexion": 0.0, "convenio": 0.0, "acuerdos_asamblea": 0.0,
        "multa": 50.0, "total_pagado": total, "total_pagado_normal": total,
        "abono_rezagado": 0.0, "abono_rezagado_cerrado": 0.0,
        "abono_rezagado_vigente": 0.0, "mes_ano": "2099-01",
    }


def main():
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        repo.SEGUIMIENTO_PATH = raiz / "seguimiento_pueblo.xlsx"
        cobranza.repo.SEGUIMIENTO_PATH = repo.SEGUIMIENTO_PATH
        cobranza.OUTPUTS_DIR = raiz / "outputs"
        cobranza.OUTPUTS_DIR.mkdir()
        cobranza.GENESIS_TARDIA_PATH = raiz / "genesis_inexistente.xlsx"

        objetivos = cobranza._objetivos_ledger([_usuario(30.0)], "2099-01")
        hash_a = cobranza._exportar_snapshot_ledger([_usuario(30.0)], "2099-01")
        assert not repo.SEGUIMIENTO_PATH.exists(), "generar snapshot no debe escribir el ledger"
        assert hash_a == cobranza._exportar_snapshot_ledger([_usuario(30.0)], "2099-01")

        repo.registrar_cargo("A", "1", "MULTA", "2099-01", 50.0,
                             source="test", audit_ref="cargo")
        primero = repo.reconciliar_objetivos_batch("2099-01", hash_a, objetivos)
        retry = repo.reconciliar_objetivos_batch("2099-01", hash_a, objetivos)
        assert primero["eventos"] == 1 and retry["eventos"] == 0
        assert repo.get_saldo("A", "1", "MULTA", "2099-01") == 20.0

        q5 = {
            "mz": "Q", "lt": "5", "mes_ano": "2026-08",
            "mes_actual": 13.0, "mantenimiento": 3.0, "mes_anterior": 20.0,
            "blanco_final": 0.0, "devolucion": 0.0, "corte_reconexion": 0.0,
            "corte_reconexion_base": 0.0, "convenio": 25.0,
            "acuerdos_asamblea": 50.0, "multa": 19.0,
            "total_pagado": 114.0, "total_pagado_normal": 0.0,
            "abono_rezagado": 114.0, "abono_rezagado_cerrado": 114.0,
            "abono_rezagado_vigente": 0.0,
        }
        for concepto, monto in (("CONVENIO", 25), ("ACUERDOS", 50), ("MULTA", 19)):
            repo.registrar_cargo("Q", "5", concepto, "2026-07", monto,
                                 source="test", audit_ref=f"q5-{concepto}")
        objetivos_q5 = cobranza._objetivos_ledger([q5], "2026-08")
        cargos_q5 = cobranza._cargos_cuenta_snapshot([q5], "2026-08")
        repo.reconciliar_objetivos_batch("2026-08", "q" * 64, objetivos_q5, cargos_q5)
        assert repo.get_saldo("Q", "5", "AGUA", "2026-08") == 13.0
        assert repo.get_saldo("Q", "5", "MANTENIMIENTO", "2026-08") == 3.0
        assert repo.get_saldo("Q", "5", "CONVENIO", "2026-08") == 0.0
        assert repo.get_saldo("Q", "5", "ACUERDOS", "2026-08") == 0.0
        assert repo.get_saldo("Q", "5", "MULTA", "2026-08") == 0.0

        obj_b = cobranza._objetivos_ledger([_usuario(20.0)], "2099-01")
        repo.reconciliar_objetivos_batch("2099-01", "b" * 64, obj_b)
        assert repo.get_saldo("A", "1", "MULTA", "2099-01") == 30.0
        repo.reconciliar_objetivos_batch("2099-01", hash_a, objetivos)
        assert repo.get_saldo("A", "1", "MULTA", "2099-01") == 20.0

        estado_path = raiz / "estado_ciclo.json"
        estado.marcar_generado("2099-01", estado_path, hash_a)
        assert not estado.ciclo_validado("2099-01", estado_path, hash_a)
        assert estado.sellar_validado(estado_path, "2099-01", hash_a) == ["2099-01"]
        assert estado.ciclo_validado("2099-01", estado_path, hash_a)
        estado.marcar_generado("2099-01", estado_path, "c" * 64)
        assert not estado.ciclo_validado("2099-01", estado_path, hash_a)
        estado.marcar_ledger_comprometido("2099-01", "c" * 64, 3, estado_path)
        assert estado.ultimo_ledger_comprometido(estado_path) == "2099-01"

    print("OK commit único: cuenta completa, Q-5, hash estable, delta e idempotencia")


if __name__ == "__main__":
    main()
