from pathlib import Path
import sys

import pandas as pd
import pytest


SHARED = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SHARED))
import abonos_rezagados_repo as repo  # noqa: E402


def _guardar(path: Path, rows: list[dict]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="Abonos_Raw", startrow=1, index=False)


def _fila(**cambios) -> dict:
    row = {
        "ID_ABONO": "corr-a-1-multa", "MZ": "A", "LT": "1", "MONTO": 10,
        "MES_CICLO": "2026-06", "MES_ANO_APLICA": "2026-07",
        "ESTADO": "CONFIRMADO", "MODO_APLICACION": "DIRIGIDO",
        "CONCEPTO_DESTINO": "MULTA", "APLICADO": "NO",
    }
    row.update(cambios)
    return row


def test_solo_devuelve_confirmados(tmp_path):
    path = tmp_path / "abonos.xlsx"
    _guardar(path, [_fila(), _fila(ID_ABONO="hist-1", ESTADO="DESCARTADO")])

    df = repo.leer_abonos(path)

    assert df["ID_ABONO"].tolist() == ["corr-a-1-multa"]


def test_rechaza_confirmado_duplicado(tmp_path):
    path = tmp_path / "abonos.xlsx"
    _guardar(path, [_fila(), _fila()])

    with pytest.raises(RuntimeError, match="ID_ABONO duplicado"):
        repo.leer_abonos(path)


def test_rechaza_aplicado_invalido(tmp_path):
    path = tmp_path / "abonos.xlsx"
    _guardar(path, [_fila(APLICADO="TAL VEZ")])

    with pytest.raises(RuntimeError, match="APLICADO invalido"):
        repo.leer_abonos(path)


def test_rechaza_destino_invalido(tmp_path):
    path = tmp_path / "abonos.xlsx"
    _guardar(path, [_fila(CONCEPTO_DESTINO="INVENTADO")])

    with pytest.raises(RuntimeError, match="CONCEPTO_DESTINO invalido"):
        repo.leer_abonos(path)
