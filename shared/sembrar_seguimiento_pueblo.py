"""
shared/sembrar_seguimiento_pueblo.py — DESECHABLE, uso único (o por tanda nueva de altas)

Siembra la deuda inicial de MULTA, ACUERDOS y CONVENIO como el primer evento CARGO
de cada predio en shared/seguimiento_pueblo.xlsx (vía seguimiento_repo).

Reglas de mapeo — ver docs/diagrama_siembra_seguimiento_pueblo.html:
    MULTA     ← 3_boletas/inputs/DATA_boletas.xlsx, columna "Multa (faena + reunión)"
    ACUERDOS  ← 3_boletas/inputs/DATA_boletas.xlsx, columna "Cuota directa"
    CONVENIO  ← shared/genesis_inputs/medidor_saldo.xlsx (Saldo) +
                shared/genesis_inputs/inscripcion_saldo.xlsx (SALDO), suma por predio.
                Excluye los 6 predios de instalación ya completos en el arrastre
                (B-20, C-43, C-35, F1-11, G-21, W-2).

Esquema de inputs y qué pasa si falta un archivo/columna: shared/README.md
("Esquema de inputs — sembrar_seguimiento_pueblo.py").

Idempotente: cada evento usa audit_ref único por (concepto, mz, lt) — re-correr el
script no duplica cargos ya sembrados.
"""

import logging
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import seguimiento_repo as repo

log = logging.getLogger(__name__)

# ── Rutas ─────────────────────────────────────────────────────────────────────
_SHARED = Path(__file__).parent
_BASE   = _SHARED.parent

DATA_BOLETAS_PATH   = _BASE / "3_boletas" / "inputs" / "DATA_boletas.xlsx"
MEDIDOR_SALDO_PATH  = _SHARED / "genesis_inputs" / "medidor_saldo.xlsx"
INSCRIPCION_PATH    = _SHARED / "genesis_inputs" / "inscripcion_saldo.xlsx"
LOG_PATH            = _SHARED / "sembrar_seguimiento_pueblo_run.log"

MES_SIEMBRA = "2026-06"  # último ciclo cerrado — 5_cobranza sigue procesando
                          # 2026-06 (julio bloqueado por D3, falta lecturas).
                          # Debe fechar <= el mes de los PAGO que reconciliará
                          # 5_cobranza, o get_saldo("MES<=X") no ve el CARGO.
SOURCE      = "sembrar_seguimiento_pueblo"

# Predios ya completos en arrastre_consolidado — no re-sembrar (fuente única
# en seguimiento_repo, también la usa 5_cobranza para no reconciliarles pagos
# de CONVENIO). Verificado 2026-07-02: ninguno solapa hoy con medidor/inscripción.
PREDIOS_INSTALACION_EXCLUIDOS = repo.PREDIOS_INSTALACION_EXCLUIDOS


# str(nan) = 'NAN' es truthy y pasaba el filtro `if mz and lt` — las filas de
# TOTALES de las hojas de génesis (MZ/LT vacíos) se sembraban como predio
# fantasma NAN-NAN (visto 2026-07-04: cargo CONVENIO de S/1,334).
_INVALIDOS = {"", "NAN", "NONE", "NAT"}


def _norm_mz(v) -> str:
    s = str(v).strip().upper()
    return "" if s in _INVALIDOS else s


def _norm_lt(v) -> str:
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except ValueError:
        pass
    s = s.upper()
    return "" if s in _INVALIDOS else s


# ── Validación de inputs ─────────────────────────────────────────────────────

def _validar_inputs() -> None:
    if not DATA_BOLETAS_PATH.exists():
        raise FileNotFoundError(f"Falta: {DATA_BOLETAS_PATH}")
    if not MEDIDOR_SALDO_PATH.exists():
        raise FileNotFoundError(f"Falta: {MEDIDOR_SALDO_PATH}")
    if not INSCRIPCION_PATH.exists():
        raise FileNotFoundError(f"Falta: {INSCRIPCION_PATH}")

    df = pd.read_excel(DATA_BOLETAS_PATH, sheet_name="Data", nrows=0)
    requeridas = {"MZ", "LT", "Multa (faena + reunión)", "Cuota directa"}
    faltantes = requeridas - set(df.columns)
    if faltantes:
        raise ValueError(f"DATA_boletas.xlsx: columnas faltantes {faltantes}")

    dfm = pd.read_excel(MEDIDOR_SALDO_PATH, sheet_name="Cobro medidores", nrows=0)
    faltantes_m = {"MZ", "LT", "Saldo"} - set(dfm.columns)
    if faltantes_m:
        raise ValueError(f"medidor_saldo.xlsx: columnas faltantes {faltantes_m}")

    dfi = pd.read_excel(INSCRIPCION_PATH, sheet_name="NUEVAS INSTALACIONES", nrows=0)
    faltantes_i = {"MZ", "LT", "SALDO"} - set(dfi.columns)
    if faltantes_i:
        raise ValueError(f"inscripcion_saldo.xlsx: columnas faltantes {faltantes_i}")


# ── Siembra por concepto ─────────────────────────────────────────────────────

def _sembrar_multa_acuerdos() -> tuple[int, int]:
    df = pd.read_excel(DATA_BOLETAS_PATH, sheet_name="Data")
    df["MZ"] = df["MZ"].map(_norm_mz)
    df["LT"] = df["LT"].map(_norm_lt)
    df["Multa (faena + reunión)"] = pd.to_numeric(df["Multa (faena + reunión)"], errors="coerce")
    df["Cuota directa"] = pd.to_numeric(df["Cuota directa"], errors="coerce")

    n_multa = n_acuerdos = 0
    for _, fila in df.iterrows():
        mz, lt = fila["MZ"], fila["LT"]
        if not mz or not lt:
            continue

        multa = fila["Multa (faena + reunión)"]
        if pd.notna(multa) and multa > 0:
            r = repo.registrar_cargo(
                mz, lt, "MULTA", MES_SIEMBRA, float(multa),
                source=SOURCE, audit_ref=f"siembra_{MES_SIEMBRA}_MULTA_{mz}_{lt}",
            )
            if not r["skipped"]:
                n_multa += 1

        acuerdos = fila["Cuota directa"]
        if pd.notna(acuerdos) and acuerdos > 0:
            r = repo.registrar_cargo(
                mz, lt, "ACUERDOS", MES_SIEMBRA, float(acuerdos),
                source=SOURCE, audit_ref=f"siembra_{MES_SIEMBRA}_ACUERDOS_{mz}_{lt}",
            )
            if not r["skipped"]:
                n_acuerdos += 1

    return n_multa, n_acuerdos


def _sembrar_convenio() -> int:
    dfm = pd.read_excel(MEDIDOR_SALDO_PATH, sheet_name="Cobro medidores")
    dfm["MZ"] = dfm["MZ"].map(_norm_mz)
    dfm["LT"] = dfm["LT"].map(_norm_lt)
    dfm["Saldo"] = pd.to_numeric(dfm["Saldo"], errors="coerce").fillna(0)
    saldo_medidor = {
        (r["MZ"], r["LT"]): float(r["Saldo"])
        for _, r in dfm.iterrows() if r["MZ"] and r["LT"] and r["Saldo"] > 0
    }

    dfi = pd.read_excel(INSCRIPCION_PATH, sheet_name="NUEVAS INSTALACIONES")
    dfi["MZ"] = dfi["MZ"].map(_norm_mz)
    dfi["LT"] = dfi["LT"].map(_norm_lt)
    dfi["SALDO"] = pd.to_numeric(dfi["SALDO"], errors="coerce").fillna(0)
    saldo_inscripcion = {
        (r["MZ"], r["LT"]): float(r["SALDO"])
        for _, r in dfi.iterrows() if r["MZ"] and r["LT"] and r["SALDO"] > 0
    }

    poblacion = set(saldo_medidor) | set(saldo_inscripcion)
    poblacion -= PREDIOS_INSTALACION_EXCLUIDOS

    n = 0
    for mz, lt in sorted(poblacion):
        monto = saldo_medidor.get((mz, lt), 0.0) + saldo_inscripcion.get((mz, lt), 0.0)
        if monto <= 0:
            continue
        r = repo.registrar_cargo(
            mz, lt, "CONVENIO", MES_SIEMBRA, monto,
            source=SOURCE, audit_ref=f"siembra_{MES_SIEMBRA}_CONVENIO_{mz}_{lt}",
        )
        if not r["skipped"]:
            n += 1

    return n


# ── Punto de entrada ──────────────────────────────────────────────────────────

def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
        ],
        force=True,
    )
    log.info(f"Iniciando siembra — mes={MES_SIEMBRA}")

    _validar_inputs()

    n_multa, n_acuerdos = _sembrar_multa_acuerdos()
    log.info(f"MULTA: {n_multa} eventos CARGO nuevos")
    log.info(f"ACUERDOS: {n_acuerdos} eventos CARGO nuevos")

    n_convenio = _sembrar_convenio()
    log.info(f"CONVENIO: {n_convenio} eventos CARGO nuevos")

    log.info("Siembra completa.")


if __name__ == "__main__":
    main()
