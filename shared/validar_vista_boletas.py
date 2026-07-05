"""
shared/validar_vista_boletas.py — Candado pre-mesa: la lista consultable y las
boletas impresas deben decir EXACTAMENTE lo mismo.

Valida, por (MZ, LT), que el saldo por concepto del ledger (lo que muestran
vista_seguimiento_pueblo.xlsx/.pdf) coincida con lo impreso en la boleta
(DATA_boletas.xlsx):

    MULTA    ↔ "Multa (faena + reunión)"
    ACUERDOS ↔ "Cuota directa"
    CONVENIO ↔ "Convenio"

Además compara el NOMBRE de la boleta contra el padrón reconciliado (detecta
descalces de identidad como el caso F-3A/F-3B) e informa los predios con deuda
en el ledger que NO reciben boleta este mes (sin servicio — en mesa solo se
consultan por la vista).

Uso:  py shared/validar_vista_boletas.py [MES_CIERRE]   (default: 2026-06)
Sale con código 1 si hay discrepancias de monto.
"""

import sys
import unicodedata
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import seguimiento_repo as repo  # noqa: E402

DATA_BOLETAS_PATH = Path(__file__).parent.parent / "3_boletas" / "inputs" / "DATA_boletas.xlsx"
PADRON_PATH = Path(__file__).parent.parent / "0_padron" / "02_matching" / "outputs" / "padron_reconciliado.xlsx"

# concepto del ledger → columna de la boleta
MAPEO = {
    "MULTA": "Multa (faena + reunión)",
    "ACUERDOS": "Cuota directa",
    "CONVENIO": "Convenio",
}
TOL = 0.005


def _nombre_norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.upper().split())


def main() -> int:
    mes = sys.argv[1] if len(sys.argv) > 1 else "2026-06"
    print(f"Validando vista/ledger (cierre {mes}) vs DATA_boletas...\n")

    db = pd.read_excel(DATA_BOLETAS_PATH, sheet_name="Data")
    db["_K"] = db["MZ"].astype(str).map(repo._norm) + "|" + db["LT"].astype(str).map(repo._norm)

    errores = 0

    # ── 1. Monto por concepto: boleta == saldo del ledger ────────────────────
    for concepto, col in MAPEO.items():
        saldos = repo.get_saldos_bulk(concepto, mes)
        difs = []
        for _, r in db.iterrows():
            mz, lt = repo._norm(r["MZ"]), repo._norm(r["LT"])
            boleta = float(pd.to_numeric(r[col], errors="coerce") or 0)
            ledger = float(saldos.get((mz, lt), 0.0))
            if abs(boleta - ledger) > TOL:
                difs.append((mz, lt, str(r["NOMBRES"]), boleta, ledger))
        if difs:
            errores += len(difs)
            print(f"[FALLA] {concepto}: {len(difs)} boleta(s) no coinciden con el ledger")
            for mz, lt, nom, b, l in difs:
                print(f"    {mz}-{lt} {nom}: boleta={b} ledger={l}")
        else:
            print(f"[OK] {concepto}: {len(db)} boletas == ledger (columna \"{col}\")")

    # ── 2. Identidad: NOMBRE de la boleta vs padrón reconciliado ─────────────
    if PADRON_PATH.exists():
        pad = pd.read_excel(PADRON_PATH, dtype=str)
        nombres_pad = {
            (repo._norm(r["MZ"]), repo._norm(r["LT"])): _nombre_norm(r["Nombres"])
            for _, r in pad.iterrows() if pd.notna(r.get("MZ")) and pd.notna(r.get("LT"))
        }
        difs_nom = []
        for _, r in db.iterrows():
            k = (repo._norm(r["MZ"]), repo._norm(r["LT"]))
            n_boleta = _nombre_norm(r["NOMBRES"])
            n_pad = nombres_pad.get(k)
            if n_pad is not None and n_boleta and n_boleta != n_pad:
                difs_nom.append((k[0], k[1], str(r["NOMBRES"]), n_pad))
        if difs_nom:
            print(f"\n[AVISO] {len(difs_nom)} nombre(s) difieren entre boleta y padrón (revisar identidad):")
            for mz, lt, nb, np_ in difs_nom:
                print(f"    {mz}-{lt}: boleta='{nb}' padrón='{np_}'")
        else:
            print("\n[OK] Nombres: boleta == padrón reconciliado en los que existen en ambos")
    else:
        print("\n[AVISO] padrón reconciliado no encontrado — chequeo de nombres omitido")

    # ── 3. Deudores del ledger SIN boleta este mes (solo informativo) ────────
    claves_boletas = set(db["_K"])
    print("\nDeudores del ledger sin boleta este mes (sin servicio — consultar solo por la vista):")
    for concepto in MAPEO:
        saldos = repo.get_saldos_bulk(concepto, mes)
        sin_boleta = {k: v for k, v in saldos.items()
                      if v > TOL and f"{k[0]}|{k[1]}" not in claves_boletas}
        total = sum(sin_boleta.values())
        print(f"    {concepto}: {len(sin_boleta)} predios · S/ {total:,.2f}")

    print("\n" + ("VALIDACION FALLÓ" if errores else "VALIDACION OK — vista y boletas dicen lo mismo"))
    return 1 if errores else 0


if __name__ == "__main__":
    sys.exit(main())
