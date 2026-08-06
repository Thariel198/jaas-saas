"""
4b_reclamos/agregar_lote_al_reporte_convenio.py — Agrega al reporte oficial
`reporte_convenio_multa_referencias_2026-07.pdf` los 10 predios del lote de
SALDO negativo (investigacion 2026-08-01, ver
3_boletas/inputs/reclamos_2026-08-01/CONSOLIDADO.md Bloque B) que hoy no
aparecen porque su CONVENIO en el ledger sigue negativo (el bug no se
corrigio en seguimiento_pueblo.xlsx todavia, solo en DATA_boletas.xlsx).
Se les fuerza el SALDO_CONVENIO correcto (verificado evento por evento,
antes del pago fantasma de julio) para que salgan con su deuda real.

No toca el ledger -- overlay solo para este reporte, igual que ya se hizo
con DATA_boletas.xlsx.

Uso: py 4b_reclamos/agregar_lote_al_reporte_convenio.py
"""

import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR.parent / "shared"))
sys.path.insert(0, str(BASE_DIR))
import seguimiento_repo as repo  # noqa: E402
import reporte_historico as rh  # noqa: E402
import reporte_convenio_multa as rcm  # noqa: E402
import reporte_referencias_pago as rrp  # noqa: E402

MES_ANO = "2026-07"

# (MZ, LT) -> SALDO_CONVENIO real, verificado con eventos crudos antes del
# pago fantasma de julio (mismo valor que CONSOLIDADO.md Bloque B).
LOTE_CONVENIO = {
    ("A", "8"): 50,
    ("B", "5"): 50,
    ("C", "1"): 50,
    ("C", "7"): 25,
    ("E", "12"): 26,
    ("I", "11"): 25,
    ("J", "3"): 50,
    ("K", "17"): 25,
    ("K", "2"): 25,
    ("P", "12"): 50,
}

# ── Deteccion automatica de pago fantasma de julio ───────────────────────────
# seguimiento_pueblo.xlsx todavia tiene PAGOs de julio sin respaldo real
# (investigacion 2026-08-01). tabla_predio() los mostraria como pagados. En vez
# de una lista fija (se quedaba corta: L-5 y W-5 no estaban), se calcula por
# predio cuanta plata REAL entro en el ciclo y se recorta el excedente.
#
#   plata_real = MONTO_YAPE + MONTO_EFECTIVO (planilla_cobrado)
#              + abonos_rezagados + blancos_efectivo + devoluciones_aplicadas
#              - aportes_tanque_manuales
#   Los overlays son plata real que NO aparece en planilla_cobrado -- sin
#   sumarlos se borrarian pagos legitimos (D-16, D1-3, D1-6, L-4, P-6).
#
#   disponible_pueblo = max(0, plata_real - agua - mant - mes_ant - corte)
#   Si la tabla muestra mas que eso en MULTA/ACUERDOS/CONVENIO, el exceso es
#   fantasma y se recorta en orden inverso a la cascada (P5 CONVENIO primero,
#   luego P4 ACUERDOS, luego P3 MULTA): lo ultimo que la cascada habria pagado
#   es lo primero que se queda sin respaldo.
_TOL = 0.005
_ORDEN_RECORTE = ("CONVENIO", "ACUERDOS", "MULTA")
_cache_disponible: dict | None = None


def _num(v) -> float:
    n = pd.to_numeric(v, errors="coerce")
    return 0.0 if pd.isna(n) else float(n)


def _clave(mz, lt) -> tuple:
    def n(v):
        s = str(v).strip()
        return (s[:-2] if s.endswith(".0") else s).upper()
    return (n(mz), n(lt))


def _disponible_pueblo() -> dict:
    """{(mz,lt): soles que de verdad podian llegar a multa/acuerdos/convenio}"""
    global _cache_disponible
    if _cache_disponible is not None:
        return _cache_disponible

    extra: dict = {}
    for nombre, signo in (("abonos_rezagados.xlsx", 1), ("blancos_efectivo.xlsx", 1),
                          ("devoluciones_aplicadas.xlsx", 1),
                          ("aportes_tanque_manuales.xlsx", -1)):
        p = BASE_DIR.parent / "shared" / nombre
        if not p.exists():
            continue
        d = pd.read_excel(p, header=1)
        if "MES_ANO_APLICA" not in d.columns or "MONTO" not in d.columns:
            continue
        d = d[d["MES_ANO_APLICA"].astype(str).str.strip() == MES_ANO]
        for _, r in d.iterrows():
            k = _clave(r.get("MZ"), r.get("LT"))
            extra[k] = extra.get(k, 0.0) + signo * _num(r.get("MONTO"))

    dfp = pd.read_excel(BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx",
                        sheet_name="planilla_cobrado", header=1)
    out: dict = {}
    for _, r in dfp.iterrows():
        k = _clave(r.get("MZ"), r.get("LT"))
        real = _num(r.get("MONTO_YAPE")) + _num(r.get("MONTO_EFECTIVO")) + extra.get(k, 0.0)
        agua = (_num(r.get("MES_ACTUAL")) + _num(r.get("MANTENIMIENTO")) +
                _num(r.get("MES_ANTERIOR")) + _num(r.get("CORTE_RECONEXION")))
        out[k] = max(0.0, round(real - agua, 2))
    _cache_disponible = out
    return out


def quitar_pago_fantasma_julio(tabla: pd.DataFrame, mz: str, lt: str) -> pd.DataFrame:
    idx = tabla.index[tabla["MES"] == MES_ANO]
    if len(idx) == 0:
        return tabla
    disponible = _disponible_pueblo().get(_clave(mz, lt))
    if disponible is None:
        return tabla  # sin fila en planilla_cobrado: no se puede afirmar nada

    tabla = tabla.copy()
    for i in idx:
        mostrado = sum(_num(tabla.loc[i, c]) for c in _ORDEN_RECORTE)
        exceso = round(mostrado - disponible, 2)
        if exceso <= _TOL:
            continue
        for campo in _ORDEN_RECORTE:
            if exceso <= _TOL:
                break
            actual = _num(tabla.loc[i, campo])
            quita = min(actual, exceso)
            if quita > _TOL:
                tabla.loc[i, campo] = round(actual - quita, 2)
                exceso = round(exceso - quita, 2)
        tabla.loc[i, "TOTAL"] = float(sum(_num(tabla.loc[i, c]) for c in rh.CONCEPTOS_TABLA))
        tabla.loc[i, "PAGO_COMPLETO"] = tabla.loc[i, "TOTAL"] > _TOL
    return tabla


# Bienes del pueblo con ESTADO=EXONERADO en registro_cortes.xlsx (ver
# LEER_ANTES.md) -- no deberian salir en reportes de cobranza. Sacado a
# pedido 01/08/2026.
EXCLUIDOS = {("J", "6")}


def calcular_tabla_con_lote(mes_ano: str) -> pd.DataFrame:
    df = rcm.calcular_tabla(mes_ano)
    df = df[~df.apply(lambda r: (r["MZ"], r["LT"]) in EXCLUIDOS, axis=1)].reset_index(drop=True)
    ya_presentes = set(zip(df["MZ"], df["LT"]))

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()
    redirects = rcm._cargar_redirects()

    filas_extra = []
    for (mz, lt), saldo in LOTE_CONVENIO.items():
        if (mz, lt) in ya_presentes:
            continue
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        mp = float(tabla["MULTA"].sum())
        filas_extra.append({
            "MZ": mz, "LT": lt, "NOMBRE": nombres.get((mz, lt), ""),
            "CONVENIO_SALDO": saldo, "MULTA_PAGO": mp,
            "CUBRIRIA": mp >= saldo - 0.005,
            "DIFERENCIA": round(mp - saldo, 2),
        })

    if filas_extra:
        df = pd.concat([df, pd.DataFrame(filas_extra)], ignore_index=True)
        df["_orden"] = df["CUBRIRIA"].map({True: 0, False: 1}) + (df["MULTA_PAGO"] <= 0.005).astype(int)
        df = df.sort_values(["_orden", "CONVENIO_SALDO"], ascending=[True, False]).drop(columns="_orden")
        df = df.reset_index(drop=True)
    return df


def generar(salida: Path | None = None) -> Path:
    df = calcular_tabla_con_lote(MES_ANO)

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()
    redirects = rcm._cargar_redirects()

    import fitz
    doc = fitz.open()
    rcm._dibujar_portada(doc, df, MES_ANO)
    for _, r in df.iterrows():
        mz, lt = r["MZ"], r["LT"]
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        tabla = quitar_pago_fantasma_julio(tabla, mz, lt)
        refs = rrp.referencias_pago(mz, lt, tabla=tabla)
        rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
        page = doc[-1]
        w = rh._PAGE_W - 2 * rh._M
        y_tabla_fin = rh._M + 42 + 14 + 18 + (18 * len(tabla))
        rrp._dibujar_tabla_referencias(page, rh._M, y_tabla_fin + 45, w, refs)

    salida = salida or (BASE_DIR / "outputs" / f"reporte_convenio_multa_referencias_{MES_ANO}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"{len(df)} predios -> PDF {salida} ({len(LOTE_CONVENIO)} del lote de saldo negativo agregados)")
    return salida


if __name__ == "__main__":
    generar()
