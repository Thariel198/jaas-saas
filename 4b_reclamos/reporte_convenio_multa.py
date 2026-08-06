"""
4b_reclamos/reporte_convenio_multa.py — Auditoria: predios con deuda de CONVENIO
(medidor) pendiente, cruzada contra lo que ya pagaron de MULTA. La cascada
aplica el pago primero a MULTA (prioridad mas alta) y despues a CONVENIO; muchos
vecinos pagaron pensando que cubrian convenio, pero la cascada lo aplico a
multa primero. Este reporte responde: si se redirige ese pago de MULTA hacia
CONVENIO, ¿a cuantos les queda el convenio saldado?

Reusa tabla_predio()/_dibujar_pagina() de reporte_historico.py para la pagina
de historial de cada predio -- agrega una portada resumen antes.

Uso: py 4b_reclamos/reporte_convenio_multa.py
"""

import sys
from pathlib import Path

import fitz
import pandas as pd

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR.parent / "shared"))
sys.path.insert(0, str(BASE_DIR))
import seguimiento_repo as repo  # noqa: E402
import reporte_historico as rh  # noqa: E402

REASIGNACIONES_PATH = BASE_DIR.parent / "shared" / "reasignaciones_aplicacion.xlsx"

_CAMPO_TABLA = {"MULTA": "MULTA", "ACUERDOS": "ACUERDOS", "CONVENIO": "CONVENIO"}


def _cargar_redirects() -> pd.DataFrame:
    if not REASIGNACIONES_PATH.exists():
        return pd.DataFrame(columns=["MZ", "LT", "CONCEPTO_ORIGEN", "CONCEPTO_DESTINO", "MES_ANO", "MONTO"])
    df = pd.read_excel(REASIGNACIONES_PATH, header=1)
    return df[["MZ", "LT", "CONCEPTO_ORIGEN", "CONCEPTO_DESTINO", "MES_ANO", "MONTO"]]


def corregir_tabla_por_redirects(tabla: pd.DataFrame, mz: str, lt: str, redirects: pd.DataFrame) -> pd.DataFrame:
    """Un redirect (ver reasignaciones_aplicacion.xlsx) mueve un PAGO ya
    registrado de un concepto a otro dentro del MISMO mes -- pero el PAGO
    original queda intacto en el ledger (append-only, se estabiliza con un
    AJUSTE parejo que la tabla no muestra). Sin esta correccion, el reporte
    cuenta esa plata dos veces (una en el concepto origen, otra en el destino)
    y TOTAL PAGADO sale inflado. Se resta del origen -- el destino ya trae el
    monto correcto porque su PAGO nuevo SI se escribio en el ledger."""
    mios = redirects[(redirects["MZ"] == mz) & (redirects["LT"].astype(str) == str(lt))]
    if mios.empty:
        return tabla
    tabla = tabla.copy()
    for _, r in mios.iterrows():
        campo = _CAMPO_TABLA.get(str(r["CONCEPTO_ORIGEN"]).strip().upper())
        mes = str(r["MES_ANO"]).strip()
        monto = float(r["MONTO"])
        if not campo:
            continue
        idx = tabla.index[tabla["MES"] == mes]
        for i in idx:
            tabla.loc[i, campo] = max(0.0, tabla.loc[i, campo] - monto)
        tabla.loc[idx, "TOTAL"] = tabla.loc[idx, rh.CONCEPTOS_TABLA].sum(axis=1)
    return tabla


_AZUL = (26/255, 82/255, 118/255)
_AZUL_BG = (235/255, 245/255, 251/255)
_GRIS = (0.42, 0.45, 0.5)
_NEGRO = (0.12, 0.16, 0.22)
_VERDE = (0.02, 0.37, 0.27)
_ROJO = (0.55, 0.13, 0.13)
_ZEBRA = (243/255, 244/255, 246/255)
_PAGE_W, _PAGE_H = 842, 595
_M = 30


def calcular_tabla(mes_ano: str = "2026-07") -> pd.DataFrame:
    """MULTA_PAGO suma el historico COMPLETO (oct 2025 -> jul 2026), no solo el
    ledger (seguimiento_pueblo arranca en junio 2026) -- reusa tabla_predio(),
    que ya combina ambas fuentes, igual que reporte_historico_CONFIRMACION."""
    conv_saldo = repo.get_saldos_bulk("CONVENIO", mes_ano)
    nombres = repo._lookup_nombres()

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)

    redirects = _cargar_redirects()

    filas = []
    for (mz, lt), saldo in conv_saldo.items():
        if saldo <= 0.005:
            continue
        if mz == "NAN" or lt == "NAN":
            continue  # fila fantasma de siembra (totales), no es un predio real
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        mp = float(tabla["MULTA"].sum())
        filas.append({
            "MZ": mz, "LT": lt, "NOMBRE": nombres.get((mz, lt), ""),
            "CONVENIO_SALDO": saldo, "MULTA_PAGO": mp,
            "CUBRIRIA": mp >= saldo - 0.005,
            "DIFERENCIA": round(mp - saldo, 2),
        })

    df = pd.DataFrame(filas)
    df["_orden"] = df["CUBRIRIA"].map({True: 0, False: 1}) + (df["MULTA_PAGO"] <= 0.005).astype(int)
    df = df.sort_values(["_orden", "CONVENIO_SALDO"], ascending=[True, False]).drop(columns="_orden")
    return df.reset_index(drop=True)


def _dibujar_portada(doc, df: pd.DataFrame, mes_ano: str) -> None:
    total = len(df)
    con_multa = int((df["MULTA_PAGO"] > 0.005).sum())
    cubriria = int(df["CUBRIRIA"].sum())

    headers = ["Predio", "Nombre", "Convenio (debe)", "Multa (ya pagó)", "¿Cubriría convenio?", "Diferencia"]
    anchos = [55, 260, 100, 100, 130, 100]
    rh_row = 16
    filas_por_pag = int((_PAGE_H - _M * 2 - 90) // rh_row)

    n_paginas = -(-total // filas_por_pag)
    for pagina in range(max(n_paginas, 1)):
        page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
        w = _PAGE_W - 2 * _M
        y = _M

        page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
        page.insert_text((_M + 10, y + 20),
                          f"Convenio pendiente vs. Multa ya pagada — {mes_ano} (pág. {pagina + 1}/{max(n_paginas,1)})",
                          fontsize=12, fontname="hebo", color=(1, 1, 1))
        y += 42

        if pagina == 0:
            resumen = (f"{total} predios con CONVENIO (medidor) pendiente. De esos, {con_multa} ya "
                       f"pagaron algo de MULTA. Si ese pago de multa se redirige a convenio, "
                       f"{cubriria} de {total} quedarían con el convenio saldado por completo.")
            page.insert_text((_M, y), resumen, fontsize=9, fontname="hebo", color=_NEGRO)
            y += 20

        page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_AZUL_BG, color=None)
        cx = _M
        for h, cw in zip(headers, anchos):
            page.insert_text((cx + 4, y + rh_row - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
            cx += cw
        y += rh_row

        bloque = df.iloc[pagina * filas_por_pag:(pagina + 1) * filas_por_pag]
        for n, (_, r) in enumerate(bloque.iterrows()):
            if n % 2 == 1:
                page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_ZEBRA, color=None)
            cx = _M
            page.insert_text((cx + 4, y + rh_row - 5), f"{r['MZ']}-{r['LT']}", fontsize=8, fontname="hebo", color=_NEGRO)
            cx += anchos[0]
            page.insert_text((cx + 4, y + rh_row - 5), str(r["NOMBRE"])[:40], fontsize=8, fontname="helv", color=_NEGRO)
            cx += anchos[1]
            page.insert_text((cx + 4, y + rh_row - 5), f"S/ {r['CONVENIO_SALDO']:,.2f}", fontsize=8, fontname="helv", color=_NEGRO)
            cx += anchos[2]
            texto_mp = f"S/ {r['MULTA_PAGO']:,.2f}" if r["MULTA_PAGO"] > 0.005 else "—"
            page.insert_text((cx + 4, y + rh_row - 5), texto_mp, fontsize=8, fontname="helv", color=_NEGRO)
            cx += anchos[3]
            texto_cub = "SI, completo" if r["CUBRIRIA"] else ("parcial" if r["MULTA_PAGO"] > 0.005 else "no pagó multa")
            color_cub = _VERDE if r["CUBRIRIA"] else _GRIS
            page.insert_text((cx + 4, y + rh_row - 5), texto_cub, fontsize=8, fontname="hebo", color=color_cub)
            cx += anchos[4]
            signo = "+" if r["DIFERENCIA"] >= 0 else ""
            color_dif = _VERDE if r["DIFERENCIA"] >= 0 else _ROJO
            page.insert_text((cx + 4, y + rh_row - 5), f"{signo}{r['DIFERENCIA']:,.2f}", fontsize=8, fontname="helv", color=color_dif)
            y += rh_row


def generar(mes_ano: str = "2026-07", salida: Path | None = None) -> Path:
    df = calcular_tabla(mes_ano)

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()

    redirects = _cargar_redirects()

    doc = fitz.open()
    _dibujar_portada(doc, df, mes_ano)
    for _, r in df.iterrows():
        mz, lt = r["MZ"], r["LT"]
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)

    salida = salida or (BASE_DIR / "outputs" / f"reporte_convenio_multa_{mes_ano}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"{len(df)} predios con convenio pendiente | {int(df['CUBRIRIA'].sum())} quedarían saldados si se redirige la multa")
    print(f"PDF -> {salida} ({doc.page_count if False else 1 + len(df)} páginas aprox.)")
    return salida


if __name__ == "__main__":
    generar()
