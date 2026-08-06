"""
4b_reclamos/reporte_deuda_ledger.py — Todos los predios que TODAVIA deben algo en
el ledger de pueblo (shared/seguimiento_pueblo.xlsx), en cualquiera de los tres
conceptos: MULTA · ACUERDOS · CONVENIO.

Diferencia con reporte_convenio_multa.py: ese arranca de CONVENIO pendiente y
responde una pregunta puntual (¿la multa ya pagada cubriria el convenio?). Este
no filtra por concepto -- es el padron completo de deuda viva, para decidir caso
por caso con el historial y la referencia de pago de cada uno a la vista.

Estructura del PDF:
  portada (varias paginas)  tabla de los N deudores, ordenada por deuda total
  1 pagina por predio       historial mes a mes + "de donde vino cada pago"

Uso: py 4b_reclamos/reporte_deuda_ledger.py [MES_ANO]
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
import reporte_convenio_multa as rcm  # noqa: E402
import reporte_referencias_pago as rrp  # noqa: E402

CONCEPTOS = ("MULTA", "ACUERDOS", "CONVENIO")

_AZUL, _AZUL_BG = rcm._AZUL, rcm._AZUL_BG
_GRIS, _NEGRO, _ROJO, _ZEBRA = rcm._GRIS, rcm._NEGRO, rcm._ROJO, rcm._ZEBRA
_PAGE_W, _PAGE_H, _M = rcm._PAGE_W, rcm._PAGE_H, rcm._M


def calcular_tabla(mes_ano: str = "2026-07") -> pd.DataFrame:
    """Union de los predios con saldo > 0 en cualquier concepto. El saldo sale
    del ledger (get_saldos_bulk), que ya aplica CARGO/PAGO/AJUSTE en orden."""
    nombres = repo._lookup_nombres()
    saldos = {c: repo.get_saldos_bulk(c, mes_ano) for c in CONCEPTOS}

    predios = set()
    for c in CONCEPTOS:
        for k, v in saldos[c].items():
            # "NAN" = fila de totales de la siembra, no es un predio real
            if v > 0.005 and k[0] != "NAN" and k[1] != "NAN":
                predios.add(k)

    filas = []
    for mz, lt in predios:
        fila = {"MZ": mz, "LT": lt, "NOMBRE": nombres.get((mz, lt), "")}
        for c in CONCEPTOS:
            fila[c] = round(max(0.0, saldos[c].get((mz, lt), 0.0)), 2)
        fila["TOTAL"] = round(sum(fila[c] for c in CONCEPTOS), 2)
        filas.append(fila)

    df = pd.DataFrame(filas)
    return df.sort_values(["TOTAL", "MZ", "LT"], ascending=[False, True, True]).reset_index(drop=True)


def _dibujar_portada(doc, df: pd.DataFrame, mes_ano: str) -> None:
    headers = ["Predio", "Nombre", "Multa", "Acuerdos", "Convenio", "Deuda total"]
    anchos = [55, 330, 95, 95, 95, 112]
    rh_row = 16
    filas_por_pag = int((_PAGE_H - _M * 2 - 95) // rh_row)
    n_paginas = max(-(-len(df) // filas_por_pag), 1)

    for pagina in range(n_paginas):
        page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
        w = _PAGE_W - 2 * _M
        y = _M

        page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
        page.insert_text((_M + 10, y + 20),
                          f"Deuda viva en el ledger de pueblo — {mes_ano} "
                          f"(pág. {pagina + 1}/{n_paginas})",
                          fontsize=12, fontname="hebo", color=(1, 1, 1))
        y += 42

        if pagina == 0:
            for c in CONCEPTOS:
                n = int((df[c] > 0.005).sum())
                page.insert_text((_M, y), f"{c}: {n} predios · S/ {df[c].sum():,.2f}",
                                  fontsize=9, fontname="helv", color=_NEGRO)
                y += 13
            page.insert_text((_M, y),
                              f"{len(df)} predios con deuda en al menos un concepto · "
                              f"S/ {df['TOTAL'].sum():,.2f} en total. Un predio puede deber en varios.",
                              fontsize=9, fontname="hebo", color=_NEGRO)
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
            page.insert_text((cx + 4, y + rh_row - 5), f"{r['MZ']}-{r['LT']}",
                              fontsize=8, fontname="hebo", color=_NEGRO)
            cx += anchos[0]
            page.insert_text((cx + 4, y + rh_row - 5), str(r["NOMBRE"])[:52],
                              fontsize=8, fontname="helv", color=_NEGRO)
            cx += anchos[1]
            for c, cw in zip(CONCEPTOS, anchos[2:5]):
                texto = f"S/ {r[c]:,.2f}" if r[c] > 0.005 else "—"
                page.insert_text((cx + 4, y + rh_row - 5), texto, fontsize=8, fontname="helv",
                                  color=_NEGRO if r[c] > 0.005 else _GRIS)
                cx += cw
            page.insert_text((cx + 4, y + rh_row - 5), f"S/ {r['TOTAL']:,.2f}",
                              fontsize=8, fontname="hebo", color=_ROJO)
            y += rh_row


def generar(mes_ano: str = "2026-07", salida: Path | None = None) -> Path:
    df = calcular_tabla(mes_ano)

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()
    redirects = rcm._cargar_redirects()

    doc = fitz.open()
    _dibujar_portada(doc, df, mes_ano)
    for n, (_, r) in enumerate(df.iterrows(), 1):
        mz, lt = r["MZ"], r["LT"]
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        refs = rrp.referencias_pago(mz, lt, tabla=tabla)
        rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
        page = doc[-1]
        w = rh._PAGE_W - 2 * rh._M
        y_tabla_fin = rh._M + 42 + 14 + 18 + (18 * len(tabla))
        rrp._dibujar_tabla_referencias(page, rh._M, y_tabla_fin + 45, w, refs)
        if n % 25 == 0:
            print(f"  {n}/{len(df)} predios...")

    salida = salida or (BASE_DIR / "outputs" / f"reporte_deuda_ledger_{mes_ano}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"{len(df)} predios con deuda · S/ {df['TOTAL'].sum():,.2f}")
    for c in CONCEPTOS:
        print(f"   {c:<9} {int((df[c] > 0.005).sum()):>4} predios · S/ {df[c].sum():>10,.2f}")
    print(f"PDF -> {salida}")
    return salida


if __name__ == "__main__":
    mes = sys.argv[1] if len(sys.argv) > 1 else "2026-07"
    generar(mes, BASE_DIR / "outputs" / f"reporte_convenio_multa_referencias_{mes}.pdf")
