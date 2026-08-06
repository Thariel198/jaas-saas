"""
4b_reclamos/reporte_lote_saldo_negativo.py — Reporte ad-hoc para el lote de 11
predios que quedaron con SALDO negativo en seguimiento_pueblo.xlsx tras el
recon de 5_cobranza del 31/07/2026 (investigacion 2026-08-01, no es el bug de
carrera del 13/07 -- ver docs/RETOMAR_fix_race_condition_yape_seguimiento_pueblo_2026-07-27.md,
descartado con evidencia).

Portada: para cada predio, el PAGO que el ledger registro el 06/07, el AJUSTE
negativo del 31/07 que lo revirtio, y el saldo actual -- para decidir la
correccion manual (mismo patron que F-12/D-1 en RETOMAR_limpieza_ledger...).

Por pagina: reusa tabla_predio() (reporte_historico) + referencias_pago()
(reporte_referencias_pago) -- la tabla de historial muestra lo que el LEDGER
dice que se pago (puede incluir el credito sin respaldo); la tabla de
"Referencia de pago" muestra la plata real encontrada en pagos_yape_tepago.xlsx
/ pagos_efectivo.xlsx / trazabilidad -- ahi es donde se ve, sin el bug, lo que
de verdad pago cada uno.

Uso: py 4b_reclamos/reporte_lote_saldo_negativo.py
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

MES_ANO = "2026-07"

# (MZ, LT, CONCEPTO, PAGO registrado 06/07, AJUSTE 31/07, SALDO actual)
LOTE = [
    ("A", "8",  "CONVENIO", 50, -50, -50),
    ("B", "5",  "ACUERDOS", 25, -25, -25),
    ("B", "5",  "CONVENIO", 50, -50, -50),
    ("C", "1",  "ACUERDOS", 25, -25, -25),
    ("C", "1",  "CONVENIO", 50, -50, -50),
    ("C", "7",  "CONVENIO", 25, -25, -25),
    ("E", "12", "CONVENIO", 21, -21, -16),
    ("I", "11", "CONVENIO", 25, -25, -25),
    ("I", "16", "MULTA",    18, -18, -18),
    ("I", "16", "ACUERDOS", 14, -14,  47),
    ("J", "3",  "CONVENIO", 40, -40, -30),
    ("K", "17", "CONVENIO", 25, -25, -25),
    ("K", "2",  "CONVENIO", 25, -25, -25),
    ("P", "12", "CONVENIO", 50, -50, -50),
]

_AZUL = (26 / 255, 82 / 255, 118 / 255)
_AZUL_BG = (235 / 255, 245 / 255, 251 / 255)
_GRIS = (0.42, 0.45, 0.5)
_NEGRO = (0.12, 0.16, 0.22)
_ROJO = (0.55, 0.13, 0.13)
_ZEBRA = (243 / 255, 244 / 255, 246 / 255)
_PAGE_W, _PAGE_H = 842, 595
_M = 30


def _dibujar_portada(doc, nombres: dict) -> None:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M

    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text((_M + 10, y + 20),
                      f"Lote con SALDO negativo en seguimiento_pueblo — {MES_ANO}",
                      fontsize=12, fontname="hebo", color=(1, 1, 1))
    y += 42

    resumen = (f"{len(LOTE)} filas (predio+concepto), {len(set((m, l) for m, l, *_ in LOTE))} predios. "
               f"Investigado 01/08/2026: NO es el bug de carrera del 13/07 (sin eventos con esa fecha "
               f"en ninguno). El PAGO del 06/07 no tiene respaldo verificable en ningun archivo de pagos "
               f"(actual ni copias historicas de 09/07 y 15/07) -- el AJUSTE del 31/07 revierte ese PAGO "
               f"pero deja el SALDO negativo en vez de restaurar la deuda real. Comparar contra la tabla "
               f"de Referencia de pago de cada predio (pagina siguiente) para decidir la correccion.")
    page.insert_text((_M, y), resumen, fontsize=8.5, fontname="helv", color=_NEGRO)
    y += 30

    headers = ["Predio", "Nombre", "Concepto", "PAGO 06/07", "AJUSTE 31/07", "SALDO hoy"]
    anchos = [55, 300, 100, 100, 110, 100]
    row_h = 16
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + row_h), fill=_AZUL_BG, color=None)
    cx = _M
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + row_h - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += row_h

    for n, (mz, lt, concepto, pago, ajuste, saldo) in enumerate(LOTE):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(_M, y, _M + w, y + row_h), fill=_ZEBRA, color=None)
        cx = _M
        page.insert_text((cx + 4, y + row_h - 5), f"{mz}-{lt}", fontsize=8, fontname="hebo", color=_NEGRO)
        cx += anchos[0]
        page.insert_text((cx + 4, y + row_h - 5), str(nombres.get((mz, lt), ""))[:45], fontsize=8,
                          fontname="helv", color=_NEGRO)
        cx += anchos[1]
        page.insert_text((cx + 4, y + row_h - 5), concepto, fontsize=8, fontname="helv", color=_NEGRO)
        cx += anchos[2]
        page.insert_text((cx + 4, y + row_h - 5), f"S/ {pago:,.2f}", fontsize=8, fontname="helv", color=_NEGRO)
        cx += anchos[3]
        page.insert_text((cx + 4, y + row_h - 5), f"S/ {ajuste:,.2f}", fontsize=8, fontname="helv", color=_ROJO)
        cx += anchos[4]
        color_saldo = _ROJO if saldo < 0 else _NEGRO
        page.insert_text((cx + 4, y + row_h - 5), f"S/ {saldo:,.2f}", fontsize=8, fontname="hebo",
                          color=color_saldo)
        y += row_h


def generar(salida: Path | None = None) -> Path:
    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()
    redirects = rcm._cargar_redirects()

    doc = fitz.open()
    _dibujar_portada(doc, nombres)

    predios = sorted(set((mz, lt) for mz, lt, *_ in LOTE))
    for mz, lt in predios:
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        refs = rrp.referencias_pago(mz, lt, tabla=tabla)
        rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
        page = doc[-1]
        w = rh._PAGE_W - 2 * rh._M
        y_tabla_fin = rh._M + 42 + 14 + 18 + (18 * len(tabla))
        rrp._dibujar_tabla_referencias(page, rh._M, y_tabla_fin + 45, w, refs)

    salida = salida or (BASE_DIR / "outputs" / f"reporte_lote_saldo_negativo_{MES_ANO}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"{len(predios)} predios -> PDF {salida}")
    return salida


if __name__ == "__main__":
    generar()
