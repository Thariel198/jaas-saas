"""
4b_reclamos/reporte_correccion_bug.py — REPORTE DE LECTURA (no escribe nada en el
ledger): los AJUSTE que todavia no dicen POR QUE pasaron, y contra que pago real
se pueden contrastar.

Un AJUSTE es la unica correccion que el ledger acepta: nunca borra ni pisa una
fila vieja, agrega una que compensa. Pero a diferencia de CARGO y PAGO, un AJUSTE
no tiene fuente propia -- su unica justificacion es lo que escribio quien lo hizo.
Hasta el 03/08/2026 registrar_ajuste() EXIGIA el motivo y despues lo descartaba
sin escribirlo, asi que quedaron 38 filas mudas.

Este PDF las junta con su evidencia para poder cerrarlas una por una:
  · pagina 1     resumen por FORMA del bug + totales
  · pagina 2..   tabla de los predio-concepto afectados
  · 1 pag/predio historial mensual + REFERENCIA DE PAGO (de donde vino la plata,
                 leida del crudo de yape/efectivo) + el detalle de sus AJUSTE

La causa NO se escribe aca: va al MOTIVO del ledger, que es su casa. Este reporte
solo la lee -- si fuera una columna de un PDF regenerable seria trabajo humano en
un archivo que se pisa solo (Regla 9 del CLAUDE.md).

Uso: py 4b_reclamos/reporte_correccion_bug.py [MES_ANO]
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

_AZUL, _AZUL_BG = rcm._AZUL, rcm._AZUL_BG
_GRIS, _NEGRO, _VERDE, _ROJO, _ZEBRA = rcm._GRIS, rcm._NEGRO, rcm._VERDE, rcm._ROJO, rcm._ZEBRA
_PAGE_W, _PAGE_H, _M = rcm._PAGE_W, rcm._PAGE_H, rcm._M

TOL = 0.005

# El orden importa: se evalua de arriba hacia abajo y gana la primera que aplica.
FORMAS = {
    1: "PAGO FANTASMA — el ledger acredita mas plata de la que entro",
    2: "CARGO ANULADO — nunca hubo pago, el ajuste borra la deuda",
    3: "PAGO PARCIAL — pago menos y el resto se ajusto",
    4: "PAR QUE NETEA 0 — entra y sale, no mueve nada",
    5: "GENESIS TARDIA — el cargo se imputo a un mes ya cerrado",
    6: "DEUDA REABIERTA — el ajuste dejo saldo vivo",
}


# ── CALCULO ───────────────────────────────────────────────────────────────────
def _clasificar(cargo, pago, saldo, ajustes_sin_motivo, hay_genesis_tardia) -> int:
    if hay_genesis_tardia:
        return 5
    if saldo > TOL:
        return 6
    if pago > cargo + TOL:
        return 1
    if len(ajustes_sin_motivo) >= 2 and abs(sum(ajustes_sin_motivo)) <= TOL:
        return 4
    if pago <= TOL:
        return 2
    return 3


_COLS_TXT = ("MZ", "LT", "CONCEPTO", "MES", "TIPO_EVENTO", "SOURCE", "AUDIT_REF",
             "TIMESTAMP", "CLASE", "MOTIVO")


def eventos_limpios(eventos: pd.DataFrame | None = None) -> pd.DataFrame:
    """Los eventos con las columnas de texto normalizadas a str, sin NA.
    Las de monto quedan intactas -- se suman como float mas abajo. Sin el
    fillna, una celda vacia queda como pd.NA y `== ""` no matchea ninguna."""
    ev = (eventos if eventos is not None else repo._leer_eventos()).copy()
    for c in _COLS_TXT:
        ev[c] = ev[c].fillna("").astype(str).str.strip()
    return ev


def calcular_tabla(eventos: pd.DataFrame | None = None) -> pd.DataFrame:
    """Una fila por (predio, concepto) con al menos un AJUSTE sin MOTIVO."""
    ev = eventos_limpios(eventos)
    nombres = repo._lookup_nombres()

    mudos = ev[(ev["TIPO_EVENTO"] == "AJUSTE") & (ev["MOTIVO"] == "")]
    pares = sorted({(r["MZ"], r["LT"], r["CONCEPTO"]) for _, r in mudos.iterrows()})

    filas = []
    for mz, lt, concepto in pares:
        s = ev[(ev["MZ"] == mz) & (ev["LT"] == lt) & (ev["CONCEPTO"] == concepto)]
        cargos = s[s["TIPO_EVENTO"] == "CARGO"]
        pagos = s[s["TIPO_EVENTO"] == "PAGO"]
        ajustes = s[s["TIPO_EVENTO"] == "AJUSTE"]
        mudos_par = ajustes[ajustes["MOTIVO"] == ""]

        cargo = round(float(cargos["CARGO"].fillna(0).sum()), 2)
        pago = round(float(pagos["PAGO"].fillna(0).sum()), 2)
        saldo = round(float(s.sort_values("TIMESTAMP").iloc[-1]["SALDO"]), 2)
        hay_gt = bool(cargos["SOURCE"].str.contains("genesis_tardia").any())

        forma = _clasificar(cargo, pago, saldo,
                            [float(v) for v in mudos_par["AJUSTE"].fillna(0)], hay_gt)

        filas.append({
            "MZ": mz, "LT": lt, "NOMBRE": nombres.get((mz, lt), ""), "CONCEPTO": concepto,
            "CARGO": cargo,
            "PAGO_LEDGER": pago,
            "N_PAGOS": len(pagos),
            "EXCESO": round(max(0.0, pago - cargo), 2),
            "AJUSTE_TOTAL": round(float(ajustes["AJUSTE"].fillna(0).sum()), 2),
            "AJUSTE_MUDO": round(float(mudos_par["AJUSTE"].fillna(0).sum()), 2),
            "N_MUDOS": len(mudos_par),
            "SALDO": saldo,
            "FORMA": forma,
        })

    df = pd.DataFrame(filas)
    return df.sort_values(["FORMA", "EXCESO", "MZ", "LT"],
                          ascending=[True, False, True, True]).reset_index(drop=True)


def ajustes_del_par(ev: pd.DataFrame, mz: str, lt: str, concepto: str) -> pd.DataFrame:
    """TODOS los AJUSTE del par -- tambien los que ya tienen motivo, que son el
    contexto de por que el mudo esta ahi. `ev` viene de eventos_limpios()."""
    s = ev[(ev["MZ"] == mz) & (ev["LT"] == lt) &
           (ev["CONCEPTO"] == concepto) & (ev["TIPO_EVENTO"] == "AJUSTE")]
    return s.sort_values("TIMESTAMP")


# ── DIBUJO ────────────────────────────────────────────────────────────────────
def _dibujar_resumen(doc, df: pd.DataFrame, n_mudos: int, n_total: int) -> None:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text((_M + 10, y + 20), "Corrección de bug — los AJUSTE que todavía no dicen por qué",
                      fontsize=12, fontname="hebo", color=(1, 1, 1))
    y += 40

    intro = (f"SOLO LECTURA — este PDF no modifica el ledger. {n_mudos} de {n_total} AJUSTE quedaron sin MOTIVO: "
             "registrar_ajuste() lo exigía y lo descartaba sin escribirlo. La causa se escribe en el ledger, no acá.")
    page.insert_text((_M, y), intro, fontsize=8.5, fontname="helv", color=_GRIS)
    y += 24

    page.insert_text((_M, y), "Por forma del bug", fontsize=10, fontname="hebo", color=_AZUL)
    y += 18
    headers = ["Forma", "Pares", "Filas mudas", "Cargado", "Pagado", "Exceso", "Saldo vivo"]
    anchos = [352, 55, 75, 75, 75, 75, 75]
    rh_row = 17
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = _M
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh_row - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row

    for n, (forma, etiqueta) in enumerate(sorted(FORMAS.items())):
        sub = df[df["FORMA"] == forma]
        if sub.empty:
            continue
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_ZEBRA, color=None)
        color = _ROJO if forma in (1, 6) else _NEGRO
        cx = _M
        page.insert_text((cx + 4, y + rh_row - 5), f"{forma}. {etiqueta}", fontsize=8, fontname="helv", color=color)
        cx += anchos[0]
        for val, cw in ((len(sub), anchos[1]), (int(sub["N_MUDOS"].sum()), anchos[2])):
            page.insert_text((cx + 4, y + rh_row - 5), f"{val}", fontsize=8, fontname="hebo", color=_NEGRO)
            cx += cw
        for campo, cw in (("CARGO", anchos[3]), ("PAGO_LEDGER", anchos[4]),
                          ("EXCESO", anchos[5]), ("SALDO", anchos[6])):
            v = float(sub[campo].sum())
            page.insert_text((cx + 4, y + rh_row - 5), f"{v:,.2f}" if v > TOL else "·",
                              fontsize=8, fontname="helv", color=_NEGRO if v > TOL else _GRIS)
            cx += cw
        y += rh_row

    page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = _M
    page.insert_text((cx + 4, y + rh_row - 5), "TOTAL", fontsize=8, fontname="hebo", color=_AZUL)
    cx += anchos[0]
    for val, cw in ((len(df), anchos[1]), (int(df["N_MUDOS"].sum()), anchos[2])):
        page.insert_text((cx + 4, y + rh_row - 5), f"{val}", fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    for campo, cw in (("CARGO", anchos[3]), ("PAGO_LEDGER", anchos[4]),
                      ("EXCESO", anchos[5]), ("SALDO", anchos[6])):
        page.insert_text((cx + 4, y + rh_row - 5), f"{float(df[campo].sum()):,.2f}",
                          fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row + 24

    page.insert_text((_M, y), "Cómo leer cada página de predio", fontsize=10, fontname="hebo", color=_AZUL)
    y += 16
    for linea in (
        "1. Historial mensual — lo que el sistema cree que pasó, mes a mes.",
        "2. Referencia de pago — de dónde vino la plata según el crudo de yape/efectivo. Es la prueba independiente:",
        "   si el ledger dice dos pagos y acá aparece uno solo, el otro es fantasma.",
        "3. Ajustes del predio — cada AJUSTE con su source, audit_ref y motivo. Los que dicen (SIN MOTIVO) son los pendientes.",
    ):
        page.insert_text((_M, y), linea, fontsize=8.5, fontname="helv", color=_NEGRO)
        y += 14


def _dibujar_portada(doc, df: pd.DataFrame) -> None:
    headers = ["Predio", "Nombre", "Concepto", "Cargado", "Pagado", "N pagos",
               "Exceso", "Ajuste total", "Mudo", "N mudos", "Saldo hoy", "Forma"]
    anchos = [46, 168, 66, 58, 58, 46, 55, 62, 55, 48, 58, 62]
    rh_row = 15
    filas_por_pag = int((_PAGE_H - _M * 2 - 60) // rh_row)
    n_paginas = max(-(-len(df) // filas_por_pag), 1)

    for pagina in range(n_paginas):
        page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
        w = _PAGE_W - 2 * _M
        y = _M
        page.draw_rect(fitz.Rect(_M, y, _M + w, y + 26), fill=_AZUL, color=None)
        page.insert_text((_M + 10, y + 18),
                          f"Predio y concepto con AJUSTE sin causa (pág. {pagina + 1}/{n_paginas})",
                          fontsize=11, fontname="hebo", color=(1, 1, 1))
        y += 34

        page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_AZUL_BG, color=None)
        cx = _M
        for h, cw in zip(headers, anchos):
            page.insert_text((cx + 3, y + rh_row - 5), h, fontsize=7, fontname="hebo", color=_AZUL)
            cx += cw
        y += rh_row

        bloque = df.iloc[pagina * filas_por_pag:(pagina + 1) * filas_por_pag]
        for n, (_, r) in enumerate(bloque.iterrows()):
            if n % 2 == 1:
                page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_ZEBRA, color=None)
            cx = _M
            page.insert_text((cx + 3, y + rh_row - 5), f"{r['MZ']}-{r['LT']}", fontsize=7, fontname="hebo", color=_NEGRO)
            cx += anchos[0]
            page.insert_text((cx + 3, y + rh_row - 5), str(r["NOMBRE"])[:32], fontsize=7, fontname="helv", color=_NEGRO)
            cx += anchos[1]
            page.insert_text((cx + 3, y + rh_row - 5), r["CONCEPTO"], fontsize=7, fontname="helv", color=_GRIS)
            cx += anchos[2]
            for campo, cw in (("CARGO", anchos[3]), ("PAGO_LEDGER", anchos[4])):
                v = r[campo]
                page.insert_text((cx + 3, y + rh_row - 5), f"{v:,.2f}" if v > TOL else "·",
                                  fontsize=7, fontname="helv", color=_NEGRO if v > TOL else _GRIS)
                cx += cw
            page.insert_text((cx + 3, y + rh_row - 5), str(int(r["N_PAGOS"])),
                              fontsize=7, fontname="hebo" if r["N_PAGOS"] > 1 else "helv",
                              color=_ROJO if r["N_PAGOS"] > 1 else _GRIS)
            cx += anchos[5]
            v = r["EXCESO"]
            page.insert_text((cx + 3, y + rh_row - 5), f"{v:,.2f}" if v > TOL else "·",
                              fontsize=7, fontname="hebo", color=_ROJO if v > TOL else _GRIS)
            cx += anchos[6]
            for campo, cw in (("AJUSTE_TOTAL", anchos[7]), ("AJUSTE_MUDO", anchos[8])):
                v = r[campo]
                page.insert_text((cx + 3, y + rh_row - 5), f"{v:+,.2f}" if abs(v) > TOL else "·",
                                  fontsize=7, fontname="helv", color=_NEGRO if abs(v) > TOL else _GRIS)
                cx += cw
            page.insert_text((cx + 3, y + rh_row - 5), str(int(r["N_MUDOS"])), fontsize=7, fontname="hebo", color=_NEGRO)
            cx += anchos[9]
            v = r["SALDO"]
            page.insert_text((cx + 3, y + rh_row - 5), f"{v:,.2f}" if abs(v) > TOL else "0",
                              fontsize=7, fontname="hebo", color=_ROJO if v > TOL else _VERDE)
            cx += anchos[10]
            page.insert_text((cx + 3, y + rh_row - 5), f"{int(r['FORMA'])}. {FORMAS[r['FORMA']].split(' — ')[0][:9]}",
                              fontsize=6.5, fontname="helv", color=_GRIS)
            y += rh_row


def _dibujar_tabla_ajustes(page, x: float, y: float, w: float, ajustes: pd.DataFrame) -> float:
    headers = ["Fecha", "Concepto", "Mes", "Ajuste", "Saldo", "Source", "Audit ref", "Clase", "Motivo"]
    anchos = [76, 58, 46, 50, 48, 84, 176, 84, w - 622]
    rh_row = 16

    page.insert_text((x, y), "Ajustes del predio (los que dicen SIN MOTIVO son los pendientes)",
                      fontsize=9, fontname="hebo", color=_AZUL)
    y += 16

    page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh_row - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row

    for n, (_, r) in enumerate(ajustes.iterrows()):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_ZEBRA, color=None)
        motivo = str(r["MOTIVO"]).strip()
        mudo = motivo == "" or motivo.lower() == "nan"
        cx = x
        celdas = [
            (str(r["TIMESTAMP"])[:16], "helv", _GRIS),
            (str(r["CONCEPTO"]), "helv", _NEGRO),
            (str(r["MES"]), "helv", _GRIS),
            (f"{float(r['AJUSTE']):+,.2f}", "hebo", _NEGRO),
            (f"{float(r['SALDO']):,.2f}", "helv", _GRIS),
            (str(r["SOURCE"])[:15], "helv", _GRIS),
            (str(r["AUDIT_REF"])[:36], "helv", _GRIS),
            (str(r["CLASE"])[:14], "helv", _ROJO if str(r["CLASE"]) == "SIN_CLASIFICAR" else _NEGRO),
            ("(SIN MOTIVO)" if mudo else motivo[:60], "hebo" if mudo else "helv", _ROJO if mudo else _NEGRO),
        ]
        for (txt, fuente, color), cw in zip(celdas, anchos):
            page.insert_text((cx + 4, y + rh_row - 5), txt, fontsize=7, fontname=fuente, color=color)
            cx += cw
        y += rh_row

    return y


def generar(mes_ano: str = "2026-07", salida: Path | None = None) -> Path:
    eventos = repo._leer_eventos()
    ev = eventos_limpios(eventos)
    df = calcular_tabla(eventos)
    n_total = int((ev["TIPO_EVENTO"] == "AJUSTE").sum())
    n_mudos = int(df["N_MUDOS"].sum())

    historicos = rh._cargar_historicos()
    mapa_raw = rh._cargar_mapa_raw()
    dfp = pd.read_excel(BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx",
                        sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()
    redirects = rcm._cargar_redirects()

    doc = fitz.open()
    _dibujar_resumen(doc, df, n_mudos, n_total)
    _dibujar_portada(doc, df)

    # Una pagina por PREDIO (no por par): un predio con dos conceptos tocados
    # comparte historial y referencias de pago, solo cambia la tabla de ajustes.
    predios = sorted({(r["MZ"], r["LT"]) for _, r in df.iterrows()})
    for n, (mz, lt) in enumerate(predios, 1):
        nombre = nombres.get((mz, lt), "")
        tabla = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombre)
        tabla = rcm.corregir_tabla_por_redirects(tabla, mz, lt, redirects)
        refs = rrp.referencias_pago(mz, lt, tabla=tabla)

        rh._dibujar_pagina(doc, mz, lt, nombre, tabla)
        page = doc[-1]
        w = rh._PAGE_W - 2 * rh._M
        y = rh._M + 42 + 14 + 18 + (18 * len(tabla)) + 45
        y = rrp._dibujar_tabla_referencias(page, rh._M, y, w, refs) + 26

        conceptos = sorted(df[(df["MZ"] == mz) & (df["LT"] == lt)]["CONCEPTO"])
        ajustes = pd.concat([ajustes_del_par(ev, mz, lt, c) for c in conceptos])
        if y + 32 + 16 * len(ajustes) > rh._PAGE_H - rh._M:
            page = doc.new_page(width=rh._PAGE_W, height=rh._PAGE_H)
            page.draw_rect(fitz.Rect(rh._M, rh._M, rh._M + w, rh._M + 30), fill=_AZUL, color=None)
            page.insert_text((rh._M + 10, rh._M + 20), f"Predio {mz}-{lt} — {nombre or '(sin nombre)'} — ajustes",
                              fontsize=12, fontname="hebo", color=(1, 1, 1))
            y = rh._M + 52
        _dibujar_tabla_ajustes(page, rh._M, y, w, ajustes)
        if n % 5 == 0:
            print(f"  paginas {n}/{len(predios)}...")

    salida = salida or (BASE_DIR / "outputs" / f"reporte_correccion_bug_{mes_ano}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()

    print()
    for forma, etiqueta in sorted(FORMAS.items()):
        sub = df[df["FORMA"] == forma]
        if sub.empty:
            continue
        print(f"  {forma}. {etiqueta:<58} {len(sub):>2} pares · {int(sub['N_MUDOS'].sum()):>2} filas · "
              f"exceso S/ {sub['EXCESO'].sum():>8,.2f} · saldo S/ {sub['SALDO'].sum():>8,.2f}")
    print(f"\n  {len(df)} pares · {n_mudos} de {n_total} AJUSTE sin motivo · {len(predios)} predios")
    print(f"\nPDF  -> {salida}")
    df.to_excel(salida.with_suffix(".xlsx"), index=False)
    print(f"XLSX -> {salida.with_suffix('.xlsx')}")
    return salida


if __name__ == "__main__":
    generar(sys.argv[1] if len(sys.argv) > 1 else "2026-07")
