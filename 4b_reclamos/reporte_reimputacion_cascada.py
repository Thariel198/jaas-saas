"""
4b_reclamos/reporte_reimputacion_cascada.py — SIMULACION (no escribe nada en el
ledger): que pasaria si los pagos de feb-jul 2026 se hubieran imputado con el
orden CA1 (CONVENIO -> ACUERDOS -> MULTA) en vez del orden que aplica el codigo
hoy (MULTA -> ACUERDOS -> CONVENIO).

Por que importa: la multa es el unico concepto que se puede cubrir con faena o
exonerar, asi que deberia ser el ULTIMO en llevarse la plata. Al ser el primero,
dejo dos reclamos recurrentes:
    "ya pague mi medidor"        -> el convenio quedo al final de la cola
    "ya pague techado y campo"   -> acuerdos quedo detras de multa

Dos movimientos por predio, en orden:
  (1) solo si el convenio es MEDIDOR:  MULTA -> CONVENIO, y si no alcanza,
      ACUERDOS -> CONVENIO, hasta saldar el convenio.
  (2) para TODOS:  lo que sobre de MULTA -> ACUERDOS, hasta saldar acuerdos.

Clasificacion del convenio (decide solo si el convenio puede RECIBIR plata; el
movimiento (2) corre para todos):
    MEDIDOR       saldo <= 100 y no figura en las listas de instalacion/reactivacion
    INSTALACION   saldo > 100, o figura en NUEVAS INSTALACIONES / ANTERIOR DIRECTIVA
    REACTIVACION  figura en la hoja REACTIVACION
    SIN_CONVENIO  no debe convenio

Uso: py 4b_reclamos/reporte_reimputacion_cascada.py [MES_ANO]
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

SEGUIMIENTO_XLSX = (BASE_DIR.parent / "obligaciones" / "inputs" /
                    "SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx")

CONCEPTOS = ("MULTA", "ACUERDOS", "CONVENIO")
MESES_VENTANA = ("2026-02", "2026-03", "2026-04", "2026-05", "2026-06", "2026-07")
TOPE_MEDIDOR = 100.0

_AZUL, _AZUL_BG = rcm._AZUL, rcm._AZUL_BG
_GRIS, _NEGRO, _VERDE, _ROJO, _ZEBRA = rcm._GRIS, rcm._NEGRO, rcm._VERDE, rcm._ROJO, rcm._ZEBRA
_PAGE_W, _PAGE_H, _M = rcm._PAGE_W, rcm._PAGE_H, rcm._M


# ── LISTAS DE INSTALACION / REACTIVACION ──────────────────────────────────────
def _norm_predio(mz, lt) -> tuple[str, str]:
    """A-8C, 'A 8 C' y 'A-8-C' son el mismo predio en fuentes distintas."""
    s = str(lt).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return str(mz).strip().upper(), s.upper().replace(" ", "").replace("-", "")


def _leer_hoja_predios(hoja: str) -> set[tuple[str, str]]:
    """Las hojas tienen filas de titulo antes de la cabecera real y el numero de
    filas de relleno varia por hoja -- se busca la fila que tenga MZ y LT."""
    df = pd.read_excel(SEGUIMIENTO_XLSX, hoja, header=None)
    fila_cab = None
    for i in range(min(6, len(df))):
        celdas = [str(v).strip().upper() for v in df.iloc[i]]
        if "MZ" in celdas and "LT" in celdas:
            fila_cab = i
            break
    if fila_cab is None:
        return set()
    d = pd.read_excel(SEGUIMIENTO_XLSX, hoja, header=fila_cab).dropna(subset=["MZ", "LT"])
    return {_norm_predio(r["MZ"], r["LT"]) for _, r in d.iterrows()}


def cargar_listas() -> tuple[set, set]:
    instalacion = _leer_hoja_predios("NUEVAS INSTALACIONES") | _leer_hoja_predios("INSTALACIONES ANTERIOR DIRECTIV")
    reactivacion = _leer_hoja_predios("REACTIVACION")
    return instalacion, reactivacion


def clasificar_convenio(clave, saldo_convenio: float, instalacion: set, reactivacion: set) -> str:
    k = _norm_predio(*clave)
    if k in reactivacion:
        return "REACTIVACION"
    if saldo_convenio <= 0.005:
        return "SIN_CONVENIO"
    if k in instalacion or saldo_convenio > TOPE_MEDIDOR + 0.005:
        return "INSTALACION"
    return "MEDIDOR"


# ── CALCULO ───────────────────────────────────────────────────────────────────
_cache_tabla: dict[tuple[str, str], pd.DataFrame] = {}


def _tabla_corregida(mz, lt, historicos, eventos, dfp, mapa_raw, nombre, redirects):
    """tabla_predio + corregir_tabla_por_redirects, cacheada por predio.

    generar() recorre los predios DOS veces —una en calcular_tabla y otra al
    dibujar cada pagina— y recalculaba la misma tabla en cada pasada. Son ~1.9 s
    por predio: con --todos (570 predios) eso es 37 min contra 18."""
    k = (mz, lt)
    if k not in _cache_tabla:
        t = rh.tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombre)
        _cache_tabla[k] = rcm.corregir_tabla_por_redirects(t, mz, lt, redirects)
    return _cache_tabla[k]


def predios_con_reclamo(mes_ciclo: str, tipo: str = "mes_anterior") -> set[tuple[str, str]]:
    """Los predios que tienen un reclamo abierto de ese tipo en el ciclo.

    Es la lista corta y util: un predio que solo debe mes anterior no aparece en
    el PDF normal (que se arma con los que deben MULTA/ACUERDOS/CONVENIO), asi
    que buscar O-28 ahi no da nada. Con esto el PDF trae exactamente los predios
    que hay que resolver."""
    ruta = BASE_DIR / "outputs" / f"reclamos_{mes_ciclo}.xlsx"
    if not ruta.exists():
        raise FileNotFoundError(f"Falta {ruta} -> correr 4b_reclamos/main.py --mes {mes_ciclo}")
    df = pd.read_excel(ruta, sheet_name="Reclamos", header=1)
    df = df[df["TIPO_RECLAMO"].astype(str).str.strip() == tipo]
    out = set()
    for _, r in df.iterrows():
        mz, lt = str(r["MZ"]).strip().upper(), str(r["LT"]).strip().upper()
        if lt.endswith(".0"):
            lt = lt[:-2]
        if mz and lt and mz != "NAN" and lt != "NAN":
            out.add((mz, lt))
    return out


def calcular_tabla(mes_ano: str = "2026-07",
                   predios: set[tuple[str, str]] | None = None) -> pd.DataFrame:
    nombres = repo._lookup_nombres()
    saldos = {c: repo.get_saldos_bulk(c, mes_ano) for c in CONCEPTOS}
    instalacion, reactivacion = cargar_listas()

    if predios is None:
        # Por defecto: los que deben alguno de los 3 conceptos que la
        # re-imputacion mueve.
        predios = set()
        for c in CONCEPTOS:
            for k, v in saldos[c].items():
                if v > 0.005 and k[0] != "NAN" and k[1] != "NAN":
                    predios.add(k)
    else:
        predios = {k for k in predios if k[0] != "NAN" and k[1] != "NAN"}

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    dfp = pd.read_excel(BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx",
                        sheet_name="planilla_cobrado", header=1)
    redirects = rcm._cargar_redirects()

    filas = []
    for n, (mz, lt) in enumerate(sorted(predios), 1):
        nombre = nombres.get((mz, lt), "")
        tabla = _tabla_corregida(mz, lt, historicos, eventos, dfp, mapa_raw, nombre, redirects)
        vent = tabla[tabla["MES"].isin(MESES_VENTANA)]

        sm = round(max(0.0, saldos["MULTA"].get((mz, lt), 0.0)), 2)
        sa = round(max(0.0, saldos["ACUERDOS"].get((mz, lt), 0.0)), 2)
        sc = round(max(0.0, saldos["CONVENIO"].get((mz, lt), 0.0)), 2)
        pm = round(float(vent["MULTA"].sum()), 2)
        pa = round(float(vent["ACUERDOS"].sum()), 2)
        pc = round(float(vent["CONVENIO"].sum()), 2)

        clase = clasificar_convenio((mz, lt), sc, instalacion, reactivacion)

        # (1) el convenio solo recibe si es MEDIDOR
        if clase == "MEDIDOR":
            r_m_c = round(min(pm, sc), 2)                 # multa    -> convenio
            r_a_c = round(min(pa, sc - r_m_c), 2)         # acuerdos -> convenio
        else:
            r_m_c = r_a_c = 0.0
        # (2) lo que sobra de multa cubre acuerdos -- incluida la deuda que el
        #     paso (1) acaba de devolverle a acuerdos
        r_m_a = round(min(pm - r_m_c, sa + r_a_c), 2)     # multa    -> acuerdos

        filas.append({
            "MZ": mz, "LT": lt, "NOMBRE": nombre, "CLASE": clase,
            "MULTA_ANTES": sm, "ACUERDOS_ANTES": sa, "CONVENIO_ANTES": sc,
            "PAGO_MULTA": pm, "PAGO_ACUERDOS": pa, "PAGO_CONVENIO": pc,
            "MULTA_A_CONVENIO": r_m_c, "ACUERDOS_A_CONVENIO": r_a_c, "MULTA_A_ACUERDOS": r_m_a,
            "MULTA_DESPUES":    round(sm + r_m_c + r_m_a, 2),
            "ACUERDOS_DESPUES": round(sa + r_a_c - r_m_a, 2),
            "CONVENIO_DESPUES": round(sc - r_m_c - r_a_c, 2),
            "PAGO_MULTA_DESPUES":    round(pm - r_m_c - r_m_a, 2),
            "PAGO_ACUERDOS_DESPUES": round(pa - r_a_c + r_m_a, 2),
            "PAGO_CONVENIO_DESPUES": round(pc + r_m_c + r_a_c, 2),
        })
        if n % 25 == 0:
            print(f"  calculando {n}/{len(predios)}...")

    df = pd.DataFrame(filas)
    df["MOVIDO"] = (df["MULTA_A_CONVENIO"] + df["ACUERDOS_A_CONVENIO"] + df["MULTA_A_ACUERDOS"]).round(2)
    df["DEUDA_ANTES"] = (df["MULTA_ANTES"] + df["ACUERDOS_ANTES"] + df["CONVENIO_ANTES"]).round(2)
    df["DEUDA_DESPUES"] = (df["MULTA_DESPUES"] + df["ACUERDOS_DESPUES"] + df["CONVENIO_DESPUES"]).round(2)

    def _resultado(r):
        partes = []
        if r["CLASE"] == "MEDIDOR":
            if r["CONVENIO_DESPUES"] <= 0.005:
                partes.append("medidor SALDADO")
            elif r["CONVENIO_DESPUES"] < r["CONVENIO_ANTES"] - 0.005:
                partes.append("medidor parcial")
        if r["ACUERDOS_ANTES"] > 0.005:
            if r["ACUERDOS_DESPUES"] <= 0.005:
                partes.append("techado/campo SALDADO")
            elif r["ACUERDOS_DESPUES"] < r["ACUERDOS_ANTES"] - 0.005:
                partes.append("techado/campo parcial")
        return " · ".join(partes) if partes else "sin efecto"

    df["RESULTADO"] = df.apply(_resultado, axis=1)
    orden_clase = {"MEDIDOR": 0, "SIN_CONVENIO": 1, "INSTALACION": 2, "REACTIVACION": 3}
    df["_o"] = df["CLASE"].map(orden_clase)
    df = df.sort_values(["_o", "MOVIDO", "DEUDA_ANTES"], ascending=[True, False, False]).drop(columns="_o")
    return df.reset_index(drop=True)


def validar(df: pd.DataFrame) -> list[dict]:
    """Nada se crea ni se pierde: deuda total y pagos totales identicos antes y
    despues, por predio y en el agregado."""
    chequeos = []

    def _add(nombre, antes, despues, tol=0.005):
        ok = abs(antes - despues) <= tol
        chequeos.append({"CHEQUEO": nombre, "ANTES": antes, "DESPUES": despues,
                          "DIF": round(despues - antes, 2), "OK": ok})

    deudores_antes = int((df["DEUDA_ANTES"] > 0.005).sum())
    deudores_despues = int((df["DEUDA_DESPUES"] > 0.005).sum())
    chequeos.append({"CHEQUEO": "predios con deuda en algun concepto",
                      "ANTES": deudores_antes, "DESPUES": deudores_despues,
                      "DIF": deudores_despues - deudores_antes,
                      "OK": deudores_despues <= deudores_antes})

    for c in CONCEPTOS:
        _add(f"deuda {c}", round(df[f"{c}_ANTES"].sum(), 2), round(df[f"{c}_DESPUES"].sum(), 2), tol=1e9)
    _add("DEUDA TOTAL", round(df["DEUDA_ANTES"].sum(), 2), round(df["DEUDA_DESPUES"].sum(), 2))

    for c in CONCEPTOS:
        _add(f"pagos {c} (feb-jul)", round(df[f"PAGO_{c}"].sum(), 2),
             round(df[f"PAGO_{c}_DESPUES"].sum(), 2), tol=1e9)
    pa_t = round(sum(df[f"PAGO_{c}"].sum() for c in CONCEPTOS), 2)
    pd_t = round(sum(df[f"PAGO_{c}_DESPUES"].sum() for c in CONCEPTOS), 2)
    _add("PAGOS TOTALES (feb-jul)", pa_t, pd_t)

    desc_deuda = df[(df["DEUDA_ANTES"] - df["DEUDA_DESPUES"]).abs() > 0.005]
    chequeos.append({"CHEQUEO": "predios donde la deuda NO se conserva",
                      "ANTES": 0, "DESPUES": len(desc_deuda), "DIF": len(desc_deuda),
                      "OK": len(desc_deuda) == 0})
    neg = df[(df[[f"{c}_DESPUES" for c in CONCEPTOS]] < -0.005).any(axis=1)]
    chequeos.append({"CHEQUEO": "predios con saldo negativo despues",
                      "ANTES": 0, "DESPUES": len(neg), "DIF": len(neg), "OK": len(neg) == 0})
    return chequeos


# ── PDF ───────────────────────────────────────────────────────────────────────
def _dibujar_validacion(doc, df: pd.DataFrame, chequeos: list[dict], mes_ano: str) -> None:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text((_M + 10, y + 20),
                      f"Re-imputación de pagos con el orden CONVENIO → ACUERDOS → MULTA — {mes_ano}",
                      fontsize=12, fontname="hebo", color=(1, 1, 1))
    y += 40

    intro = ("SIMULACION — este PDF no modifica el ledger. Muestra como quedaria la deuda si los pagos de "
             "feb-jul 2026 se hubieran aplicado con la multa al final (es lo unico que se puede cubrir con faena).")
    page.insert_text((_M, y), intro, fontsize=8.5, fontname="helv", color=_GRIS)
    y += 24

    page.insert_text((_M, y), "Validación — nada se crea, nada se pierde", fontsize=10, fontname="hebo", color=_AZUL)
    y += 18
    headers = ["Chequeo", "Antes", "Después", "Diferencia", ""]
    anchos = [330, 110, 110, 110, 122]
    rh_row = 17
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = _M
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh_row - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row
    for n, ch in enumerate(chequeos):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(_M, y, _M + w, y + rh_row), fill=_ZEBRA, color=None)
        es_total = ch["CHEQUEO"].isupper() or "TOTAL" in ch["CHEQUEO"]
        fuente = "hebo" if es_total else "helv"
        cx = _M
        page.insert_text((cx + 4, y + rh_row - 5), ch["CHEQUEO"], fontsize=8, fontname=fuente, color=_NEGRO)
        cx += anchos[0]
        for val, cw in ((ch["ANTES"], anchos[1]), (ch["DESPUES"], anchos[2]), (ch["DIF"], anchos[3])):
            txt = f"{val:,.2f}" if isinstance(val, float) else f"{val:,}"
            page.insert_text((cx + 4, y + rh_row - 5), txt, fontsize=8, fontname=fuente, color=_NEGRO)
            cx += cw
        page.insert_text((cx + 4, y + rh_row - 5), "OK" if ch["OK"] else "REVISAR",
                          fontsize=8, fontname="hebo", color=_VERDE if ch["OK"] else _ROJO)
        y += rh_row

    y += 22
    page.insert_text((_M, y), "Alcance por clase de convenio", fontsize=10, fontname="hebo", color=_AZUL)
    y += 18
    for clase in ("MEDIDOR", "SIN_CONVENIO", "INSTALACION", "REACTIVACION"):
        sub = df[df["CLASE"] == clase]
        if sub.empty:
            continue
        entra = "convenio recibe plata" if clase == "MEDIDOR" else "convenio NO entra al cálculo"
        page.insert_text((_M, y),
                          f"{clase:<14} {len(sub):>4} predios · deuda S/ {sub['DEUDA_ANTES'].sum():>10,.2f} · "
                          f"se mueve S/ {sub['MOVIDO'].sum():>9,.2f} · {entra}",
                          fontsize=8.5, fontname="helv", color=_NEGRO)
        y += 14

    y += 14
    saldados_med = int(((df["CLASE"] == "MEDIDOR") & (df["CONVENIO_DESPUES"] <= 0.005)).sum())
    tot_med = int((df["CLASE"] == "MEDIDOR").sum())
    con_acu = int((df["ACUERDOS_ANTES"] > 0.005).sum())
    saldados_acu = int(((df["ACUERDOS_ANTES"] > 0.005) & (df["ACUERDOS_DESPUES"] <= 0.005)).sum())
    page.insert_text((_M, y), "Resultado", fontsize=10, fontname="hebo", color=_AZUL)
    y += 16
    for linea in (
        f"medidor saldado por completo:      {saldados_med} de {tot_med} predios MEDIDOR",
        f"techado y campo saldado:           {saldados_acu} de {con_acu} predios con deuda de acuerdos",
        f"predios sin ningun movimiento:     {int((df['MOVIDO'] <= 0.005).sum())} de {len(df)}",
    ):
        page.insert_text((_M, y), linea, fontsize=8.5, fontname="helv", color=_NEGRO)
        y += 14


def _dibujar_portada(doc, df: pd.DataFrame, mes_ano: str) -> None:
    headers = ["Predio", "Nombre", "Clase", "Multa", "Acuer.", "Conv.",
               "M→Conv", "A→Conv", "M→Acu", "Multa'", "Acuer.'", "Conv.'", "Resultado"]
    anchos = [42, 150, 60, 45, 45, 45, 48, 48, 48, 45, 45, 45, 116]
    rh_row = 15
    filas_por_pag = int((_PAGE_H - _M * 2 - 60) // rh_row)
    n_paginas = max(-(-len(df) // filas_por_pag), 1)

    for pagina in range(n_paginas):
        page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
        w = _PAGE_W - 2 * _M
        y = _M
        page.draw_rect(fitz.Rect(_M, y, _M + w, y + 26), fill=_AZUL, color=None)
        page.insert_text((_M + 10, y + 18),
                          f"Deuda antes → después de re-imputar — {mes_ano} (pág. {pagina + 1}/{n_paginas})",
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
            page.insert_text((cx + 3, y + rh_row - 5), str(r["NOMBRE"])[:28], fontsize=7, fontname="helv", color=_NEGRO)
            cx += anchos[1]
            page.insert_text((cx + 3, y + rh_row - 5), r["CLASE"][:12], fontsize=6.5, fontname="helv",
                              color=_NEGRO if r["CLASE"] == "MEDIDOR" else _GRIS)
            cx += anchos[2]
            for campo, cw in zip(("MULTA_ANTES", "ACUERDOS_ANTES", "CONVENIO_ANTES"), anchos[3:6]):
                v = r[campo]
                page.insert_text((cx + 3, y + rh_row - 5), f"{v:,.0f}" if v > 0.005 else "·",
                                  fontsize=7, fontname="helv", color=_NEGRO if v > 0.005 else _GRIS)
                cx += cw
            for campo, cw in zip(("MULTA_A_CONVENIO", "ACUERDOS_A_CONVENIO", "MULTA_A_ACUERDOS"), anchos[6:9]):
                v = r[campo]
                page.insert_text((cx + 3, y + rh_row - 5), f"{v:,.0f}" if v > 0.005 else "·",
                                  fontsize=7, fontname="hebo", color=_AZUL if v > 0.005 else _GRIS)
                cx += cw
            for campo, cw in zip(("MULTA_DESPUES", "ACUERDOS_DESPUES", "CONVENIO_DESPUES"), anchos[9:12]):
                v = r[campo]
                page.insert_text((cx + 3, y + rh_row - 5), f"{v:,.0f}" if v > 0.005 else "·",
                                  fontsize=7, fontname="hebo", color=_NEGRO if v > 0.005 else _VERDE)
                cx += cw
            color_res = _VERDE if "SALDADO" in r["RESULTADO"] else (_GRIS if r["RESULTADO"] == "sin efecto" else _NEGRO)
            page.insert_text((cx + 3, y + rh_row - 5), r["RESULTADO"][:30], fontsize=6.5, fontname="helv", color=color_res)
            y += rh_row


def generar(mes_ano: str = "2026-07", salida: Path | None = None,
            predios: set[tuple[str, str]] | None = None) -> Path:
    df = calcular_tabla(mes_ano, predios)
    chequeos = validar(df)

    historicos = rh._cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = rh._cargar_mapa_raw()
    dfp = pd.read_excel(BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx",
                        sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()
    redirects = rcm._cargar_redirects()

    doc = fitz.open()
    _dibujar_validacion(doc, df, chequeos, mes_ano)
    _dibujar_portada(doc, df, mes_ano)
    for n, (_, r) in enumerate(df.iterrows(), 1):
        mz, lt = r["MZ"], r["LT"]
        tabla = _tabla_corregida(mz, lt, historicos, eventos, dfp, mapa_raw,
                                  nombres.get((mz, lt), ""), redirects)
        refs = rrp.referencias_pago(mz, lt, tabla=tabla)
        saldo_pendiente = {"MULTA": r["MULTA_DESPUES"], "ACUERDOS": r["ACUERDOS_DESPUES"],
                            "CONVENIO": r["CONVENIO_DESPUES"], "TOTAL": r["DEUDA_DESPUES"]}
        y_fin = rh._dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla,
                                    saldo_pendiente=saldo_pendiente)
        page = doc[-1]
        w = rh._PAGE_W - 2 * rh._M
        rrp._dibujar_tabla_referencias(page, rh._M, y_fin + 15, w, refs)
        if n % 25 == 0:
            print(f"  paginas {n}/{len(df)}...")

    salida = salida or (BASE_DIR / "outputs" / f"reporte_reimputacion_cascada_{mes_ano}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()

    print()
    for ch in chequeos:
        a, d = ch["ANTES"], ch["DESPUES"]
        fa = f"{a:,.2f}" if isinstance(a, float) else f"{a:,}"
        fd = f"{d:,.2f}" if isinstance(d, float) else f"{d:,}"
        print(f"  {'OK ' if ch['OK'] else 'REV'}  {ch['CHEQUEO']:<45} {fa:>13} -> {fd:>13}")
    print(f"\nPDF -> {salida}")
    df.to_excel(salida.with_suffix(".xlsx"), index=False)
    print(f"XLSX -> {salida.with_suffix('.xlsx')}")
    return salida


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    mes_ano = args[0] if args else "2026-07"

    if "--reclamos" in sys.argv:
        # Solo los predios con un reclamo abierto de mes anterior: es la lista
        # que se va a resolver, y son ~28 predios (~1 min) contra los 570 del
        # padron (~18 min). El PDF va a un archivo aparte para no pisar el
        # normal, que sigue siendo el de los 3 conceptos de la re-imputacion.
        sys.path.insert(0, str(BASE_DIR.parent / "shared"))
        import ciclo  # noqa: E402

        mes_ciclo = ciclo.activo(default=None) or mes_ano
        predios = predios_con_reclamo(mes_ciclo)
        print(f"  {len(predios)} predios con reclamo de mes anterior en {mes_ciclo}")
        generar(mes_ano, BASE_DIR / "outputs" /
                f"reporte_reimputacion_cascada_{mes_ano}_reclamos.pdf", predios)
    else:
        generar(mes_ano)
