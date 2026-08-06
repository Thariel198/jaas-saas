"""
4b_reclamos/reporte_seguimiento.py — Reporte de cuenta imprimible, 1 página PDF por predio

Herramienta on-demand (no es parte del pipeline mensual): arma un reporte tipo
boleta para cada predio que notas_2026-07.xlsx marca TIPO_RECLAMO=CONFIRMACION,
con 3 secciones — estado actual, evolución mensual de la deuda y pagos
registrados — y lo dibuja en PDF con PyMuPDF (mismo enfoque que
seguimiento_repo.exportar_vista_pdf), para que se pueda imprimir y entregar.

Fuentes:
  - shared/seguimiento_pueblo.xlsx  (estado + evolución de MULTA/ACUERDOS/CONVENIO)
  - 5_cobranza/outputs/trazabilidad_cobranza.xlsx  (pagos reales con fecha)
  - shared/abonos_rezagados.xlsx    (pagos declarados sin registro en trazabilidad)

No incluye AGUA/MANTENIMIENTO/CORTE_RECONEXION — eso vive en arrastre_consolidado/
planilla_cobrado, fuera de seguimiento_pueblo. No es el extracto_predio() del
backfill (libro_mayor/estado_cuenta, decisión ⑫) — versión chica, solo lo que hoy
vive en seguimiento_pueblo. No linkea cada pago a un concepto específico cuando
hay 2+ pagos el mismo mes (eso necesita el ABONO_ID del ledger completo, que no
existe hoy) — sección ③ solo lista los pagos, sección ② solo la evolución de
deuda; el lector cruza los dos si le sirve.

Uso: py 4b_reclamos/reporte_seguimiento.py
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR.parent / "shared"))
import seguimiento_repo as repo  # noqa: E402

NOTAS_PATH = BASE_DIR / "pendientes_secretaria" / "notas_2026-07.xlsx"
TRAZAB_PATH = BASE_DIR.parent / "5_cobranza" / "outputs" / "trazabilidad_cobranza.xlsx"
ABONOS_REZAGADOS_PATH = BASE_DIR.parent / "shared" / "abonos_rezagados.xlsx"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_PATH = OUTPUT_DIR / "reporte_confirmacion_2026-07.pdf"

CONCEPTOS = ("CONVENIO", "ACUERDOS", "MULTA")  # orden de cascada P3-P5

log = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")

# ── Paleta (misma que exportar_vista_pdf, para consistencia visual) ──────────
_AZUL = (26/255, 82/255, 118/255)
_AZUL_BG = (235/255, 245/255, 251/255)
_GRIS = (0.42, 0.45, 0.5)
_NEGRO = (0.12, 0.16, 0.22)
_ROJO = (0.7, 0.1, 0.1)
_VERDE = (0.02, 0.37, 0.27)
_ZEBRA = (243/255, 244/255, 246/255)

_PAGE_W, _PAGE_H = 595, 842  # A4 vertical (pts)
_M = 36


def _predios_confirmacion() -> list[tuple[str, str]]:
    df = pd.read_excel(NOTAS_PATH, dtype=str)
    df = df[df["TIPO_RECLAMO"].astype(str).str.strip().str.upper() == "CONFIRMACION"]
    df = df.dropna(subset=["MZ", "LT"])
    pares = sorted({(str(r["MZ"]).strip(), str(r["LT"]).strip()) for _, r in df.iterrows()})
    return pares


def _cargar_pagos_trazabilidad() -> pd.DataFrame:
    if not TRAZAB_PATH.exists():
        return pd.DataFrame(columns=["MZ", "LT", "FECHA", "MONTO", "ORIGEN"])
    df = pd.read_excel(TRAZAB_PATH, header=1, dtype=str)
    origen = df["REFERENCIA"].fillna("").astype(str).str.strip()
    coment = df.get("COMENTARIO", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    salida = pd.DataFrame({
        "MZ": df["MZ"].astype(str).str.strip(),
        "LT": df["LT"].astype(str).str.strip(),
        "FECHA": df["FECHA"],
        "MONTO": pd.to_numeric(df["MONTO"], errors="coerce"),
        "ORIGEN": (origen + (" — " + coment).where(coment.ne(""), "")).str.strip(" —"),
    })
    return salida


def _cargar_pagos_abonos_rezagados() -> pd.DataFrame:
    if not ABONOS_REZAGADOS_PATH.exists():
        return pd.DataFrame(columns=["MZ", "LT", "FECHA", "MONTO", "ORIGEN"])
    df = pd.read_excel(ABONOS_REZAGADOS_PATH, sheet_name="abonos_rezagados", header=1, dtype=str)
    fecha = df["FECHA_REAL"].fillna("").astype(str).str.strip()
    mes = df["MES_CICLO"].fillna("").astype(str).str.strip()
    retenido = df.get("RETENIDO_POR", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    salida = pd.DataFrame({
        "MZ": df["MZ"].astype(str).str.strip(),
        "LT": df["LT"].astype(str).str.strip(),
        "FECHA": fecha.where(fecha.ne(""), mes),
        "MONTO": pd.to_numeric(df["MONTO"], errors="coerce"),
        "ORIGEN": "confirmado por directiva" + (" (" + retenido + ")").where(retenido.ne(""), ""),
    })
    return salida


def _fecha_orden(fecha: str):
    """Clave de orden cronológico real — 'DD/MM/YYYY' y 'YYYY-MM' mezclados no
    ordenan bien como texto (p.ej. '04/07/2026' < '05/06/2026' alfabéticamente,
    aunque junio es antes que julio)."""
    fecha = str(fecha).strip()
    try:
        return datetime.strptime(fecha, "%d/%m/%Y")
    except ValueError:
        pass
    try:
        return datetime.strptime(fecha, "%Y-%m")
    except ValueError:
        return datetime.min


def _pagos_predio(mz: str, lt: str, traz: pd.DataFrame, abonos: pd.DataFrame) -> pd.DataFrame:
    p1 = traz[(traz["MZ"] == mz) & (traz["LT"] == lt)]
    p2 = abonos[(abonos["MZ"] == mz) & (abonos["LT"] == lt)]
    todos = pd.concat([p1, p2], ignore_index=True)
    if todos.empty:
        return pd.DataFrame(columns=["FECHA", "MONTO", "ORIGEN"])
    todos = todos[["FECHA", "MONTO", "ORIGEN"]].copy()
    todos["_ORDEN"] = todos["FECHA"].map(_fecha_orden)
    return todos.sort_values("_ORDEN").drop(columns=["_ORDEN"]).reset_index(drop=True)


# Orden de cascada VIEJO — el que corre hoy en 5_cobranza/main.py::_descomponer_saldo
# (P3 MULTA -> P4 ACUERDOS -> P5 CONVENIO). El orden nuevo (CONVENIO->ACUERDOS->MULTA,
# dominio CA1) es para el backfill/ledger futuro, no para el código vivo.
_CASCADA_VIEJA = ("MULTA", "ACUERDOS", "CONVENIO")


def _mes_de(fecha: str) -> str:
    """'05/06/2026' -> '2026-06'; '2026-06' se queda igual."""
    fecha = str(fecha).strip()
    if len(fecha) == 7 and fecha[4] == "-":
        return fecha
    try:
        return datetime.strptime(fecha, "%d/%m/%Y").strftime("%Y-%m")
    except ValueError:
        return ""


def _repartir_pagos_por_cascada(historial: pd.DataFrame, pagos: pd.DataFrame) -> pd.DataFrame:
    """Para cada pago real, arma el desglose de la cascada vieja (MULTA->ACUERDOS->
    CONVENIO) contra el PAGO real ya confirmado por seguimiento_pueblo ese mes:
    por cada concepto que este pago tocó, cuánto necesitaba el concepto en TOTAL
    ese mes y cuánto cubrió ESTE pago puntual (para dibujar la barra de cascada,
    "necesitaba X, cubrió Y, se corta acá"). RESTO_CONSUMO = lo que sobra del pago
    sin concepto pendiente (P1/P2 agua-corte, fuera de este reporte)."""
    if pagos.empty:
        return pagos.assign(DESGLOSE=None, RESTO_CONSUMO=0.0)

    total_necesita = {}  # (mes, concepto) -> monto TOTAL confirmado ese mes (fijo)
    for _, r in historial.iterrows():
        if r["PAGO"] > 0.005:
            total_necesita[(r["MES"], r["CONCEPTO"])] = r["PAGO"]

    pagos = pagos.copy()
    pagos["MES"] = pagos["FECHA"].map(_mes_de)
    desgloses: dict = {}
    restos: dict = {}
    for mes, grupo in pagos.groupby("MES", sort=False):
        restante_concepto = {c: total_necesita.get((mes, c), 0.0) for c in _CASCADA_VIEJA}
        for idx, fila in grupo.iterrows():
            restante_pago = fila["MONTO"] if not pd.isna(fila["MONTO"]) else 0.0
            desglose = []
            for concepto in _CASCADA_VIEJA:
                necesidad_total = total_necesita.get((mes, concepto), 0.0)
                if necesidad_total <= 0.005:
                    continue  # no debía nada de este concepto ese mes, no se muestra
                if restante_pago <= 0.005 or restante_concepto[concepto] <= 0.005:
                    desglose.append({"concepto": concepto, "necesitaba": necesidad_total, "cubrio": 0.0})
                    continue
                aplicar = min(restante_pago, restante_concepto[concepto])
                desglose.append({"concepto": concepto, "necesitaba": necesidad_total, "cubrio": aplicar})
                restante_concepto[concepto] -= aplicar
                restante_pago -= aplicar
            desgloses[idx] = desglose
            restos[idx] = restante_pago if restante_pago > 0.005 else 0.0

    pagos["DESGLOSE"] = pagos.index.map(desgloses)
    pagos["RESTO_CONSUMO"] = pagos.index.map(restos).fillna(0.0)
    return pagos.drop(columns=["MES"])


def _resumen_y_historial(mz: str, lt: str, eventos: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """resumen: 1 fila por concepto con CARGADO/PAGADO/DEBE (estado actual).
    historial: 1 fila por (mes, concepto) en lenguaje simple — DEBIA/PAGO/QUEDO,
    agrupando lo que haya pasado ese mes (nada de CARGO/PAGO/AJUSTE, el usuario
    no sabe de eventos internos, solo "cuánto debía / pagó / quedó")."""
    propio = eventos[
        (eventos["MZ"].astype(str).str.strip() == mz) &
        (eventos["LT"].astype(str).str.strip() == lt)
    ].copy()

    resumen_filas = []
    historial_filas = []
    for concepto in CONCEPTOS:
        sub = propio[propio["CONCEPTO"].astype(str).str.strip().str.upper() == concepto].sort_values(
            ["MES", "TIMESTAMP"])
        if sub.empty:
            continue
        cargado = float(sub["CARGO"].fillna(0).sum())
        pagado = float(sub["PAGO"].fillna(0).sum())
        debe = float(sub.iloc[-1]["SALDO"])
        resumen_filas.append({"CONCEPTO": concepto, "CARGADO": cargado, "PAGADO": pagado, "DEBE": debe})

        for mes in sorted(sub["MES"].astype(str).unique()):
            del_mes = sub[sub["MES"].astype(str) == mes]
            debia = float(del_mes["CARGO"].fillna(0).sum() + del_mes["AJUSTE"].fillna(0).sum())
            pago = float(del_mes["PAGO"].fillna(0).sum())
            quedo = float(del_mes.iloc[-1]["SALDO"])
            historial_filas.append({"MES": mes, "CONCEPTO": concepto, "DEBIA": debia,
                                     "PAGO": pago, "QUEDO": quedo})

    resumen = pd.DataFrame(resumen_filas, columns=["CONCEPTO", "CARGADO", "PAGADO", "DEBE"])
    historial = pd.DataFrame(historial_filas, columns=["MES", "CONCEPTO", "DEBIA", "PAGO", "QUEDO"])
    if not historial.empty:
        orden = {c: i for i, c in enumerate(CONCEPTOS)}
        historial = historial.sort_values(
            ["MES", "CONCEPTO"], key=lambda s: s.map(orden) if s.name == "CONCEPTO" else s
        ).reset_index(drop=True)
    return resumen, historial


def _cargos_por_concepto(mz: str, lt: str, eventos: pd.DataFrame, resumen: pd.DataFrame) -> pd.DataFrame:
    """1 fila por CARGO real (no agregado) — CONCEPTO | MES_CARGO | MONTO | APLICADO |
    ESTADO, como formato_vista_estado_cuenta.html. Reparte el PAGADO agregado del
    concepto en orden FIFO contra sus cargos (más viejo primero) — aproximación
    razonable porque en la práctica casi siempre hay 1 solo cargo por concepto."""
    propio = eventos[
        (eventos["MZ"].astype(str).str.strip() == mz) &
        (eventos["LT"].astype(str).str.strip() == lt) &
        (eventos["TIPO_EVENTO"] == "CARGO")
    ]
    pagado_por_concepto = {r["CONCEPTO"]: r["PAGADO"] for _, r in resumen.iterrows()}

    filas = []
    for concepto in CONCEPTOS:
        cargos = propio[propio["CONCEPTO"].astype(str).str.strip().str.upper() == concepto].sort_values("MES")
        restante = pagado_por_concepto.get(concepto, 0.0)
        for _, c in cargos.iterrows():
            monto = float(c["CARGO"] or 0.0)
            aplicado = min(monto, restante) if restante > 0 else 0.0
            restante -= aplicado
            estado = "CUBIERTO" if aplicado >= monto - 0.005 else "ABIERTO"
            filas.append({"CONCEPTO": concepto, "MES_CARGO": c["MES"], "MONTO": monto,
                          "APLICADO": aplicado, "ESTADO": estado})
    return pd.DataFrame(filas, columns=["CONCEPTO", "MES_CARGO", "MONTO", "APLICADO", "ESTADO"])


def _historia_aplicaciones(pagos: pd.DataFrame) -> pd.DataFrame:
    """Aplana el DESGLOSE de cada pago a 1 fila por (pago, concepto cubierto) —
    FECHA | CANAL/ORIGEN | CONCEPTO cubierto, como la tabla de aplicaciones del
    formato real (sin la barra visual, tabla simple)."""
    filas = []
    for _, p in pagos.iterrows():
        for d in (p["DESGLOSE"] or []):
            if d["cubrio"] <= 0.005:
                continue
            filas.append({"FECHA": p["FECHA"], "ORIGEN": p["ORIGEN"],
                          "CONCEPTO": f"{d['concepto']} — S/ {d['cubrio']:,.2f}"})
    return pd.DataFrame(filas, columns=["FECHA", "ORIGEN", "CONCEPTO"])


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, (int, float)):
        return f"{v:,.2f}"
    return str(v)


def _tabla(page, x: float, y: float, w: float, headers: list[str], anchos: list[float],
           filas: list[list], col_saldo: int | None = None) -> float:
    """Dibuja una tabla simple, devuelve el y donde terminó."""
    rh = 16
    page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh
    for n, fila in enumerate(filas):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_ZEBRA, color=None)
        cx = x
        for i, (v, cw) in enumerate(zip(fila, anchos)):
            texto = _fmt(v)
            color = _NEGRO
            font = "helv"
            if col_saldo is not None and i == col_saldo and isinstance(v, (int, float)) and not pd.isna(v):
                color = _ROJO if v > 0.005 else _VERDE
                font = "hebo"
            if i == 0:
                page.insert_text((cx + 4, y + rh - 5), texto[:40], fontsize=8, fontname=font, color=color)
            else:
                tw = fitz.get_text_length(texto, fontname=font, fontsize=8)
                page.insert_text((cx + cw - 4 - tw, y + rh - 5), texto, fontsize=8, fontname=font, color=color)
            cx += cw
        y += rh
    if not filas:
        page.insert_text((x + 4, y + rh - 5), "(sin datos)", fontsize=8, fontname="helv", color=_GRIS)
        y += rh
    return y


_BARRA_W = 160  # ancho máximo de la barra de cascada (pts) — escala relativa a la necesidad más grande de la página
_AMBAR = (0.72, 0.45, 0.02)
_ROJO_BG = (0.98, 0.90, 0.90)


def _dibujar_evolucion_barras(page, x: float, y: float, w: float,
                               resumen: pd.DataFrame, historial: pd.DataFrame) -> float:
    """1 barra horizontal por concepto: cuánto se cargó (100% del ancho) vs cuánto
    queda debiendo hoy (la porción roja) — de un vistazo, cuánto bajó y cuánto falta."""
    if resumen.empty:
        page.insert_text((x, y + 9), "(sin conceptos registrados)", fontsize=8, fontname="helv", color=_GRIS)
        return y + 14

    barra_w = w - 210
    for _, r in resumen.iterrows():
        concepto, cargado, debe = r["CONCEPTO"], r["CARGADO"], r["DEBE"]
        meses = sorted(historial[historial["CONCEPTO"] == concepto]["MES"].astype(str).unique())
        rango = f"({meses[0]} → {meses[-1]})" if meses else ""

        page.insert_text((x, y + 9), concepto, fontsize=8, fontname="hebo", color=_NEGRO)
        bx = x + 90
        page.draw_rect(fitz.Rect(bx, y, bx + barra_w, y + 11), fill=_ROJO_BG, color=None)
        if cargado > 0.005:
            frac_pagado = max(0.0, min(1.0, (cargado - debe) / cargado))
            if frac_pagado > 0:
                page.draw_rect(fitz.Rect(bx, y, bx + barra_w * frac_pagado, y + 11), fill=_VERDE, color=None)
        etiqueta = f"{cargado:,.2f} -> {debe:,.2f}  {rango}"
        page.insert_text((bx + barra_w + 8, y + 9), etiqueta, fontsize=7.5, fontname="helv",
                          color=_ROJO if debe > 0.005 else _VERDE)
        y += 17
    return y


def _dibujar_cascada_pago(page, x: float, y: float, w: float, fila) -> float:
    """1 pago = 1 bloque: cabecera (fecha/monto/origen) + una barra por concepto
    que tocó (necesitaba vs cubrió) + nota de resto a consumo si sobró plata."""
    fecha, monto, origen = fila["FECHA"], fila["MONTO"], fila["ORIGEN"]
    if pd.isna(fecha) or str(fecha).strip().lower() == "nan":
        fecha = "(fecha no registrada)"
    desglose = fila["DESGLOSE"] or []
    resto = fila["RESTO_CONSUMO"]

    cab = f"{fecha}  ·  S/ {monto:,.2f}  ·  {origen}"
    page.insert_text((x, y + 10), cab, fontsize=8.5, fontname="hebo", color=_NEGRO)
    y += 16

    if not desglose and resto <= 0.005:
        page.insert_text((x + 12, y + 9), "(sin multa/acuerdos/convenio pendiente ese mes)",
                          fontsize=7.5, fontname="helv", color=_GRIS)
        y += 14

    for d in desglose:
        necesitaba, cubrio = d["necesitaba"], d["cubrio"]
        frac = min(cubrio / necesitaba, 1.0) if necesitaba > 0.005 else 0.0
        barra_w = _BARRA_W
        bx = x + 100
        page.draw_rect(fitz.Rect(bx, y, bx + barra_w, y + 10), fill=(0.96, 0.93, 0.93), color=None)
        if frac > 0:
            color_lleno = _VERDE if frac >= 0.999 else _AMBAR
            page.draw_rect(fitz.Rect(bx, y, bx + barra_w * frac, y + 10), fill=color_lleno, color=None)
        page.insert_text((x + 12, y + 9), d["concepto"], fontsize=7.5, fontname="helv", color=_NEGRO)
        etiqueta = f"cubrió {cubrio:,.2f} de {necesitaba:,.2f}"
        page.insert_text((bx + barra_w + 8, y + 9), etiqueta, fontsize=7.5, fontname="helv", color=_GRIS)
        y += 13

    if resto > 0.005:
        page.insert_text((x + 12, y + 9), f"resto S/ {resto:,.2f} a tu consumo de agua (no en este reporte)",
                          fontsize=7.5, fontname="helv", color=_GRIS)
        y += 13

    return y + 10


def _tabla_estado(page, x: float, y: float, w: float, headers: list[str], anchos: list[float],
                   filas: list[list], col_estado: int) -> float:
    """Como _tabla pero con una columna ESTADO tipo semáforo (CUBIERTO=verde,
    ABIERTO=rojo), igual que formato_vista_estado_cuenta.html."""
    rh = 16
    page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh
    for n, fila in enumerate(filas):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_ZEBRA, color=None)
        cx = x
        for i, (v, cw) in enumerate(zip(fila, anchos)):
            if i == col_estado:
                bg = _VERDE if v == "CUBIERTO" else (0.98, 0.85, 0.85)
                fg = (1, 1, 1) if v == "CUBIERTO" else _ROJO
                page.draw_rect(fitz.Rect(cx + 2, y + 2, cx + cw - 2, y + rh - 2), fill=bg, color=None)
                tw = fitz.get_text_length(v, fontname="hebo", fontsize=7.5)
                page.insert_text((cx + cw/2 - tw/2, y + rh - 5), v, fontsize=7.5, fontname="hebo", color=fg)
            else:
                texto = _fmt(v)
                if i == 0:
                    page.insert_text((cx + 4, y + rh - 5), texto[:40], fontsize=8, fontname="helv", color=_NEGRO)
                else:
                    tw = fitz.get_text_length(texto, fontname="helv", fontsize=8)
                    page.insert_text((cx + cw - 4 - tw, y + rh - 5), texto, fontsize=8, fontname="helv", color=_NEGRO)
            cx += cw
        y += rh
    if not filas:
        page.insert_text((x + 4, y + rh - 5), "(sin datos)", fontsize=8, fontname="helv", color=_GRIS)
        y += rh
    return y


def _dibujar_pagina(doc, mz: str, lt: str, nombre: str, resumen: pd.DataFrame,
                     historial: pd.DataFrame, pagos: pd.DataFrame, cargos: pd.DataFrame,
                     aplicaciones: pd.DataFrame) -> None:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M

    # Header oscuro — igual a .predio-header de formato_vista_estado_cuenta.html
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text((_M + 10, y + 20), f"Predio {mz}-{lt} — {nombre or '(sin nombre)'}",
                      fontsize=12, fontname="hebo", color=(1, 1, 1))
    hoy = datetime.now().strftime("%d/%m/%Y")
    tw = fitz.get_text_length(f"Emitido {hoy}", fontname="helv", fontsize=8)
    page.insert_text((_M + w - 10 - tw, y + 20), f"Emitido {hoy}", fontsize=8, fontname="helv", color=(1, 1, 1))
    y += 42

    page.insert_text((_M, y), "CARGO — deuda por concepto", fontsize=10, fontname="hebo", color=_AZUL)
    y += 6
    filas = [[c["CONCEPTO"], c["MES_CARGO"], c["MONTO"], c["APLICADO"], c["ESTADO"]] for _, c in cargos.iterrows()]
    y = _tabla_estado(page, _M, y, w, ["Concepto", "Mes cargo", "Monto", "Aplicado", "Estado"],
                       [w - 320, 90, 80, 80, 70], filas, col_estado=4)
    y += 20

    page.insert_text((_M, y), "HISTORIA DE APLICACIONES — ¿a dónde fue cada sol?", fontsize=10,
                      fontname="hebo", color=_AZUL)
    y += 6
    filas = [[a["FECHA"], a["ORIGEN"], a["CONCEPTO"]] for _, a in aplicaciones.iterrows()]
    y = _tabla(page, _M, y, w, ["Fecha", "Origen", "Concepto cubierto"],
               [90, w - 340, 250], filas)
    y += 20

    saldo_total = float(resumen["DEBE"].sum()) if not resumen.empty else 0.0
    color_saldo = _ROJO if saldo_total > 0.005 else _VERDE
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 24), fill=color_saldo, color=None)
    texto_saldo = f"SALDO PENDIENTE: S/ {saldo_total:,.2f}"
    page.insert_text((_M + w - 10 - fitz.get_text_length(texto_saldo, fontname="hebo", fontsize=11),
                      y + 17), texto_saldo, fontsize=11, fontname="hebo", color=(1, 1, 1))
    y += 40

    page.insert_text((_M, y), "EVOLUCIÓN DE TU DEUDA (de cuánto a cuánto bajó)", fontsize=10,
                      fontname="hebo", color=_AZUL)
    y += 12
    y = _dibujar_evolucion_barras(page, _M, y, w, resumen, historial)
    y += 20

    page.insert_text((_M, y), "DETALLE DE CADA PAGO (cascada de aplicación)", fontsize=10,
                      fontname="hebo", color=_AZUL)
    y += 12
    if pagos.empty:
        page.insert_text((_M, y), "(sin pagos registrados)", fontsize=8, fontname="helv", color=_GRIS)
        y += 14
    for _, fila in pagos.iterrows():
        y = _dibujar_cascada_pago(page, _M, y, w, fila)

    page.insert_text((_M, _PAGE_H - 20),
                      "No incluye agua/mantenimiento/corte — vive fuera de seguimiento_pueblo. "
                      "Formato basado en libro_mayor/estado_cuenta/docs/formato_vista_estado_cuenta.html.",
                      fontsize=6.5, fontname="helv", color=_GRIS)


def generar(limite: list[tuple[str, str]] | None = None, salida: Path | None = None) -> Path:
    predios = limite if limite is not None else _predios_confirmacion()
    log.info(f"{len(predios)} predio(s) a reportar")

    nombres = repo._lookup_nombres()
    eventos = repo._leer_eventos()
    traz = _cargar_pagos_trazabilidad()
    abonos = _cargar_pagos_abonos_rezagados()

    doc = fitz.open()
    for mz, lt in predios:
        resumen, historial = _resumen_y_historial(mz, lt, eventos)
        pagos = _pagos_predio(mz, lt, traz, abonos)
        pagos = _repartir_pagos_por_cascada(historial, pagos)
        cargos = _cargos_por_concepto(mz, lt, eventos, resumen)
        aplicaciones = _historia_aplicaciones(pagos)
        _dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), resumen, historial, pagos, cargos, aplicaciones)

    salida = salida or OUTPUT_PATH
    OUTPUT_DIR.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    log.info(f"Reporte guardado -> {salida} ({len(predios)} página(s))")
    return salida


if __name__ == "__main__":
    generar()
