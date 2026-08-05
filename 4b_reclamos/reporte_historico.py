"""
4b_reclamos/reporte_historico.py — Tabla horizontal mes a mes por predio, desde
octubre 2025 hasta julio 2026.

Dos fuentes distintas, unidas en una sola tabla:
  - Oct 2025 -> mayo 2026: obligaciones/inputs/planillas anteriores/*.xlsx,
    hoja "Cobranza" — cada mes tiene su propio esquema de columnas (nombres
    cambian: "Convenio" vs "Medidor", "Mant." vs "Mantenimiento", etc.), se
    leen con match flexible por palabra clave, no por nombre exacto.
  - Junio 2026 en adelante: shared/seguimiento_pueblo.xlsx (MULTA/ACUERDOS/
    CONVENIO) + 5_cobranza/outputs/planilla_cobrado.xlsx (consumo), vía la
    cascada ya armada en reporte_seguimiento.py.

Reglas:
  - Si lo pagado (Yape+Efectivo) es MENOR a lo debido, se muestra tal cual
    viene del archivo (así se manejaba hasta mayo 2026: el exceso/deuda no
    cerrado pasaba a MES_ANTERIOR del mes siguiente, incluso negativo).
  - Desde junio 2026 la política cambió: un exceso ya NO se aplica solo ni se
    muestra — se retiene hasta que el vecino reclame. No se expone ningún
    sobrante en el reporte para estos meses.

Uso: py 4b_reclamos/reporte_historico.py
"""

import sys
import unicodedata
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR.parent / "shared"))
import seguimiento_repo as repo  # noqa: E402
import ciclo  # noqa: E402

HIST_DIR = BASE_DIR.parent / "obligaciones" / "inputs" / "planillas anteriores"
REASIGNACIONES_PATH = BASE_DIR.parent / "0_padron" / "reasignaciones_candidata.xlsx"

# (archivo, MES_ANO) — orden cronológico real, no alfabético del nombre de archivo
_ARCHIVOS_HISTORICOS = [
    ("octubre-planilla 2025-08-11 A 2025-09-11.xlsx", "2025-10"),
    ("noviembre-planilla 2025-09-11 A 2025-10-11.xlsx", "2025-11"),
    ("diciembre-planilla 2025-10-11 A 2025-11-11 - 26.xlsx", "2025-12"),
    ("enero-planilla 2025-11-11 A 2025-12-11.xlsx", "2026-01"),
    ("febrero-planilla 2025-12-11 A 2026-01-10.xlsx", "2026-02"),
    ("marzo-planilla 2026-01-11 A 2026.xlsx", "2026-03"),
    ("abril-planilla 2026-02-11 A 2026-03-10.xlsx", "2026-04"),
    ("mayo-planilla 2026-03-11 A 2026-04-10 (2).xlsx", "2026-05"),
]

CONCEPTOS_TABLA = ["CONSUMO", "MANT", "MES_ANT", "CORTE", "CONVENIO", "MULTA", "ACUERDOS"]


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return s.lower().replace("\n", " ").strip()


def _col(df: pd.DataFrame, *candidatos: str):
    """Busca una columna por coincidencia flexible (normalizada, sin tildes/mayúsculas).
    Prueba los candidatos en orden, devuelve el primer nombre de columna real que matchea.
    Si dos columnas normalizan igual (ej. diciembre trae 'MES ANTERIOR' Y 'Mes anterior'
    duplicadas), se queda con la PRIMERA en orden de aparición — la duplicada tardía
    suele ser una columna suelta casi siempre vacía, no la que se sumó al Total real."""
    normed = {}
    for c in df.columns:
        k = _norm(c)
        if k not in normed:
            normed[k] = c
    for cand in candidatos:
        cn = _norm(cand)
        if cn in normed:
            return normed[cn]
    # match parcial (contiene) si no hubo exacto
    for cand in candidatos:
        cn = _norm(cand)
        for k, real in normed.items():
            if cn in k:
                return real
    return None


def _val(row, df, *candidatos) -> float:
    c = _col(df, *candidatos)
    if c is None:
        return 0.0
    v = row.get(c)
    return float(v) if pd.notna(v) else 0.0


def _numf(v) -> float:
    """float NaN-safe — 'v or 0' no sirve porque NaN es truthy en Python."""
    return float(v) if pd.notna(v) else 0.0


# Cruces confirmados a mano (nombre exacto, sin ambigüedad) que no están en
# reasignaciones_candidata.xlsx — ir agregando aquí solo los que se verifiquen 1 a 1.
_MAPA_RAW_EXTRA = {
    ("G1", "13"): ("H1", "1"),  # YOLIT VILLANUEVA ALEGRIA, nombre exacto sin ambigüedad
    ("G", "16C"): ("L", "10"),  # ELOY (ALBERTO) SIGUEÑAS UGARTE, confirmado por el usuario
}


def _cargar_mapa_raw() -> dict[tuple[str, str], tuple[str, str]]:
    """(MZ,LT) actual -> (MZ,LT) como aparecía en las planillas anteriores a la
    corrección COFOPRI/manual — B-20 hoy fue B-29 en esos archivos, C-43 fue
    C-45, B-19 fue B-15, etc. (0_padron/reasignaciones_candidata.xlsx)."""
    df = pd.read_excel(REASIGNACIONES_PATH, sheet_name="REASIGNACIONES", dtype=str)
    mapa = {(str(r["SYS_MZ"]).strip(), str(r["SYS_LT"]).strip()):
            (str(r["RAW_MZ"]).strip(), str(r["RAW_LT"]).strip()) for _, r in df.iterrows()}
    mapa.update(_MAPA_RAW_EXTRA)
    return mapa


def _cargar_historicos() -> dict[str, pd.DataFrame]:
    """Lee cada uno de los 8 archivos de planillas anteriores UNA sola vez
    (evita reabrir el mismo Excel por cada predio en una corrida por lote)."""
    dfs = {}
    for archivo, mes in _ARCHIVOS_HISTORICOS:
        ruta = HIST_DIR / archivo
        if ruta.exists():
            dfs[mes] = pd.read_excel(ruta, sheet_name="Cobranza", header=0)
    return dfs


def _nombres_comparten_palabra(a: str, b: str) -> bool:
    if not a or not b:
        return True  # sin dato de uno de los dos lados, no se puede descartar -- no bloquea
    norm = lambda s: set(unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().upper().split())
    return bool(norm(a) & norm(b))


def _fila_historica(mz: str, lt: str, df: pd.DataFrame | None, mes: str, nombre_actual: str = "") -> dict | None:
    if df is None:
        return None
    mzcol, ltcol = _col(df, "MZ"), _col(df, "LT")
    fila = df[(df[mzcol].astype(str).str.strip() == mz) & (df[ltcol].astype(str).str.strip() == lt)]
    if fila.empty:
        return None
    r = fila.iloc[0]

    nombre_col = _col(df, "NOMBRES", "Nombres")
    nombre_hist = r.get(nombre_col, "") if nombre_col else ""
    if not _nombres_comparten_palabra(nombre_actual, nombre_hist):
        # La etiqueta MZ-LT fue reasignada y no tenemos su traducción — mostrar
        # esto sería el historial de OTRA persona. Mejor sin dato que equivocado.
        return {"MES": mes, "CONSUMO": 0, "MANT": 0, "MES_ANT": 0, "CORTE": 0,
                "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "TOTAL": 0.0, "PAGO_COMPLETO": False,
                "NOTA": f"lote reasignado, historial de {nombre_hist} no es tuyo"}

    estado_col = _col(df, "Estado")
    estado = str(r.get(estado_col, "")).strip().lower() if estado_col else ""
    pagado = estado == "c"

    consumo = _val(r, df, "Total mes actual", "Mes actual")
    mant = _val(r, df, "Mant.", "Mantenimiento")
    mes_ant = _val(r, df, "MES ANTERIOR", "Mes anterior")
    corte = (_val(r, df, "Corte y reconexion", "Corte y reconexión")
             + _val(r, df, "Reactivacion", "Reactivación"))  # P-9 2026-03: fee de reconexión con nombre distinto
    convenio = _val(r, df, "Convenio") + _val(r, df, "Medidor")
    multa = (_val(r, df, "Multa (faena + reunión)") or _val(r, df, "Multa")
             or (_val(r, df, "Reunión") + _val(r, df, "Faena")))
    acuerdos = _val(r, df, "Techado")

    if not pagado:
        # No pagó nada ese mes -> no se muestra ningun concepto como cubierto
        return {"MES": mes, "CONSUMO": 0, "MANT": 0, "MES_ANT": 0, "CORTE": 0,
                "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "TOTAL": 0.0, "PAGO_COMPLETO": False,
                "NOTA": "No pago nada"}

    total = consumo + mant + mes_ant + corte + convenio + multa + acuerdos
    return {"MES": mes, "CONSUMO": consumo, "MANT": mant, "MES_ANT": mes_ant, "CORTE": corte,
            "CONVENIO": convenio, "MULTA": multa, "ACUERDOS": acuerdos, "TOTAL": total,
            "PAGO_COMPLETO": True, "NOTA": ""}


REPO_JUNIO = Path(r"C:\Users\wilde\PycharmProjects\Junio\jass_system - junio")


def _planilla_junio() -> Path:
    """Ciclo cerrado: el nombre canónico lleva el periodo
    (planilla_cobrado_2026-06.xlsx). shared/ciclo.resolver acepta también los
    nombres históricos, así que un rename del archivo ya no vuelve a dejar todas
    las filas de junio sin consumo ni referencia de pago (bug del 05/08/2026)."""
    return ciclo.resolver(REPO_JUNIO / "5_cobranza" / "outputs", "planilla_cobrado", "2026-06")


_JUNIO_PLANILLA_COBRADO = _planilla_junio()
_dfp_junio_cache: pd.DataFrame | None = None


def _cargar_dfp_junio() -> pd.DataFrame | None:
    """planilla_cobrado.xlsx de junio vive en el repo 'jass_system - junio'
    (ciclo ya cerrado, congelado ahí) -- no en el activo, que solo tiene el
    ciclo vigente. Cacheado a nivel modulo: se reusa entre predios en una
    corrida por lote.

    Si el repo cerrado no esta donde dice el path, se AVISA: cuando fallaba en
    silencio (04/08/2026, el repo se movio a PycharmProjects\\Junio\\) toda fila
    de junio salia sin consumo/mant y sin referencia de pago -- se leia como un
    pago fantasma (caso K-9: mostraba multa 30 en vez de los S/38 reales)."""
    global _dfp_junio_cache
    if _dfp_junio_cache is None:
        if _JUNIO_PLANILLA_COBRADO.exists():
            _dfp_junio_cache = pd.read_excel(_JUNIO_PLANILLA_COBRADO, sheet_name="planilla_cobrado", header=1)
        else:
            print(f"  AVISO: no se encontro el ciclo cerrado de junio -> {_JUNIO_PLANILLA_COBRADO}\n"
                  f"         las filas de 2026-06 van a salir sin consumo/mantenimiento.")
            _dfp_junio_cache = pd.DataFrame()
    return _dfp_junio_cache


def _datos_ciclo(mz: str, lt: str, dfp: pd.DataFrame) -> dict | None:
    """Consumo/mant/mes_ant/corte de UN ciclo (junio o julio), capados por
    cascada P1(agua)->P2(mes anterior)->P2b(corte) contra lo realmente pagado
    ese ciclo -- si pagó menos de lo debido, eso es lo que se acredita, no de
    mas."""
    fila = dfp[(dfp["MZ"].astype(str).str.strip() == mz) & (dfp["LT"].astype(str).str.strip() == lt)]
    if fila.empty:
        return None
    r = fila.iloc[0]
    mes_ano = str(r.get("MES_ANO", "")).strip()
    consumo_debido = _numf(r.get("MES_ACTUAL")) + _numf(r.get("MANTENIMIENTO"))
    mes_ant_debido = _numf(r.get("MES_ANTERIOR"))
    corte_debido = _numf(r.get("CORTE_RECONEXION"))
    total_pagado = _numf(r.get("MONTO_YAPE")) + _numf(r.get("MONTO_EFECTIVO"))
    consumo = min(consumo_debido, total_pagado)
    restante = max(0.0, total_pagado - consumo)
    mes_ant = min(mes_ant_debido, restante)
    restante = max(0.0, restante - mes_ant)
    corte = min(corte_debido, restante)
    return {"mes_ano": mes_ano, "consumo": consumo, "mes_ant": mes_ant, "corte": corte,
            "hubo_pago": total_pagado > 0.005}


def _filas_recientes(mz: str, lt: str, eventos: pd.DataFrame, dfp: pd.DataFrame) -> list[dict]:
    """Junio en adelante: seguimiento_pueblo + planilla_cobrado. Reusa la logica
    ya construida en reporte_seguimiento.py (cascada vieja, sin exponer exceso)."""
    sys.path.insert(0, str(BASE_DIR))
    import reporte_seguimiento as rs

    resumen, historial = rs._resumen_y_historial(mz, lt, eventos)

    # Un ciclo por cada planilla_cobrado disponible -- julio (activo, pasado
    # por el llamador) y junio (repo 'jass_system - junio', ciclo cerrado).
    # Antes solo se calculaba consumo/mant/mes_ant/corte para julio ("el
    # ciclo actual"); junio, aunque ya vive en el ledger, siempre quedaba en
    # 0 en esos 4 campos -- bug: el total de junio no cuadraba contra lo
    # realmente pagado (ej. C-16: pago 33 en junio, la tabla solo mostraba 25
    # de multa, los 8 de consumo+mant desaparecian).
    datos_por_ciclo: dict[str, dict] = {}
    for fuente in (dfp, _cargar_dfp_junio()):
        if fuente is None or fuente.empty:
            continue
        d = _datos_ciclo(mz, lt, fuente)
        if d and d["mes_ano"]:
            datos_por_ciclo[d["mes_ano"]] = d

    # el mes de cada ciclo entra a la tabla si hubo PAGO real (yape/efectivo),
    # aunque no haya evento MULTA/ACUERDOS/CONVENIO ese mes en seguimiento_pueblo
    # (antes: se exigia un evento de pueblo para crear la fila, y un pago de solo
    # consumo -como A-6 en julio- nunca generaba fila, aunque si hubiera pagado)
    meses = set(historial["MES"].astype(str).unique())
    for mes_ano, d in datos_por_ciclo.items():
        if d["hubo_pago"]:
            meses.add(mes_ano)

    filas = []
    for mes in sorted(meses):
        del_mes = historial[historial["MES"] == mes]
        fila_d = {"MES": mes, "CONSUMO": 0, "MANT": 0, "MES_ANT": 0, "CORTE": 0,
                  "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "NOTA": ""}
        for _, r in del_mes.iterrows():
            concepto = r["CONCEPTO"]
            if concepto == "MULTA":
                fila_d["MULTA"] = r["PAGO"]
            elif concepto == "ACUERDOS":
                fila_d["ACUERDOS"] = r["PAGO"]
            elif concepto == "CONVENIO":
                fila_d["CONVENIO"] = r["PAGO"]
        d = datos_por_ciclo.get(mes)
        if d:
            fila_d["CONSUMO"] = d["consumo"]
            fila_d["MES_ANT"] = d["mes_ant"]
            fila_d["CORTE"] = d["corte"]
        fila_d["TOTAL"] = sum(fila_d[c] for c in CONCEPTOS_TABLA)
        fila_d["PAGO_COMPLETO"] = fila_d["TOTAL"] > 0.005
        filas.append(fila_d)
    return filas


def tabla_predio(mz: str, lt: str, historicos: dict | None = None,
                  eventos: pd.DataFrame | None = None, dfp: pd.DataFrame | None = None,
                  mapa_raw: dict | None = None, nombre_actual: str = "") -> pd.DataFrame:
    historicos = historicos if historicos is not None else _cargar_historicos()
    eventos = eventos if eventos is not None else repo._leer_eventos()
    mapa_raw = mapa_raw if mapa_raw is not None else _cargar_mapa_raw()
    if not nombre_actual:
        nombre_actual = repo._lookup_nombres().get((mz, lt), "")
    if dfp is None:
        f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
        dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)

    mz_hist, lt_hist = mapa_raw.get((mz, lt), (mz, lt))
    filas = []
    for archivo, mes in _ARCHIVOS_HISTORICOS:
        f = _fila_historica(mz_hist, lt_hist, historicos.get(mes), mes, nombre_actual)
        if f is None:
            filas.append({"MES": mes, "CONSUMO": 0, "MANT": 0, "MES_ANT": 0, "CORTE": 0,
                          "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "TOTAL": 0.0, "PAGO_COMPLETO": False,
                          "NOTA": "(sin dato en el archivo de este mes)"})
        else:
            filas.append(f)
    filas.extend(_filas_recientes(mz, lt, eventos, dfp))
    return pd.DataFrame(filas, columns=["MES", "CONSUMO", "MANT", "MES_ANT", "CORTE",
                                         "CONVENIO", "MULTA", "ACUERDOS", "TOTAL", "PAGO_COMPLETO", "NOTA"])


def _predios_a_reportar() -> list[tuple[str, str]]:
    """Filtra CONFIRMACION a los predios que realmente necesitan el reporte —
    si ya está al día (SALDO=0 en seguimiento_pueblo), no hace falta imprimirle
    nada, su boleta normal ya lo deja conforme."""
    sys.path.insert(0, str(BASE_DIR))
    import reporte_seguimiento as rs

    eventos = repo._leer_eventos()
    predios = rs._predios_confirmacion()
    a_reportar = []
    for mz, lt in predios:
        resumen, _ = rs._resumen_y_historial(mz, lt, eventos)
        saldo = float(resumen["DEBE"].sum()) if not resumen.empty else 0.0
        if saldo > 0.005:
            a_reportar.append((mz, lt))
    return a_reportar


_AZUL = (26/255, 82/255, 118/255)
_AZUL_BG = (235/255, 245/255, 251/255)
_GRIS = (0.42, 0.45, 0.5)
_NEGRO = (0.12, 0.16, 0.22)
_VERDE = (0.02, 0.37, 0.27)
_ZEBRA = (243/255, 244/255, 246/255)
_PAGE_W, _PAGE_H = 842, 595  # A4 horizontal (pts) — tabla ancha
_M = 30


def _dibujar_tabla_historico(page, x: float, y: float, w: float, tabla: pd.DataFrame) -> float:
    headers = ["Mes", "Pago", "Consumo", "Mant.", "Mes ant.", "Corte", "Convenio", "Multa", "Acuerdos", "TOTAL PAGADO"]
    resto = w - 70 - 40 - 110
    col = resto / 7
    anchos = [70, 40, col, col, col, col, col, col, col, 110]

    rh = 18
    page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh - 6), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh

    for n, (_, r) in enumerate(tabla.iterrows()):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_ZEBRA, color=None)
        cx = x
        page.insert_text((cx + 4, y + rh - 6), str(r["MES"]), fontsize=8, fontname="hebo", color=_NEGRO)
        cx += anchos[0]

        check = "OK" if r["PAGO_COMPLETO"] else "-"
        color_chk = _VERDE if r["PAGO_COMPLETO"] else _GRIS
        page.insert_text((cx + 8, y + rh - 6), check, fontsize=8, fontname="hebo", color=color_chk)
        cx += anchos[1]

        for i, campo in enumerate(CONCEPTOS_TABLA):
            v = r[campo]
            texto = f"{v:,.2f}" if abs(v) > 0.005 else "—"
            tw = fitz.get_text_length(texto, fontname="helv", fontsize=8)
            page.insert_text((cx + anchos[2 + i] - 6 - tw, y + rh - 6), texto, fontsize=8, fontname="helv",
                              color=_NEGRO)
            cx += anchos[2 + i]

        texto_total = f"S/ {r['TOTAL']:,.2f}" if r["TOTAL"] > 0.005 else "—"
        tw = fitz.get_text_length(texto_total, fontname="hebo", fontsize=8)
        page.insert_text((cx + anchos[-1] - 6 - tw, y + rh - 6), texto_total, fontsize=8, fontname="hebo",
                          color=_NEGRO)

        if r["NOTA"]:
            pass  # nota se omite en la tabla, ya lo dice el check "-"

        y += rh
    return y


def _dibujar_pagina(doc, mz: str, lt: str, nombre: str, tabla: pd.DataFrame) -> None:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M

    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text((_M + 10, y + 20), f"Predio {mz}-{lt} — {nombre or '(sin nombre)'} — historial mensual",
                      fontsize=12, fontname="hebo", color=(1, 1, 1))
    y += 42

    page.insert_text((_M, y), "Pago por mes y concepto (octubre 2025 -> julio 2026)", fontsize=9,
                      fontname="helv", color=_GRIS)
    y += 14
    y = _dibujar_tabla_historico(page, _M, y, w, tabla)
    y += 20

    nota = ("Desde junio 2026 el exceso pagado ya no se muestra ni se aplica solo — se retiene hasta "
            "que el vecino reclame. Meses sin \"OK\" son meses sin pago registrado en el archivo de esa fecha.")
    page.insert_text((_M, y), nota, fontsize=7.5, fontname="helv", color=_GRIS)


def generar_pdf(mz: str, lt: str, salida: Path | None = None) -> Path:
    tabla = tabla_predio(mz, lt)
    nombres = repo._lookup_nombres()
    doc = fitz.open()
    _dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
    salida = salida or (BASE_DIR / "outputs" / f"reporte_historico_{mz}-{lt}.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    return salida


def generar_boletas_individuales(predios: list[tuple[str, str]] | None = None,
                                  carpeta: Path | None = None) -> Path:
    """1 archivo PDF por predio (no un solo PDF combinado) — para poder entregar
    cada boleta por separado. Por defecto: los de CONFIRMACION con saldo pendiente."""
    predios = predios if predios is not None else _predios_a_reportar()
    carpeta = carpeta or (BASE_DIR / "outputs" / "boletas_2026-07")
    carpeta.mkdir(parents=True, exist_ok=True)

    historicos = _cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = _cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()

    for mz, lt in predios:
        tabla = tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        doc = fitz.open()
        _dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)
        doc.save(str(carpeta / f"boleta_{mz}-{lt}.pdf"))
        doc.close()

    print(f"{len(predios)} boleta(s) individuales -> {carpeta}")
    return carpeta


def generar_lote(predios: list[tuple[str, str]] | None = None, salida: Path | None = None) -> Path:
    """1 PDF con 1 página por predio. Por defecto: solo los de CONFIRMACION que
    todavía tienen SALDO pendiente en seguimiento_pueblo — los que ya están al
    día no necesitan este reporte, su boleta normal les alcanza."""
    predios = predios if predios is not None else _predios_a_reportar()
    log_msg = f"{len(predios)} predio(s) a reportar (con saldo pendiente)"
    print(log_msg)

    historicos = _cargar_historicos()
    eventos = repo._leer_eventos()
    mapa_raw = _cargar_mapa_raw()
    f = BASE_DIR.parent / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
    dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
    nombres = repo._lookup_nombres()

    doc = fitz.open()
    for mz, lt in predios:
        tabla = tabla_predio(mz, lt, historicos, eventos, dfp, mapa_raw, nombres.get((mz, lt), ""))
        _dibujar_pagina(doc, mz, lt, nombres.get((mz, lt), ""), tabla)

    salida = salida or (BASE_DIR / "outputs" / "reporte_historico_2026-07.pdf")
    salida.parent.mkdir(exist_ok=True)
    doc.save(str(salida))
    doc.close()
    print(f"PDF -> {salida} ({len(predios)} páginas)")
    return salida


if __name__ == "__main__":
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 300)
    print(tabla_predio("Q", "5").to_string(index=False))
    ruta = generar_pdf("Q", "5")
    print(f"\nPDF -> {ruta}")
