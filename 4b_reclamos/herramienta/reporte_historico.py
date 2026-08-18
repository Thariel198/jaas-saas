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
REPO_ROOT = BASE_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "shared"))
import seguimiento_repo as repo  # noqa: E402
import ciclo  # noqa: E402

HIST_DIR = REPO_ROOT / "obligaciones" / "inputs" / "planillas anteriores"
REASIGNACIONES_PATH = REPO_ROOT / "0_padron" / "reasignaciones_candidata.xlsx"

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

    deuda = {"DEUDA_CONSUMO": consumo, "DEUDA_MANT": mant, "DEUDA_MES_ANT": mes_ant,
             "DEUDA_CORTE": corte, "DEUDA_CONVENIO": convenio, "DEUDA_MULTA": multa,
             "DEUDA_ACUERDOS": acuerdos}

    if not pagado:
        # No pagó nada ese mes -> no se muestra ningun concepto como cubierto
        return {"MES": mes, "CONSUMO": 0, "MANT": 0, "MES_ANT": 0, "CORTE": 0,
                "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "TOTAL": 0.0, "PAGO_COMPLETO": False,
                "NOTA": "No pago nada", **deuda}

    total = consumo + mant + mes_ant + corte + convenio + multa + acuerdos
    return {"MES": mes, "CONSUMO": consumo, "MANT": mant, "MES_ANT": mes_ant, "CORTE": corte,
            "CONVENIO": convenio, "MULTA": multa, "ACUERDOS": acuerdos, "TOTAL": total,
            "PAGO_COMPLETO": True, "NOTA": "", **deuda}


# Cada ciclo posterior a mayo (post-ledger) que ya cerró vive congelado en su
# propio repo -- el pipeline se copia entero por mes (ver CLAUDE.md). Cuando un
# mes nuevo cierra, se agrega su línea acá; no hace falta tocar ninguna función.
REPOS_CICLO_CERRADO = {
    "2026-06": Path(r"C:\Users\wilde\PycharmProjects\Junio\jass_system - junio"),
    "2026-07": Path(r"C:\Users\wilde\PycharmProjects\Julio\jass_system - Julio"),
}


def _planilla_ciclo_cerrado(mes_ano: str) -> Path:
    """Ciclo cerrado: el nombre canónico lleva el periodo
    (planilla_cobrado_2026-06.xlsx). shared/ciclo.resolver acepta también los
    nombres históricos (mes en español, ej. planilla_cobrado_julio.xlsx), así
    que un rename del archivo ya no vuelve a dejar todas las filas de ese mes
    sin consumo ni referencia de pago (bug del 05/08/2026)."""
    return ciclo.resolver(REPOS_CICLO_CERRADO[mes_ano] / "5_cobranza" / "outputs",
                          "planilla_cobrado", mes_ano)


_dfp_ciclo_cerrado_cache: dict[str, pd.DataFrame] = {}


def _cargar_dfp_ciclo_cerrado(mes_ano: str) -> pd.DataFrame | None:
    """planilla_cobrado.xlsx de un ciclo cerrado vive en su propio repo
    congelado (jass_system - <mes>) -- no en el activo, que solo tiene el
    ciclo vigente. Cacheado a nivel modulo por mes_ano: se reusa entre predios
    en una corrida por lote.

    Si el repo cerrado no esta donde dice REPOS_CICLO_CERRADO, se AVISA: cuando
    fallaba en silencio (04/08/2026, el repo de junio se movio a
    PycharmProjects\\Junio\\) toda fila de ese mes salia sin consumo/mant y sin
    referencia de pago -- se leia como un pago fantasma (caso K-9: mostraba
    multa 30 en vez de los S/38 reales)."""
    if mes_ano not in _dfp_ciclo_cerrado_cache:
        ruta = _planilla_ciclo_cerrado(mes_ano)
        if ruta.exists():
            _dfp_ciclo_cerrado_cache[mes_ano] = pd.read_excel(ruta, sheet_name="planilla_cobrado", header=1)
        else:
            print(f"  AVISO: no se encontro el ciclo cerrado de {mes_ano} -> {ruta}\n"
                  f"         las filas de {mes_ano} van a salir sin consumo/mantenimiento.")
            _dfp_ciclo_cerrado_cache[mes_ano] = pd.DataFrame()
    return _dfp_ciclo_cerrado_cache[mes_ano]


_PLANILLA_MES_DIR = REPO_ROOT / "shared" / "planilla_mes"
_planilla_correcta_cache: dict[str, pd.DataFrame] = {}


def _cargar_planilla_correcta(mes_ano: str) -> pd.DataFrame | None:
    """shared/planilla_mes/planilla_<mes_ano>.xlsx -- el CARGO real (2_planilla),
    verificado exacto contra DATA_boletas.xlsx. planilla_cobrado.xlsx puede quedar
    desactualizado si 5_cobranza no se re-corrio despues de una correccion de
    2_planilla (caso confirmado: D1-6/S-5 julio, ver LEER_ANTES.md) -- de ahi
    viene el CARGO (mes_actual/mantenimiento/mes_anterior/corte), nunca de
    planilla_cobrado."""
    if mes_ano in _planilla_correcta_cache:
        return _planilla_correcta_cache[mes_ano]
    ruta = _PLANILLA_MES_DIR / f"planilla_{mes_ano}.xlsx"
    if not ruta.exists():
        _planilla_correcta_cache[mes_ano] = None
        return None
    try:
        df = pd.read_excel(ruta, header=1, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        df["MZ"] = df["MZ"].astype(str).str.strip()
        df["LT"] = df["LT"].astype(str).str.strip()
    except Exception:
        df = None
    _planilla_correcta_cache[mes_ano] = df
    return df


_abonos_rezagados_cache: pd.DataFrame | None = None


def _abono_rezagado_predio(mz: str, lt: str, mes_ano: str) -> float:
    """Suma de abonos_rezagados.xlsx para este predio+mes_ano_aplica -- plata
    real que 5_cobranza nunca alcanzo a aplicar (archivo posterior a las
    corridas del ciclo, ver LEER_ANTES.md). El reporte la tiene que sumar acá
    para no mostrar un TOTAL PAGADO menor a la plata real (ya se ve aparte en
    la tabla de Referencia de pago, pero el total de arriba tiene que cuadrar)."""
    global _abonos_rezagados_cache
    if _abonos_rezagados_cache is None:
        ruta = REPO_ROOT / "shared" / "abonos_rezagados.xlsx"
        if ruta.exists():
            df = pd.read_excel(ruta, sheet_name="Abonos_Raw", header=1, dtype=str)
            df["MZ"] = df["MZ"].astype(str).str.strip()
            df["LT"] = df["LT"].astype(str).str.strip()
            df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce").fillna(0.0)
            _abonos_rezagados_cache = df
        else:
            _abonos_rezagados_cache = pd.DataFrame()
    df = _abonos_rezagados_cache
    if df.empty:
        return 0.0
    sub = df[(df["MZ"] == mz) & (df["LT"] == lt)
             & (df["MES_ANO_APLICA"].astype(str).str.strip() == mes_ano)]
    return round(float(sub["MONTO"].sum()), 2)


_ajustes_cargo_cache: pd.DataFrame | None = None


def _corte_exonerado(mz: str, lt: str) -> bool:
    """ajustes_cargo.xlsx con CONCEPTO=CORTE_RECONEXION para este predio -- no
    corresponde, sin importar en que mes_ano_aplica se vaya a ejecutar el
    efecto (puede ser el siguiente ciclo si julio ya cerro, ver S-5/F1-4 en
    LEER_ANTES.md) ni si tiene CLASE=EXONERACION escrita (F1-4 no la tiene,
    el motivo dice "anula" igual). Toda fila CORTE_RECONEXION vista hasta hoy
    es una anulacion, nunca un cargo nuevo. Para el HISTORIAL no hay que
    esperar a que se aplique: ya sabemos que no es deuda real, no se puede
    mostrar como pagada con plata que en realidad se acredito en otro lado
    (multa/acuerdos)."""
    global _ajustes_cargo_cache
    if _ajustes_cargo_cache is None:
        ruta = REPO_ROOT / "shared" / "ajustes_cargo.xlsx"
        if ruta.exists():
            df = pd.read_excel(ruta, header=1, dtype=str)
            df["MZ"] = df["MZ"].astype(str).str.strip()
            df["LT"] = df["LT"].astype(str).str.strip()
            _ajustes_cargo_cache = df
        else:
            _ajustes_cargo_cache = pd.DataFrame()
    df = _ajustes_cargo_cache
    if df.empty:
        return False
    sub = df[(df["MZ"] == mz) & (df["LT"] == lt)
             & (df["CONCEPTO"].astype(str).str.strip().str.upper() == "CORTE_RECONEXION")]
    return not sub.empty


def _datos_ciclo(mz: str, lt: str, dfp: pd.DataFrame,
                 incluir_abonos_rezagados: bool = True) -> dict | None:
    """Consumo/mant/mes_ant/corte de UN ciclo (junio o julio), capados por
    cascada P1(agua)->P2(mes anterior)->P2b(corte) contra lo realmente pagado
    ese ciclo -- si pagó menos de lo debido, eso es lo que se acredita, no de
    mas."""
    fila = dfp[(dfp["MZ"].astype(str).str.strip() == mz) & (dfp["LT"].astype(str).str.strip() == lt)]
    if fila.empty:
        return None
    r = fila.iloc[0]
    mes_ano = str(r.get("MES_ANO", "")).strip()

    # CARGO real: shared/planilla_mes (2_planilla), no planilla_cobrado -- ver
    # docstring de _cargar_planilla_correcta. Si no hay copia disponible para
    # ese mes_ano, cae al valor de dfp (comportamiento anterior).
    planilla_ok = _cargar_planilla_correcta(mes_ano) if mes_ano else None
    if planilla_ok is not None:
        fila_ok = planilla_ok[(planilla_ok["MZ"] == mz) & (planilla_ok["LT"] == lt)]
    else:
        fila_ok = None
    if fila_ok is not None and not fila_ok.empty:
        r_ok = fila_ok.iloc[0]
        fuente = r_ok
    else:
        fuente = r
    consumo_debido = _numf(fuente.get("MES_ACTUAL"))
    mant_debido = _numf(fuente.get("MANTENIMIENTO"))
    mes_ant_debido = _numf(fuente.get("MES_ANTERIOR"))
    corte_debido = 0.0 if _corte_exonerado(mz, lt) else _numf(fuente.get("CORTE_RECONEXION"))

    abono_extra = (_abono_rezagado_predio(mz, lt, mes_ano)
                   if incluir_abonos_rezagados and mes_ano else 0.0)
    total_pagado = _numf(r.get("MONTO_YAPE")) + _numf(r.get("MONTO_EFECTIVO")) + abono_extra
    consumo = min(consumo_debido, total_pagado)
    restante = max(0.0, total_pagado - consumo)
    mant = min(mant_debido, restante)
    restante = max(0.0, restante - mant)
    mes_ant = min(mes_ant_debido, restante)
    restante = max(0.0, restante - mes_ant)
    corte = min(corte_debido, restante)
    restante = max(0.0, restante - corte)
    convenio = min(_numf(fuente.get("CONVENIO")), restante)
    restante = max(0.0, restante - convenio)
    multa = min(_numf(fuente.get("MULTA")), restante)
    restante = max(0.0, restante - multa)
    acuerdos = min(_numf(fuente.get("ACUERDOS_ASAMBLEA")), restante)
    return {"mes_ano": mes_ano, "consumo": consumo, "mant": mant, "mes_ant": mes_ant,
            "corte": corte, "hubo_pago": total_pagado > 0.005,
            "deuda_consumo": consumo_debido, "deuda_mant": mant_debido,
            "deuda_mes_ant": mes_ant_debido, "deuda_corte": corte_debido,
            "deuda_convenio": _numf(fuente.get("CONVENIO")),
            "deuda_multa": _numf(fuente.get("MULTA")),
            "deuda_acuerdos": _numf(fuente.get("ACUERDOS_ASAMBLEA")),
            "deuda_total": (consumo_debido + mant_debido + mes_ant_debido + corte_debido
                            + _numf(fuente.get("CONVENIO")) + _numf(fuente.get("MULTA"))
                            + _numf(fuente.get("ACUERDOS_ASAMBLEA"))),
            "pago_consumo": consumo, "pago_mant": mant,
            "pago_mes_ant": mes_ant, "pago_corte": corte,
            "pago_convenio": convenio, "pago_multa": multa, "pago_acuerdos": acuerdos}


def _filas_recientes(mz: str, lt: str, eventos: pd.DataFrame, dfp: pd.DataFrame,
                     incluir_abonos_rezagados: bool = True,
                     deuda_conceptos_desde_ledger: bool = False) -> list[dict]:
    """Junio en adelante: seguimiento_pueblo + planilla_cobrado. Reusa la logica
    ya construida en reporte_seguimiento.py (cascada vieja, sin exponer exceso)."""
    sys.path.insert(0, str(BASE_DIR.parent))
    import reporte_seguimiento as rs

    resumen, historial = rs._resumen_y_historial(mz, lt, eventos)

    ledger_por_mes: dict[str, dict[str, float]] = {}
    ledger_deuda_por_mes: dict[str, dict[str, float]] = {}
    if deuda_conceptos_desde_ledger:
        ev_predio = eventos[(eventos["MZ"].astype(str).str.strip() == mz) &
                            (eventos["LT"].astype(str).str.strip() == lt)].copy()
        ev_predio["MES"] = ev_predio["MES"].astype(str).str.strip()
        ev_predio["CONCEPTO"] = ev_predio["CONCEPTO"].astype(str).str.strip().str.upper()
        running: dict[str, float] = {}
        for mes in sorted(ev_predio["MES"].unique()):
            deuda_mes: dict[str, float] = running.copy()
            for concepto in ("MULTA", "ACUERDOS", "CONVENIO"):
                sub = ev_predio[(ev_predio["MES"] == mes) &
                                (ev_predio["CONCEPTO"] == concepto)]
                if not sub.empty:
                    antes_del_pago = (running.get(concepto, 0.0)
                                      + pd.to_numeric(sub["CARGO"], errors="coerce").fillna(0).sum()
                                      + pd.to_numeric(sub["AJUSTE"], errors="coerce").fillna(0).sum())
                    deuda_mes[concepto] = max(0.0, round(antes_del_pago, 2))
                    running[concepto] = max(0.0, round(
                        antes_del_pago
                        - pd.to_numeric(sub["PAGO"], errors="coerce").fillna(0).sum(), 2))
            ledger_deuda_por_mes[mes] = deuda_mes.copy()
            ledger_por_mes[mes] = running.copy()

    def _saldo_ledger(mes: str, concepto: str) -> float | None:
        anteriores = [m for m in ledger_por_mes if m <= mes]
        if not anteriores:
            return None
        return ledger_por_mes[max(anteriores)].get(concepto, 0.0)

    # Un ciclo por cada planilla_cobrado disponible -- el activo (pasado por
    # el llamador, hoy agosto) y cada ciclo posterior a mayo que ya cerró
    # (REPOS_CICLO_CERRADO: junio, julio, ...). Antes solo se calculaba
    # consumo/mant/mes_ant/corte para "el ciclo actual" resuelto a mano por
    # cada llamador; un mes que cerraba y dejaba de ser "el activo"
    # desaparecía en silencio de la tabla (bug real: julio se cayó cuando el
    # repo activo rodó a agosto). Con REPOS_CICLO_CERRADO ningún ciclo cerrado
    # depende de cuál sea el repo activo hoy.
    datos_por_ciclo: dict[str, dict] = {}
    fuentes = [dfp] + [_cargar_dfp_ciclo_cerrado(m) for m in REPOS_CICLO_CERRADO]
    for fuente in fuentes:
        if fuente is None or fuente.empty:
            continue
        d = _datos_ciclo(mz, lt, fuente, incluir_abonos_rezagados)
        if d and d["mes_ano"]:
            datos_por_ciclo[d["mes_ano"]] = d

    # el mes de cada ciclo entra a la tabla si hubo PAGO real (yape/efectivo),
    # aunque no haya evento MULTA/ACUERDOS/CONVENIO ese mes en seguimiento_pueblo
    # (antes: se exigia un evento de pueblo para crear la fila, y un pago de solo
    # consumo -como A-6 en julio- nunca generaba fila, aunque si hubiera pagado)
    meses = set(historial["MES"].astype(str).unique())
    for mes_ano, d in datos_por_ciclo.items():
        if d["hubo_pago"] or d["deuda_total"] > 0.005:
            meses.add(mes_ano)

    filas = []
    for mes in sorted(meses):
        del_mes = historial[historial["MES"] == mes]
        fila_d = {"MES": mes, "CONSUMO": 0, "MANT": 0, "MES_ANT": 0, "CORTE": 0,
                  "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "NOTA": ""}
        for c in ("CONSUMO", "MANT", "MES_ANT", "CORTE", "CONVENIO", "MULTA", "ACUERDOS"):
            fila_d[f"DEUDA_{c}"] = 0.0
            fila_d[f"PAGO_{c}"] = 0.0
        for _, r in del_mes.iterrows():
            concepto = r["CONCEPTO"]
            c = str(concepto).strip().upper()
            if c in ("MULTA", "ACUERDOS", "CONVENIO"):
                fila_d[f"DEUDA_{c}"] += max(0.0, _numf(r.get("DEBIA")))
                fila_d[f"PAGO_{c}"] += max(0.0, _numf(r["PAGO"]))
            # DEBIA = CARGO + AJUSTE del mes (ver reporte_seguimiento._resumen_y_historial).
            # Cuando un AJUSTE perdona/salda deuda (condonacion, correccion del bug de
            # signo), DEBIA queda negativo -- esa parte tambien salda la deuda aunque
            # no sea un PAGO en efectivo, y hay que sumarla para que el total cuadre
            # contra la plata real (ver LEER_ANTES.md, caso L-4/D-16/D1-6/S-5 09/08/2026).
            # Un DEBIA positivo es CARGO nuevo, no plata: no se resta del pagado.
            saldado = float(r["PAGO"]) + max(0.0, -float(r["DEBIA"]))
            if concepto == "MULTA":
                fila_d["MULTA"] = saldado
            elif concepto == "ACUERDOS":
                fila_d["ACUERDOS"] = saldado
            elif concepto == "CONVENIO":
                fila_d["CONVENIO"] = saldado
        d = datos_por_ciclo.get(mes)
        if d:
            fila_d["CONSUMO"] = d["consumo"]
            fila_d["MANT"] = d["mant"]
            fila_d["MES_ANT"] = d["mes_ant"]
            fila_d["CORTE"] = d["corte"]
            for c in ("CONSUMO", "MANT", "MES_ANT", "CORTE"):
                fila_d[f"DEUDA_{c}"] = d[f"deuda_{c.lower()}"]
                fila_d[f"PAGO_{c}"] = d[f"pago_{c.lower()}"]
            for c in ("CONVENIO", "MULTA", "ACUERDOS"):
                fila_d[f"DEUDA_{c}"] = d[f"deuda_{c.lower()}"]
            if deuda_conceptos_desde_ledger:
                for c in ("CONVENIO", "MULTA", "ACUERDOS"):
                    if mes in ledger_deuda_por_mes:
                        fila_d[f"DEUDA_{c}"] = ledger_deuda_por_mes[mes].get(c, 0.0)
                    else:
                        meses_saldo = [m for m in ledger_por_mes if m <= mes]
                        if meses_saldo:
                            fila_d[f"DEUDA_{c}"] = ledger_por_mes[max(meses_saldo)].get(c, 0.0)
            if incluir_abonos_rezagados and not deuda_conceptos_desde_ledger:
                for c in ("CONVENIO", "MULTA", "ACUERDOS"):
                    fila_d[c] = d[f"pago_{c.lower()}"]
                    fila_d[f"PAGO_{c}"] = d[f"pago_{c.lower()}"]
        fila_d["DEUDA_TOTAL"] = sum(fila_d[f"DEUDA_{c}"] for c in CONCEPTOS_TABLA)
        fila_d["PAGO_TOTAL"] = sum(fila_d[f"PAGO_{c}"] for c in CONCEPTOS_TABLA)
        fila_d["TOTAL"] = sum(fila_d[c] for c in CONCEPTOS_TABLA)
        fila_d["PAGO_COMPLETO"] = fila_d["TOTAL"] > 0.005
        filas.append(fila_d)
    return filas


def tabla_predio(mz: str, lt: str, historicos: dict | None = None,
                  eventos: pd.DataFrame | None = None, dfp: pd.DataFrame | None = None,
                  mapa_raw: dict | None = None, nombre_actual: str = "",
                  incluir_abonos_rezagados: bool = True,
                  deuda_conceptos_desde_ledger: bool = False) -> pd.DataFrame:
    historicos = historicos if historicos is not None else _cargar_historicos()
    eventos = eventos if eventos is not None else repo._leer_eventos()
    mapa_raw = mapa_raw if mapa_raw is not None else _cargar_mapa_raw()
    if not nombre_actual:
        nombre_actual = repo._lookup_nombres().get((mz, lt), "")
    if dfp is None:
        f = REPO_ROOT / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
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
    filas.extend(_filas_recientes(
        mz, lt, eventos, dfp, incluir_abonos_rezagados,
        deuda_conceptos_desde_ledger,
    ))
    tabla = pd.DataFrame(filas, columns=["MES", "CONSUMO", "MANT", "MES_ANT", "CORTE",
                                          "CONVENIO", "MULTA", "ACUERDOS", "TOTAL", "PAGO_COMPLETO", "NOTA",
                                          *[f"DEUDA_{c}" for c in CONCEPTOS_TABLA],
                                          *[f"PAGO_{c}" for c in CONCEPTOS_TABLA], "DEUDA_TOTAL", "PAGO_TOTAL"])
    for c in CONCEPTOS_TABLA:
        deuda = f"DEUDA_{c}"
        pago = f"PAGO_{c}"
        tabla[deuda] = tabla[deuda].fillna(0.0)
        tabla[pago] = tabla[pago].fillna(tabla[c]).fillna(0.0)
    tabla["DEUDA_TOTAL"] = tabla["DEUDA_TOTAL"].fillna(
        tabla[[f"DEUDA_{c}" for c in CONCEPTOS_TABLA]].sum(axis=1))
    tabla["PAGO_TOTAL"] = tabla["PAGO_TOTAL"].fillna(
        tabla[[f"PAGO_{c}" for c in CONCEPTOS_TABLA]].sum(axis=1))
    return tabla


def _predios_a_reportar() -> list[tuple[str, str]]:
    """Filtra CONFIRMACION a los predios que realmente necesitan el reporte —
    si ya está al día (SALDO=0 en seguimiento_pueblo), no hace falta imprimirle
    nada, su boleta normal ya lo deja conforme."""
    sys.path.insert(0, str(BASE_DIR.parent))
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
    headers = ["Mes", "Tipo", "Consumo", "Mant.", "Mes ant.", "Corte",
               "Convenio", "Multa", "Acuerdos", "Total"]
    resto = w - 58 - 48 - 78
    col = resto / 7
    anchos = [58, 48, col, col, col, col, col, col, col, 78]

    rh = 13
    page.draw_rect(fitz.Rect(x, y, x + w, y + rh), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh - 4), h, fontsize=7, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh

    meses = tabla
    for n, (_, r) in enumerate(meses.iterrows()):
        es_ultimo_mes = n == len(meses) - 1
        filas_mes = 3 if es_ultimo_mes else 2
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh * filas_mes), fill=_ZEBRA, color=None)
        tipos = [("DEUDA", "DEUDA_", _NEGRO), ("PAGO", "PAGO_", _VERDE)]
        if es_ultimo_mes:
            tipos.append(("SALDO", "DEUDA_", _ROJO))
        saldo = max(0.0, float(r.get("DEUDA_TOTAL", 0.0) or 0.0)
                    - float(r.get("PAGO_TOTAL", 0.0) or 0.0))
        for tipo, prefix, color in tipos:
            cx = x
            page.insert_text((cx + 4, y + rh - 6), str(r["MES"]) if tipo == "DEUDA" else "",
                             fontsize=6.5, fontname="hebo", color=_NEGRO)
            cx += anchos[0]
            page.insert_text((cx + 4, y + rh - 4), tipo, fontsize=6.5, fontname="hebo", color=color)
            cx += anchos[1]
            total = 0.0
            for c in CONCEPTOS_TABLA:
                v = float(r.get(f"{prefix}{c}", 0.0) or 0.0)
                if tipo == "SALDO":
                    v = max(0.0, v - float(r.get(f"PAGO_{c}", 0.0) or 0.0))
                total += v
                texto = f"{v:,.2f}" if abs(v) > 0.005 else "—"
                tw = fitz.get_text_length(texto, fontname="helv", fontsize=6.5)
                page.insert_text((cx + anchos[2] - 5 - tw, y + rh - 4), texto,
                                 fontsize=6.5, fontname="helv", color=color)
                cx += anchos[2]
            if tipo == "SALDO":
                total = saldo
            texto = f"S/ {total:,.2f}" if total > 0.005 else "—"
            tw = fitz.get_text_length(texto, fontname="hebo", fontsize=6.5)
            page.insert_text((cx + anchos[-1] - 5 - tw, y + rh - 4), texto,
                             fontsize=6.5, fontname="hebo", color=color)
            y += rh

        if r["NOTA"]:
            pass  # nota se omite en la tabla, ya lo dice el check "-"

    return y


_ROJO = (0.62, 0.10, 0.10)


def _dibujar_saldo_pendiente(page, x: float, y: float, w: float, saldo: dict) -> float:
    """Caja con la deuda de pueblo (Multa/Acuerdos/Convenio) que todavia falta
    pagar, al cierre del mes que reporta la tabla — no incluye agua ni
    mantenimiento, que viven en otro precursor. Los valores se leen en vivo
    de seguimiento_pueblo.xlsx (via calcular_tabla), asi que reflejan cualquier
    correccion ya aplicada al ledger."""
    rh_box = 22
    total = float(saldo.get("TOTAL", 0.0))
    al_dia = total <= 0.005
    color_total = _VERDE if al_dia else _ROJO
    page.draw_rect(fitz.Rect(x, y, x + w, y + rh_box),
                    fill=_AZUL_BG if al_dia else (0.99, 0.94, 0.90), color=None)
    page.insert_text((x + 6, y + rh_box - 7), "SALDO PENDIENTE (Multa + Acuerdos + Convenio)",
                      fontsize=9, fontname="hebo", color=_NEGRO)

    partes = [("Multa", saldo.get("MULTA", 0.0)), ("Acuerdos", saldo.get("ACUERDOS", 0.0)),
              ("Convenio", saldo.get("CONVENIO", 0.0))]
    texto_partes = "  ·  ".join(f"{n}: S/ {v:,.2f}" for n, v in partes if abs(v) > 0.005) or "al día"
    tw_partes = fitz.get_text_length(texto_partes, fontname="helv", fontsize=8.5)

    texto_total = f"TOTAL: S/ {total:,.2f}"
    tw_total = fitz.get_text_length(texto_total, fontname="hebo", fontsize=10)

    page.insert_text((x + w - tw_total - 10, y + rh_box - 7), texto_total, fontsize=10, fontname="hebo",
                      color=color_total)
    page.insert_text((x + w - tw_total - 20 - tw_partes, y + rh_box - 7), texto_partes, fontsize=8.5,
                      fontname="helv", color=_GRIS)
    return y + rh_box


def _dibujar_pagina(doc, mz: str, lt: str, nombre: str, tabla: pd.DataFrame,
                     saldo_pendiente: dict | None = None) -> float:
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M

    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_AZUL, color=None)
    page.insert_text((_M + 10, y + 20), f"Predio {mz}-{lt} — {nombre or '(sin nombre)'} — historial mensual",
                      fontsize=12, fontname="hebo", color=(1, 1, 1))
    y += 42

    page.insert_text((_M, y), "Historial mensual: deuda, pago y saldo por concepto (octubre 2025 -> ciclo vigente)", fontsize=9,
                      fontname="helv", color=_GRIS)
    y += 14
    y = _dibujar_tabla_historico(page, _M, y, w, tabla)
    y += 12

    if saldo_pendiente is not None:
        y = _dibujar_saldo_pendiente(page, _M, y, w, saldo_pendiente)
        y += 10
    else:
        y += 8

    nota = ("Desde junio 2026 el exceso pagado ya no se muestra ni se aplica solo — se retiene hasta que el "
            "vecino reclame. Se conserva todo el historial mensual, incluso meses sin pago.")
    page.insert_text((_M, y), nota, fontsize=7.5, fontname="helv", color=_GRIS)
    return y + 12


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
    f = REPO_ROOT / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
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
    f = REPO_ROOT / "5_cobranza" / "outputs" / "planilla_cobrado.xlsx"
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
