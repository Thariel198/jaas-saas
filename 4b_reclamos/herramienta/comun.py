"""4b_reclamos/herramienta/comun.py — helpers compartidos por las herramientas
de reporte de 4b_reclamos (reporte_historico.py, reporte_referencias_pago.py,
buscar_pago.py).

Antes vivian pegados (copy-paste) en reporte_referencias_pago.py y en
buscar_pago.py por separado -- dos copias identicas que había que mantener
sincronizadas a mano. Este modulo es la unica fuente ahora.

tabla_predio()/_dibujar_pagina() (y toda su cadena interna: _fila_historica,
_datos_ciclo, _filas_recientes, _cargar_historicos, etc.) vivían en
reporte_historico.py, pero reporte_referencias_pago.py y buscar_pago.py las
consumían igual -- reporte_historico.py terminó siendo la "hoja de abajo" de
facto sin serlo (es una herramienta más, genera sus propios PDFs). Todo eso
se centralizó acá; reporte_historico.py ahora también lee de comun.py, como
cualquier otro script del módulo.
"""

import sys
import unicodedata
from pathlib import Path

import fitz
import pandas as pd

HERRAMIENTA_DIR = Path(__file__).parent          # 4b_reclamos/herramienta/
BASE_DIR = HERRAMIENTA_DIR.parent                # 4b_reclamos/
REPO_DIR = BASE_DIR.parent                       # raíz del repo activo
SHARED_DIR = REPO_DIR / "shared"

sys.path.insert(0, str(HERRAMIENTA_DIR))
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(SHARED_DIR))

import ciclo                            # noqa: E402
import seguimiento_repo as repo         # noqa: E402

HIST_DIR = REPO_DIR / "obligaciones" / "inputs" / "planillas anteriores"
REASIGNACIONES_PATH = REPO_DIR / "0_padron" / "reasignaciones_candidata.xlsx"
REASIGNACIONES_APLICACION_PATH = SHARED_DIR / "reasignaciones_aplicacion.xlsx"
_ABONOS_REZAGADOS = SHARED_DIR / "abonos_rezagados.xlsx"
_BLANCOS_EFECTIVO = SHARED_DIR / "blancos_efectivo.xlsx"
_PLANILLA_MES_DIR = SHARED_DIR / "planilla_mes"

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
_CAMPO_TABLA = {"MULTA": "MULTA", "ACUERDOS": "ACUERDOS", "CONVENIO": "CONVENIO"}
_CONCEPTOS_RESUMEN = ("CONVENIO", "ACUERDOS", "MULTA")  # orden de cascada P3-P5

# Cada ciclo posterior a mayo (post-ledger) que ya cerró vive congelado en su
# propio repo -- el pipeline se copia entero por mes (ver CLAUDE.md). Cuando un
# mes nuevo cierra, se agrega su línea acá; no hace falta tocar ninguna función.
REPOS_CICLO_CERRADO = {
    "2026-06": Path(r"C:\Users\wilde\PycharmProjects\Junio\jass_system - junio"),
    "2026-07": Path(r"C:\Users\wilde\PycharmProjects\Julio\jass_system - Julio"),
}

_AZUL = (26/255, 82/255, 118/255)
_AZUL_BG = (235/255, 245/255, 251/255)
_GRIS = (0.42, 0.45, 0.5)
_NEGRO = (0.12, 0.16, 0.22)
_VERDE = (0.02, 0.37, 0.27)
_ROJO = (0.62, 0.10, 0.10)
_ZEBRA = (243/255, 244/255, 246/255)
_REFERENCIAS_BG = (218/255, 223/255, 228/255)
_PAGE_W, _PAGE_H = 842, 595  # A4 horizontal (pts) — tabla ancha
_M = 30


def _norm_col(s: str) -> str:
    """Normaliza NOMBRE DE COLUMNA (sin tildes/mayus) para _col(). Distinto de
    un normalizador de VALORES de MZ/LT/nombre -- ver comentario historico en
    el buscar_pago.py viejo, que tenía los dos con el mismo nombre "_norm" en
    un solo archivo y tuvo que renombrar este para que convivieran."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return s.lower().replace("\n", " ").strip()


def _col(df: pd.DataFrame, *candidatos: str):
    normed = {}
    for c in df.columns:
        k = _norm_col(c)
        if k not in normed:
            normed[k] = c
    for cand in candidatos:
        cn = _norm_col(cand)
        if cn in normed:
            return normed[cn]
    for cand in candidatos:
        cn = _norm_col(cand)
        for k, real in normed.items():
            if cn in k:
                return real
    return None


def _norm_lote(v) -> str:
    """pagos_yape_tepago.xlsx guarda LOTE como numero (4.0) para lotes sin
    letra -- comparar como texto plano ("4") sin esto nunca matchea, aunque
    la fila SI exista (bug encontrado 01/08/2026, caso I-4: el dato estaba,
    la comparacion de texto lo escondia)."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _val(row, df, *candidatos) -> float:
    c = _col(df, *candidatos)
    if c is None:
        return 0.0
    v = row.get(c)
    return float(v) if pd.notna(v) else 0.0


def _numf(v) -> float:
    """float NaN-safe — 'v or 0' no sirve porque NaN es truthy en Python."""
    return float(v) if pd.notna(v) else 0.0


_avisados: set[str] = set()


def _avisar_falta(p: Path, que: str) -> None:
    """Un archivo de ciclo cerrado que no esta donde dice el path se avisa UNA
    vez por corrida (esto se llama por predio). Fallaba en silencio hasta el
    04/08/2026 y las filas sin origen se leian como pagos fantasma."""
    if str(p) in _avisados:
        return
    _avisados.add(str(p))
    print(f"  AVISO: falta {que} -> {p}")


_cache_hojas: dict[str, dict] = {}


def _cargar_hojas_historicas(archivo: str) -> dict:
    """Cachea Cobranza/Reporte de UN archivo historico a nivel modulo -- sin
    esto, una corrida por lote reabre los mismos 8 excels una vez POR PREDIO
    (~500 predios x 8 archivos = ~4000 lecturas en vez de 8)."""
    if archivo in _cache_hojas:
        return _cache_hojas[archivo]
    p = HIST_DIR / archivo
    hojas = {}
    if p.exists():
        hojas["Cobranza"] = pd.read_excel(p, sheet_name="Cobranza", header=0)
        try:
            hojas["Reporte"] = pd.read_excel(p, sheet_name="Reporte", header=0)
        except Exception:
            hojas["Reporte"] = None
    _cache_hojas[archivo] = hojas
    return hojas


_cache_overlays: dict[str, pd.DataFrame] = {}


def _cargar_overlay(p: Path) -> pd.DataFrame:
    if str(p) not in _cache_overlays:
        if not p.exists():
            _avisar_falta(p, "overlay de pagos")
            _cache_overlays[str(p)] = pd.DataFrame()
        else:
            df = pd.read_excel(p, header=1)
            df.columns = [str(c).strip().upper() for c in df.columns]
            _cache_overlays[str(p)] = df
    return _cache_overlays[str(p)]


def _overlays_de_plata(mz: str, lt: str, mes_ciclo: str) -> list[dict]:
    """Plata real que entra a la cascada del ciclo pero NO esta en
    MONTO_YAPE/MONTO_EFECTIVO de planilla_cobrado: abonos que el cobrador
    retuvo y se regularizaron despues, y efectivo que habia entrado como
    blanco. Sin estas lineas el pago que salda la deuda queda sin origen
    visible y se lee como pago fantasma (caso encontrado 05/08/2026: los 4
    abonos retenidos por Wagner en julio -- S-5, D-16, D1-6, L-4)."""
    out = []
    for p, medio, quien_cols in ((_ABONOS_REZAGADOS, "ABONO REZ.", ("RETENIDO_POR", "CANAL_ORIGEN")),
                                  (_BLANCOS_EFECTIVO, "BLANCO EF.", ("ORIGEN", "CANAL"))):
        df = _cargar_overlay(p)
        if df.empty or "MES_ANO_APLICA" not in df.columns:
            continue
        sub = df[(df["MZ"].astype(str).str.strip() == mz) &
                 (df["LT"].apply(_norm_lote) == _norm_lote(lt)) &
                 (df["MES_ANO_APLICA"].astype(str).str.strip() == mes_ciclo)]
        for _, r in sub.iterrows():
            quien = " · ".join(str(r[c]).strip() for c in quien_cols
                               if c in df.columns and pd.notna(r.get(c)))
            fecha_raw = r.get("FECHA_REAL", "")
            fecha = str(fecha_raw).strip() if pd.notna(fecha_raw) else "sin fecha registrada"
            mes_pago = str(r.get("MES_CICLO", "")).strip()[:7] or mes_ciclo
            out.append({"MES": mes_pago, "MES_APLICA": mes_ciclo, "MEDIO": medio,
                        "FECHA_HORA": f"{fecha} · aplicado en {mes_ciclo} · {quien}",
                        "MONTO": float(r["MONTO"])})
    return out


def _repo_de_ciclo(mes_ano: str) -> Path:
    """Repo donde vive el output crudo de ese ciclo: el suyo propio congelado
    si ya cerró (REPOS_CICLO_CERRADO), o el repo activo si es el ciclo
    vigente -- así ningún ciclo depende de hardcodear "cuál mes es hoy"."""
    return REPOS_CICLO_CERRADO.get(mes_ano, REPO_DIR)


def _planilla_cobrado_path(mes_ano: str | None = None) -> Path:
    mes_ano = mes_ano or ciclo.activo()
    carpeta = _repo_de_ciclo(mes_ano) / "5_cobranza" / "outputs"
    return ciclo.resolver(
        carpeta, "planilla_cobrado", mes_ano,
        legacy_sin_periodo=ciclo.acepta_legacy(mes_ano),
    )


def _ciclos_recientes() -> list[tuple[str, Path]]:
    """(mes_ano, ruta a planilla_cobrado) de cada ciclo posterior a mayo: los
    cerrados de REPOS_CICLO_CERRADO + el vigente de shared/ciclo_activo.json
    (si no es ya uno de los cerrados). Se recalcula en cada llamada -- barato,
    y así recoge el ciclo activo sin reiniciar el proceso cuando el mes rueda."""
    ciclos = {m: _planilla_ciclo_cerrado(m) for m in REPOS_CICLO_CERRADO}
    activo = ciclo.activo(default=None)
    if activo and activo not in ciclos:
        ciclos[activo] = _planilla_cobrado_path(activo)
    return sorted(ciclos.items())


def _cargar_redirects() -> pd.DataFrame:
    if not REASIGNACIONES_APLICACION_PATH.exists():
        return pd.DataFrame(columns=["MZ", "LT", "CONCEPTO_ORIGEN", "CONCEPTO_DESTINO", "MES_ANO", "MONTO"])
    df = pd.read_excel(REASIGNACIONES_APLICACION_PATH, header=1)
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
        tabla.loc[idx, "TOTAL"] = tabla.loc[idx, CONCEPTOS_TABLA].sum(axis=1)
    return tabla


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
    for concepto in _CONCEPTOS_RESUMEN:
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
        orden = {c: i for i, c in enumerate(_CONCEPTOS_RESUMEN)}
        historial = historial.sort_values(
            ["MES", "CONCEPTO"], key=lambda s: s.map(orden) if s.name == "CONCEPTO" else s
        ).reset_index(drop=True)
    return resumen, historial


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


def _abonos_rezagados_predio(mz: str, lt: str, mes_ano: str) -> tuple[float, float]:
    """Suma de abonos_rezagados.xlsx para este predio+mes_ano_aplica -- plata
    real que 5_cobranza nunca alcanzo a aplicar (archivo posterior a las
    corridas del ciclo, ver LEER_ANTES.md). El reporte la tiene que sumar acá
    para no mostrar un TOTAL PAGADO menor a la plata real (ya se ve aparte en
    la tabla de Referencia de pago, pero el total de arriba tiene que cuadrar)."""
    global _abonos_rezagados_cache
    if _abonos_rezagados_cache is None:
        ruta = _ABONOS_REZAGADOS
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
        return 0.0, 0.0
    sub = df[(df["MZ"] == mz) & (df["LT"] == lt)
             & (df["MES_ANO_APLICA"].astype(str).str.strip() == mes_ano)]
    mes_origen = sub.get("MES_CICLO", pd.Series(index=sub.index, dtype=str)).astype(str).str.strip()
    cerrado = sub[mes_origen < mes_ano]
    vigente = sub[mes_origen >= mes_ano]
    return round(float(cerrado["MONTO"].sum()), 2), round(float(vigente["MONTO"].sum()), 2)


def _abono_rezagado_predio(mz: str, lt: str, mes_ano: str) -> float:
    cerrado, vigente = _abonos_rezagados_predio(mz, lt, mes_ano)
    return round(cerrado + vigente, 2)


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
        ruta = SHARED_DIR / "ajustes_cargo.xlsx"
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
    """Deuda vigente y reparto histórico de un ciclo.

    La deuda usa planilla_mes corregida. El reparto conserva la foto de
    planilla_cobrado que existía al cobrar: una corrección posterior cambia el
    saldo, pero no reescribe retroactivamente cómo se imputó el dinero.
    """
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
        fuente_deuda = r_ok
    else:
        fuente_deuda = r
    fuente_pago = r
    consumo_debido = _numf(fuente_deuda.get("MES_ACTUAL"))
    mant_debido = _numf(fuente_deuda.get("MANTENIMIENTO"))
    mes_ant_debido = _numf(fuente_deuda.get("MES_ANTERIOR"))
    corte_debido = 0.0 if _corte_exonerado(mz, lt) else _numf(fuente_deuda.get("CORTE_RECONEXION"))

    abono_cerrado, abono_vigente = (_abonos_rezagados_predio(mz, lt, mes_ano)
                                     if incluir_abonos_rezagados and mes_ano else (0.0, 0.0))
    saldos_pago = {
        "consumo": _numf(fuente_pago.get("MES_ACTUAL")),
        "mant": _numf(fuente_pago.get("MANTENIMIENTO")),
        "mes_ant": _numf(fuente_pago.get("MES_ANTERIOR")),
        "corte": 0.0 if _corte_exonerado(mz, lt) else _numf(fuente_pago.get("CORTE_RECONEXION")),
        "convenio": _numf(fuente_pago.get("CONVENIO")),
        "acuerdos": _numf(fuente_pago.get("ACUERDOS_ASAMBLEA")),
        "multa": _numf(fuente_pago.get("MULTA")),
    }
    pagos = {k: 0.0 for k in saldos_pago}

    def aplicar(monto: float, orden: tuple[str, ...]) -> None:
        restante = max(monto, 0.0)
        for concepto in orden:
            usado = min(saldos_pago[concepto], restante)
            saldos_pago[concepto] = round(saldos_pago[concepto] - usado, 2)
            pagos[concepto] = round(pagos[concepto] + usado, 2)
            restante = round(restante - usado, 2)

    # Un abono de un ciclo cerrado no puede pagar consumo nuevo del mes en que se regularizó.
    aplicar(abono_cerrado, ("mes_ant", "corte", "convenio", "acuerdos", "multa"))
    pago_ciclo = _numf(r.get("MONTO_YAPE")) + _numf(r.get("MONTO_EFECTIVO")) + abono_vigente
    aplicar(pago_ciclo, ("consumo", "mant", "mes_ant", "corte", "convenio", "acuerdos", "multa"))
    total_pagado = round(pago_ciclo + abono_cerrado, 2)
    diferencias_snapshot = []
    etiquetas = {"MES_ACTUAL": "consumo", "MANTENIMIENTO": "mantenimiento",
                 "MES_ANTERIOR": "mes anterior", "CORTE_RECONEXION": "corte"}
    for campo, etiqueta in etiquetas.items():
        antes = _numf(fuente_pago.get(campo))
        ahora = _numf(fuente_deuda.get(campo))
        if abs(antes - ahora) > 0.005:
            diferencias_snapshot.append(f"{etiqueta} S/{antes:,.2f} -> S/{ahora:,.2f}")
    nota_snapshot = ""
    if diferencias_snapshot:
        nota_snapshot = ("Reparto histórico según snapshot de cobro; deuda corregida después: "
                         + ", ".join(diferencias_snapshot) + ".")
    return {"mes_ano": mes_ano, "consumo": pagos["consumo"], "mant": pagos["mant"],
            "mes_ant": pagos["mes_ant"], "corte": pagos["corte"],
            "hubo_pago": total_pagado > 0.005,
            "aplicacion_pendiente": abono_cerrado > 0.005,
            "deuda_consumo": consumo_debido, "deuda_mant": mant_debido,
            "deuda_mes_ant": mes_ant_debido, "deuda_corte": corte_debido,
            "deuda_convenio": _numf(fuente_deuda.get("CONVENIO")),
            "deuda_multa": _numf(fuente_deuda.get("MULTA")),
            "deuda_acuerdos": _numf(fuente_deuda.get("ACUERDOS_ASAMBLEA")),
            "deuda_total": (consumo_debido + mant_debido + mes_ant_debido + corte_debido
                            + _numf(fuente_deuda.get("CONVENIO")) + _numf(fuente_deuda.get("MULTA"))
                            + _numf(fuente_deuda.get("ACUERDOS_ASAMBLEA"))),
            "pago_consumo": pagos["consumo"], "pago_mant": pagos["mant"],
            "pago_mes_ant": pagos["mes_ant"], "pago_corte": pagos["corte"],
            "pago_convenio": pagos["convenio"], "pago_multa": pagos["multa"],
            "pago_acuerdos": pagos["acuerdos"], "total_pagado": total_pagado,
            "nota_snapshot": nota_snapshot}


def _filas_recientes(mz: str, lt: str, eventos: pd.DataFrame, dfp: pd.DataFrame,
                     incluir_abonos_rezagados: bool = True,
                     deuda_conceptos_desde_ledger: bool = False,
                     proyectar_mes: str | None = None) -> list[dict]:
    """Junio en adelante: seguimiento_pueblo + planilla_cobrado. Reusa la logica
    de _resumen_y_historial() (cascada vieja, sin exponer exceso)."""
    resumen, historial = _resumen_y_historial(mz, lt, eventos)

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
                  "CONVENIO": 0, "MULTA": 0, "ACUERDOS": 0, "NOTA": "",
                  "APLICACION_PENDIENTE": False}
        for c in ("CONSUMO", "MANT", "MES_ANT", "CORTE", "CONVENIO", "MULTA", "ACUERDOS"):
            fila_d[f"DEUDA_{c}"] = 0.0
            fila_d[f"PAGO_{c}"] = 0.0
        for _, r in del_mes.iterrows():
            concepto = r["CONCEPTO"]
            c = str(concepto).strip().upper()
            if c in ("MULTA", "ACUERDOS", "CONVENIO"):
                fila_d[f"DEUDA_{c}"] += max(0.0, _numf(r.get("DEBIA")))
                fila_d[f"PAGO_{c}"] += max(0.0, _numf(r["PAGO"]))
            # DEBIA = CARGO + AJUSTE del mes (ver _resumen_y_historial). Cuando
            # un AJUSTE perdona/salda deuda (condonacion, correccion del bug de
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
            fila_d["NOTA"] = d.get("nota_snapshot", "")
            fila_d["APLICACION_PENDIENTE"] = d.get("aplicacion_pendiente", False)
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
            if (incluir_abonos_rezagados and not deuda_conceptos_desde_ledger) or mes == proyectar_mes:
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
                  deuda_conceptos_desde_ledger: bool = False,
                  proyectar_mes: str | None = None) -> pd.DataFrame:
    historicos = historicos if historicos is not None else _cargar_historicos()
    eventos = eventos if eventos is not None else repo._leer_eventos()
    mapa_raw = mapa_raw if mapa_raw is not None else _cargar_mapa_raw()
    if not nombre_actual:
        nombre_actual = repo._lookup_nombres().get((mz, lt), "")
    if dfp is None:
        f = _planilla_cobrado_path()
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
        deuda_conceptos_desde_ledger, proyectar_mes,
    ))
    tabla = pd.DataFrame(filas, columns=["MES", "CONSUMO", "MANT", "MES_ANT", "CORTE",
                                          "CONVENIO", "MULTA", "ACUERDOS", "TOTAL", "PAGO_COMPLETO", "NOTA",
                                          *[f"DEUDA_{c}" for c in CONCEPTOS_TABLA],
                                          *[f"PAGO_{c}" for c in CONCEPTOS_TABLA], "DEUDA_TOTAL", "PAGO_TOTAL",
                                          "APLICACION_PENDIENTE"])
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


def _dibujar_tabla_historico(page, x: float, y: float, w: float, tabla: pd.DataFrame) -> float:
    headers = ["Mes", "Tipo", "Consumo", "Mant.", "Mes ant.", "Corte",
               "Convenio", "Multa", "Acuerdos", "Total"]
    resto = w - 58 - 48 - 78
    col = resto / 7
    anchos = [58, 48, col, col, col, col, col, col, col, 78]

    rh_row = 13
    page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh_row - 4), h, fontsize=7, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row

    meses = tabla
    for n, (_, r) in enumerate(meses.iterrows()):
        es_ultimo_mes = n == len(meses) - 1
        estado_ledger = str(r.get("ESTADO_LEDGER", ""))
        provisional = estado_ledger in {"NO_ASENTADO", "APLICACION_PENDIENTE"}
        filas_mes = 3 if es_ultimo_mes else 2
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row * filas_mes), fill=_ZEBRA, color=None)
        tipos = [("DEUDA", "DEUDA_", _NEGRO), ("PAGO", "PAGO_", _VERDE)]
        if es_ultimo_mes:
            tipos.append(("SALDO", "DEUDA_", _ROJO))
        saldo = max(0.0, float(r.get("DEUDA_TOTAL", 0.0) or 0.0)
                    - float(r.get("PAGO_TOTAL", 0.0) or 0.0))
        for tipo, prefix, color in tipos:
            cx = x
            mes_label = f"{r['MES']}*" if provisional else str(r["MES"])
            page.insert_text((cx + 4, y + rh_row - 6), mes_label if tipo == "DEUDA" else "",
                             fontsize=6.5, fontname="hebo", color=_ROJO if provisional else _NEGRO)
            cx += anchos[0]
            page.insert_text((cx + 4, y + rh_row - 4), tipo, fontsize=6.5, fontname="hebo", color=color)
            cx += anchos[1]
            total = 0.0
            for c in CONCEPTOS_TABLA:
                v = float(r.get(f"{prefix}{c}", 0.0) or 0.0)
                if tipo == "SALDO":
                    v = max(0.0, v - float(r.get(f"PAGO_{c}", 0.0) or 0.0))
                total += v
                texto = f"{v:,.2f}" if abs(v) > 0.005 else "—"
                tw = fitz.get_text_length(texto, fontname="helv", fontsize=6.5)
                page.insert_text((cx + anchos[2] - 5 - tw, y + rh_row - 4), texto,
                                 fontsize=6.5, fontname="helv", color=color)
                cx += anchos[2]
            if tipo == "SALDO":
                total = saldo
            texto = f"S/ {total:,.2f}" if total > 0.005 else "—"
            tw = fitz.get_text_length(texto, fontname="hebo", fontsize=6.5)
            page.insert_text((cx + anchos[-1] - 5 - tw, y + rh_row - 4), texto,
                             fontsize=6.5, fontname="hebo", color=color)
            y += rh_row

        if r["NOTA"]:
            pass  # nota se omite en la tabla, ya lo dice el check "-"

    return y


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

    notas_snapshot = [str(n) for n in tabla.get("NOTA", pd.Series(dtype=str)).dropna()
                      if str(n).startswith("Reparto histórico según snapshot")]
    for nota_snapshot in notas_snapshot:
        page.insert_text((_M, y), nota_snapshot, fontsize=7, fontname="hebo", color=_ROJO)
        y += 10

    if "ESTADO_LEDGER" in tabla.columns and tabla["ESTADO_LEDGER"].isin(
            ["NO_ASENTADO", "APLICACION_PENDIENTE"]).any():
        pendiente = (tabla["ESTADO_LEDGER"] == "APLICACION_PENDIENTE").any()
        texto = ("* Pago asentado; aplicación por concepto pendiente de corrección."
                 if pendiente else "* Ciclo abierto: pago y saldo proyectados; NO ASENTADO EN LEDGER.")
        page.insert_text((_M, y), texto,
                         fontsize=8, fontname="hebo", color=_ROJO)
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


# ── Absorbido de reporte_referencias_pago.py (borrado -- generar_pdf() ahí era
# un duplicado desactualizado de reporte_historico.py::generar_pdf(), y todo
# lo demás ya lo consumían reporte_historico.py/verificar_yape.py/buscar_pago.py
# -- este último tenía además su PROPIA copia de referencias_pago()/
# verificar_predio(), nunca deduplicada) ───────────────────────────────────────

def referencias_pago(mz: str, lt: str, tabla: pd.DataFrame | None = None,
                     incluir_overlays: bool = True) -> list[dict]:
    """Lista de {MES, MEDIO, FECHA_HORA, MONTO} -- una entrada por pago detectado.

    El MONTO de cada mes SIEMPRE es el TOTAL que ya calculo _fila_historica()
    (mismo numero que aparece en la tabla de arriba) -- no una relectura aparte
    de una columna suelta. Hasta mayo 2026 ese total ya incluye mantenimiento
    arrastrado y exceso devuelto via "mes anterior" (asi se manejaba antes);
    la hoja Reporte/Efectivo solo aporta DONDE vino la plata (yape con hora,
    o efectivo), no el monto -- eso evita que un monto parcial de una sola
    transaccion descuadre contra el total real del mes."""
    out = []

    mapa_raw = _cargar_mapa_raw()
    mz_hist, lt_hist = mapa_raw.get((mz, lt), (mz, lt))

    for archivo, mes in _ARCHIVOS_HISTORICOS:
        hojas = _cargar_hojas_historicas(archivo)
        dfc = hojas.get("Cobranza")
        if dfc is None:
            continue
        cmz, clt, cest, cmed = (_col(dfc, "mz"), _col(dfc, "lt", "lote"),
                                 _col(dfc, "estado"), _col(dfc, "medio"))
        sub = dfc[(dfc[cmz].astype(str).str.strip() == mz_hist) & (dfc[clt].astype(str).str.strip() == lt_hist)]
        if sub.empty or str(sub.iloc[0][cest]).strip().lower() != "c":
            continue

        fila_mes = None
        if tabla is not None:
            m = tabla[tabla["MES"] == mes]
            if not m.empty:
                fila_mes = m.iloc[0]
        monto_mes = float(fila_mes["TOTAL"]) if fila_mes is not None else None
        if monto_mes is None or monto_mes <= 0.005:
            continue

        # Monto real por medio (columnas "Efectivo"/"Yape" del propio sheet
        # Cobranza, presentes en los 8 archivos y siempre Total==Efectivo+Yape,
        # 0 diferencias verificadas) -- monto_mes (TOTAL de conceptos) solo se
        # usa como gate de "hubo pago", no como el monto a mostrar.
        monto_efectivo = _val(sub.iloc[0], dfc, "Efectivo")
        monto_yape = _val(sub.iloc[0], dfc, "Yape")

        medio = str(sub.iloc[0][cmed]).strip().lower()
        if medio == "y":
            dfr = hojas.get("Reporte")
            if dfr is None:
                out.append({"MES": mes, "MEDIO": "YAPE", "FECHA_HORA": "sin hoja Reporte", "MONTO": monto_yape})
                continue
            cmzr, cltr = _col(dfr, "mz"), _col(dfr, "lote", "lt")
            ctr, cfr = _col(dfr, "tipo de transaccion"), _col(dfr, "fecha de operacion")
            subr = dfr[(dfr[cmzr].astype(str).str.strip() == mz_hist) & (dfr[cltr].astype(str).str.strip() == lt_hist)]
            subr = subr[subr[ctr].astype(str).str.upper().str.startswith("TE PAG", na=False)]
            fecha = str(subr.iloc[0][cfr]) if not subr.empty else "sin match en Reporte"
            out.append({"MES": mes, "MEDIO": "YAPE", "FECHA_HORA": fecha, "MONTO": monto_yape})
        else:
            out.append({"MES": mes, "MEDIO": "EFECTIVO", "FECHA_HORA": "--", "MONTO": monto_efectivo})

    # Cada ciclo posterior a mayo: total real pagado (yape+efectivo) desde el
    # planilla_cobrado DE SU PROPIO CICLO (los cerrados en su repo congelado,
    # el vigente en el repo activo -- ver _ciclos_recientes()) vs. lo que la
    # tabla de arriba muestra aplicado a conceptos -- desde junio 2026 el
    # exceso se RETIENE (no se aplica ni se muestra solo, ver nota del
    # reporte), asi que puede haber una diferencia real y legitima. Se
    # muestra como linea aparte, nunca mezclada.
    _PLANILLAS_RECIENTES = _ciclos_recientes()
    for mes_ciclo, f in _PLANILLAS_RECIENTES:
        # Los overlays entran a la cascada del ciclo pero NO viven en
        # planilla_cobrado -- van primero para que aparezcan aunque el ciclo
        # no tenga fila de yape/efectivo.
        if incluir_overlays:
            out.extend(_overlays_de_plata(mz, lt, mes_ciclo))
        if not f.exists():
            _avisar_falta(f, f"referencias de pago de {mes_ciclo}")
            continue
        dfp = pd.read_excel(f, sheet_name="planilla_cobrado", header=1)
        subp = dfp[(dfp["MZ"].astype(str).str.strip() == mz) & (dfp["LT"].astype(str).str.strip() == lt)]
        if subp.empty:
            continue
        r = subp.iloc[0]
        yape = _numf(r.get("MONTO_YAPE"))
        efectivo = _numf(r.get("MONTO_EFECTIVO"))
        total_real = yape + efectivo
        total_aplicado = 0.0
        if tabla is not None:
            m = tabla[tabla["MES"] == mes_ciclo]
            if not m.empty:
                total_aplicado = float(m.iloc[0]["TOTAL"])
        if yape > 0.005:
            # pagos_yape_tepago.xlsx directo -- ya viene resuelto 1 a 1 por
            # MZ/LOTE, no hace falta adivinar nada. _yape_fecha_hora (cruce
            # de nombre por prefijo contra el banco crudo) se sacó del
            # camino: cuando el origen registrado en maestro_yape.xlsx es
            # generico (ej. J-1: origen = "Janet Villanueva", la tesorera --
            # su nombre aparece en decenas de transacciones ajenas) pegaba
            # fechas de pagos de otros predios (bug encontrado 01/08/2026).
            fecha_yape = _fecha_yape_cruda(mz, lt, mes_ciclo)
            out.append({"MES": mes_ciclo, "MEDIO": "YAPE",
                        "FECHA_HORA": fecha_yape or "sin match en pagos_yape_tepago.xlsx (total del ciclo)",
                        "MONTO": yape})
        if efectivo > 0.005:
            dia, cobrador = _cobrador_efectivo(mz, lt, mes_ciclo)
            detalle = f"{dia} · {cobrador}" if cobrador else "sin detalle de transaccion (total del ciclo)"
            out.append({"MES": mes_ciclo, "MEDIO": "EFECTIVO", "FECHA_HORA": detalle, "MONTO": efectivo})
        # Si total_real y total_aplicado no coinciden (en cualquier sentido),
        # no se afirma de donde viene ni a donde fue la diferencia -- esto
        # solia mostrar una nota especulando el origen (blancos/abonos/
        # condonacion/convenio instalacion) y generaba confusion, porque esa
        # explicacion no esta confirmada (a veces es solo lo que la secretaria
        # declaro despues). Se deja en blanco a proposito: se muestra
        # unicamente el pago real encontrado, sin inventar el resto.

    return out


def _pagos_efectivo_crudo_path(mes_ciclo: str) -> Path:
    """Repo propio si el ciclo cerró, repo activo si es el vigente (ver
    _repo_de_ciclo). legacy_sin_periodo=True porque el ciclo activo todavía
    escribe pagos_efectivo.xlsx sin periodo además del nombre canónico
    (migra en la tanda B) y los repos cerrados de junio/julio quedaron con
    esa convención -- ciclo.resolver prueba el nombre con periodo primero,
    así que no cambia nada para los ciclos que ya migraron."""
    return ciclo.resolver(_repo_de_ciclo(mes_ciclo) / "4_pagos" / "efectivo" / "outputs",
                          "pagos_efectivo", mes_ciclo, legacy_sin_periodo=True)


_cache_pagos_efectivo: dict[str, pd.DataFrame] = {}


def _cargar_pagos_efectivo_crudo(mes_ciclo: str) -> pd.DataFrame | None:
    """FECHA real por transaccion -- trazabilidad_{mes}.xlsx (hoja
    solo_un_cobrador) tiene FECHA_COBRO vacio para TODO julio (bug de
    construccion, 221/221 filas verificado 01/08/2026) aunque el archivo
    crudo que la alimenta si trae fecha. Se usa como fuente de la fecha en
    vez de trazabilidad; trazabilidad sigue siendo la fuente del COBRADOR
    resuelto (para pagos divididos entre mesas)."""
    if mes_ciclo not in _cache_pagos_efectivo:
        p = _pagos_efectivo_crudo_path(mes_ciclo)
        if not p.exists():
            _avisar_falta(p, f"pagos_efectivo de {mes_ciclo}")
            _cache_pagos_efectivo[mes_ciclo] = pd.DataFrame()
        else:
            df = pd.read_excel(p, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            _cache_pagos_efectivo[mes_ciclo] = df
    return _cache_pagos_efectivo[mes_ciclo]


def _fecha_efectivo_cruda(mz: str, lt: str, mes_ciclo: str, cobrador: str) -> str | None:
    df = _cargar_pagos_efectivo_crudo(mes_ciclo)
    if df is None or df.empty or "FECHA" not in df.columns:
        return None
    sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LT"].astype(str).str.strip() == str(lt))]
    if sub.empty:
        return None
    if cobrador and "COBRADOR" in sub.columns:
        con_cobrador = sub[sub["COBRADOR"].astype(str).str.strip() == cobrador]
        if not con_cobrador.empty:
            sub = con_cobrador
    fechas = sub["FECHA"].dropna()
    if fechas.empty:
        return None
    return " · ".join(sorted({str(f) for f in fechas}))


def _cobrador_fecha_efectivo_crudo(mz: str, lt: str, mes_ciclo: str) -> tuple[str, str] | None:
    """Respaldo cuando el predio NO aparece en trazabilidad_{mes}.xlsx en
    absoluto (135 de 357 casos de julio, verificado 01/08/2026 -- gap
    distinto al de FECHA_COBRO vacio: acá falta la fila entera). Se toma
    directo de pagos_efectivo.xlsx filtrando ESTADO=solo_un_cobrador (mismo
    criterio que usaría trazabilidad si el predio estuviera ahí)."""
    df = _cargar_pagos_efectivo_crudo(mes_ciclo)
    if df is None or df.empty or "ESTADO" not in df.columns:
        return None
    sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LT"].astype(str).str.strip() == str(lt)) &
              (df["ESTADO"] == "solo_un_cobrador")]
    if sub.empty:
        return None
    r = sub.iloc[0]
    fecha = r.get("FECHA")
    dia = str(fecha) if pd.notna(fecha) else "(fecha no registrada)"
    cobrador = str(r.get("COBRADOR", "")).strip()
    return dia, cobrador


def _pagos_yape_crudo_path(mes_ciclo: str) -> Path:
    """Mismo criterio que _pagos_efectivo_crudo_path: repo propio si el ciclo
    cerró, repo activo si es el vigente."""
    return ciclo.resolver(_repo_de_ciclo(mes_ciclo) / "4_pagos" / "yape" / "motor_matching" / "outputs",
                          "pagos_yape_tepago", mes_ciclo, legacy_sin_periodo=True)


_cache_pagos_yape_crudo: dict[str, pd.DataFrame] = {}


def _cargar_pagos_yape_crudo(mes_ciclo: str) -> pd.DataFrame | None:
    """FECHA real por transaccion, directo de pagos_yape_tepago.xlsx (lo que
    alimenta maestro_yape.xlsx) -- respaldo cuando _yape_fecha_hora() no
    encuentra match cruzando maestro_yape contra el banco crudo (cruce por
    nombre truncado + fecha, se cae seguido). Un paso mas atras que ese
    cruce: esta fecha ya viene resuelta contra MZ/LOTE, sin necesidad de
    adivinar el origen."""
    if mes_ciclo not in _cache_pagos_yape_crudo:
        p = _pagos_yape_crudo_path(mes_ciclo)
        if not p.exists():
            _avisar_falta(p, f"pagos_yape_tepago de {mes_ciclo}")
            _cache_pagos_yape_crudo[mes_ciclo] = pd.DataFrame()
        else:
            df = pd.read_excel(p, header=1)
            df.columns = [str(c).strip() for c in df.columns]
            _cache_pagos_yape_crudo[mes_ciclo] = df
    return _cache_pagos_yape_crudo[mes_ciclo]


def _fecha_yape_cruda(mz: str, lt: str, mes_ciclo: str) -> str | None:
    df = _cargar_pagos_yape_crudo(mes_ciclo)
    if df is None or df.empty or "FECHA" not in df.columns or "LOTE" not in df.columns:
        return None
    sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LOTE"].apply(_norm_lote) == _norm_lote(lt))]
    fechas = sub["FECHA"].dropna()
    if fechas.empty:
        return None
    return " · ".join(sorted({str(f) for f in fechas}))


def _cobrador_efectivo(mz: str, lt: str, mes_ciclo: str) -> tuple[str, str]:
    """Dia de cobro y cobrador para un pago en efectivo de junio/julio 2026.
    Fuente primaria: 4_pagos/efectivo/trazabilidad/trazabilidad_{mes}.xlsx
    (hoja solo_un_cobrador). Dos gaps verificados en julio (01/08/2026), los
    dos con respaldo en pagos_efectivo.xlsx (el crudo que alimenta la
    trazabilidad):
      1. El predio SI esta en trazabilidad pero FECHA_COBRO vacio (221/221
         filas de julio) -> _fecha_efectivo_cruda().
      2. El predio NO esta en trazabilidad en absoluto (135/357 de julio,
         gap distinto) -> _cobrador_fecha_efectivo_crudo().
    Un solo fix acá cubre cualquier reporte que llame referencias_pago()."""
    f = REPO_DIR / "4_pagos" / "efectivo" / "trazabilidad" / f"trazabilidad_{mes_ciclo}.xlsx"
    df = None
    if f.exists():
        try:
            df = pd.read_excel(f, sheet_name="solo_un_cobrador", header=1)
        except Exception:
            df = None

    sub = None
    if df is not None:
        sub = df[(df["MZ"].astype(str).str.strip() == mz) & (df["LT"].astype(str).str.strip() == str(lt))]

    if sub is None or sub.empty:
        crudo = _cobrador_fecha_efectivo_crudo(mz, lt, mes_ciclo)
        return crudo if crudo else ("", "")

    r = sub.iloc[0]
    cobrador = str(r.get("COBRADOR", "")).strip()
    fecha = r.get("FECHA_COBRO")
    if pd.notna(fecha):
        dia = str(fecha)
    else:
        dia = _fecha_efectivo_cruda(mz, lt, mes_ciclo, cobrador) or "(fecha no registrada)"
    return dia, cobrador


def _dibujar_pagina_referencias(doc, mz: str, lt: str, nombre: str, refs: list[dict]) -> None:
    """Pagina aparte para la referencia de pago -- separada del historial
    mensual para que un predio con muchos meses (11+) no desborde una sola
    pagina A4 fija y corte filas (bug encontrado 18/08/2026)."""
    page = doc.new_page(width=_PAGE_W, height=_PAGE_H)
    w = _PAGE_W - 2 * _M
    y = _M
    page.draw_rect(fitz.Rect(_M, y, _M + w, y + 30), fill=_REFERENCIAS_BG, color=None)
    page.insert_text((_M + 10, y + 20), f"Referencias del predio {mz}-{lt} - {nombre or '(sin nombre)'}",
                      fontsize=12, fontname="hebo", color=_NEGRO)
    y += 42
    _dibujar_tabla_referencias(page, _M, y, w, refs)


def _dibujar_tabla_referencias(page, x: float, y: float, w: float, refs: list[dict]) -> float:
    headers = ["Mes", "Medio", "Fecha / hora", "Estado ledger", "Monto"]
    anchos = [65, 75, w - 65 - 75 - 160 - 95, 160, 95]
    rh_row = 16

    page.insert_text((x, y), "Referencia de pago (de donde vino cada pago)", fontsize=9, fontname="hebo", color=_AZUL)
    y += 16

    page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_AZUL_BG, color=None)
    cx = x
    for h, cw in zip(headers, anchos):
        page.insert_text((cx + 4, y + rh_row - 5), h, fontsize=8, fontname="hebo", color=_AZUL)
        cx += cw
    y += rh_row

    for n, r in enumerate(refs):
        if n % 2 == 1:
            page.draw_rect(fitz.Rect(x, y, x + w, y + rh_row), fill=_ZEBRA, color=None)
        cx = x
        page.insert_text((cx + 4, y + rh_row - 5), str(r["MES"]), fontsize=8, fontname="hebo", color=_NEGRO)
        cx += anchos[0]
        es_nota = r["MEDIO"] == "(nota)"
        color_medio = _GRIS if es_nota else (_VERDE if r["MEDIO"] == "YAPE" else _GRIS)
        page.insert_text((cx + 4, y + rh_row - 5), r["MEDIO"], fontsize=8, fontname="hebo", color=color_medio)
        cx += anchos[1]
        page.insert_text((cx + 4, y + rh_row - 5), str(r["FECHA_HORA"]), fontsize=7.5 if es_nota else 8,
                          fontname="helv", color=_GRIS if es_nota else _NEGRO)
        cx += anchos[2]
        estado_ledger = str(r.get("ESTADO_LEDGER", "ASENTADO"))
        page.insert_text((cx + 4, y + rh_row - 5), estado_ledger, fontsize=7.5,
                          fontname="hebo", color=_ROJO if estado_ledger != "ASENTADO" else _GRIS)
        cx += anchos[3]
        if not es_nota:
            texto_monto = f"S/ {r['MONTO']:,.2f}"
            tw = fitz.get_text_length(texto_monto, fontname="hebo", fontsize=8)
            page.insert_text((cx + anchos[-1] - 6 - tw, y + rh_row - 5), texto_monto, fontsize=8, fontname="hebo", color=_NEGRO)
        y += rh_row

    return y


def verificar_predio(mz: str, lt: str, tabla: pd.DataFrame | None = None, refs: list[dict] | None = None) -> list[dict]:
    """Compara, mes a mes, el TOTAL de la tabla de historial contra la suma de
    referencias de pago de ese mes. Devuelve solo los meses que NO cuadran
    (lista vacia = todo cuadra)."""
    if tabla is None:
        tabla = tabla_predio(mz, lt)
        tabla = corregir_tabla_por_redirects(tabla, mz, lt, _cargar_redirects())
    if refs is None:
        # sin tabla=tabla, referencias_pago() nunca entra al tramo historico
        # (su "if tabla is not None" se salta) y devuelve refs=[] -- cada mes
        # se marcaba "problema" en falso, aunque el pago si estaba (bug
        # encontrado 19/08/2026).
        refs = referencias_pago(mz, lt, tabla=tabla)

    ref_por_mes: dict[str, float] = {}
    for r in refs:
        ref_por_mes[r["MES"]] = ref_por_mes.get(r["MES"], 0.0) + r["MONTO"]

    _MESES_RECIENTES = {"2026-06", "2026-07"}

    problemas = []
    for _, row in tabla.iterrows():
        mes = row["MES"]
        total_tabla = float(row["TOTAL"])
        total_ref = ref_por_mes.get(mes, 0.0)
        # jun-jul: no se afirma el origen de una diferencia (ver
        # referencias_pago) -- un descuadre en cualquier sentido ahi es
        # esperado, no se marca como "problema" del reporte.
        if mes in _MESES_RECIENTES:
            continue
        if abs(total_tabla - total_ref) > 0.5:
            problemas.append({"MES": mes, "TOTAL_TABLA": total_tabla, "TOTAL_REFERENCIA": total_ref,
                               "DIFERENCIA": round(total_tabla - total_ref, 2)})
    return problemas
