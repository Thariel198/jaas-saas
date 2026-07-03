"""5_cobranza/main.py — Carga pagos en planilla · genera estado de cobro

Lee planilla (de 2_planilla) + pagos_yape + pagos_efectivo (de 4_pagos).
Genera 5 outputs:
  · planilla_cobrado.xlsx          — copia enriquecida con pagos + SALDO + ESTADO
  · trazabilidad_cobranza.xlsx     — un registro por pago cargado (acumulada)
  · resumen_recaudacion.xlsx       — totales del mes
  · arrastre_deuda_YYYY-MM.xlsx    — SALDO>0  → 2_planilla del próximo mes
  · arrastre_devolucion_YYYY-MM.xlsx — SALDO<0 → excesos pendientes de reclamo

SALDO sale como columna explícita en planilla_cobrado — la lista de corte la
genera 6_corte/generar_lista.py leyendo SALDO + MES_ANTERIOR desde acá.

Idempotente: si los pagos no cambiaron respecto a la trazabilidad existente,
sale sin modificar nada.
"""
import json
import logging
import re
import shutil
import sys
import pandas as pd
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
ROOT          = Path(__file__).parent

# seguimiento_pueblo: writer único es shared/seguimiento_repo.py (patrón repo,
# igual que data_boletas_repo). Se importa por ruta física — no depende de
# SHARED_DIR (que sí se monkey-patchea en tests) porque el módulo vive siempre
# en el shared/ real, sea cual sea la data que apunte ahí en un test.
sys.path.insert(0, str(ROOT.parent / "shared"))
import seguimiento_repo as repo  # noqa: E402
import utils_estado_ciclo as repo_estado  # noqa: E402
INPUTS_DIR    = ROOT / "inputs"
OUTPUTS_DIR   = ROOT / "outputs"
SHARED_DIR    = ROOT.parent / "shared"

# Single source of truth: la planilla del mes vive en shared/planilla_mes/.
# 2_planilla la publica ahí, 6_corte/aplicar_penalidad le suma CORTE_RECONEXION,
# y 4_pagos también la lee desde ahí. Leemos el master para que las penalidades
# (y cualquier cargo posterior) se reflejen sin copiar la planilla a mano.
PLAN_DIR      = SHARED_DIR / "planilla_mes"
# Single source of truth: los pagos se leen en vivo desde los outputs de 4_pagos,
# no desde copias en inputs/ (que quedaban stale y perdían pagos agregados después
# de copiar — ej. el 2º pago Yape de un predio). 5_cobranza retroescribe
# CICLO_COBRANZA en estos archivos (per contrato, le corresponde llenarla).
YAPE_DIR      = ROOT.parent / "4_pagos" / "yape" / "motor_matching" / "outputs"
EFEC_DIR      = ROOT.parent / "4_pagos" / "efectivo" / "outputs"
BLANCOS_PATH  = SHARED_DIR / "blancos_acumulados.xlsx"

# Overlay de penalidad (Modelo A · writer único). 6_corte y 6b ya NO escriben
# CORTE_RECONEXION en shared/planilla_mes — la penalidad vive SOLO en sus audits.
# 5_cobranza la re-deriva en vivo sumando el neto de PENALIDAD_APLICADA (col 5,
# ya viene con signo +/-) por predio. shared queda como base pura de 2_planilla.
AUDIT_CORTE_PATH = ROOT.parent / "6_corte" / "outputs" / "audit_penalidad.xlsx"
AUDIT_MULTA_PATH = ROOT.parent / "6b_corte_multas" / "outputs" / "audit_penalidad_multas.xlsx"

# Gate del arrastre_consolidado: solo se emite tras el sello de 5b_validacion.
ESTADO_CICLO_PATH = SHARED_DIR / "reporte_acumulado_procesado" / "estado_ciclo.json"

YAPE_FILE         = "pagos_yape_tepago.xlsx"
EFEC_FILE         = "pagos_efectivo.xlsx"
YAPE_DEV_FILE     = "pagos_yape_devolucion.xlsx"
YAPE_RETORNO_FILE = "pagos_yape_retorno.xlsx"
EFEC_DEV_FILE     = "pagos_efectivo_devolucion.xlsx"

CORR_LOTE_PATH = INPUTS_DIR / "correcciones_lote.xlsx"

PENALIDAD     = 20.0   # S/ por reconexión
ARRASTRE_MIN  = 8.0    # MES_ANTERIOR ≥ 8 confirma no pago anterior
TOL           = 0.005  # tolerancia de redondeo

# Posiciones en shared/blancos_acumulados.xlsx (1-indexed)
_BL_MZ   = 13
_BL_LOTE = 14
_BL_EST  = 18
_BL_MES  = 19

# ── PALETA — coincide con planilla_cobrado_diseno.html ───────────────────────
GH_QUIEN = ("EBF5FB", "1A5276")
GH_LEC   = ("E6F1FB", "0C447C")
GH_COB   = ("E9F7EF", "1E5C3A")
GH_DESC  = ("EDE9FE", "4C1D95")
GH_TOTAL = ("1E8449", "FFFFFF")
GH_PAGO  = ("F3E8FF", "5B21B6")
GH_TRAZ  = ("FEF9E7", "7D6608")

TD_QUIEN = "F4FAFF"
TD_LEC   = "F0F8FF"
TD_COB   = "F4FBF7"
TD_DESC  = "F5F3FF"
TD_TOTAL = "D5F5E3"
TD_PAGO  = "FAF5FF"
TD_TRAZ  = "FEFCE8"

ESTADO_BG  = {"CANCELADO": "E1F5EE", "EXCESO": "EFF6FF",
              "PARCIAL":   "FAEEDA", "PENDIENTE": "FEF2F2"}
ESTADO_TXT = {"CANCELADO": "085041", "EXCESO": "1D4ED8",
              "PARCIAL":   "854F0B", "PENDIENTE": "991B1B"}

# Badges para RETORNO (medio en que se devolvio el pago — puntero a archivo retorno)
RETORNO_BG  = {"yape": "E1F5EE", "efectivo": "EFF6FF", "mixto": "FEF9E7"}
RETORNO_TXT = {"yape": "085041", "efectivo": "1D4ED8", "mixto": "854F0B"}

# Paleta trazabilidad — ¿Quién es? (morado/lila)
GH_TZ_QUIEN  = ("F4ECF7", "5B21B6")
TD_TZ_QUIEN  = "FAF5FF"

# Paleta trazabilidad — ¿Cómo verificarlo? (ámbar) · ¿De qué ciclo? (azul)
GH_TZ_VERIF  = ("FFF3CD", "7C2D12")   # REFERENCIA · FECHA · COMENTARIO
TD_TZ_VERIF  = "FFFBEC"
GH_TZ_CICLO  = ("EFF6FF", "1E40AF")   # CICLO_CORRECCION_ORIGEN · CICLO_COBRANZA · FECHA_CARGA
TD_TZ_CICLO  = "EFF6FF"

# Paleta arrastre_deuda (coincide con arrastre_deuda_diseno.html)
GH_AD_QUIEN  = ("E8F8F5", "0E6655")
GH_AD_MONTO  = ("EAF2FF", "1A5276")
GH_AD_TRAZ   = ("E8F8F5", "0E6655")
TD_AD_QUIEN  = "F0FFF8"
TD_AD_MONTO  = "EBF5FB"
TD_AD_TRAZ   = "F0FFF8"

# arrastre_consolidado — P1..P5 por prioridad (formato_arrastre_consolidado.html)
GH_AC_QUIEN = ("E8F8F5", "0E6655")
GH_AC_TOTAL = ("1E8449", "FFFFFF")
# (header_bg, header_txt, cell_bg, cell_txt) por componente en orden de prioridad
_AC_P = [
    ("DEUDA_AGUA",        "EAF2FF", "1A5276", "F4F9FF", "1A5276"),  # P1
    ("CORTE_RECONEXION",  "FEF3C7", "92400E", "FFFBEB", "92400E"),  # P2
    ("MULTA",             "FEF2F2", "991B1B", "FFF5F5", "991B1B"),  # P3
    ("ACUERDOS_ASAMBLEA", "EDE9FE", "4C1D95", "F5F3FF", "4C1D95"),  # P4
    ("CONVENIO",          "E0F2FE", "0C4A6E", "F0F9FF", "0369A1"),  # P5
]
TD_AC_QUIEN = "F0FFF8"
TD_AC_TOTAL = "D5F5E3"
TD_AC_ZERO  = "F3F4F6"

# Paleta arrastre_devolucion (paleta EXCESO — azul)
GH_AV_QUIEN  = ("EFF6FF", "1D4ED8")
GH_AV_MONTO  = ("EFF6FF", "1D4ED8")
GH_AV_TRAZ   = ("EFF6FF", "1D4ED8")
TD_AV_QUIEN  = "F5F9FF"
TD_AV_MONTO  = "F5F9FF"
TD_AV_TRAZ   = "F5F9FF"
GH_AV_REVIS  = ("ECFDF5", "065F46")   # verde — revisión editable por el operador (¿legítimo o error?)
TD_AV_REVIS  = "F9FAFB"               # gris claro — celda vacía esperando input

# Paleta discrepancias_cobranza (coincide con formato_discrepancias_cobranza.html)
GH_DC_PREDIO = ("FEF2F2", "991B1B")   # rojo — el MZ+LT que no existe
GH_DC_PAGO   = ("EBF5FB", "1A5276")   # azul — monto y fecha
GH_DC_ORIGEN = ("FEF9E7", "7D6608")   # ambar — pista de origen (ORIGEN o MESA+COBRADOR)
GH_DC_TRAZ   = ("F4ECF7", "5B21B6")   # morado — ciclo y motivo
GH_DC_CORR   = ("ECFDF5", "065F46")   # verde — corrección editable por el operador
TD_DC_PREDIO = "FFF5F5"
TD_DC_PAGO   = "F4FAFF"
TD_DC_ORIGEN = "FFFDF5"
TD_DC_TRAZ   = "FAF5FF"
TD_DC_CORR   = "ECFDF5"   # verde claro — celda con corrección ingresada
TD_DC_CORR_V = "F9FAFB"   # gris muy claro — celda vacía esperando input


# ── LOGGING ───────────────────────────────────────────────────────────────────
def _init_logging():
    OUTPUTS_DIR.mkdir(exist_ok=True)
    # force=True: si el root logger ya esta configurado (tests, imports previos),
    # basicConfig sin force es no-op y el FileHandler nunca se agrega (metodologia v1.9).
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(OUTPUTS_DIR / "run.log", encoding="utf-8"),
        ],
        force=True,
    )

log = logging.getLogger(__name__)


# ── ESTILO HELPERS ────────────────────────────────────────────────────────────
def _borde():
    b = Side(style="thin", color="CCCCCC")
    return Border(left=b, right=b, top=b, bottom=b)

def _c(ws, row, col, value=None, bg=None, txt="333333",
       bold=False, align="left", mono=False, size=9, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="Consolas" if mono else "Arial",
                       size=size, bold=bold, color=txt)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border    = _borde()
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    if fmt:
        c.number_format = fmt
    return c

def _gh(ws, row, cs, ce, texto, bg, txt):
    ws.merge_cells(start_row=row, start_column=cs, end_row=row, end_column=ce)
    c = ws.cell(row=row, column=cs, value=texto)
    c.font      = Font(name="Arial", size=8, bold=True, color=txt)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill      = PatternFill("solid", start_color=bg)
    c.border    = _borde()

def _ch(ws, row, col, texto, bg, txt):
    _c(ws, row, col, texto, bg=bg, txt=txt, bold=True, align="center")

def _sep(ws, col, row_end):
    letra = get_column_letter(col)
    ws.column_dimensions[letra].width = 0.8
    b_sep = Border(left=Side(style="thin", color="D1D5DB"),
                   right=Side(style="thin", color="D1D5DB"))
    for r in range(1, row_end + 1):
        c = ws.cell(row=r, column=col)
        c.fill   = PatternFill("solid", start_color="F3F4F6")
        c.border = b_sep

def _w(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width


# ── UTILIDADES ────────────────────────────────────────────────────────────────
def _norm_mz(val) -> str:
    s = str(val).strip().upper()
    return "" if not s or s in ("NAN", "NONE") else s

def _norm_lt(val) -> str:
    s = str(val).strip()
    if not s or s.upper() in ("NONE", "NAN"):
        return ""
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s.upper().replace(" ", "")

def _float(val) -> float:
    if val is None:
        return 0.0
    try:
        f = float(str(val).replace(",", ".").strip())
        return 0.0 if f != f else f  # NaN guard
    except (ValueError, TypeError):
        return 0.0

def _txt(val) -> str:
    """Texto limpio: '' para None/nan/nat/none (evita 'nan' literal en celdas)."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s

def _ref_efectivo(mesa, cobrador) -> str:
    """REFERENCIA de pago en efectivo: 'MESA / COBRADOR' (omite el lado vacío)."""
    m, c = _txt(mesa), _txt(cobrador)
    if m and c:
        return f"{m} / {c}"
    return m or c

def _norm_cols(df) -> list[str]:
    cols = []
    for i, c in enumerate(df.columns):
        s = str(c).strip().upper()
        if not s or s.startswith("UNNAMED"):
            cols.append(f"_C{i}")
        else:
            cols.append(s)
    return cols

def _estado(saldo: float, total_pagado: float) -> str:
    if saldo < -TOL:
        return "EXCESO"
    if abs(saldo) <= TOL:
        return "CANCELADO"
    return "PARCIAL" if total_pagado > TOL else "PENDIENTE"

def _fecha_str(val) -> str:
    """Convierte fecha a string DD/MM/YYYY. Soporta str, Timestamp, datetime, NaN."""
    if val is None:
        return ""
    if isinstance(val, str):
        s = val.strip()
        if not s or s.upper() in ("NAN", "NAT", "NONE"):
            return ""
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
        if m:
            d, mes, a = m.groups()
            return f"{int(d):02d}/{int(mes):02d}/{a}"
        try:
            return pd.to_datetime(s).strftime("%d/%m/%Y")
        except Exception:
            return s[:10]
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    try:
        return pd.to_datetime(val).strftime("%d/%m/%Y")
    except Exception:
        return str(val)[:10]

def _fecha_hora_str(val) -> str:
    """Como _fecha_str pero preserva la hora si existe: DD/MM/YYYY HH:MM.
    Yape trae '21/06/2026 12:08:29' → '21/06/2026 12:08'. Efectivo sin hora → DD/MM/YYYY."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    # Timestamp/datetime ya parseado: formatear directo. NO hacer str()+dayfirst,
    # eso reinterpreta '2026-06-05' (5 jun) como 6 may y rompe la idempotencia.
    if isinstance(val, (pd.Timestamp, datetime)):
        dt = val
    else:
        s = str(val).strip()
        if not s or s.upper() in ("NAN", "NAT", "NONE"):
            return ""
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return _fecha_str(val)
    # sin componente horario → solo fecha
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
        return dt.strftime("%d/%m/%Y")
    return dt.strftime("%d/%m/%Y %H:%M")


def _fecha_hora_seg_str(val) -> str:
    """Como _fecha_hora_str pero con segundos: DD/MM/YYYY HH:MM:SS.
    Usado en REFERENCIA de arrastre_devolucion — ahí sí importa el segundo exacto."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(val, (pd.Timestamp, datetime)):
        dt = val
    else:
        s = str(val).strip()
        if not s or s.upper() in ("NAN", "NAT", "NONE"):
            return ""
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return _fecha_str(val)
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
        return dt.strftime("%d/%m/%Y")
    return dt.strftime("%d/%m/%Y %H:%M:%S")

def _fecha_max(fechas: list[str]) -> str:
    """Devuelve la fecha más reciente en formato DD/MM/YYYY."""
    fechas = [f for f in fechas if f]
    if not fechas:
        return ""
    dts = []
    for f in fechas:
        try:
            dts.append(datetime.strptime(f, "%d/%m/%Y"))
        except ValueError:
            pass
    if not dts:
        return fechas[0]
    return max(dts).strftime("%d/%m/%Y")


# ── VALIDACIÓN ────────────────────────────────────────────────────────────────
def _localizar_planilla() -> Path:
    candidatos = sorted(PLAN_DIR.glob("planilla_*.xlsx"))
    if not candidatos:
        raise FileNotFoundError(
            f"Falta planilla_YYYY-MM.xlsx en {PLAN_DIR}\n"
            f"  → Correr 2_planilla (publica la planilla en shared/planilla_mes/)"
        )
    if len(candidatos) > 1:
        log.warning(f"Múltiples planillas en {PLAN_DIR} → usando {candidatos[-1].name}")
    return candidatos[-1]

def _validar_inputs() -> Path:
    plan = _localizar_planilla()
    requeridos = [
        (YAPE_DIR / YAPE_FILE, "Copiar desde 4_pagos/yape/motor_matching/outputs/"),
        (EFEC_DIR / EFEC_FILE, "Copiar desde 4_pagos/efectivo/outputs/"),
    ]
    errores = []
    for ruta, sug in requeridos:
        if not ruta.exists():
            errores.append(f"Falta: {ruta}\n  → {sug}")
    if errores:
        for e in errores:
            log.error(e)
        raise FileNotFoundError("Inputs faltantes — ver errores arriba")
    log.info(f"Inputs OK · planilla = {plan.name}")
    return plan


# ── CARGA: PENALIDADES (overlay Modelo A) ────────────────────────────────────
def _cargar_penalidades(mes_ano: str) -> dict[tuple[str, str], float]:
    """Neto de penalidad de corte por predio, sumado de los audits de 6_corte y 6b.

    Fuente única de la penalidad: sus audit_*.xlsx. Se suma PENALIDAD_APLICADA
    (col 5, ya trae signo: +APLICADO / −REVERTIDO) → net delta. Robusto a que la
    base cambie: no depende del absoluto (col 6). Solo el ciclo MES_ANO actual.
    """
    neto: dict[tuple[str, str], float] = {}
    for p in (AUDIT_CORTE_PATH, AUDIT_MULTA_PATH):
        if not p.exists():
            continue
        df = pd.read_excel(p, header=1)
        df.columns = _norm_cols(df)
        if "PENALIDAD_APLICADA" not in df.columns:
            continue
        for _, f in df.iterrows():
            if str(f.get("MES_ANO", "")).strip() != mes_ano:
                continue
            mz = _norm_mz(f.get("MZ"))
            lt = _norm_lt(f.get("LT"))
            if not mz or not lt:
                continue
            key = (mz, lt)
            neto[key] = round(neto.get(key, 0.0) + _float(f.get("PENALIDAD_APLICADA")), 2)
    return {k: v for k, v in neto.items() if abs(v) > TOL}


# ── CARGA: PLANILLA ──────────────────────────────────────────────────────────
def _cargar_planilla(plan_path: Path) -> tuple[list[dict], str]:
    df = pd.read_excel(plan_path, header=1)
    df.columns = _norm_cols(df)
    requeridas = {"MZ", "LT", "NOMBRE", "MES_ANO", "MARC_ANT", "MARC_ACT", "M3",
                  "MES_ACTUAL", "MANTENIMIENTO", "MES_ANTERIOR", "CORTE_RECONEXION",
                  "CONVENIO", "MULTA", "ACUERDOS_ASAMBLEA",
                  "BLANCO", "DEVOLUCION", "TOTAL_A_PAGAR"}
    faltantes = requeridas - set(df.columns)
    if faltantes:
        raise ValueError(f"Planilla — columnas faltantes: {sorted(faltantes)}")

    usuarios, mes_ano = [], ""
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LT"))
        if not mz or not lt:
            continue
        if not mes_ano:
            mes_ano = str(f.get("MES_ANO", "")).strip()
        usuarios.append({
            "mz":       mz,
            "lt":       lt,
            "key":      f"{mz}-{lt}",
            "nombre":   str(f.get("NOMBRE", "")).strip(),
            "mes_ano":  str(f.get("MES_ANO", "")).strip(),
            "marc_ant": _float(f.get("MARC_ANT")),
            "marc_act": _float(f.get("MARC_ACT")),
            "m3":       _float(f.get("M3")),
            "mes_actual":        _float(f.get("MES_ACTUAL")),
            "mantenimiento":     _float(f.get("MANTENIMIENTO")),
            "mes_anterior":      _float(f.get("MES_ANTERIOR")),
            "corte_reconexion":  _float(f.get("CORTE_RECONEXION")),
            "convenio":          _float(f.get("CONVENIO")),
            "multa":             _float(f.get("MULTA")),
            "acuerdos_asamblea": _float(f.get("ACUERDOS_ASAMBLEA")),
            "blanco_inicial":    _float(f.get("BLANCO")),
            "devolucion":        _float(f.get("DEVOLUCION")),
        })
    # Overlay Modelo A: sumar penalidad de corte (audits de 6_corte/6b) sobre la
    # base leída de shared. shared ya no la trae — se re-deriva acá en cada corrida.
    penalidades = _cargar_penalidades(mes_ano)
    if penalidades:
        n = 0
        for u in usuarios:
            d = penalidades.get((u["mz"], u["lt"]), 0.0)
            if abs(d) > TOL:
                u["corte_reconexion"] = round(u["corte_reconexion"] + d, 2)
                n += 1
        log.info(f"Overlay penalidad · {n} predios con delta de audit (CORTE_RECONEXION)")

    log.info(f"Planilla {mes_ano} → {len(usuarios)} usuarios")
    return usuarios, mes_ano


# ── CARGA: PAGOS YAPE ────────────────────────────────────────────────────────
def _cargar_pagos_yape() -> list[dict]:
    df = pd.read_excel(YAPE_DIR / YAPE_FILE, header=1)
    df.columns = _norm_cols(df)
    ciclo_col = "CICLO_CORRECCION" if "CICLO_CORRECCION" in df.columns else "CICLO"

    pagos, sin_id = [], 0
    for idx, f in df.iterrows():
        if str(f.get("TIPO", "")).strip().upper() != "TE PAGÓ":
            continue
        mz   = _norm_mz(f.get("MZ"))
        lt   = _norm_lt(f.get("LOTE"))
        conc = str(f.get("CONCEPTO", "")).strip().upper()
        if not mz or not lt:
            sin_id += 1
            continue
        monto = _float(f.get("MONTO_ASIGNADO"))
        if monto <= TOL:
            monto = _float(f.get("MONTO_PAGO"))
        if monto <= TOL:
            continue
        concepto = conc.lower() if conc and conc not in ("NAN", "NONE") else ""
        pagos.append({
            "row":    idx + 3,                # fila Excel (filas 1-2 son cabeceras)
            "mz":     mz,
            "lt":     lt,
            "key":    f"{mz}-{lt}",
            "nombre": str(f.get("NOMBRE", "")).strip(),
            "monto":  round(monto, 2),
            "fecha":  _fecha_str(f.get("FECHA")),
            "ciclo_correccion": int(_float(f.get(ciclo_col)) or 1),
            "fuente": "yape",
            "concepto": concepto,
            # ORIGEN: nombre que el banco asigna al remitente (ej: "Wilder Tru*").
            # Sirve para rastrear pagos huérfanos en discrepancias_cobranza.xlsx.
            "origen": str(f.get("ORIGEN", "")).strip(),
            # Trazabilidad para auditoría (trazabilidad_cobranza.xlsx):
            # REFERENCIA = ORIGEN del banco · COMENTARIO = MENSAJE del pagador.
            "referencia": str(f.get("ORIGEN", "")).strip(),
            "comentario": str(f.get("MENSAJE", "")).strip(),
            "fecha_hora": _fecha_hora_str(f.get("FECHA")),
            "fecha_hora_seg": _fecha_hora_seg_str(f.get("FECHA")),
        })
    log.info(f"Pagos Yape → {len(pagos)} filas · {sin_id} sin identificar")
    return pagos


# ── CARGA: PAGOS EFECTIVO ────────────────────────────────────────────────────
def _cargar_pagos_efectivo() -> list[dict]:
    df = pd.read_excel(EFEC_DIR / EFEC_FILE, header=1)
    df.columns = _norm_cols(df)
    ciclo_col = "CICLO_CORRECCION" if "CICLO_CORRECCION" in df.columns else "CICLO"

    pagos = []
    for idx, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LT"))
        if not mz or not lt:
            continue
        monto = _float(f.get("MONTO"))
        if monto <= TOL:
            continue
        concepto = str(f.get("CONCEPTO") or "").strip().lower()
        if concepto.upper() in ("NAN", "NONE"):
            concepto = ""
        pagos.append({
            "row":    idx + 3,
            "mz":     mz,
            "lt":     lt,
            "key":    f"{mz}-{lt}",
            "nombre": "",
            "monto":  round(monto, 2),
            "fecha":  _fecha_str(f.get("FECHA")),
            "ciclo_correccion": int(_float(f.get(ciclo_col)) or 1),
            "fuente": "efectivo",
            "concepto": concepto,
            # MESA + COBRADOR: pista física para rastrear pagos huérfanos.
            # MESA dice en cuál mesa_N.xlsx buscar, COBRADOR quién lo registró.
            "mesa":     str(f.get("MESA", "")).strip(),
            "cobrador": str(f.get("COBRADOR", "")).strip(),
            # Trazabilidad para auditoría (trazabilidad_cobranza.xlsx):
            # REFERENCIA = "MESA / COBRADOR" · COMENTARIO = nota del cobrador.
            "referencia": _ref_efectivo(f.get("MESA"), f.get("COBRADOR")),
            "comentario": _txt(f.get("COMENTARIO")),
            "fecha_hora": _fecha_hora_str(f.get("FECHA")),
        })
    log.info(f"Pagos Efectivo → {len(pagos)} filas")
    return pagos


# ── CARGA: BLANCOS ───────────────────────────────────────────────────────────
def _cargar_blancos() -> dict:
    if not BLANCOS_PATH.exists():
        log.info("blancos_acumulados.xlsx no encontrado → sin blancos")
        return {}
    df = pd.read_excel(BLANCOS_PATH, header=1)
    df.columns = _norm_cols(df)
    blancos = {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LOTE") if "LOTE" in df.columns else f.get("LT"))
        if not mz or not lt:
            continue
        est = str(f.get("ESTADO", "")).strip().lower()
        if est == "aplicado":
            continue
        key = f"{mz}-{lt}"
        blancos[key] = blancos.get(key, 0.0) + _float(f.get("MONTO"))
    if blancos:
        log.info(f"Blancos pendientes → {len(blancos)} lotes")
    return blancos


# ── CARGA: RETORNOS (yape + efectivo) ────────────────────────────────────────
# Los retornos reducen MONTO_YAPE (cualquier medio de devolucion).
# El badge RETORNO en planilla_cobrado y trazabilidad apunta al archivo origen.
# Ambos archivos son opcionales — si no existen, no hay retornos este ciclo.

def _cargar_retornos_yape() -> dict:
    """Retorna {key MZ-LT: monto_total}. Lee pagos_yape_retorno.xlsx. Archivo opcional."""
    path = YAPE_DIR / YAPE_RETORNO_FILE
    if not path.exists():
        log.info(f"{YAPE_RETORNO_FILE} no encontrado → sin retornos Yape")
        return {}
    df = pd.read_excel(path, header=1)
    df.columns = _norm_cols(df)
    devs = {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LOTE"))
        if not mz or not lt:
            continue
        monto = _float(f.get("MONTO"))
        if monto <= TOL:
            continue
        k = f"{mz}-{lt}"
        devs[k] = devs.get(k, 0.0) + round(monto, 2)
    if devs:
        log.info(f"Retornos Yape → {len(devs)} lotes · S/ {sum(devs.values()):.2f}")
    return devs


def _cargar_devueltos_yape() -> dict:
    """Retorna {key MZ-LT: monto_total}. Lee pagos_yape_devolucion.xlsx. Archivo opcional."""
    path = YAPE_DIR / YAPE_DEV_FILE
    if not path.exists():
        log.info(f"{YAPE_DEV_FILE} no encontrado → sin devoluciones Yape")
        return {}
    df = pd.read_excel(path, header=1)
    df.columns = _norm_cols(df)
    devs = {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LOTE"))
        if not mz or not lt:
            continue
        monto = _float(f.get("MONTO"))
        if monto <= TOL:
            continue
        k = f"{mz}-{lt}"
        devs[k] = devs.get(k, 0.0) + round(monto, 2)
    if devs:
        log.info(f"Devoluciones Yape → {len(devs)} lotes · S/ {sum(devs.values()):.2f}")
    return devs


def _cargar_retornos_efectivo() -> dict:
    """Retorna {key MZ-LT: monto_total}. Archivo opcional."""
    path = EFEC_DIR / EFEC_DEV_FILE
    if not path.exists():
        log.info(f"{EFEC_DEV_FILE} no encontrado → sin retornos Efectivo")
        return {}
    df = pd.read_excel(path, header=1)
    df.columns = _norm_cols(df)
    devs = {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LOTE"))
        if not mz or not lt:
            continue
        monto = _float(f.get("MONTO"))
        if monto <= TOL:
            continue
        k = f"{mz}-{lt}"
        devs[k] = devs.get(k, 0.0) + round(monto, 2)
    if devs:
        log.info(f"Retornos Efectivo → {len(devs)} lotes · S/ {sum(devs.values()):.2f}")
    return devs


def _retorno_badge(yape_dev: float, efec_dev: float) -> str | None:
    """Calcula el badge RETORNO para un lote a partir de sus retornos."""
    has_y = yape_dev > TOL
    has_e = efec_dev > TOL
    if has_y and has_e:
        return "mixto"
    if has_y:
        return "yape"
    if has_e:
        return "efectivo"
    return None


def _retornos_por_lote(dev_yape: dict, dev_efec: dict) -> dict:
    """Combina los retornos yape y efectivo en {key: badge}. Solo incluye lotes con retorno."""
    keys = set(dev_yape) | set(dev_efec)
    out = {}
    for k in keys:
        badge = _retorno_badge(dev_yape.get(k, 0.0), dev_efec.get(k, 0.0))
        if badge:
            out[k] = badge
    return out


def _devueltos_por_lote(dev_devuelto: dict) -> dict:
    """Devuelve {key: 'yape'} para lotes con devolución Yape. Solo incluye lotes con devuelto."""
    return {k: "yape" for k, v in dev_devuelto.items() if v > TOL}


def _retornos_planilla_previa() -> dict:
    """Lee el estado RETORNO de la planilla_cobrado anterior — para detectar cambios e idempotencia."""
    path = OUTPUTS_DIR / "planilla_cobrado.xlsx"
    if not path.exists():
        return {}
    try:
        df = pd.read_excel(path, header=1)
        df.columns = _norm_cols(df)
    except Exception:
        return {}
    if "RETORNO" not in df.columns:
        return {}
    out = {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LT"))
        if not mz or not lt:
            continue
        val = f.get("RETORNO")
        if val is None:
            continue
        badge = str(val).strip().lower()
        if badge in ("yape", "efectivo", "mixto"):
            out[f"{mz}-{lt}"] = badge
    return out


# ── CARGA: TRAZABILIDAD PREVIA (para idempotencia + ciclo) ───────────────────
def _identidad_pago(p: dict) -> tuple:
    # Usa el lote original como identidad para que la idempotencia funcione
    # aunque en runs futuros el key se remapee a otro lote vía correcciones_lote.
    return (p.get("mz_origen", p["mz"]),
            p.get("lt_origen", p["lt"]),
            p["monto"], p["fuente"],
            p["fecha"], p["ciclo_correccion"])

def _cargar_trazabilidad_previa() -> tuple[set, int]:
    """Retorna (set de identidades ya cargadas, max CICLO_COBRANZA usado)."""
    p = OUTPUTS_DIR / "trazabilidad_cobranza.xlsx"
    if not p.exists():
        return set(), 0
    df = pd.read_excel(p, header=1)
    df.columns = _norm_cols(df)
    ids, mx = set(), 0
    for _, f in df.iterrows():
        # Si el pago fue corregido, MZ_ORIGEN/LT_ORIGEN guardan la identidad real
        # (lo que el cobrador escribió). Usar eso para que la idempotencia coincida.
        mz = _norm_mz(f.get("MZ_ORIGEN")) or _norm_mz(f.get("MZ"))
        lt = _norm_lt(f.get("LT_ORIGEN")) or _norm_lt(f.get("LT"))
        if not mz or not lt:
            continue
        ident = (mz, lt,
                 round(_float(f.get("MONTO")), 2),
                 str(f.get("FUENTE", "")).strip().lower(),
                 _fecha_str(f.get("FECHA")) if "FECHA" in df.columns else "",
                 int(_float(f.get("CICLO_CORRECCION_ORIGEN")) or 0))
        ids.add(ident)
        mx = max(mx, int(_float(f.get("CICLO_COBRANZA")) or 0))
    return ids, mx


# ── CORRECCIONES DE LOTE ─────────────────────────────────────────────────────
# El operador llena MZ_CORRECTO + LT_CORRECTO en discrepancias_cobranza.xlsx.
# El módulo las absorbe, las persiste en inputs/correcciones_lote.xlsx y las
# aplica a todos los pagos antes del matching — el registro de origen no se toca.

def _leer_correcciones() -> dict:
    """Lee correcciones persistidas en correcciones_lote.xlsx.
    Retorna {(mz_origen, lt_origen): (mz_destino, lt_destino)}.
    """
    if not CORR_LOTE_PATH.exists():
        return {}
    df = pd.read_excel(CORR_LOTE_PATH, header=0)
    df.columns = _norm_cols(df)
    corr = {}
    for _, f in df.iterrows():
        mo = _norm_mz(f.get("MZ_ORIGEN"))
        lo = _norm_lt(f.get("LT_ORIGEN"))
        md = _norm_mz(f.get("MZ_DESTINO"))
        ld = _norm_lt(f.get("LT_DESTINO"))
        if mo and lo and md and ld:
            corr[(mo, lo)] = (md, ld)
    if corr:
        log.info(f"correcciones_lote.xlsx → {len(corr)} remapeos activos")
    return corr


def _escribir_correcciones_lote(filas: list[dict]):
    """Escribe correcciones_lote.xlsx — delega el dibujo al primitivo compartido
    (shared/utils_lote.py) para que 7_cierre pueda resetearlo con el mismo formato."""
    from utils_lote import escribir_correcciones_lote
    escribir_correcciones_lote(CORR_LOTE_PATH, filas)


def _recuperar_correcciones_trazabilidad(existentes: dict, ciclo: int) -> dict:
    """Auto-sana correcciones_lote desde trazabilidad_cobranza.xlsx.
    Cada corrección aplicada queda grabada en la trazabilidad (MZ_ORIGEN/LT_ORIGEN
    → MZ/LT). Si correcciones_lote se revirtió (git) y perdió alguna, se recupera
    desde ahí — el trabajo manual no depende de un solo archivo mutable.
    """
    traz = OUTPUTS_DIR / "trazabilidad_cobranza.xlsx"
    if not traz.exists():
        return existentes
    try:
        df = pd.read_excel(traz, header=1, dtype=str)
    except Exception as e:
        log.warning(f"No se pudo leer trazabilidad para recuperar correcciones: {e}")
        return existentes
    df.columns = _norm_cols(df)
    if not {"MZ_ORIGEN", "LT_ORIGEN", "MZ", "LT"} <= set(df.columns):
        return existentes

    nuevas = {}
    for _, f in df.iterrows():
        mo = _norm_mz(f.get("MZ_ORIGEN"))
        lo = _norm_lt(f.get("LT_ORIGEN"))
        md = _norm_mz(f.get("MZ"))
        ld = _norm_lt(f.get("LT"))
        if mo and lo and md and ld and (mo, lo) not in existentes:
            nuevas[(mo, lo)] = (md, ld)
    if not nuevas:
        return existentes

    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    filas_prev = []
    if CORR_LOTE_PATH.exists():
        df_prev = pd.read_excel(CORR_LOTE_PATH, header=0)
        df_prev.columns = _norm_cols(df_prev)
        filas_prev = df_prev.to_dict("records")
    filas_nuevas = [
        {"MZ_ORIGEN": mo, "LT_ORIGEN": lo, "MZ_DESTINO": md, "LT_DESTINO": ld,
         "MOTIVO": "Recuperado desde trazabilidad_cobranza.xlsx",
         "CICLO": ciclo, "FECHA": ahora}
        for (mo, lo), (md, ld) in nuevas.items()
    ]
    _escribir_correcciones_lote(filas_prev + filas_nuevas)
    log.info(f"correcciones_lote.xlsx → {len(nuevas)} recuperada(s) desde trazabilidad")
    for (mo, lo), (md, ld) in nuevas.items():
        log.info(f"  recuperada: {mo}-{lo} → {md}-{ld}")
    return {**existentes, **nuevas}


def _absorber_correcciones_discrepancias(existentes: dict, ciclo: int) -> dict:
    """Lee MZ_CORRECTO+LT_CORRECTO llenados en discrepancias_cobranza.xlsx.
    Guarda las nuevas en correcciones_lote.xlsx y retorna el mapa combinado.
    """
    ruta = OUTPUTS_DIR / "discrepancias_cobranza.xlsx"
    if not ruta.exists():
        return existentes

    nuevas = {}
    try:
        wb_disc = load_workbook(ruta, data_only=True)
        for sheet in wb_disc.sheetnames:
            ws = wb_disc[sheet]
            hdrs = {str(ws.cell(2, c).value or "").strip().upper(): c
                    for c in range(1, ws.max_column + 1)}
            col_mzo = hdrs.get("MZ")
            col_lto = hdrs.get("LT")
            col_mzc = hdrs.get("MZ_CORRECTO")
            col_ltc = hdrs.get("LT_CORRECTO")
            if not all([col_mzo, col_lto, col_mzc, col_ltc]):
                continue
            for r in range(3, ws.max_row + 1):
                mo = _norm_mz(ws.cell(r, col_mzo).value)
                lo = _norm_lt(ws.cell(r, col_lto).value)
                mc = _norm_mz(ws.cell(r, col_mzc).value)
                lc = _norm_lt(ws.cell(r, col_ltc).value)
                if mo and lo and mc and lc and (mo, lo) not in existentes:
                    nuevas[(mo, lo)] = (mc, lc)
    except Exception as e:
        log.warning(f"No se pudo leer correcciones de discrepancias_cobranza.xlsx: {e}")
        return existentes

    if not nuevas:
        return existentes

    # Persistir en correcciones_lote.xlsx
    CORR_LOTE_PATH.parent.mkdir(exist_ok=True)
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    filas_prev = []
    if CORR_LOTE_PATH.exists():
        df_prev = pd.read_excel(CORR_LOTE_PATH, header=0)
        df_prev.columns = _norm_cols(df_prev)
        filas_prev = df_prev.to_dict("records")

    filas_nuevas = [
        {"MZ_ORIGEN": mo, "LT_ORIGEN": lo,
         "MZ_DESTINO": md, "LT_DESTINO": ld,
         "MOTIVO": "Corregido desde discrepancias_cobranza.xlsx",
         "CICLO": ciclo, "FECHA": ahora}
        for (mo, lo), (md, ld) in nuevas.items()
    ]

    _escribir_correcciones_lote(filas_prev + filas_nuevas)
    log.info(f"correcciones_lote.xlsx → {len(nuevas)} nueva(s) guardada(s) · "
             f"total {len(filas_prev) + len(filas_nuevas)}")

    return {**existentes, **nuevas}


def _aplicar_correcciones_lote(pagos: list[dict], correcciones: dict) -> list[dict]:
    """Remapea MZ+LT de pagos según correcciones. Preserva mz_origen/lt_origen."""
    if not correcciones:
        return pagos
    for p in pagos:
        key = (p["mz"], p["lt"])
        if key in correcciones:
            mzd, ltd = correcciones[key]
            p["mz_origen"] = p["mz"]
            p["lt_origen"] = p["lt"]
            p["mz"]  = mzd
            p["lt"]  = ltd
            p["key"] = f"{mzd}-{ltd}"
            log.info(f"  corrección aplicada: {key[0]}-{key[1]} → {mzd}-{ltd}")
    return pagos


# ── CÁLCULO ──────────────────────────────────────────────────────────────────
def _calcular(usuarios: list[dict],
              pagos_yape: list[dict],
              pagos_efectivo: list[dict],
              blancos: dict,
              dev_yape: dict,
              dev_efec: dict,
              dev_devuelto: dict,
              ciclo_nuevo: int,
              pagos_nuevos: set) -> tuple[list[dict], set]:
    yape_por_key: dict[str, list[dict]] = {}
    efec_por_key: dict[str, list[dict]] = {}
    for p in pagos_yape:
        yape_por_key.setdefault(p["key"], []).append(p)
    for p in pagos_efectivo:
        efec_por_key.setdefault(p["key"], []).append(p)

    keys_validos = {u["key"] for u in usuarios}
    huerfanos = (set(yape_por_key) | set(efec_por_key)) - keys_validos
    for k in sorted(huerfanos):
        log.warning(f"Pago para {k} pero no está en planilla → discrepancias_cobranza.xlsx")

    # Lotes con retornos pero sin estar en planilla — anomalía
    huerfanos_dev = (set(dev_yape) | set(dev_efec)) - keys_validos
    for k in sorted(huerfanos_dev):
        log.warning(f"Retorno para {k} pero no está en planilla — ignorado")

    resultado = []
    blancos_usados = set()
    for u in usuarios:
        k  = u["key"]
        ys = yape_por_key.get(k, [])
        es = efec_por_key.get(k, [])
        ys_agua = [p for p in ys if not p.get("concepto")]
        es_agua = [p for p in es if not p.get("concepto")]
        yape_sum = round(sum(p["monto"] for p in ys_agua), 2)
        efec_sum = round(sum(p["monto"] for p in es_agua), 2)

        # Retornos y devoluciones — ambos reducen MONTO_YAPE.
        # RETORNO badge: desde pagos_yape_retorno.xlsx + pagos_efectivo_devolucion.xlsx
        # DEVUELTO badge: desde pagos_yape_devolucion.xlsx
        yape_dev_lote     = round(dev_yape.get(k, 0.0), 2)
        efec_dev_lote     = round(dev_efec.get(k, 0.0), 2)
        devuelto_lote     = round(dev_devuelto.get(k, 0.0), 2)
        total_dev         = round(yape_dev_lote + efec_dev_lote + devuelto_lote, 2)
        yape_neto         = round(yape_sum - total_dev, 2)
        if yape_neto < -TOL:
            log.warning(f"Lote {k}: retorno+devuelto (S/ {total_dev:.2f}) excede pago Yape "
                        f"(S/ {yape_sum:.2f}) → MONTO_YAPE quedará negativo")
        retorno_badge = _retorno_badge(yape_dev_lote, efec_dev_lote)
        devuelto_badge = "yape" if devuelto_lote > TOL else None

        pagado = round(yape_neto + efec_sum, 2)

        blanco_aplicar = round(blancos.get(k, 0.0), 2)
        if blanco_aplicar > TOL:
            blancos_usados.add(k)
        # BLANCO en planilla queda negativo (reduce el total)
        blanco_final = round(u["blanco_inicial"] - blanco_aplicar, 2)

        total = round(
            u["mes_actual"] + u["mantenimiento"]
            + u["mes_anterior"] + u["corte_reconexion"]
            + u["convenio"] + u["multa"] + u["acuerdos_asamblea"]
            + blanco_final + u["devolucion"],
            2,
        )
        saldo = round(total - pagado, 2)
        fechas = [p["fecha"] for p in (ys_agua + es_agua)]

        # CICLO_COBRANZA del usuario: ciclo del pago más reciente cargado.
        # Si alguno de sus pagos es "nuevo" este run → ciclo_nuevo.
        # Si solo tiene pagos viejos → ciclo lo dejamos en None (se respeta el previo).
        ciclo_user = ciclo_nuevo if any(
            _identidad_pago(p) in pagos_nuevos for p in (ys + es)
        ) else None
        # Si no tiene pagos → vacío
        if not (ys or es):
            ciclo_user = None

        resultado.append({
            **u,
            "blanco_final":   blanco_final,
            "total_a_pagar":  total,
            "monto_yape":     yape_neto,
            "monto_efectivo": efec_sum,
            "total_pagado":   pagado,
            "saldo":          saldo,
            "estado":         _estado(saldo, pagado),
            "fecha_pago":     _fecha_max(fechas),
            "ciclo_cobranza": ciclo_user,
            "retorno":        retorno_badge,
            "devuelto":       devuelto_badge,
            "pagos_yape":     ys,
            "pagos_efectivo": es,
        })
    cnt = {e: sum(1 for r in resultado if r["estado"] == e)
           for e in ("CANCELADO", "EXCESO", "PARCIAL", "PENDIENTE")}
    log.info(f"Estados → CANCELADO={cnt['CANCELADO']} EXCESO={cnt['EXCESO']} "
             f"PARCIAL={cnt['PARCIAL']} PENDIENTE={cnt['PENDIENTE']}")
    n_retornos  = sum(1 for r in resultado if r["retorno"])
    n_devueltos = sum(1 for r in resultado if r["devuelto"])
    if n_retornos:
        log.info(f"Lotes con retorno → {n_retornos}")
    if n_devueltos:
        log.info(f"Lotes con devuelto → {n_devueltos}")
    return resultado, blancos_usados


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 1 — planilla_cobrado.xlsx
# ─────────────────────────────────────────────────────────────────────────────
# Layout (matching planilla_cobrado_diseno.html):
#   1- 4  ¿Quién es?     MZ LT NOMBRE MES_ANO
#   5      sep
#   6- 8  Lectura        MARC_ANT MARC_ACT M3
#   9      sep
#  10-16  Cobro—cargos   MES_ACTUAL MANTENIMIENTO MES_ANTERIOR CORTE_RECONEXION
#                         CONVENIO MULTA ACUERDOS_ASAMBLEA
#  17-18  Descuentos     BLANCO DEVOLUCION
#  19     Total          TOTAL_A_PAGAR
#  20     sep
#  21-26  Pago→5_cob     MONTO_YAPE MONTO_EFECTIVO SALDO RETORNO ESTADO FECHA_PAGO
#  27     sep
#  28     ¿Cuándo?       CICLO_COBRANZA
#
# SALDO se expone como columna del contrato — la consumen 6_corte y futuras
# tools sin re-implementar la fórmula (total − pagado).
# ─────────────────────────────────────────────────────────────────────────────

_PC_GRUPOS = [
    (1,  4,  "¿Quién es?",          *GH_QUIEN),
    (6,  8,  "Lectura",              *GH_LEC),
    (10, 16, "Cobro — cargos",       *GH_COB),
    (17, 18, "Descuentos",           *GH_DESC),
    (19, 19, "Total",                *GH_TOTAL),
    (21, 27, "Pago → 5_cobranza",    *GH_PAGO),
    (29, 29, "¿Cuándo?",             *GH_TRAZ),
]
_PC_COLS = [
    (1,  "MZ",                *GH_QUIEN,   6),
    (2,  "LT",                *GH_QUIEN,   6),
    (3,  "NOMBRE",            *GH_QUIEN,  26),
    (4,  "MES_ANO",           *GH_QUIEN,  10),
    (6,  "MARC_ANT",          *GH_LEC,     9),
    (7,  "MARC_ACT",          *GH_LEC,     9),
    (8,  "M3",                *GH_LEC,     6),
    (10, "MES_ACTUAL",        *GH_COB,    11),
    (11, "MANTENIMIENTO",     *GH_COB,    13),
    (12, "MES_ANTERIOR",      *GH_COB,    12),
    (13, "CORTE_RECONEXION",  *GH_COB,    16),
    (14, "CONVENIO",          *GH_COB,    10),
    (15, "MULTA",             *GH_COB,     8),
    (16, "ACUERDOS_ASAMBLEA", *GH_COB,    17),
    (17, "BLANCO",            *GH_DESC,    9),
    (18, "DEVOLUCION",        *GH_DESC,   11),
    (19, "TOTAL_A_PAGAR",     *GH_TOTAL,  13),
    (21, "MONTO_YAPE",        *GH_PAGO,   11),
    (22, "MONTO_EFECTIVO",    *GH_PAGO,   14),
    (23, "SALDO",             *GH_PAGO,   12),
    (24, "RETORNO",           *GH_PAGO,   10),
    (25, "DEVUELTO",          *GH_PAGO,   10),
    (26, "ESTADO",            *GH_PAGO,   11),
    (27, "FECHA_PAGO",        *GH_PAGO,   11),
    (29, "CICLO_COBRANZA",    *GH_TRAZ,   14),
]
_PC_SEP_COLS = [5, 9, 20, 28]


def _exportar_planilla_cobrado(resultado: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "planilla_cobrado"
    ws.freeze_panes = "A3"
    last_row = len(resultado) + 2

    for cs, ce, texto, bg, txt in _PC_GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)
    for sc in _PC_SEP_COLS:
        _sep(ws, sc, last_row)
    for col, nombre, bg, txt, ancho in _PC_COLS:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    MONEY = '"S/ "#,##0.00'

    for ri, r in enumerate(resultado, 3):
        # Quién es?
        _c(ws, ri,  1, r["mz"],      TD_QUIEN, "1A5276", mono=True, align="center")
        _c(ws, ri,  2, r["lt"],      TD_QUIEN, "1A5276", mono=True, align="center")
        _c(ws, ri,  3, r["nombre"],  TD_QUIEN, "333333", align="left")
        _c(ws, ri,  4, r["mes_ano"], TD_QUIEN, "1A5276", mono=True, align="center")

        # Lectura
        _c(ws, ri,  6, r["marc_ant"], TD_LEC, "0C447C", mono=True, align="right")
        _c(ws, ri,  7, r["marc_act"], TD_LEC, "0C447C", mono=True, align="right")
        _c(ws, ri,  8, r["m3"],       TD_LEC, "065F46", mono=True, align="right", bold=True)

        # Cobro — cargos (None si es 0 → muestra vacío)
        def _cob(col, val):
            val_disp = val if val > TOL else None
            _c(ws, ri, col, val_disp, TD_COB, "1E5C3A",
               mono=True, align="right", fmt=MONEY if val_disp else None)

        _cob(10, r["mes_actual"])
        _cob(11, r["mantenimiento"])
        _cob(12, r["mes_anterior"])
        _cob(13, r["corte_reconexion"])
        _cob(14, r["convenio"])
        _cob(15, r["multa"])
        _cob(16, r["acuerdos_asamblea"])

        # Descuentos (negativos)
        def _desc(col, val):
            val_disp = val if abs(val) > TOL else None
            _c(ws, ri, col, val_disp, TD_DESC, "4C1D95",
               mono=True, align="right", bold=True,
               fmt=MONEY if val_disp else None)

        _desc(17, r["blanco_final"])
        _desc(18, r["devolucion"] if r["devolucion"] >= 0 else -r["devolucion"])

        # Total (fórmula Excel: J:R = cols 10-18)
        formula = f"=SUM(J{ri}:R{ri})"
        c_tot = ws.cell(row=ri, column=19, value=formula)
        c_tot.font          = Font(name="Consolas", size=10, bold=True, color="1E5C3A")
        c_tot.fill          = PatternFill("solid", start_color=TD_TOTAL)
        c_tot.alignment     = Alignment(horizontal="right", vertical="center")
        c_tot.border        = _borde()
        c_tot.number_format = MONEY

        # Pago → 5_cobranza
        def _pag(col, val):
            val_disp = val if val > TOL else None
            _c(ws, ri, col, val_disp, TD_PAGO, "5B21B6",
               mono=True, align="right", fmt=MONEY if val_disp else None)

        # MONTO_YAPE puede ser negativo si el retorno excede el pago — no lo ocultes
        _pag_yape_val = r["monto_yape"]
        if abs(_pag_yape_val) > TOL:
            _c(ws, ri, 21, _pag_yape_val, TD_PAGO, "5B21B6",
               mono=True, align="right", bold=(_pag_yape_val < 0),
               fmt=MONEY)
        else:
            _c(ws, ri, 21, None, TD_PAGO, "5B21B6", mono=True, align="right")
        _pag(22, r["monto_efectivo"])

        # SALDO — total − pagado. Puede ser positivo (debe), 0 (cancelado),
        # negativo (exceso). Lo expone como columna del contrato para que
        # 6_corte y otras tools lean directo sin re-computar.
        saldo_val = r["saldo"]
        if abs(saldo_val) > TOL:
            _c(ws, ri, 23, saldo_val, TD_PAGO, "5B21B6",
               mono=True, align="right", bold=True, fmt=MONEY)
        else:
            _c(ws, ri, 23, 0, TD_PAGO, "5B21B6",
               mono=True, align="right", fmt=MONEY)

        # RETORNO badge (vacio si no hubo retorno)
        if r["retorno"]:
            ret_bg  = RETORNO_BG.get(r["retorno"], "FFFFFF")
            ret_txt = RETORNO_TXT.get(r["retorno"], "333333")
            c_ret = ws.cell(row=ri, column=24, value=r["retorno"])
            c_ret.font      = Font(name="Arial", size=9, bold=True, color=ret_txt)
            c_ret.fill      = PatternFill("solid", start_color=ret_bg)
            c_ret.alignment = Alignment(horizontal="center", vertical="center")
            c_ret.border    = _borde()
        else:
            _c(ws, ri, 24, None, TD_PAGO, "5B21B6", mono=True, align="center")

        # DEVUELTO badge — PAGASTE devolucion aplicado a este lote
        if r.get("devuelto"):
            dev_bg  = RETORNO_BG.get(r["devuelto"], "FFFFFF")
            dev_txt = RETORNO_TXT.get(r["devuelto"], "333333")
            c_dev = ws.cell(row=ri, column=25, value=r["devuelto"])
            c_dev.font      = Font(name="Arial", size=9, bold=True, color=dev_txt)
            c_dev.fill      = PatternFill("solid", start_color=dev_bg)
            c_dev.alignment = Alignment(horizontal="center", vertical="center")
            c_dev.border    = _borde()
        else:
            _c(ws, ri, 25, None, TD_PAGO, "5B21B6", mono=True, align="center")

        # ESTADO badge
        est_bg  = ESTADO_BG.get(r["estado"], "FFFFFF")
        est_txt = ESTADO_TXT.get(r["estado"], "333333")
        c_est = ws.cell(row=ri, column=26, value=r["estado"])
        c_est.font      = Font(name="Arial", size=9, bold=True, color=est_txt)
        c_est.fill      = PatternFill("solid", start_color=est_bg)
        c_est.alignment = Alignment(horizontal="center", vertical="center")
        c_est.border    = _borde()

        _c(ws, ri, 27, r["fecha_pago"] or None, TD_PAGO, "7C3AED",
           mono=True, align="center")

        # CICLO_COBRANZA
        _c(ws, ri, 29, r["ciclo_cobranza"], TD_TRAZ, "7D6608",
           mono=True, align="center", bold=True)

        ws.row_dimensions[ri].height = 17

    wb.save(OUTPUTS_DIR / "planilla_cobrado.xlsx")
    log.info(f"planilla_cobrado.xlsx → {len(resultado)} filas")


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 2 — trazabilidad_cobranza.xlsx (acumulada)
# ─────────────────────────────────────────────────────────────────────────────
# Layout (matching trazabilidad_cobranza.html v2.0):
#   1-3   ¿Quién es?          MZ LT NOMBRE
#   4     sep
#   5-7   ¿Qué pagó?          MONTO FUENTE RETORNO
#   8     sep
#   9-11  ¿Cómo verificarlo?  REFERENCIA FECHA COMENTARIO
#   12    sep
#   13-15 ¿De qué ciclo?      CICLO_CORRECCION_ORIGEN CICLO_COBRANZA FECHA_CARGA
#   16    sep
#   17-18 Lote corregido      MZ_ORIGEN LT_ORIGEN
# REFERENCIA: ORIGEN del banco (Yape) ó "MESA / COBRADOR" (efectivo).
# FECHA: con hora para Yape (DD/MM/YYYY HH:MM), solo fecha para efectivo.
# COMENTARIO: MENSAJE del pagador (Yape) ó nota del cobrador (efectivo).

_TZ_GRUPOS = [
    (1,  3,  "¿Quién es?",          *GH_TZ_QUIEN),
    (5,  9,  "¿Qué pagó?",          *GH_COB),
    (11, 13, "¿Cómo verificarlo?",  *GH_TZ_VERIF),
    (15, 17, "¿De qué ciclo?",      *GH_TZ_CICLO),
    (19, 20, "Lote corregido",      *GH_DC_CORR),
]
_TZ_COLS = [
    (1,  "MZ",                       *GH_TZ_QUIEN,   6),
    (2,  "LT",                       *GH_TZ_QUIEN,   6),
    (3,  "NOMBRE",                   *GH_TZ_QUIEN, 26),
    (5,  "MONTO",                    *GH_COB,      11),
    (6,  "FUENTE",                   *GH_COB,      10),
    (7,  "CONCEPTO",                 *GH_COB,      14),
    (8,  "RETORNO",                  *GH_COB,      10),
    (9,  "DEVUELTO",                 *GH_COB,      10),
    (11, "REFERENCIA",               *GH_TZ_VERIF, 24),
    (12, "FECHA",                    *GH_TZ_VERIF, 18),
    (13, "COMENTARIO",               *GH_TZ_VERIF, 30),
    (15, "CICLO_CORRECCION_ORIGEN",  *GH_TZ_CICLO, 22),
    (16, "CICLO_COBRANZA",           *GH_TZ_CICLO, 16),
    (17, "FECHA_CARGA",              *GH_TZ_CICLO, 18),
    (19, "MZ_ORIGEN",                *GH_DC_CORR,   9),
    (20, "LT_ORIGEN",                *GH_DC_CORR,   9),
]
_TZ_SEP_COLS = [4, 10, 14, 18]

FUENTE_BG  = {"yape": "E1F5EE", "efectivo": "EFF6FF"}
FUENTE_TXT = {"yape": "085041", "efectivo": "1D4ED8"}
CONCEPTO_BG  = {"tanque": "FFF7ED", "honorario": "F5F3FF", "gasto": "ECFDF5", "comunitario": "FEF9E7"}
CONCEPTO_TXT = {"tanque": "9A3412", "honorario": "5B21B6", "gasto": "065F46", "comunitario": "7D6608"}


def _exportar_trazabilidad_cobranza(
    resultado: list[dict],
    pagos_yape: list[dict],
    pagos_efectivo: list[dict],
    ciclo_nuevo: int,
    pagos_nuevos: set,
    trazabilidad_path: Path,
    retornos_por_lote: dict,
    devueltos_por_lote: dict,
):
    """Append: lee filas previas, agrega solo las nuevas con CICLO_COBRANZA=ciclo_nuevo.
    RETORNO se recalcula en cada run desde el estado actual de retornos_por_lote —
    todas las filas del mismo lote muestran el mismo badge."""
    # Mapa key → nombre (de planilla)
    nombre_de = {r["key"]: r["nombre"] for r in resultado}

    # Filas previas (preservar)
    previas = []
    if trazabilidad_path.exists():
        df = pd.read_excel(trazabilidad_path, header=1)
        df.columns = _norm_cols(df)
        for _, f in df.iterrows():
            mz = _norm_mz(f.get("MZ"))
            lt = _norm_lt(f.get("LT"))
            if not mz or not lt:
                continue
            previas.append({
                "mz":     mz,
                "lt":     lt,
                "nombre": str(f.get("NOMBRE", "")).strip(),
                "monto":  round(_float(f.get("MONTO")), 2),
                "fuente": str(f.get("FUENTE", "")).strip().lower(),
                "concepto":   _txt(f.get("CONCEPTO")) if "CONCEPTO" in df.columns else "",
                "referencia": _txt(f.get("REFERENCIA")) if "REFERENCIA" in df.columns else "",
                "fecha_hora":  _txt(f.get("FECHA")) if "FECHA" in df.columns else "",
                "comentario":  _txt(f.get("COMENTARIO")) if "COMENTARIO" in df.columns else "",
                "ciclo_correccion_origen": int(_float(f.get("CICLO_CORRECCION_ORIGEN")) or 0),
                "ciclo_cobranza":          int(_float(f.get("CICLO_COBRANZA")) or 0),
                "fecha_carga":             str(f.get("FECHA_CARGA", "")).strip(),
                "mz_origen":               _norm_mz(f.get("MZ_ORIGEN")) if "MZ_ORIGEN" in df.columns else "",
                "lt_origen":               _norm_lt(f.get("LT_ORIGEN")) if "LT_ORIGEN" in df.columns else "",
            })

    # Backfill in-place: las filas de ciclos antiguos no tenían REFERENCIA/COMENTARIO.
    # Se rellenan cruzando contra los pagos fuente por (mz, lt, monto, fuente) sin
    # tocar su ciclo ni fecha_carga. Idempotente: re-correr rellena igual.
    enriq = {}
    for p in (pagos_yape + pagos_efectivo):
        enriq[(p["mz"], p["lt"], round(p["monto"], 2), p["fuente"])] = {
            "referencia": p.get("referencia", ""),
            "comentario": p.get("comentario", ""),
            "fecha_hora": p.get("fecha_hora", "") or p.get("fecha", ""),
        }
    for t in previas:
        if t.get("referencia") and t.get("comentario"):
            continue
        src = enriq.get((t["mz"], t["lt"], round(t["monto"], 2), t["fuente"]))
        if not src:
            continue
        if not t.get("referencia"):
            t["referencia"] = src["referencia"]
        if not t.get("comentario"):
            t["comentario"] = src["comentario"]
        if not t.get("fecha_hora") or len(t["fecha_hora"]) <= 10:
            t["fecha_hora"] = src["fecha_hora"] or t.get("fecha_hora", "")

    # Filas nuevas (solo pagos imputados — huérfanos van a discrepancias, no aquí)
    ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
    nuevas = []
    for p in (pagos_yape + pagos_efectivo):
        if _identidad_pago(p) not in pagos_nuevos:
            continue
        if p["key"] not in nombre_de:
            continue  # huérfano — no está en planilla, no se imputa
        nuevas.append({
            "mz":     p["mz"],
            "lt":     p["lt"],
            "nombre": nombre_de.get(p["key"], p["nombre"]),
            "monto":  p["monto"],
            "fuente": p["fuente"],
            "concepto":   p.get("concepto", ""),
            "referencia": p.get("referencia", ""),
            "fecha_hora": p.get("fecha_hora", "") or p.get("fecha", ""),
            "comentario": p.get("comentario", ""),
            "ciclo_correccion_origen": p["ciclo_correccion"],
            "ciclo_cobranza":          ciclo_nuevo,
            "fecha_carga":             ahora,
            "mz_origen":               p.get("mz_origen", ""),
            "lt_origen":               p.get("lt_origen", ""),
        })

    todas = previas + nuevas
    last_row = max(len(todas) + 2, 3)

    wb = Workbook()
    ws = wb.active
    ws.title = "trazabilidad_cobranza"
    ws.freeze_panes = "A3"

    for cs, ce, texto, bg, txt in _TZ_GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)
    for sc in _TZ_SEP_COLS:
        _sep(ws, sc, last_row)
    for col, nombre, bg, txt, ancho in _TZ_COLS:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    MONEY = '"S/ "#,##0.00'
    for ri, t in enumerate(todas, 3):
        _c(ws, ri, 1, t["mz"],     TD_TZ_QUIEN, "5B21B6", mono=True, align="center")
        _c(ws, ri, 2, t["lt"],     TD_TZ_QUIEN, "5B21B6", mono=True, align="center")
        _c(ws, ri, 3, t["nombre"], TD_TZ_QUIEN, "333333", align="left")
        _c(ws, ri, 5, t["monto"],  TD_COB,      "065F46",
           mono=True, align="right", bold=True, fmt=MONEY)

        # FUENTE como badge
        f_bg  = FUENTE_BG.get(t["fuente"], "F3F4F6")
        f_txt = FUENTE_TXT.get(t["fuente"], "374151")
        c_f = ws.cell(row=ri, column=6, value=t["fuente"])
        c_f.font      = Font(name="Arial", size=9, bold=True, color=f_txt)
        c_f.fill      = PatternFill("solid", start_color=f_bg)
        c_f.alignment = Alignment(horizontal="center", vertical="center")
        c_f.border    = _borde()

        # CONCEPTO badge (tipo de pago: tanque, honorario, gasto, comunitario, o vacío=agua)
        conc_val = t.get("concepto", "") or ""
        if conc_val:
            c_conc = ws.cell(row=ri, column=7, value=conc_val)
            c_conc.font      = Font(name="Arial", size=9, bold=True,
                                    color=CONCEPTO_TXT.get(conc_val, "9A3412"))
            c_conc.fill      = PatternFill("solid",
                                           start_color=CONCEPTO_BG.get(conc_val, "FFF7ED"))
            c_conc.alignment = Alignment(horizontal="center", vertical="center")
            c_conc.border    = _borde()
        else:
            _c(ws, ri, 7, None, TD_COB, "065F46", mono=True, align="center")

        # RETORNO badge (puntero al archivo de retornos — vacio si no hubo)
        retorno_lote = retornos_por_lote.get(f"{t['mz']}-{t['lt']}")
        if retorno_lote:
            r_bg  = RETORNO_BG.get(retorno_lote, "FFFFFF")
            r_txt = RETORNO_TXT.get(retorno_lote, "333333")
            c_r = ws.cell(row=ri, column=8, value=retorno_lote)
            c_r.font      = Font(name="Arial", size=9, bold=True, color=r_txt)
            c_r.fill      = PatternFill("solid", start_color=r_bg)
            c_r.alignment = Alignment(horizontal="center", vertical="center")
            c_r.border    = _borde()
        else:
            _c(ws, ri, 8, None, TD_COB, "065F46", mono=True, align="center")

        # DEVUELTO badge — pagos_yape_devolucion aplicado a este lote
        devuelto_lote = devueltos_por_lote.get(f"{t['mz']}-{t['lt']}")
        if devuelto_lote:
            dv_bg  = RETORNO_BG.get(devuelto_lote, "FFFFFF")
            dv_txt = RETORNO_TXT.get(devuelto_lote, "333333")
            c_dv = ws.cell(row=ri, column=9, value=devuelto_lote)
            c_dv.font      = Font(name="Arial", size=9, bold=True, color=dv_txt)
            c_dv.fill      = PatternFill("solid", start_color=dv_bg)
            c_dv.alignment = Alignment(horizontal="center", vertical="center")
            c_dv.border    = _borde()
        else:
            _c(ws, ri, 9, None, TD_COB, "065F46", mono=True, align="center")

        # ¿Cómo verificarlo? — REFERENCIA · FECHA (con hora si Yape) · COMENTARIO
        _c(ws, ri, 11, t.get("referencia") or None, TD_TZ_VERIF, "92400E",
           mono=True, align="left")
        _c(ws, ri, 12, t.get("fecha_hora") or None, TD_TZ_VERIF, "7C2D12",
           mono=True, align="center")
        _c(ws, ri, 13, t.get("comentario") or None, TD_TZ_VERIF, "78350F",
           align="left")
        # ¿De qué ciclo? — corrección origen · cobranza · timestamp de carga
        _c(ws, ri, 15, t["ciclo_correccion_origen"], TD_TZ_CICLO, "1E40AF",
           mono=True, align="center")
        _c(ws, ri, 16, t["ciclo_cobranza"],          TD_TZ_CICLO, "1E40AF",
           mono=True, align="center", bold=True)
        _c(ws, ri, 17, t["fecha_carga"],             TD_TZ_CICLO, "1E40AF",
           mono=True, align="center")
        # Lote corregido — solo para pagos con remapeo de correcciones_lote
        mz_orig = t.get("mz_origen") or ""
        lt_orig = t.get("lt_origen") or ""
        bg_orig = TD_DC_CORR if mz_orig else TD_DC_CORR_V
        _c(ws, ri, 19, mz_orig or None, bg_orig, GH_DC_CORR[1], mono=True, align="center")
        _c(ws, ri, 20, lt_orig or None, bg_orig, GH_DC_CORR[1], mono=True, align="center")
        ws.row_dimensions[ri].height = 17

    wb.save(trazabilidad_path)
    log.info(f"trazabilidad_cobranza.xlsx → {len(todas)} filas "
             f"({len(nuevas)} nuevas en ciclo {ciclo_nuevo})")


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 3 — resumen_recaudacion.xlsx
# ─────────────────────────────────────────────────────────────────────────────
def _exportar_resumen(resultado: list[dict], n_corte: int,
                      mes_ano: str, ciclo_nuevo: int):
    tot = {
        "deuda":    round(sum(r["total_a_pagar"]  for r in resultado), 2),
        "yape":     round(sum(r["monto_yape"]     for r in resultado), 2),
        "efectivo": round(sum(r["monto_efectivo"] for r in resultado), 2),
        "pagado":   round(sum(r["total_pagado"]   for r in resultado), 2),
        "saldo":    round(sum(r["saldo"] for r in resultado if r["saldo"] > 0), 2),
        "exceso":   round(sum(-r["saldo"] for r in resultado if r["saldo"] < 0), 2),
    }
    cnt = {e: sum(1 for r in resultado if r["estado"] == e)
           for e in ("CANCELADO", "EXCESO", "PARCIAL", "PENDIENTE")}

    filas = [
        ("── RECAUDACIÓN ──────────────────",         None,           None),
        ("Total a pagar (planilla)",                  tot["deuda"],    "S/"),
        ("Recaudado Yape",                            tot["yape"],     "S/"),
        ("Recaudado Efectivo",                        tot["efectivo"], "S/"),
        ("Recaudado total",                           tot["pagado"],   "S/"),
        ("Saldo pendiente",                           tot["saldo"],    "S/"),
        ("Exceso pagado (a devolver)",                tot["exceso"],   "S/"),
        ("── ESTADOS ──────────────────────",         None,           None),
        ("CANCELADO",                                 cnt["CANCELADO"], "usuarios"),
        ("EXCESO",                                    cnt["EXCESO"],    "usuarios"),
        ("PARCIAL",                                   cnt["PARCIAL"],   "usuarios"),
        ("PENDIENTE",                                 cnt["PENDIENTE"], "usuarios"),
        ("── CORTE ────────────────────────",         None,           None),
        ("Elegibles para corte (SALDO>0 & MES_ANT>=8)", n_corte,        "usuarios"),
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "resumen"
    ws.merge_cells("A1:C1")
    t = ws.cell(row=1, column=1,
                value=f"Resumen cobranza · {mes_ano} · ciclo {ciclo_nuevo}")
    t.font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    t.fill      = PatternFill("solid", start_color="4C1D95")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 10

    for ri, (concepto, valor, unidad) in enumerate(filas, 2):
        es_sep = valor is None
        bg = "EDE9FE" if es_sep else "FAF5FF"
        _c(ws, ri, 1, concepto, bg=bg, bold=es_sep, align="left",
           txt="4C1D95" if es_sep else "333333")
        if isinstance(valor, float):
            _c(ws, ri, 2, valor, bg=bg, align="right", mono=True,
               fmt='"S/ "#,##0.00' if unidad == "S/" else "#,##0")
        else:
            _c(ws, ri, 2, valor, bg=bg, align="right", mono=not es_sep)
        _c(ws, ri, 3, unidad, bg=bg, align="left", txt="888888")
        ws.row_dimensions[ri].height = 17

    wb.save(OUTPUTS_DIR / "resumen_recaudacion.xlsx")
    log.info("resumen_recaudacion.xlsx generado")


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 4 — arrastre_deuda_YYYY-MM.xlsx
# ─────────────────────────────────────────────────────────────────────────────
_AD_GRUPOS = [
    (1, 3, "¿Quién es?",      *GH_AD_QUIEN),
    (5, 5, "¿Cuánto debe?",   *GH_AD_MONTO),
    (7, 7, "¿De qué mes?",    *GH_AD_TRAZ),
]
_AD_COLS = [
    (1, "MZ",             *GH_AD_QUIEN,   6),
    (2, "LT",             *GH_AD_QUIEN,   6),
    (3, "NOMBRE",         *GH_AD_QUIEN, 26),
    (5, "monto",          *GH_AD_MONTO, 12),
    (7, "MES_ANO_ORIGEN", *GH_AD_TRAZ,  16),
]
_AD_SEP_COLS = [4, 6]


def _exportar_arrastre_deuda(resultado: list[dict], mes_ano: str):
    pendientes = [r for r in resultado if r["saldo"] > TOL]
    last_row = max(len(pendientes) + 2, 3)

    wb = Workbook()
    ws = wb.active
    ws.title = f"arrastre_deuda_{mes_ano}"[:31]
    ws.freeze_panes = "A3"

    for cs, ce, texto, bg, txt in _AD_GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)
    for sc in _AD_SEP_COLS:
        _sep(ws, sc, last_row)
    for col, nombre, bg, txt, ancho in _AD_COLS:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    MONEY = '"S/ "#,##0.00'
    for ri, r in enumerate(pendientes, 3):
        _c(ws, ri, 1, r["mz"],         TD_AD_QUIEN, "065F46", mono=True, align="center")
        _c(ws, ri, 2, r["lt"],         TD_AD_QUIEN, "065F46", mono=True, align="center")
        _c(ws, ri, 3, r["nombre"],     TD_AD_QUIEN, "333333", align="left")
        _c(ws, ri, 5, r["saldo"],      TD_AD_MONTO, "1A5276",
           mono=True, align="right", bold=True, fmt=MONEY)
        _c(ws, ri, 7, mes_ano,         TD_AD_TRAZ,  "0E6655", mono=True, align="center")
        ws.row_dimensions[ri].height = 17

    nombre = f"arrastre_deuda_{mes_ano}.xlsx"
    wb.save(OUTPUTS_DIR / nombre)
    log.info(f"{nombre} → {len(pendientes)} usuarios con SALDO>0")


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 5 — arrastre_devolucion_YYYY-MM.xlsx
#  Paralelo a arrastre_deuda · misma estructura · paleta azul EXCESO.
#  monto = |saldo| (positivo · lo que la JASS le debe al usuario).
# ─────────────────────────────────────────────────────────────────────────────
_AV_GRUPOS = [
    (1, 3, "¿Quién es?",      *GH_AV_QUIEN),
    (5, 5, "¿Cuánto sobra?",  *GH_AV_MONTO),
    (7, 7, "¿De qué mes?",    *GH_AV_TRAZ),
    (9, 10, "¿Cómo ubicarlo?", *GH_AV_TRAZ),
    (12, 12, "¿Revisado?",    *GH_AV_REVIS),
]
_AV_COLS = [
    (1, "MZ",             *GH_AV_QUIEN,   6),
    (2, "LT",             *GH_AV_QUIEN,   6),
    (3, "NOMBRE",         *GH_AV_QUIEN, 26),
    (5, "monto",          *GH_AV_MONTO, 12),
    (7, "MES_ANO_ORIGEN", *GH_AV_TRAZ,  16),
    (9, "REFERENCIA",     *GH_AV_TRAZ,  32),
    (10, "COMENTARIO",    *GH_AV_TRAZ,  28),
    (12, "REVISION",      *GH_AV_REVIS, 32),
]
_AV_SEP_COLS = [4, 6, 8, 11]


def _backup_arrastre_devolucion(ruta: Path) -> Path | None:
    """Copia el archivo actual a outputs/backups/ antes de sobreescribirlo —
    capa 1 de preservación de trabajo manual (columna REVISION, Regla 9)."""
    if not ruta.exists():
        return None
    bk_dir = OUTPUTS_DIR / "backups"
    bk_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bk = bk_dir / f"{ruta.stem}_{ts}{ruta.suffix}"
    shutil.copy2(ruta, bk)
    return bk


def _leer_revision_previa(ruta: Path) -> dict:
    """Lee REVISION ya tipeada en arrastre_devolucion, keyed por (mz, lt),
    para no borrar el trabajo manual del operador al regenerar el archivo."""
    revisiones = {}
    if not ruta.exists():
        return revisiones
    try:
        wb = load_workbook(ruta, data_only=True)
    except Exception as e:
        log.warning(f"No se pudo leer REVISION previa: {e}")
        return revisiones
    ws = wb.active
    hdrs = {str(ws.cell(2, c).value or "").strip().upper(): c
            for c in range(1, ws.max_column + 1)}
    cmz, clt, crev = hdrs.get("MZ"), hdrs.get("LT"), hdrs.get("REVISION")
    if not all([cmz, clt, crev]):
        return revisiones
    for r in range(3, ws.max_row + 1):
        mz = _norm_mz(ws.cell(r, cmz).value)
        lt = _norm_lt(ws.cell(r, clt).value)
        rev = ws.cell(r, crev).value
        if mz and lt and rev not in (None, ""):
            revisiones[(mz, lt)] = rev
    return revisiones


def _exportar_arrastre_devolucion(resultado: list[dict], mes_ano: str):
    excesos = [r for r in resultado if r["saldo"] < -TOL]
    last_row = max(len(excesos) + 2, 3)

    ruta = OUTPUTS_DIR / f"arrastre_devolucion_{mes_ano}.xlsx"
    _backup_arrastre_devolucion(ruta)
    revisiones = _leer_revision_previa(ruta)

    wb = Workbook()
    ws = wb.active
    ws.title = f"arrastre_devolucion_{mes_ano}"[:31]
    ws.freeze_panes = "A3"

    for cs, ce, texto, bg, txt in _AV_GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)
    for sc in _AV_SEP_COLS:
        _sep(ws, sc, last_row)
    for col, nombre, bg, txt, ancho in _AV_COLS:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    MONEY = '"S/ "#,##0.00'
    for ri, r in enumerate(excesos, 3):
        monto = round(abs(r["saldo"]), 2)
        _c(ws, ri, 1, r["mz"],     TD_AV_QUIEN, "1D4ED8", mono=True, align="center")
        _c(ws, ri, 2, r["lt"],     TD_AV_QUIEN, "1D4ED8", mono=True, align="center")
        _c(ws, ri, 3, r["nombre"], TD_AV_QUIEN, "333333", align="left")
        _c(ws, ri, 5, monto,       TD_AV_MONTO, "1D4ED8",
           mono=True, align="right", bold=True, fmt=MONEY)
        _c(ws, ri, 7, mes_ano,     TD_AV_TRAZ,  "1D4ED8", mono=True, align="center")

        # REFERENCIA/COMENTARIO — solo pagos agua (sin CONCEPTO), son los que generan el exceso.
        refs, coments = [], []
        for p in r["pagos_efectivo"]:
            if p.get("concepto"):
                continue
            refs.append(f"{p.get('mesa','')} / {p.get('cobrador','')} / {p.get('fecha','')}")
            coments.append(p.get("comentario", ""))
        for p in r["pagos_yape"]:
            if p.get("concepto"):
                continue
            refs.append(f"{p.get('referencia','')} / S/{p['monto']:g} / {p.get('fecha_hora_seg','')}")
            coments.append(p.get("comentario", ""))
        _c(ws, ri, 9,  " · ".join(refs),    TD_AV_TRAZ, "1D4ED8", mono=True, align="left")
        _c(ws, ri, 10, " · ".join(coments), TD_AV_TRAZ, "555555", align="left")

        rev = revisiones.get((r["mz"], r["lt"]))
        bg_rev = TD_DC_CORR if rev else TD_AV_REVIS
        _c(ws, ri, 12, rev, bg_rev, GH_AV_REVIS[1], align="left")
        ws.row_dimensions[ri].height = 17

    wb.save(ruta)
    log.info(f"{ruta.name} → {len(excesos)} usuarios con SALDO<0 (esperan reclamo)")


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 6 — arrastre_consolidado_YYYY-MM.xlsx  (writer único hacia 2_planilla)
#  Consolida en un archivo lo que eran 3 arrastres separados (deuda + corte +
#  multa). Descompone el SALDO pendiente por componente en orden de prioridad
#  P1 DEUDA_AGUA → P2 CORTE → P3 MULTA → P4 ACUERDOS → P5 CONVENIO: cada columna
#  muestra lo que quedó SIN cubrir tras aplicar el pago en ese orden.
#  Solo filas TOTAL_ARRASTRE>0. Gate: estado_ciclo.json[mes].arrastre.validado.
#  2_planilla del próximo mes lo lee como fuente única de multa/acuerdos/convenio.
# ─────────────────────────────────────────────────────────────────────────────

def _ciclo_validado(mes_ano: str) -> bool:
    return repo_estado.ciclo_validado(mes_ano, ruta=ESTADO_CICLO_PATH)


def _marcar_generado(mes_ano: str):
    """Marca arrastre.generado=true (idempotente). 5b lo lee para sellar validado."""
    repo_estado.marcar_generado(mes_ano, ruta=ESTADO_CICLO_PATH)
    log.info(f"estado_ciclo.json → arrastre.generado=true · {mes_ano}")


def _descomponer_saldo(r: dict) -> tuple[list[float], list[float], float]:
    """Aplica total_pagado en prioridad P1→P5. Devuelve (comps[5], sin_cubrir[5], total)."""
    comps = [
        round(r["mes_actual"] + r["mantenimiento"] + r["mes_anterior"]
              + r["blanco_final"] + r["devolucion"], 2),   # P1 DEUDA_AGUA
        round(r["corte_reconexion"], 2),                    # P2 CORTE_RECONEXION
        round(r["multa"], 2),                               # P3 MULTA
        round(r["acuerdos_asamblea"], 2),                   # P4 ACUERDOS_ASAMBLEA
        round(r["convenio"], 2),                            # P5 CONVENIO
    ]
    restante = r["total_pagado"]
    sin_cubrir = []
    for comp in comps:
        comp = max(comp, 0.0)
        if restante >= comp:
            restante = round(restante - comp, 2)
            sin_cubrir.append(0.0)
        else:
            sin_cubrir.append(round(comp - restante, 2))
            restante = 0.0
    return comps, sin_cubrir, round(sum(sin_cubrir), 2)


# ─────────────────────────────────────────────────────────────────────────────
#  Reconciliación de pagos hacia seguimiento_pueblo (MULTA/ACUERDOS/CONVENIO)
#  5_cobranza recalcula el mes ENTERO en cada corrida — no procesa pago por
#  pago. seguimiento_pueblo es append-only (nunca "sobreescribe" un total).
#  Por eso cada corrida reconcilia por DELTA: SET_DEBE (recién calculado) −
#  SET_TIENE (Σ PAGO ya en el registro) = lo que falta anotar.
#  No gateada por _ciclo_validado — los pagos son hechos, no esperan el
#  sello de fin de mes (a diferencia de arrastre_consolidado).
#  Ver docs/decisiones/seguimiento_pueblo.md
# ─────────────────────────────────────────────────────────────────────────────

_CONCEPTOS_PUEBLO = ((2, "MULTA"), (3, "ACUERDOS"), (4, "CONVENIO"))


def _reconciliar_pagos_pueblo(resultado: list[dict], mes_ano: str) -> None:
    n_pago = n_ajuste = 0
    for r in resultado:
        comps, sin_cubrir, _ = _descomponer_saldo(r)
        for idx, concepto in _CONCEPTOS_PUEBLO:
            # Los predios de convenio de instalación no viven en seguimiento_pueblo
            # (su deuda ya está completa en arrastre_consolidado) — reconciliarles
            # un pago acá crearía un PAGO sin CARGO, saldo negativo falso.
            if concepto == "CONVENIO" and (r["mz"], r["lt"]) in repo.PREDIOS_INSTALACION_EXCLUIDOS:
                continue
            pagado_fresco = round(comps[idx] - sin_cubrir[idx], 2)
            ya = repo.pago_registrado(r["mz"], r["lt"], concepto, mes_ano)
            delta = round(pagado_fresco - ya, 2)
            if delta > TOL:
                ref = f"recon_{mes_ano}_{concepto}_{r['mz']}_{r['lt']}_{datetime.now().timestamp()}"
                repo.registrar_pago(r["mz"], r["lt"], concepto, mes_ano, delta,
                                     source="5_cobranza", audit_ref=ref)
                n_pago += 1
            elif delta < -TOL:
                ref = f"recon_{mes_ano}_{concepto}_{r['mz']}_{r['lt']}_{datetime.now().timestamp()}"
                repo.registrar_ajuste(r["mz"], r["lt"], concepto, mes_ano, delta,
                                      source="5_cobranza", audit_ref=ref,
                                      motivo="corrección: pago recalculado a la baja en 5_cobranza")
                n_ajuste += 1
    if n_pago or n_ajuste:
        log.info(f"seguimiento_pueblo: {n_pago} pago(s) · {n_ajuste} ajuste(s) registrados")


def _exportar_arrastre_consolidado(resultado: list[dict], mes_ano: str):
    if not _ciclo_validado(mes_ano):
        log.warning(f"arrastre_consolidado_{mes_ano} omitido — "
                    f"estado_ciclo.json sin validado:true (falta 5b_validacion)")
        return

    filas = []
    for r in resultado:
        _comps, sin_cubrir, total_arr = _descomponer_saldo(r)
        if total_arr > TOL:
            filas.append((r["mz"], r["lt"], r["nombre"], sin_cubrir, total_arr))

    last_row = max(len(filas) + 2, 3)
    wb = Workbook()
    ws = wb.active
    ws.title = f"arrastre_consolidado_{mes_ano}"[:31]
    ws.freeze_panes = "A3"

    # Group headers (fila 1)
    _gh(ws, 1, 1, 3, "¿Quién es?", *GH_AC_QUIEN)
    for i, (_, hbg, htx, _, _) in enumerate(_AC_P):
        _gh(ws, 1, 5 + i, 5 + i, f"P{i + 1}", hbg, htx)
    _gh(ws, 1, 11, 11, "Cierre", *GH_AC_TOTAL)

    # Column headers (fila 2)
    _ch(ws, 2, 1, "MZ",     *GH_AC_QUIEN)
    _ch(ws, 2, 2, "LT",     *GH_AC_QUIEN)
    _ch(ws, 2, 3, "NOMBRE", *GH_AC_QUIEN)
    for i, (nombre, hbg, htx, _, _) in enumerate(_AC_P):
        _ch(ws, 2, 5 + i, nombre, hbg, htx)
    _ch(ws, 2, 11, "TOTAL_ARRASTRE", *GH_AC_TOTAL)

    for sc in (4, 10):
        _sep(ws, sc, last_row)
    for col, ancho in ((1, 6), (2, 6), (3, 26),
                       (5, 12), (6, 16), (7, 8), (8, 17), (9, 10), (11, 15)):
        _w(ws, col, ancho)
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    MONEY = '"S/ "#,##0.00'
    for ri, (mz, lt, nombre, sin_cubrir, total_arr) in enumerate(filas, 3):
        _c(ws, ri, 1, mz,     TD_AC_QUIEN, "065F46", mono=True, align="center")
        _c(ws, ri, 2, lt,     TD_AC_QUIEN, "065F46", mono=True, align="center")
        _c(ws, ri, 3, nombre, TD_AC_QUIEN, "333333", align="left")
        for i, val in enumerate(sin_cubrir):
            _, _, _, cbg, ctx = _AC_P[i]
            if val > TOL:
                _c(ws, ri, 5 + i, val, cbg, ctx,
                   mono=True, align="right", bold=True, fmt=MONEY)
            else:
                _c(ws, ri, 5 + i, "—", TD_AC_ZERO, "9CA3AF", align="center")
        _c(ws, ri, 11, total_arr, TD_AC_TOTAL, "1E5C3A",
           mono=True, align="right", bold=True, fmt=MONEY)
        ws.row_dimensions[ri].height = 17

    nombre_arch = f"arrastre_consolidado_{mes_ano}.xlsx"
    wb.save(OUTPUTS_DIR / nombre_arch)
    log.info(f"{nombre_arch} → {len(filas)} usuarios con TOTAL_ARRASTRE>0")


# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT 7 — discrepancias_cobranza.xlsx
#  Pagos cuyo MZ+LT no existe en planilla — no pudieron imputarse a un usuario.
#  Layout (matching formato_discrepancias_cobranza.html):
#
#  Hoja discrepancias_pago_yape (cols 1..9):
#    1-2  predio    MZ LT                              (rojo — el que no existe)
#    3    sep
#    4-5  pago      MONTO FECHA                        (azul)
#    6    sep
#    7    origen    ORIGEN                             (ambar — pista de origen)
#    8    sep
#    9-10 traz      CICLO_CORRECCION MOTIVO            (morado)
#
#  Hoja discrepancias_pago_efectivo (cols 1..11):
#    1-2  predio    MZ LT
#    3    sep
#    4-5  pago      MONTO FECHA
#    6    sep
#    7-8  origen    MESA COBRADOR                      (pista fisica)
#    9    sep
#    10-11 traz     CICLO_CORRECCION MOTIVO
#
#  Si no hay discrepancias en ninguna fuente → borrar el archivo si existe
#  (su presencia es la senal de que hay trabajo pendiente, como en 4_pagos/efectivo).
# ─────────────────────────────────────────────────────────────────────────────

_DC_MOTIVO = "predio no encontrado en planilla"

_DC_YAPE_GRUPOS = [
    (1,  2,  "¿Dónde vive?",      *GH_DC_PREDIO),
    (4,  5,  "¿Cuánto y cuándo?", *GH_DC_PAGO),
    (7,  7,  "¿Quién pagó?",      *GH_DC_ORIGEN),
    (9,  10, "Trazabilidad",      *GH_DC_TRAZ),
    (12, 13, "¿Corrección?",      *GH_DC_CORR),
]
_DC_YAPE_COLS = [
    (1,  "MZ",               *GH_DC_PREDIO,  8),
    (2,  "LT",               *GH_DC_PREDIO,  8),
    (4,  "MONTO",            *GH_DC_PAGO,   12),
    (5,  "FECHA",            *GH_DC_PAGO,   12),
    (7,  "ORIGEN",           *GH_DC_ORIGEN, 22),
    (9,  "CICLO_CORRECCION", *GH_DC_TRAZ,   16),
    (10, "MOTIVO",           *GH_DC_TRAZ,   34),
    (12, "MZ_CORRECTO",      *GH_DC_CORR,    9),
    (13, "LT_CORRECTO",      *GH_DC_CORR,    9),
]
_DC_YAPE_SEP_COLS = [3, 6, 8, 11]

_DC_EFEC_GRUPOS = [
    (1,  2,  "¿Dónde vive?",      *GH_DC_PREDIO),
    (4,  5,  "¿Cuánto y cuándo?", *GH_DC_PAGO),
    (7,  8,  "¿De qué mesa?",     *GH_DC_ORIGEN),
    (10, 11, "Trazabilidad",      *GH_DC_TRAZ),
    (13, 14, "¿Corrección?",      *GH_DC_CORR),
]
_DC_EFEC_COLS = [
    (1,  "MZ",               *GH_DC_PREDIO,  8),
    (2,  "LT",               *GH_DC_PREDIO,  8),
    (4,  "MONTO",            *GH_DC_PAGO,   12),
    (5,  "FECHA",            *GH_DC_PAGO,   12),
    (7,  "MESA",             *GH_DC_ORIGEN, 12),
    (8,  "COBRADOR",         *GH_DC_ORIGEN, 22),
    (10, "CICLO_CORRECCION", *GH_DC_TRAZ,   16),
    (11, "MOTIVO",           *GH_DC_TRAZ,   34),
    (13, "MZ_CORRECTO",      *GH_DC_CORR,    9),
    (14, "LT_CORRECTO",      *GH_DC_CORR,    9),
]
_DC_EFEC_SEP_COLS = [3, 6, 9, 12]


def _leer_correcciones_tipeadas(ruta) -> dict:
    """Lee MZ_CORRECTO/LT_CORRECTO ya tipeados en discrepancias_cobranza.xlsx,
    keyed por (sheet, mz, lt), para re-mostrarlos al regenerar el archivo y no
    borrar el trabajo manual de un huérfano que persiste entre corridas."""
    tipeadas = {}
    if not ruta.exists():
        return tipeadas
    try:
        wb = load_workbook(ruta, data_only=True)
    except Exception as e:
        log.warning(f"No se pudo leer correcciones tipeadas: {e}")
        return tipeadas
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        hdrs = {str(ws.cell(2, c).value or "").strip().upper(): c
                for c in range(1, ws.max_column + 1)}
        cmzo, clto = hdrs.get("MZ"), hdrs.get("LT")
        cmzc, cltc = hdrs.get("MZ_CORRECTO"), hdrs.get("LT_CORRECTO")
        if not all([cmzo, clto, cmzc, cltc]):
            continue
        for r in range(3, ws.max_row + 1):
            mo = _norm_mz(ws.cell(r, cmzo).value)
            lo = _norm_lt(ws.cell(r, clto).value)
            mc = ws.cell(r, cmzc).value
            lc = ws.cell(r, cltc).value
            if mo and lo and (mc not in (None, "") or lc not in (None, "")):
                tipeadas[(sheet, mo, lo)] = (mc, lc)
    return tipeadas


def _exportar_discrepancias_cobranza(disc_yape: list[dict], disc_efec: list[dict]):
    """
    Genera 5_cobranza/outputs/discrepancias_cobranza.xlsx con dos hojas:
      - discrepancias_pago_yape:     pagos Yape huerfanos (ORIGEN como pista)
      - discrepancias_pago_efectivo: cobros en mesa huerfanos (MESA+COBRADOR como pista)
    Si no hay ninguna discrepancia → borra el archivo si existe.
    La presencia del archivo es la senal de que hay trabajo pendiente.
    """
    ruta = OUTPUTS_DIR / "discrepancias_cobranza.xlsx"

    if not disc_yape and not disc_efec:
        if ruta.exists():
            ruta.unlink()
            log.info("discrepancias_cobranza.xlsx eliminado — todo imputado")
        return

    # Preservar correcciones ya tipeadas por el operador (no borrarlas al regenerar)
    tipeadas = _leer_correcciones_tipeadas(ruta)

    MONEY = '"S/ "#,##0.00'
    wb = Workbook()

    # ── Hoja 1: discrepancias_pago_yape ──────────────────────────────────────
    ws = wb.active
    ws.title = "discrepancias_pago_yape"
    ws.freeze_panes = "A3"

    last_row = max(len(disc_yape) + 2, 3)
    for cs, ce, texto, bg, txt in _DC_YAPE_GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)
    for sc in _DC_YAPE_SEP_COLS:
        _sep(ws, sc, last_row)
    for col, nombre, bg, txt, ancho in _DC_YAPE_COLS:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    # Orden estable: MZ, LT, FECHA, ORIGEN — para que re-runs produzcan el mismo archivo
    for ri, p in enumerate(sorted(disc_yape,
                                   key=lambda x: (x["mz"], x["lt"], x["fecha"], x["origen"])),
                            3):
        _c(ws, ri, 1,  p["mz"],               TD_DC_PREDIO, "991B1B", mono=True, align="center", bold=True)
        _c(ws, ri, 2,  p["lt"],               TD_DC_PREDIO, "991B1B", mono=True, align="center", bold=True)
        _c(ws, ri, 4,  p["monto"],            TD_DC_PAGO,   "1A5276", mono=True, align="right",  fmt=MONEY)
        _c(ws, ri, 5,  p["fecha"],            TD_DC_PAGO,   "1A5276", mono=True, align="center")
        _c(ws, ri, 7,  p.get("origen", ""),   TD_DC_ORIGEN, "7D6608", align="left")
        _c(ws, ri, 9,  p["ciclo_correccion"], TD_DC_TRAZ,   "5B21B6", mono=True, align="center")
        _c(ws, ri, 10, _DC_MOTIVO,            TD_DC_TRAZ,   "5B21B6", align="left")
        mzc, ltc = tipeadas.get(("discrepancias_pago_yape", _norm_mz(p["mz"]), _norm_lt(p["lt"])), (None, None))
        _c(ws, ri, 12, mzc, TD_DC_CORR_V, GH_DC_CORR[1], mono=True, align="center")
        _c(ws, ri, 13, ltc, TD_DC_CORR_V, GH_DC_CORR[1], mono=True, align="center")
        ws.row_dimensions[ri].height = 17

    # ── Hoja 2: discrepancias_pago_efectivo ──────────────────────────────────
    ws = wb.create_sheet("discrepancias_pago_efectivo")
    ws.freeze_panes = "A3"

    last_row = max(len(disc_efec) + 2, 3)
    for cs, ce, texto, bg, txt in _DC_EFEC_GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)
    for sc in _DC_EFEC_SEP_COLS:
        _sep(ws, sc, last_row)
    for col, nombre, bg, txt, ancho in _DC_EFEC_COLS:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    for ri, p in enumerate(sorted(disc_efec,
                                   key=lambda x: (x["mz"], x["lt"], x["fecha"], x["mesa"])),
                            3):
        _c(ws, ri, 1,  p["mz"],               TD_DC_PREDIO, "991B1B", mono=True, align="center", bold=True)
        _c(ws, ri, 2,  p["lt"],               TD_DC_PREDIO, "991B1B", mono=True, align="center", bold=True)
        _c(ws, ri, 4,  p["monto"],            TD_DC_PAGO,   "1A5276", mono=True, align="right",  fmt=MONEY)
        _c(ws, ri, 5,  p["fecha"],            TD_DC_PAGO,   "1A5276", mono=True, align="center")
        _c(ws, ri, 7,  p.get("mesa", ""),     TD_DC_ORIGEN, "7D6608", align="center")
        _c(ws, ri, 8,  p.get("cobrador", ""), TD_DC_ORIGEN, "7D6608", align="left")
        _c(ws, ri, 10, p["ciclo_correccion"], TD_DC_TRAZ,   "5B21B6", mono=True, align="center")
        _c(ws, ri, 11, _DC_MOTIVO,            TD_DC_TRAZ,   "5B21B6", align="left")
        mzc, ltc = tipeadas.get(("discrepancias_pago_efectivo", _norm_mz(p["mz"]), _norm_lt(p["lt"])), (None, None))
        _c(ws, ri, 13, mzc, TD_DC_CORR_V, GH_DC_CORR[1], mono=True, align="center")
        _c(ws, ri, 14, ltc, TD_DC_CORR_V, GH_DC_CORR[1], mono=True, align="center")
        ws.row_dimensions[ri].height = 17

    wb.save(ruta)
    log.info(f"discrepancias_cobranza.xlsx → "
             f"{len(disc_yape)} yape · {len(disc_efec)} efectivo")


# ─────────────────────────────────────────────────────────────────────────────
#  RETROESCRITURA — CICLO_COBRANZA en pagos_yape y pagos_efectivo
# ─────────────────────────────────────────────────────────────────────────────
def _retroescribir_ciclo(path: Path, ciclo_col_nombre: str,
                         filas_nuevas: list[int], ciclo_nuevo: int):
    """Agrega o actualiza columna CICLO_COBRANZA en el archivo de pagos.
    filas_nuevas: filas (1-indexed) que recibieron ciclo_nuevo en este run.
    """
    wb = load_workbook(path)
    ws = wb.active

    # Detectar columna CICLO_COBRANZA en fila 2 (cabecera de columnas).
    cob_col = None
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        if str(ws.cell(row=2, column=col).value or "").strip().upper() == "CICLO_COBRANZA":
            cob_col = col
            break

    if cob_col is None:
        # Crear nueva columna al final
        cob_col = max_col + 1
        # Header grupo (fila 1) — usar el mismo grupo que la columna ciclo_correccion
        ref_col = None
        for col in range(1, max_col + 1):
            if str(ws.cell(row=2, column=col).value or "").strip().upper() == ciclo_col_nombre:
                ref_col = col
                break
        if ref_col:
            ref_header_row1 = ws.cell(row=1, column=ref_col).value
            c1 = ws.cell(row=1, column=cob_col, value=ref_header_row1 or "¿Cuándo?")
            c1.font      = Font(name="Arial", size=8, bold=True, color="7D6608")
            c1.fill      = PatternFill("solid", start_color="FEF9E7")
            c1.alignment = Alignment(horizontal="center", vertical="center")
        c2 = ws.cell(row=2, column=cob_col, value="CICLO_COBRANZA")
        c2.font      = Font(name="Arial", size=9, bold=True, color="7D6608")
        c2.fill      = PatternFill("solid", start_color="FEF9E7")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(cob_col)].width = 14

    filas_nuevas_set = set(filas_nuevas)
    for r in filas_nuevas_set:
        if r > ws.max_row:
            continue
        cell = ws.cell(row=r, column=cob_col)
        # Solo escribir si está vacío (para preservar ciclos previos)
        if cell.value in (None, ""):
            cell.value     = ciclo_nuevo
            cell.font      = Font(name="Consolas", size=9, color="7D6608")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill      = PatternFill("solid", start_color="FEFCE8")

    wb.save(path)
    log.info(f"{path.name} → CICLO_COBRANZA={ciclo_nuevo} en {len(filas_nuevas_set)} filas")


# ─────────────────────────────────────────────────────────────────────────────
#  ACTUALIZAR BLANCOS APLICADOS
# ─────────────────────────────────────────────────────────────────────────────
def _actualizar_blancos(blancos_aplicados: set, mes_ano: str):
    if not blancos_aplicados or not BLANCOS_PATH.exists():
        return
    wb = load_workbook(BLANCOS_PATH)
    ws = wb.active
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        mz = _norm_mz(row[_BL_MZ - 1].value)
        lt = _norm_lt(row[_BL_LOTE - 1].value)
        if f"{mz}-{lt}" in blancos_aplicados:
            row[_BL_EST - 1].value = "aplicado"
            row[_BL_MES - 1].value = mes_ano
    wb.save(BLANCOS_PATH)
    log.info(f"blancos_acumulados.xlsx → {len(blancos_aplicados)} marcados aplicado/{mes_ano}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("\n" + "═" * 60)
    print("  5_cobranza — Carga de pagos · estado de cobro")
    print("═" * 60)
    _init_logging()

    print("\n[1/6] Validando inputs...")
    plan_path = _validar_inputs()

    print("\n[2/6] Cargando datos...")
    usuarios, mes_ano = _cargar_planilla(plan_path)
    pagos_yape        = _cargar_pagos_yape()
    pagos_efectivo    = _cargar_pagos_efectivo()
    blancos           = _cargar_blancos()
    dev_yape          = _cargar_retornos_yape()
    dev_efec          = _cargar_retornos_efectivo()
    dev_devuelto      = _cargar_devueltos_yape()
    traz_path         = OUTPUTS_DIR / "trazabilidad_cobranza.xlsx"
    ids_previas, max_ciclo = _cargar_trazabilidad_previa()

    print("\n[2b/6] Aplicando correcciones de lote...")
    correcciones   = _leer_correcciones()
    correcciones   = _recuperar_correcciones_trazabilidad(correcciones, max_ciclo)
    correcciones   = _absorber_correcciones_discrepancias(correcciones, max_ciclo)
    pagos_yape     = _aplicar_correcciones_lote(pagos_yape,     correcciones)
    pagos_efectivo = _aplicar_correcciones_lote(pagos_efectivo, correcciones)

    print("\n[3/6] Detectando ciclo de cobranza...")
    # Huérfanos: pagos cuyo MZ+LT no existe en planilla. Nunca se escriben en
    # trazabilidad (ver _exportar_trazabilidad_cobranza) → se excluyen de la
    # identidad de idempotencia para que pagos_nuevos quede vacío sin cambios.
    _keys_validos = {u["key"] for u in usuarios}
    _huerfanos    = ({p["key"] for p in pagos_yape} |
                    {p["key"] for p in pagos_efectivo}) - _keys_validos

    ids_actuales = {_identidad_pago(p) for p in (pagos_yape + pagos_efectivo)
                    if p["key"] not in _huerfanos}
    pagos_nuevos = ids_actuales - ids_previas

    # Discrepancias: se exportan siempre — el archivo actúa como señal de trabajo
    # pendiente independientemente de si el ciclo es idempotente o no.
    disc_yape = [p for p in pagos_yape     if p["key"] in _huerfanos]
    disc_efec = [p for p in pagos_efectivo if p["key"] in _huerfanos]
    _exportar_discrepancias_cobranza(disc_yape, disc_efec)

    # Idempotencia: tambien comparar retornos contra el estado previo.
    # Si pagos no cambian pero retornos si → re-generar sin avanzar ciclo.
    retornos_actuales  = _retornos_por_lote(dev_yape, dev_efec)
    devueltos_actuales = _devueltos_por_lote(dev_devuelto)
    retornos_previos  = _retornos_planilla_previa()
    retornos_cambiados = retornos_actuales != retornos_previos

    force = "--force" in sys.argv
    if not pagos_nuevos and not retornos_cambiados and not force:
        log.info(f"Sin cambios (pagos ni retornos) · ciclo {max_ciclo} → idempotente")
        print(f"\n  Idempotencia: no hay pagos ni retornos nuevos")
        print(f"  Último ciclo cargado: {max_ciclo}")
        if disc_yape or disc_efec:
            print(f"  · discrepancias_cobranza.xlsx     "
                  f"({len(disc_yape)} yape · {len(disc_efec)} efectivo)")
            print(f"    → Pagos cuyo MZ+LT no existe en planilla — corregir el archivo de origen")
        print("\n" + "═" * 60 + "\n")
        return

    if pagos_nuevos:
        ciclo_nuevo = max_ciclo + 1
        log.info(f"Ciclo nuevo = {ciclo_nuevo} · pagos nuevos = {len(pagos_nuevos)}")
    else:
        ciclo_nuevo = max_ciclo if max_ciclo > 0 else 1
        log.info(f"Sin pagos nuevos · retornos cambiaron → re-generando en ciclo {ciclo_nuevo}")

    print("\n[4/6] Calculando cobranza...")
    resultado, blancos_usados = _calcular(
        usuarios, pagos_yape, pagos_efectivo, blancos,
        dev_yape, dev_efec, dev_devuelto,
        ciclo_nuevo, pagos_nuevos
    )
    # Reconciliación hacia seguimiento_pueblo — no gateada por _ciclo_validado
    # (los pagos ya ocurrieron, no esperan el sello de fin de mes).
    _reconciliar_pagos_pueblo(resultado, mes_ano)

    print("\n[5/6] Exportando outputs...")
    _exportar_planilla_cobrado(resultado)
    # Contar elegibles para corte — la lista_corte la genera 6_corte leyendo
    # SALDO + MES_ANTERIOR desde planilla_cobrado. Acá solo se reporta el conteo.
    n_corte = sum(1 for r in resultado
                  if r["saldo"] > TOL and r["mes_anterior"] >= ARRASTRE_MIN - TOL)
    _exportar_trazabilidad_cobranza(
        resultado, pagos_yape, pagos_efectivo,
        ciclo_nuevo, pagos_nuevos, traz_path,
        retornos_actuales, devueltos_actuales,
    )
    _exportar_resumen(resultado, n_corte, mes_ano, ciclo_nuevo)
    _exportar_arrastre_deuda(resultado, mes_ano)
    _exportar_arrastre_devolucion(resultado, mes_ano)
    _marcar_generado(mes_ano)                          # 5b lo lee para sellar validado
    _exportar_arrastre_consolidado(resultado, mes_ano)  # gate: validado:true

    print("\n[6/6] Retroescritura y blancos...")
    filas_yape_nuevas = [p["row"] for p in pagos_yape
                         if _identidad_pago(p) in pagos_nuevos]
    filas_efec_nuevas = [p["row"] for p in pagos_efectivo
                         if _identidad_pago(p) in pagos_nuevos]
    if filas_yape_nuevas:
        _retroescribir_ciclo(YAPE_DIR / YAPE_FILE, "CICLO",
                             filas_yape_nuevas, ciclo_nuevo)
    if filas_efec_nuevas:
        _retroescribir_ciclo(EFEC_DIR / EFEC_FILE, "CICLO_CORRECCION",
                             filas_efec_nuevas, ciclo_nuevo)
    _actualizar_blancos(blancos_usados, mes_ano)

    n_pend = sum(1 for r in resultado if r["estado"] in ("PARCIAL", "PENDIENTE"))
    print("\n" + "═" * 60)
    print(f"  Cobranza completada · ciclo {ciclo_nuevo} · {mes_ano}")
    print(f"  Outputs → 5_cobranza/outputs/")
    print(f"  · planilla_cobrado.xlsx  ({len(resultado)} usuarios · SALDO expuesto)")
    print(f"  · trazabilidad_cobranza.xlsx")
    print(f"  · resumen_recaudacion.xlsx")
    n_exceso = sum(1 for r in resultado if r["saldo"] < -TOL)
    print(f"  · arrastre_deuda_{mes_ano}.xlsx      ({sum(1 for r in resultado if r['saldo'] > TOL)} pendientes)")
    print(f"  · arrastre_devolucion_{mes_ano}.xlsx ({n_exceso} excesos)")
    if disc_yape or disc_efec:
        print(f"  · discrepancias_cobranza.xlsx     "
              f"({len(disc_yape)} yape · {len(disc_efec)} efectivo)")
        print(f"    → Pagos cuyo MZ+LT no existe en planilla — corregir el archivo de origen")
    if n_corte:
        print(f"\n  → {n_corte} usuarios elegibles para corte "
              f"(SALDO>0 & MES_ANT>=8) — correr 6_corte/generar_lista.py")
    print("═" * 60 + "\n")


if __name__ == "__main__":
    main()
