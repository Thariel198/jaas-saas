"""
shared/seguimiento_repo.py — Único writer de seguimiento_pueblo.xlsx

API PÚBLICA:
    registrar_cargo(mz, lt, concepto, mes, monto, *, source, audit_ref)   → dict
    registrar_pago(mz, lt, concepto, mes, monto, *, source, audit_ref)    → dict
    registrar_ajuste(mz, lt, concepto, mes, monto, *, source, audit_ref, motivo) → dict
    get_saldo(mz, lt, concepto, mes)          → float
    get_saldos_bulk(concepto, mes)            → dict {(mz,lt): saldo} (todos los predios, 1 sola lectura)
    pago_registrado(mz, lt, concepto, mes)    → float (Σ PAGO ya anotado — para reconciliación por delta)
    estado_cuenta(mz, lt, concepto)           → DataFrame (pivot ancho: DEUDA·PAGO·SALDO por mes, 1 predio)
    generar_vista(ruta=None)                  → Path (vista_seguimiento_pueblo.xlsx, 3 hojas, TODOS los predios)

INVARIANTES:
    - Único escritor de seguimiento_pueblo.xlsx (single writer).
    - Append-only real: un evento nunca se modifica ni se borra.
    - Idempotencia por (source, audit_ref, mz, lt, concepto) — mismo evento no duplica.
    - SALDO = Σcargos − Σpagos ± Σajustes, siempre derivado, nunca celda mutable a mano.
    - Génesis = el primer evento CARGO de un predio. No existe ledger de génesis aparte.

Contrato visual: docs/formato_seguimiento_pueblo.html
Decisión de diseño: docs/decisiones/seguimiento_pueblo.md
"""

import logging
import os
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

log = logging.getLogger(__name__)

# ── Rutas ─────────────────────────────────────────────────────────────────────
_SHARED = Path(__file__).parent

SEGUIMIENTO_PATH = _SHARED / "seguimiento_pueblo.xlsx"
SHEET_NAME       = "Eventos"

CONCEPTOS_VALIDOS = {"MULTA", "ACUERDOS", "CONVENIO"}
TIPOS_EVENTO      = ("CARGO", "PAGO", "AJUSTE")

# Predios de convenio de instalación (conexión nueva) — su deuda ya está
# completa en arrastre_consolidado (junio la cargó full desde DATA_boletas),
# NUNCA entran a seguimiento_pueblo para CONVENIO. Fuente única acá: tanto
# sembrar_seguimiento_pueblo.py (no los siembra) como 5_cobranza (no les
# reconcilia pagos de CONVENIO) importan esta lista — evita que ambos lados
# se desincronicen. Ver docs/decisiones/seguimiento_pueblo.md.
PREDIOS_INSTALACION_EXCLUIDOS = {("B", "20"), ("C", "43"), ("C", "35"), ("F1", "11"), ("G", "21"), ("W", "2")}

# ── Paleta (contrato: formato_seguimiento_pueblo.html) ────────────────────────
_SEC_PREDIO = ("EBF5FB", "1A5276", "F4FAFF", "1A5276")
_SEC_EVENTO = ("FEF3C7", "92400E", "FFFBEB", "92400E")
_SEC_MONTO  = ("F3E8FF", "5B21B6", "FAF5FF", "5B21B6")
_SEC_QUIEN  = ("F3F4F6", "374151", "F9FAFB", "374151")

_COLS = [
    ("MZ",          _SEC_PREDIO, 6,  "center"),
    ("LT",          _SEC_PREDIO, 7,  "center"),
    ("CONCEPTO",    _SEC_EVENTO, 12, "center"),
    ("MES",         _SEC_EVENTO, 10, "center"),
    ("TIPO_EVENTO", _SEC_EVENTO, 12, "center"),
    ("CARGO",       _SEC_MONTO,  12, "right"),
    ("PAGO",        _SEC_MONTO,  12, "right"),
    ("AJUSTE",      _SEC_MONTO,  12, "right"),
    ("SALDO",       _SEC_MONTO,  12, "right"),
    ("SOURCE",      _SEC_QUIEN,  24, "center"),
    ("AUDIT_REF",   _SEC_QUIEN,  28, "left"),
    ("TIMESTAMP",   _SEC_QUIEN,  20, "center"),
]

_SECCIONES = [
    ("Predio",         "MZ",     "LT"),
    ("Evento",         "CONCEPTO", "TIPO_EVENTO"),
    ("Movimiento",     "CARGO",  "SALDO"),
    ("Quién / por qué", "SOURCE", "TIMESTAMP"),
]

# ── Helpers internos ─────────────────────────────────────────────────────────

def _save_atomic(wb, path: Path) -> None:
    """Guarda el workbook de forma atómica: escribe a un archivo temporal en el
    mismo directorio y lo reemplaza con os.replace() (atómico en Windows y POSIX
    — o el reemplazo pasa completo, o no pasa nada). Si el proceso se corta a
    mitad de camino, el archivo real nunca queda a medio escribir (un .xlsx
    corrupto). El temporal usa el mismo directorio para que el replace sea
    atómico incluso si /tmp está en otro filesystem.

    Reintenta el replace ante PermissionError — en Windows, un antivirus o
    indexador puede tener el archivo agarrado un instante justo después de
    escribirlo; es transitorio, no un fallo real (visto en producción:
    WinError 5 en la corrida real de la siembra, 2026-07-02)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.stem}.tmp{path.suffix}")
    wb.save(tmp)
    ultimo_error = None
    for intento in range(5):
        try:
            os.replace(tmp, path)
            return
        except PermissionError as e:
            ultimo_error = e
            time.sleep(0.3 * (intento + 1))
    raise ultimo_error


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip().upper().replace(" ", "")


def _argb(hex6: str) -> str:
    return "FF" + hex6.lstrip("#")


def _fill(hex6):
    return PatternFill("solid", fgColor=_argb(hex6))


def _hdr(cell, bg, fg, texto):
    cell.value = texto
    cell.fill  = _fill(bg)
    cell.font  = Font(color=_argb(fg), bold=True, size=9)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _dat(cell, valor, bg, fg, align="center"):
    cell.value = valor
    cell.fill  = _fill(bg)
    cell.font  = Font(color=_argb(fg), size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if isinstance(valor, (int, float)):
        cell.number_format = "#,##0.00"


def _validar_concepto(concepto: str) -> str:
    c = str(concepto).strip().upper()
    if c not in CONCEPTOS_VALIDOS:
        raise ValueError(f"concepto inválido: {concepto!r} — válidos: {sorted(CONCEPTOS_VALIDOS)}")
    return c


def _write_headers(ws) -> None:
    col_idx = {c[0]: i + 1 for i, c in enumerate(_COLS)}
    for label, start, end in _SECCIONES:
        c1, c2 = col_idx[start], col_idx[end]
        sec = next(s for n, s, _, _ in _COLS if n == start)
        if c1 != c2:
            ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        _hdr(ws.cell(row=1, column=c1), sec[0], sec[1], label)
    ws.row_dimensions[1].height = 18

    for i, (nombre, sec, ancho, _align) in enumerate(_COLS, start=1):
        _hdr(ws.cell(row=2, column=i), sec[0], sec[1], nombre)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = "A3"


_COLS_TEXTO = ("MZ", "LT", "CONCEPTO", "MES", "TIPO_EVENTO", "SOURCE", "AUDIT_REF", "TIMESTAMP")
_COLS_NUM   = ("CARGO", "PAGO", "AJUSTE", "SALDO")


def _leer_eventos() -> pd.DataFrame:
    """Lee todos los eventos como DataFrame. Vacío (sin filas) si el archivo no existe.
    MZ/LT y demás columnas de texto se leen forzadas a str — sin esto, Excel infiere LT
    numérico ("6" → int64) y rompe la normalización downstream."""
    if not SEGUIMIENTO_PATH.exists():
        return pd.DataFrame(columns=[c[0] for c in _COLS])
    dtype_map = {c: str for c in _COLS_TEXTO}
    df = pd.read_excel(SEGUIMIENTO_PATH, sheet_name=SHEET_NAME, header=1, dtype=dtype_map)
    for col in _COLS_NUM:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def _ya_registrado(source: str, audit_ref: str, mz: str, lt: str, concepto: str) -> bool:
    """Idempotencia: mismo (source, audit_ref, mz, lt, concepto) no duplica."""
    df = _leer_eventos()
    if df.empty:
        return False
    mask = (
        (df["SOURCE"].astype(str).str.strip() == source) &
        (df["AUDIT_REF"].astype(str).str.strip() == audit_ref) &
        (df["MZ"].astype(str).map(_norm) == _norm(mz)) &
        (df["LT"].astype(str).map(_norm) == _norm(lt)) &
        (df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto)
    )
    return bool(mask.any())


def _saldo_previo(mz: str, lt: str, concepto: str, mes: str) -> float:
    """Saldo acumulado ANTES de aplicar un evento nuevo — última fila del predio/concepto
    con MES <= mes dado, ordenado por TIMESTAMP. 0.0 si no hay eventos previos."""
    df = _leer_eventos()
    if df.empty:
        return 0.0
    mz_n, lt_n = _norm(mz), _norm(lt)
    sub = df[
        (df["MZ"].astype(str).map(_norm) == mz_n) &
        (df["LT"].astype(str).map(_norm) == lt_n) &
        (df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto) &
        (df["MES"].astype(str) <= str(mes))
    ]
    if sub.empty:
        return 0.0
    sub = sub.sort_values(["MES", "TIMESTAMP"])
    return float(sub.iloc[-1]["SALDO"])


def _append_evento(mz, lt, concepto, mes, tipo_evento, cargo, pago, ajuste, saldo, source, audit_ref) -> None:
    if SEGUIMIENTO_PATH.exists():
        wb = load_workbook(SEGUIMIENTO_PATH)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        next_row = max(ws.max_row + 1, 3)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        _write_headers(ws)
        next_row = 3

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fila = {
        "MZ": _norm(mz), "LT": _norm(lt), "CONCEPTO": concepto, "MES": str(mes),
        "TIPO_EVENTO": tipo_evento,
        "CARGO": cargo, "PAGO": pago, "AJUSTE": ajuste, "SALDO": saldo,
        "SOURCE": source, "AUDIT_REF": audit_ref, "TIMESTAMP": ts,
    }
    for ci, (nombre, sec, _ancho, align) in enumerate(_COLS, start=1):
        _dat(ws.cell(row=next_row, column=ci), fila[nombre], sec[2], sec[3], align=align)

    _save_atomic(wb, SEGUIMIENTO_PATH)

# ── API pública: escritura ───────────────────────────────────────────────────

def _registrar(mz, lt, concepto, mes, monto, tipo_evento, *, source, audit_ref, motivo=None) -> dict:
    if not source:
        raise ValueError("source no puede ser vacío")
    if not audit_ref:
        raise ValueError("audit_ref no puede ser vacío")
    concepto = _validar_concepto(concepto)
    mz_n, lt_n = _norm(mz), _norm(lt)
    if not mz_n or not lt_n:
        raise ValueError(f"MZ/LT inválidos ({mz!r}, {lt!r})")
    if tipo_evento == "AJUSTE" and not motivo:
        raise ValueError("registrar_ajuste: motivo no puede ser vacío")

    if _ya_registrado(source, audit_ref, mz_n, lt_n, concepto):
        log.info(f"seguimiento_repo: skip (idempotente) — {source}/{audit_ref} {mz_n}-{lt_n} {concepto}")
        return {"saldo_resultante": None, "skipped": True}

    monto = float(monto)
    saldo_previo = _saldo_previo(mz_n, lt_n, concepto, mes)

    cargo = pago = ajuste = None
    if tipo_evento == "CARGO":
        cargo = monto
        saldo_nuevo = saldo_previo + monto
    elif tipo_evento == "PAGO":
        pago = monto
        saldo_nuevo = saldo_previo - monto
    else:  # AJUSTE
        ajuste = monto
        saldo_nuevo = saldo_previo + monto

    _append_evento(mz_n, lt_n, concepto, mes, tipo_evento, cargo, pago, ajuste, saldo_nuevo, source, audit_ref)
    log.info(f"seguimiento_repo: {tipo_evento} {mz_n}-{lt_n} {concepto} {mes} monto={monto} saldo={saldo_nuevo}")
    return {"saldo_resultante": saldo_nuevo, "skipped": False}


def registrar_cargo(mz, lt, concepto, mes, monto, *, source: str, audit_ref: str) -> dict:
    """Deuda nueva — siembra inicial o cargo del mes (faena/reunión/asamblea/convenio)."""
    return _registrar(mz, lt, concepto, mes, monto, "CARGO", source=source, audit_ref=audit_ref)


def registrar_pago(mz, lt, concepto, mes, monto, *, source: str, audit_ref: str) -> dict:
    """5_cobranza registra la porción de un pago que va a este concepto."""
    return _registrar(mz, lt, concepto, mes, monto, "PAGO", source=source, audit_ref=audit_ref)


def registrar_ajuste(mz, lt, concepto, mes, monto, *, source: str, audit_ref: str, motivo: str) -> dict:
    """Corrección manual — monto puede ser negativo. Nunca edita un evento pasado."""
    return _registrar(mz, lt, concepto, mes, monto, "AJUSTE", source=source, audit_ref=audit_ref, motivo=motivo)

# ── API pública: lectura ─────────────────────────────────────────────────────

def pago_registrado(mz, lt, concepto, mes) -> float:
    """Σ PAGO ya registrado para (mz, lt, concepto, mes) — no incluye AJUSTE.
    Es el "SET_TIENE" de la reconciliación por delta que hace 5_cobranza en
    cada corrida (ver docs/decisiones/seguimiento_pueblo.md)."""
    concepto = _validar_concepto(concepto)
    df = _leer_eventos()
    if df.empty:
        return 0.0
    mz_n, lt_n = _norm(mz), _norm(lt)
    sub = df[
        (df["MZ"].astype(str).map(_norm) == mz_n) &
        (df["LT"].astype(str).map(_norm) == lt_n) &
        (df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto) &
        (df["MES"].astype(str) == str(mes)) &
        (df["TIPO_EVENTO"] == "PAGO")
    ]
    if sub.empty:
        return 0.0
    return float(sub["PAGO"].fillna(0).sum())


def get_saldo(mz, lt, concepto, mes) -> float:
    """Saldo de (mz, lt, concepto) al cierre de `mes` (inclusive). 0.0 si no hay eventos."""
    concepto = _validar_concepto(concepto)
    df = _leer_eventos()
    if df.empty:
        return 0.0
    mz_n, lt_n = _norm(mz), _norm(lt)
    sub = df[
        (df["MZ"].astype(str).map(_norm) == mz_n) &
        (df["LT"].astype(str).map(_norm) == lt_n) &
        (df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto) &
        (df["MES"].astype(str) <= str(mes))
    ]
    if sub.empty:
        return 0.0
    sub = sub.sort_values(["MES", "TIMESTAMP"])
    return float(sub.iloc[-1]["SALDO"])


def get_saldos_bulk(concepto, mes) -> dict:
    """{(mz, lt): saldo} al cierre de `mes`, para TODOS los predios de una vez.

    2_planilla necesita el saldo de ~575 predios por concepto — llamar
    get_saldo() por predio sería 575 lecturas completas del registro (mismo
    patrón O(n) por llamada que hizo lenta la siembra). Esta función lee el
    registro una sola vez y calcula todos los saldos de un pase."""
    concepto = _validar_concepto(concepto)
    df = _leer_eventos()
    if df.empty:
        return {}
    sub = df[
        (df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto) &
        (df["MES"].astype(str) <= str(mes))
    ]
    if sub.empty:
        return {}
    sub = sub.sort_values(["MZ", "LT", "MES", "TIMESTAMP"])
    ultimo = sub.groupby(["MZ", "LT"], as_index=False).last()
    return {(_norm(r["MZ"]), _norm(r["LT"])): float(r["SALDO"]) for _, r in ultimo.iterrows()}


def estado_cuenta(mz, lt, concepto) -> pd.DataFrame:
    """Pivot a la vista ancha — una fila (mz, lt), columnas DEUDA/PAGO/SALDO por mes.
    Regenerable siempre desde el registro largo — nunca es la fuente."""
    concepto = _validar_concepto(concepto)
    df = _leer_eventos()
    mz_n, lt_n = _norm(mz), _norm(lt)
    sub = df[
        (df["MZ"].astype(str).map(_norm) == mz_n) &
        (df["LT"].astype(str).map(_norm) == lt_n) &
        (df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto)
    ].sort_values(["MES", "TIMESTAMP"])

    meses = sorted(sub["MES"].astype(str).unique())
    filas = []
    for mes in meses:
        del_mes = sub[sub["MES"].astype(str) == mes]
        deuda = float(del_mes["CARGO"].fillna(0).sum() + del_mes["AJUSTE"].fillna(0).sum())
        pago  = float(del_mes["PAGO"].fillna(0).sum())
        saldo = float(del_mes.iloc[-1]["SALDO"])
        filas.append({"MES": mes, "DEUDA": deuda, "PAGO": pago, "SALDO": saldo})
    return pd.DataFrame(filas, columns=["MES", "DEUDA", "PAGO", "SALDO"])


def deudores(concepto, minimo: float = 0.0) -> pd.DataFrame:
    """Lotes con saldo vigente (última fila por mz,lt,concepto) >= minimo."""
    concepto = _validar_concepto(concepto)
    df = _leer_eventos()
    if df.empty:
        return pd.DataFrame(columns=["MZ", "LT", "SALDO"])
    sub = df[df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto].copy()
    if sub.empty:
        return pd.DataFrame(columns=["MZ", "LT", "SALDO"])
    sub = sub.sort_values(["MZ", "LT", "MES", "TIMESTAMP"])
    ultimo = sub.groupby(["MZ", "LT"], as_index=False).last()[["MZ", "LT", "SALDO"]]
    return ultimo[ultimo["SALDO"] >= minimo].reset_index(drop=True)

# ── Vista ancha regenerable ────────────────────────────────────────────────────
# Contrato: docs/formato_vista_seguimiento_pueblo.html

VISTA_PATH = _SHARED / "vista_seguimiento_pueblo.xlsx"

_SEC_PREDIO_V = ("EBF5FB", "1A5276", "F4FAFF", "1A5276")
# (header_bg, header_fg, deuda_fg, pago_fg, saldo_bg, saldo_fg) — cicla cada 3 meses
_PALETAS_MES = [
    ("FEF3C7", "92400E", "92400E", "065F46", "FDE68A", "78350F"),
    ("DBEAFE", "1E3A8A", "1E3A8A", "065F46", "BFDBFE", "1E3A8A"),
    ("EDE9FE", "5B21B6", "5B21B6", "065F46", "DDD6FE", "4C1D95"),
]
_DATA_BOLETAS_PATH_VISTA = _SHARED.parent / "3_boletas" / "inputs" / "DATA_boletas.xlsx"


def _lookup_nombres() -> dict:
    """(mz, lt) → NOMBRES, leído en vivo de DATA_boletas. {} si no existe el archivo."""
    if not _DATA_BOLETAS_PATH_VISTA.exists():
        return {}
    try:
        df = pd.read_excel(_DATA_BOLETAS_PATH_VISTA, sheet_name="Data", dtype=str)
    except Exception:
        return {}
    out = {}
    for _, r in df.iterrows():
        mz, lt = r.get("MZ"), r.get("LT")
        if mz and lt:
            out[(_norm(mz), _norm(lt))] = str(r.get("NOMBRES", "")).strip()
    return out


def _escribir_hoja_vista(ws, concepto: str, df_concepto: pd.DataFrame, nombres: dict) -> None:
    predios = sorted({(r["MZ"], r["LT"]) for _, r in df_concepto.iterrows()})
    meses = sorted(df_concepto["MES"].astype(str).unique())

    col_predio = [("MZ", 6, "center"), ("LT", 7, "center"), ("NOMBRE", 26, "left")]
    cols = list(col_predio)
    for mes in meses:
        cols += [(f"{mes}|DEUDA", 11, "right"), (f"{mes}|PAGO", 11, "right"), (f"{mes}|SALDO", 11, "right")]

    # ── Fila 1: secciones (Predio + 1 por mes) ──
    _hdr(ws.cell(row=1, column=1), *_SEC_PREDIO_V[:2], "Predio")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    for i, mes in enumerate(meses):
        pal = _PALETAS_MES[i % 3]
        c1 = 4 + i * 3
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c1 + 2)
        _hdr(ws.cell(row=1, column=c1), pal[0], pal[1], mes)
    ws.row_dimensions[1].height = 18

    # ── Fila 2: columnas ──
    for i, (nombre, ancho, _align) in enumerate(cols, start=1):
        etiqueta = nombre.split("|")[-1]
        if i <= 3:
            _hdr(ws.cell(row=2, column=i), *_SEC_PREDIO_V[:2], etiqueta)
        else:
            mes_idx = (i - 4) // 3
            pal = _PALETAS_MES[mes_idx % 3]
            _hdr(ws.cell(row=2, column=i), pal[0], pal[1], etiqueta)
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = "D3"

    if not predios:
        return

    # ── Filas de datos ──
    for r_off, (mz, lt) in enumerate(predios):
        row = r_off + 3
        _dat(ws.cell(row=row, column=1), mz, *_SEC_PREDIO_V[2:], align="center")
        _dat(ws.cell(row=row, column=2), lt, *_SEC_PREDIO_V[2:], align="center")
        _dat(ws.cell(row=row, column=3), nombres.get((mz, lt), ""), *_SEC_PREDIO_V[2:], align="left")

        sub_predio = df_concepto[(df_concepto["MZ"] == mz) & (df_concepto["LT"] == lt)]
        for i, mes in enumerate(meses):
            pal = _PALETAS_MES[i % 3]
            c_deuda, c_pago, c_saldo = 4 + i * 3, 5 + i * 3, 6 + i * 3
            del_mes = sub_predio[sub_predio["MES"].astype(str) == mes]

            if del_mes.empty:
                for c in (c_deuda, c_pago, c_saldo):
                    cell = ws.cell(row=row, column=c, value="—")
                    cell.fill = _fill("F3F4F6")
                    cell.font = Font(color=_argb("9CA3AF"), italic=True, size=10)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                continue

            deuda = float(del_mes["CARGO"].fillna(0).sum() + del_mes["AJUSTE"].fillna(0).sum())
            pago = float(del_mes["PAGO"].fillna(0).sum())
            saldo = float(del_mes.sort_values("TIMESTAMP").iloc[-1]["SALDO"])

            _dat(ws.cell(row=row, column=c_deuda), deuda, "FFFBEB", pal[2], align="right")
            _dat(ws.cell(row=row, column=c_pago), pago, "FFFBEB", pal[3], align="right")
            cell_s = ws.cell(row=row, column=c_saldo)
            cell_s.value = saldo
            cell_s.fill = _fill(pal[4])
            cell_s.font = Font(color=_argb(pal[5]), bold=True, size=10)
            cell_s.alignment = Alignment(horizontal="right", vertical="center")
            cell_s.number_format = "#,##0.00"


def generar_vista(ruta: Path | None = None) -> Path:
    """Regenera vista_seguimiento_pueblo.xlsx completa (3 hojas) desde el registro.
    Nunca es la fuente — se puede borrar y recrear en cualquier momento."""
    ruta = ruta or VISTA_PATH
    df = _leer_eventos()
    nombres = _lookup_nombres()

    wb = Workbook()
    wb.remove(wb.active)
    for concepto in ("MULTA", "ACUERDOS", "CONVENIO"):
        ws = wb.create_sheet(concepto)
        df_c = df[df["CONCEPTO"].astype(str).str.strip().str.upper() == concepto]
        _escribir_hoja_vista(ws, concepto, df_c, nombres)

    _save_atomic(wb, ruta)
    log.info(f"vista_seguimiento_pueblo regenerada -> {ruta}")
    return ruta
