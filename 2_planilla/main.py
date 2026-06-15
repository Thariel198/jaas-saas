import glob
import logging
import shutil
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config

log = logging.getLogger(__name__)


# ── Normalización de clave (MZ, LT) ──────────────────────────────────────

def _norm_mz(v) -> str:
    return str(v).strip().upper()


def _norm_lt(v) -> str:
    s = str(v).strip()
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except ValueError:
        pass
    return s.upper()


def _add_key(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_mz"] = df["MZ"].map(_norm_mz)
    df["_lt"] = df["LT"].map(_norm_lt)
    return df


# ── Carga de inputs ───────────────────────────────────────────────────────

def _load_lecturas(mes: str) -> pd.DataFrame:
    path = config.lecturas_path(mes)
    df = pd.read_excel(path, dtype=str)
    missing = [c for c in config.COLS_LECTURAS if c not in df.columns]
    if missing:
        raise ValueError(f"lecturas_planilla: columnas faltantes: {missing}")
    return _add_key(df)


def _load_optional(path: Path, required_cols: list, label: str) -> pd.DataFrame | None:
    if not path.exists():
        log.warning(f"{label}: archivo no encontrado -> valores = 0")
        return None
    df = pd.read_excel(path, dtype=str)
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        log.warning(f"{label}: columnas faltantes {missing} -> valores = 0")
        return None
    return _add_key(df)


# ── Join de arrastres ─────────────────────────────────────────────────────

def _join_optional(base: pd.DataFrame, src: pd.DataFrame | None,
                   src_col: str, dest_col: str) -> pd.DataFrame:
    if src is None:
        base = base.copy()
        base[dest_col] = 0.0
        return base

    # Advertir sobre filas en src sin match en base
    check = src.merge(base[["_mz", "_lt"]], on=["_mz", "_lt"], how="left", indicator=True)
    sin_match = (check["_merge"] == "left_only").sum()
    if sin_match:
        log.warning(f"{dest_col}: {sin_match} fila(s) en arrastre sin match en lecturas -> ignoradas")

    merged = base.merge(
        src[["_mz", "_lt", src_col]].rename(columns={src_col: dest_col}),
        on=["_mz", "_lt"],
        how="left",
    )
    merged[dest_col] = pd.to_numeric(merged[dest_col], errors="coerce").fillna(0)
    return merged


# ── Build del dataframe de planilla ──────────────────────────────────────

def build_planilla(mes: str) -> pd.DataFrame:
    df = _load_lecturas(mes)
    log.info(f"Lecturas cargadas: {len(df)} usuarios · mes {mes}")

    for col in ["MARC_ANT", "MARC_ACT", "M3"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df_deuda     = _load_optional(config.deuda_path(mes), config.COLS_DEUDA,     "arrastre_deuda")
    df_corte     = _load_optional(config.corte_path(mes), config.COLS_CORTE,     "arrastre_corte")
    df_convenios = _load_optional(config.CONVENIOS_PATH,  config.COLS_CONVENIOS, "convenios")
    df_multas    = _load_optional(config.MULTAS_PATH,     config.COLS_MULTAS,    "multas")
    df_acuerdos  = _load_optional(config.ACUERDOS_PATH,   config.COLS_ACUERDOS,  "acuerdos_asamblea")

    df = _join_optional(df, df_deuda,     "monto",     "MES_ANTERIOR")
    df = _join_optional(df, df_corte,     "monto",     "CORTE_RECONEXION")
    df = _join_optional(df, df_convenios, "cuota_mes", "CONVENIO")
    df = _join_optional(df, df_multas,    "monto_mes", "MULTA")
    df = _join_optional(df, df_acuerdos,  "monto_mes", "ACUERDOS_ASAMBLEA")

    df["MES_ACTUAL"]    = df["M3"].apply(
        lambda m: max(float(m) * config.TARIFA_M3, config.TARIFA_MIN)
    )
    df["MANTENIMIENTO"] = float(config.MANT_FIJO)
    df["BLANCO"]        = 0.0
    df["DEVOLUCION"]    = 0.0

    # TOTAL_A_PAGAR y columnas de pago son fórmula/vacío — se escriben en Excel
    df["TOTAL_A_PAGAR"]  = None
    df["MONTO_YAPE"]     = None
    df["MONTO_EFECTIVO"] = None
    df["ESTADO"]         = None
    df["FECHA_PAGO"]     = None

    return df


# ── Escritura Excel ───────────────────────────────────────────────────────

_SECTION_LABELS = [
    ("¿Quién es?",      "MZ",            "MES_ANO"),
    ("Lectura",         "MARC_ANT",      "M3"),
    ("Cobro — cargos",  "MES_ACTUAL",    "ACUERDOS_ASAMBLEA"),
    ("Descuentos",      "BLANCO",        "DEVOLUCION"),
    ("Total",           "TOTAL_A_PAGAR", "TOTAL_A_PAGAR"),
    ("Pago → 4_pagos", "MONTO_YAPE",    "FECHA_PAGO"),
]

# Columnas con valores numéricos → alineación derecha en datos
_RIGHT_COLS = {
    "MARC_ANT", "MARC_ACT", "M3",
    "MES_ACTUAL", "MANTENIMIENTO", "MES_ANTERIOR", "CORTE_RECONEXION",
    "CONVENIO", "MULTA", "ACUERDOS_ASAMBLEA",
    "BLANCO", "DEVOLUCION", "TOTAL_A_PAGAR",
    "MONTO_YAPE", "MONTO_EFECTIVO",
}


def _argb(hex6: str) -> str:
    return "FF" + hex6.lstrip("#")


def _fill(hex6: str) -> PatternFill:
    return PatternFill("solid", fgColor=_argb(hex6))


def _cell_align(col: str) -> str:
    sec = config.COL_SECTION[col]
    if sec is config.SEC_PAGO:
        return "center"
    if col in _RIGHT_COLS:
        return "right"
    if col == "NOMBRE":
        return "left"
    return "center"


def write_excel(df: pd.DataFrame, mes: str) -> None:
    cols = config.OUTPUT_COLS
    col_idx = {col: i + 1 for i, col in enumerate(cols)}  # 1-indexed

    # Columnas sumandas: MES_ACTUAL … DEVOLUCION (OUTPUT_COLS índices 7–15)
    summand_cols = cols[7:16]

    wb = Workbook()
    ws = wb.active
    ws.title = config.OUTPUT_SHEET

    # ── Fila 1: secciones ──────────────────────────────────────────────
    for label, start, end in _SECTION_LABELS:
        c1 = col_idx[start]
        c2 = col_idx[end]
        sec = config.COL_SECTION[start]
        if c1 != c2:
            ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        cell = ws.cell(row=1, column=c1, value=label)
        cell.fill = _fill(sec["header_bg"])
        cell.font = Font(color=_argb(sec["header_fg"]), bold=True, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = config.ROW_HEIGHT_SECTIONS

    # ── Fila 2: nombres de columna ─────────────────────────────────────
    for i, col in enumerate(cols, start=1):
        sec = config.COL_SECTION[col]
        cell = ws.cell(row=2, column=i, value=col)
        cell.fill = _fill(sec["header_bg"])
        cell.font = Font(color=_argb(sec["header_fg"]), bold=True, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[2].height = config.ROW_HEIGHT_COLS

    # ── Filas de datos (desde fila 3) ─────────────────────────────────
    for r_offset, (_, row) in enumerate(df.iterrows()):
        xl_row = r_offset + 3

        for col in cols:
            c = col_idx[col]
            sec = config.COL_SECTION[col]
            is_total = col == "TOTAL_A_PAGAR"
            is_pago  = sec is config.SEC_PAGO

            if is_total:
                val = "=" + "+".join(
                    f"{get_column_letter(col_idx[sc])}{xl_row}"
                    for sc in summand_cols
                )
            else:
                val = row.get(col)
                try:
                    if pd.isna(val):
                        val = None
                except TypeError:
                    pass

            cell = ws.cell(row=xl_row, column=c, value=val)
            cell.fill = _fill(sec["data_bg"])
            cell.number_format = config.COL_FORMAT[col]
            cell.alignment = Alignment(horizontal=_cell_align(col), vertical="center")
            cell.font = Font(
                color=_argb(sec["data_fg"]),
                bold=is_total,
                italic=is_pago,
                size=10,
            )

    # ── Anchos de columna ──────────────────────────────────────────────
    for i, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = config.COL_WIDTH[col]

    # ── Congelar paneles ───────────────────────────────────────────────
    ws.freeze_panes = config.FREEZE_PANES

    # ── Guardar ────────────────────────────────────────────────────────
    out = config.output_path(mes)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    log.info(f"Planilla guardada: {out}")


# ── Punto de entrada ──────────────────────────────────────────────────────

def main() -> None:
    config.OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(config.OUTPUTS_DIR / "run.log", encoding="utf-8"),
        ],
        force=True,
    )

    pattern = str(config.INPUTS_DIR / "lecturas" / "lecturas_planilla_*.xlsx")
    matches = sorted(glob.glob(pattern))
    if not matches:
        log.error("No se encontró ningún archivo lecturas_planilla_YYYY-MM.xlsx en inputs/lecturas/")
        sys.exit(1)
    if len(matches) > 1:
        log.warning(f"Múltiples archivos de lecturas — usando el más reciente: {matches[-1]}")

    mes = Path(matches[-1]).stem.replace("lecturas_planilla_", "")
    log.info(f"Mes detectado: {mes}")

    df = build_planilla(mes)
    write_excel(df, mes)
    publicar_a_shared(mes)


def publicar_a_shared(mes: str) -> None:
    """
    Copia planilla_{mes}.xlsx a shared/planilla_mes/ — donde 4_pagos espera leerla.
    Sobreescribe el archivo del mismo mes; no toca archivos de otros meses.
    """
    src = config.output_path(mes)
    dest = config.shared_planilla_path(mes)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    log = logging.getLogger(__name__)
    log.info(f"Publicada a shared: {dest}")


if __name__ == "__main__":
    main()
