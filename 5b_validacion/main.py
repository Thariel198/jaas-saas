import logging
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
ROOT      = Path(__file__).parent
OUT_DIR   = ROOT / "outputs"
MOD04     = ROOT.parent / "5_cobranza"
SHARED    = ROOT.parent / "shared"

CRUDO_DIR  = SHARED / "reporte_mes_crudo"
YAPE_PATH  = MOD04 / "inputs" / "pagos_yape"    / "pagos_yape_tepago.xlsx"
EFEC_PATH  = MOD04 / "inputs" / "pagos_efectivo" / "pagos_efectivo.xlsx"
COB_PATH   = MOD04 / "outputs" / "cobranza_final.xlsx"

TOLERANCIA = 0.005  # diferencia máxima considerada OK

# ── PALETA ────────────────────────────────────────────────────────────────────
GH_PERIODO  = ("F3F4F6", "374151")
GH_QUIEN    = ("F4ECF7", "5B21B6")
GH_REPORTE  = ("E6F1FB", "0C447C")
GH_PLANILLA = ("FEF9E7", "7D6608")
GH_CUADRA   = ("E1F5EE", "085041")

TD_PERIODO  = "F9FAFB"
TD_QUIEN    = "FAF5FF"
TD_REPORTE  = "F0F8FF"
TD_PLANILLA = "FEFCE8"
TD_OK       = "E1F5EE"
TD_ERR      = "FEF2F2"

# ── LOGGING ───────────────────────────────────────────────────────────────────
def _init_logging():
    OUT_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(OUT_DIR / "run.log", encoding="utf-8"),
        ],
    )

log = logging.getLogger(__name__)

# ── ESTILO ────────────────────────────────────────────────────────────────────
def _borde():
    b = Side(style="thin", color="CCCCCC")
    return Border(left=b, right=b, top=b, bottom=b)

def _c(ws, row, col, value=None, bg=None, txt="333333",
       bold=False, align="left", mono=False, size=9):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="Consolas" if mono else "Arial",
                       size=size, bold=bold, color=txt)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border    = _borde()
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
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
    ws.column_dimensions[get_column_letter(col)].width = 0.8
    b_sep = Border(left=Side(style="thin", color="D1D5DB"),
                   right=Side(style="thin", color="D1D5DB"))
    for r in range(1, row_end + 1):
        c = ws.cell(row=r, column=col)
        c.fill   = PatternFill("solid", start_color="F3F4F6")
        c.border = b_sep

def _w(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

# ── UTILIDADES ────────────────────────────────────────────────────────────────
def _float(val) -> float:
    try:
        return float(str(val).replace(",", ".").strip())
    except (ValueError, TypeError):
        return 0.0

def _norm_mz(val) -> str:
    return str(val).strip().upper()

def _norm_cols(df) -> list:
    return [str(c).strip().upper() if str(c) not in ("NAN", "NONE") else f"_C{i}"
            for i, c in enumerate(df.columns)]

def _norm_tipo(val) -> str:
    """Normaliza tipo de transacción quitando tildes para comparar."""
    return (str(val).strip().upper()
            .replace("Ó", "O").replace("É", "E")
            .replace("Á", "A").replace("Í", "I").replace("Ú", "U"))

# ── VALIDAR INPUTS ────────────────────────────────────────────────────────────
def _validar_inputs():
    errores = []
    crudos = list(CRUDO_DIR.glob("*.xlsx")) if CRUDO_DIR.exists() else []
    if not crudos:
        errores.append(f"Falta reporte crudo en: {CRUDO_DIR}")
    for ruta in [YAPE_PATH, EFEC_PATH, COB_PATH]:
        if not ruta.exists():
            errores.append(f"Falta: {ruta}")
    if errores:
        for e in errores:
            log.error(e)
        raise FileNotFoundError("Inputs faltantes — ver errores arriba")
    log.info("Inputs validados")

# ── CARGA: PAGOS YAPE TEPAGO ──────────────────────────────────────────────────
def _cargar_yape_tepago() -> tuple[dict, float, tuple]:
    """Retorna ({mz: monto}, total, (fecha_min, fecha_max)) para TE PAGÓ identificados."""
    df = pd.read_excel(YAPE_PATH, header=1, dtype=str)
    df.columns = _norm_cols(df)

    totales = {}
    fechas  = []
    for _, f in df.iterrows():
        tipo = _norm_tipo(f.get("TIPO", ""))
        if tipo != "TE PAGO":
            continue
        mz   = _norm_mz(f.get("MZ", ""))
        lote = str(f.get("LOTE", "")).strip()
        conc = str(f.get("CONCEPTO", "")).strip().upper()
        if not mz or mz == "NAN" or not lote or lote == "NAN":
            continue
        if conc and conc not in ("NAN", "NONE", ""):
            continue
        totales[mz] = totales.get(mz, 0.0) + _float(f.get("MONTO_ASIGNA", 0))
        fecha_raw = str(f.get("FECHA", "")).strip()
        if fecha_raw and fecha_raw not in ("NAN", ""):
            try:
                fechas.append(pd.to_datetime(fecha_raw, dayfirst=True))
            except Exception:
                pass

    periodo  = (min(fechas), max(fechas)) if fechas else (None, None)
    total    = round(sum(totales.values()), 2)
    log.info(f"Yape tepago: {len(totales)} MZ · S/ {total:.2f} "
             f"· periodo {periodo[0]} → {periodo[1]}")
    return totales, total, periodo

# ── CARGA: REPORTE CRUDO DEL BANCO ────────────────────────────────────────────
def _cargar_crudo(periodo: tuple) -> float:
    """Suma MONTO del crudo (TE PAGÓ) filtrado al periodo del ciclo."""
    fecha_min, fecha_max = periodo
    crudo_file = sorted(CRUDO_DIR.glob("*.xlsx"))[-1]

    df = pd.read_excel(crudo_file, header=4, dtype=str)
    # Normalizar nombres de columna eliminando tildes
    df.columns = [
        str(c).strip().upper()
        .replace("Ó", "O").replace("É", "E").replace("Á", "A")
        .replace("Ó", "O").replace("Ú", "U").replace("Ñ", "N")
        .replace(" ", "_")
        for c in df.columns
    ]

    col_tipo  = next((c for c in df.columns if "TIPO" in c), None)
    col_monto = next((c for c in df.columns if "MONTO" in c), None)
    col_fecha = next((c for c in df.columns if "FECHA" in c), None)

    if not col_tipo or not col_monto:
        raise ValueError(f"Columnas TIPO/MONTO no encontradas en crudo: {df.columns.tolist()}")

    total    = 0.0
    sin_fecha = 0
    for _, f in df.iterrows():
        tipo = _norm_tipo(f.get(col_tipo, ""))
        if tipo != "TE PAGO":
            continue
        if col_fecha and fecha_min and fecha_max:
            try:
                fecha = pd.to_datetime(str(f[col_fecha]), dayfirst=True)
                if not (fecha_min <= fecha <= fecha_max):
                    continue
            except Exception:
                sin_fecha += 1
        total += _float(f.get(col_monto, 0))

    if sin_fecha:
        log.warning(f"Crudo: {sin_fecha} filas sin fecha válida — incluidas en total")
    total = round(total, 2)
    log.info(f"Crudo banco · {crudo_file.name} · TE PAGÓ en periodo · S/ {total:.2f}")
    return total

# ── CARGA: PAGOS EFECTIVO ─────────────────────────────────────────────────────
def _cargar_efectivo() -> tuple[dict, float]:
    """Retorna ({mz: monto}, total) desde pagos_efectivo.xlsx."""
    df = pd.read_excel(EFEC_PATH, header=1, dtype=str)
    df.columns = _norm_cols(df)
    totales = {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ", ""))
        if not mz or mz == "NAN":
            continue
        totales[mz] = totales.get(mz, 0.0) + _float(f.get("MONTO", 0))
    total = round(sum(totales.values()), 2)
    log.info(f"Efectivo: {len(totales)} MZ · S/ {total:.2f}")
    return totales, total

# ── CARGA: COBRANZA FINAL ─────────────────────────────────────────────────────
def _cargar_cobranza() -> tuple[dict, dict, float, float]:
    """Retorna ({mz: yape}, {mz: efectivo}, total_yape, total_efectivo)."""
    df = pd.read_excel(COB_PATH, header=1, dtype=str)
    df.columns = _norm_cols(df)
    yape_mz, efec_mz = {}, {}
    for _, f in df.iterrows():
        mz = _norm_mz(f.get("MZ", ""))
        if not mz or mz == "NAN":
            continue
        yape_mz[mz] = yape_mz.get(mz, 0.0) + _float(f.get("YAPE",     0))
        efec_mz[mz] = efec_mz.get(mz, 0.0) + _float(f.get("EFECTIVO", 0))
    t_yape = round(sum(yape_mz.values()), 2)
    t_efec = round(sum(efec_mz.values()), 2)
    log.info(f"Cobranza → YAPE: S/ {t_yape:.2f} · EFECTIVO: S/ {t_efec:.2f}")
    return yape_mz, efec_mz, t_yape, t_efec

# ── CALCULAR DIFERENCIAS ─────────────────────────────────────────────────────
def _diferencias_por_mz(reporte: dict, planilla: dict) -> list[dict]:
    """DIFERENCIA = PLANILLA − REPORTE. Incluye MZs de ambas fuentes."""
    mzs = sorted(set(reporte) | set(planilla))
    return [
        {
            "mz":        mz,
            "reporte":   round(reporte.get(mz, 0.0), 2),
            "planilla":  round(planilla.get(mz, 0.0), 2),
            "diferencia": round(planilla.get(mz, 0.0) - reporte.get(mz, 0.0), 2),
            "ok":        abs(planilla.get(mz, 0.0) - reporte.get(mz, 0.0)) < TOLERANCIA,
        }
        for mz in mzs
    ]

# ── EXPORT: HOJA RESUMEN ──────────────────────────────────────────────────────
def _hoja_resumen(wb, crudo_total, yape_proc_total, cob_yape_total,
                  efec_total, cob_efec_total, periodo):
    ws = wb.create_sheet("resumen", 0)
    ws.freeze_panes = "A2"
    fecha_min, fecha_max = periodo
    periodo_str = (f"{fecha_min.strftime('%Y-%m-%d')} / {fecha_max.strftime('%Y-%m-%d')}"
                   if fecha_min else "—")

    ws.merge_cells("A1:E1")
    t = ws.cell(row=1, column=1, value=f"Resumen validación — {periodo_str}")
    t.font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    t.fill      = PatternFill("solid", start_color="DC2626")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    cabeceras = ["MEDIO", "CRUDO BANCO", "PROCESADO", "PLANILLA", "DIFERENCIA"]
    for ci, h in enumerate(cabeceras, 1):
        _c(ws, 2, ci, h, bg="F3F4F6", txt="374151", bold=True, align="center")
    ws.row_dimensions[2].height = 18

    # Yape — tres niveles
    dif_yape = round(cob_yape_total - crudo_total, 2)
    ok_yape  = abs(dif_yape) < TOLERANCIA
    _c(ws, 3, 1, "Yape",          bg="EFF6FF", txt="1D4ED8", bold=True)
    _c(ws, 3, 2, crudo_total,      bg=TD_REPORTE,  txt="185FA5", mono=True, align="right")
    _c(ws, 3, 3, yape_proc_total,  bg=TD_REPORTE,  txt="185FA5", mono=True, align="right")
    _c(ws, 3, 4, cob_yape_total,   bg=TD_PLANILLA, txt="7D6608", mono=True, align="right")
    _c(ws, 3, 5, dif_yape,
       bg=TD_OK if ok_yape else TD_ERR,
       txt="085041" if ok_yape else "991B1B", bold=True, mono=True, align="right")
    ws.row_dimensions[3].height = 17

    # Efectivo — dos niveles (sin crudo separado)
    dif_efec = round(cob_efec_total - efec_total, 2)
    ok_efec  = abs(dif_efec) < TOLERANCIA
    _c(ws, 4, 1, "Efectivo",       bg="F0FFF8", txt="085041", bold=True)
    _c(ws, 4, 2, "— n/a",          bg="F9FAFB", txt="9CA3AF", align="center", size=8)
    _c(ws, 4, 3, efec_total,        bg=TD_REPORTE,  txt="185FA5", mono=True, align="right")
    _c(ws, 4, 4, cob_efec_total,    bg=TD_PLANILLA, txt="7D6608", mono=True, align="right")
    _c(ws, 4, 5, dif_efec,
       bg=TD_OK if ok_efec else TD_ERR,
       txt="085041" if ok_efec else "991B1B", bold=True, mono=True, align="right")
    ws.row_dimensions[4].height = 17

    for ci, w in enumerate([14, 16, 16, 16, 14], 1):
        _w(ws, ci, w)

# ── EXPORT: HOJA POR MZ ───────────────────────────────────────────────────────
# Col layout:
#  1    ¿Qué período?   PERIODO
#  2    sep
#  3    ¿Quién?         MZ
#  4    sep
#  5    ¿Qué entró?     REPORTE
#  6    sep
#  7    ¿Qué quedó?     PLANILLA
#  8    sep
#  9-10 ¿Cuadra?        DIFERENCIA  ESTADO

_GRUPOS = [
    (1, 1, "¿Qué período?",  GH_PERIODO[0],  GH_PERIODO[1]),
    (3, 3, "¿Quién?",        GH_QUIEN[0],    GH_QUIEN[1]),
    (5, 5, "¿Qué entró?",   GH_REPORTE[0],  GH_REPORTE[1]),
    (7, 7, "¿Qué quedó?",   GH_PLANILLA[0], GH_PLANILLA[1]),
    (9,10, "¿Cuadra?",       GH_CUADRA[0],   GH_CUADRA[1]),
]
_SEP_COLS = [2, 4, 6, 8]

def _hoja_por_mz(wb, sheet_name, lbl_reporte, lbl_planilla, filas, periodo):
    ws = wb.create_sheet(sheet_name)
    ws.freeze_panes = "A3"
    fecha_min, fecha_max = periodo
    periodo_str = (f"{fecha_min.strftime('%Y-%m-%d')} / {fecha_max.strftime('%Y-%m-%d')}"
                   if fecha_min else "—")
    last_row = len(filas) + 3  # 2 cabeceras + datos + fila total

    # Fila 1 — grupos
    for cs, ce, texto, bg, txt in _GRUPOS:
        _gh(ws, 1, cs, ce, texto, bg, txt)

    # Fila 2 — columnas
    col_defs = [
        (1,  "PERIODO",       GH_PERIODO[0],  GH_PERIODO[1],  24),
        (3,  "MZ",            GH_QUIEN[0],    GH_QUIEN[1],     8),
        (5,  lbl_reporte,     GH_REPORTE[0],  GH_REPORTE[1],  16),
        (7,  lbl_planilla,    GH_PLANILLA[0], GH_PLANILLA[1], 16),
        (9,  "DIFERENCIA",    GH_CUADRA[0],   GH_CUADRA[1],   13),
        (10, "ESTADO",        GH_CUADRA[0],   GH_CUADRA[1],   10),
    ]
    for col, nombre, bg, txt, ancho in col_defs:
        _ch(ws, 2, col, nombre, bg, txt)
        _w(ws, col, ancho)

    # Separadores (todos los rows incluyendo total)
    for sc in _SEP_COLS:
        _sep(ws, sc, last_row)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 20

    # Datos
    for ri, r in enumerate(filas, 3):
        ok      = r["ok"]
        dif_bg  = TD_OK  if ok else TD_ERR
        dif_txt = "085041" if ok else "991B1B"
        _c(ws, ri, 1,  periodo_str,    TD_PERIODO,  "374151", mono=True, size=9)
        _c(ws, ri, 3,  r["mz"],        TD_QUIEN,    "5B21B6", mono=True, align="center")
        _c(ws, ri, 5,  r["reporte"],   TD_REPORTE,  "185FA5", mono=True, align="right")
        _c(ws, ri, 7,  r["planilla"],  TD_PLANILLA, "7D6608", mono=True, align="right")
        _c(ws, ri, 9,  r["diferencia"],dif_bg,      dif_txt,  bold=True, mono=True, align="right")
        c_e = ws.cell(row=ri, column=10, value="OK" if ok else "ALERTA")
        c_e.font      = Font(name="Arial", size=9, bold=True, color=dif_txt)
        c_e.fill      = PatternFill("solid", start_color=dif_bg)
        c_e.alignment = Alignment(horizontal="center", vertical="center")
        c_e.border    = _borde()
        ws.row_dimensions[ri].height = 17

    # Fila de totales
    tr      = last_row
    t_rep   = round(sum(r["reporte"]   for r in filas), 2)
    t_pla   = round(sum(r["planilla"]  for r in filas), 2)
    t_dif   = round(t_pla - t_rep, 2)
    ok_tot  = abs(t_dif) < TOLERANCIA
    dif_bg  = TD_OK  if ok_tot else TD_ERR
    dif_txt = "085041" if ok_tot else "991B1B"

    for col in range(1, 11):
        _c(ws, tr, col, None, bg="F3F4F6")
    _c(ws, tr, 1, "TOTAL",  "F3F4F6", "374151", bold=True)
    _c(ws, tr, 5, t_rep,    TD_REPORTE,  "185FA5", bold=True, mono=True, align="right")
    _c(ws, tr, 7, t_pla,    TD_PLANILLA, "7D6608", bold=True, mono=True, align="right")
    _c(ws, tr, 9, t_dif,    dif_bg,      dif_txt,  bold=True, mono=True, align="right")
    c_e = ws.cell(row=tr, column=10, value="OK" if ok_tot else "ALERTA")
    c_e.font      = Font(name="Arial", size=9, bold=True, color=dif_txt)
    c_e.fill      = PatternFill("solid", start_color=dif_bg)
    c_e.alignment = Alignment(horizontal="center", vertical="center")
    c_e.border    = _borde()
    ws.row_dimensions[tr].height = 18

    for sc in _SEP_COLS:
        _sep(ws, sc, last_row)

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("\n" + "═" * 55)
    print("  5b_validacion — Validación de cobranza")
    print("═" * 55)
    _init_logging()

    print("\n[1/4] Validando inputs...")
    _validar_inputs()

    print("\n[2/4] Cargando datos...")
    yape_mz, yape_proc_total, periodo = _cargar_yape_tepago()
    crudo_total                        = _cargar_crudo(periodo)
    efec_mz, efec_total                = _cargar_efectivo()
    cob_yape_mz, cob_efec_mz, cob_yape_total, cob_efec_total = _cargar_cobranza()

    print("\n[3/4] Calculando diferencias...")
    dif_yape = _diferencias_por_mz(yape_mz, cob_yape_mz)
    dif_efec = _diferencias_por_mz(efec_mz, cob_efec_mz)

    alertas_yape = sum(1 for r in dif_yape if not r["ok"])
    alertas_efec = sum(1 for r in dif_efec if not r["ok"])
    ok_crudo     = abs(cob_yape_total - crudo_total) < TOLERANCIA
    ok_efec_tot  = abs(cob_efec_total - efec_total)  < TOLERANCIA

    print("\n[4/4] Exportando reporte...")
    wb = Workbook()
    wb.remove(wb.active)  # eliminar hoja vacía por defecto
    _hoja_resumen(wb, crudo_total, yape_proc_total, cob_yape_total,
                  efec_total, cob_efec_total, periodo)
    _hoja_por_mz(wb, "yape_por_mz",     "REPORTE_YAPE",     "PLANILLA_YAPE",     dif_yape, periodo)
    _hoja_por_mz(wb, "efectivo_por_mz", "REPORTE_EFECTIVO", "PLANILLA_EFECTIVO", dif_efec, periodo)
    wb.save(OUT_DIR / "validacion_diferencias.xlsx")

    print("\n" + "═" * 55)
    if alertas_yape == 0 and alertas_efec == 0 and ok_crudo and ok_efec_tot:
        print("  VALIDACION OK — todos los montos cuadran")
        log.info("VALIDACION OK")
    else:
        if not ok_crudo:
            dif = round(cob_yape_total - crudo_total, 2)
            print(f"  ALERTA  Yape total: crudo={crudo_total} vs planilla={cob_yape_total} (dif={dif:+.2f})")
            log.warning(f"Yape total no cuadra: crudo={crudo_total} planilla={cob_yape_total} dif={dif:+.2f}")
        if alertas_yape:
            print(f"  ALERTA  Yape por MZ: {alertas_yape} MZ(s) con diferencia")
            log.warning(f"Yape por MZ: {alertas_yape} alertas")
        if not ok_efec_tot:
            dif = round(cob_efec_total - efec_total, 2)
            print(f"  ALERTA  Efectivo total: reporte={efec_total} vs planilla={cob_efec_total} (dif={dif:+.2f})")
            log.warning(f"Efectivo total no cuadra: reporte={efec_total} planilla={cob_efec_total} dif={dif:+.2f}")
        if alertas_efec:
            print(f"  ALERTA  Efectivo por MZ: {alertas_efec} MZ(s) con diferencia")
            log.warning(f"Efectivo por MZ: {alertas_efec} alertas")
        print("  Revisar outputs/validacion_diferencias.xlsx")
    print("═" * 55 + "\n")


if __name__ == "__main__":
    main()
