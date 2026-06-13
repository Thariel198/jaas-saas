# Genera mesa_1.xlsx … mesa_7.xlsx en inputs/ con 3 hojas cada uno.
# También genera pagos_efectivo_devolucion.xlsx en outputs/ — template para retornos manuales.
# Correr una sola vez: python crear_templates.py
# Contrato visual: docs/formato_registro.html

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INPUTS_DIR = Path(__file__).parent / "inputs"
INPUTS_DIR.mkdir(exist_ok=True)

# ── Secciones (del contrato formato_registro.html) ─────────────────────────
# (texto, n_cols, bg, txt)
GRUPOS = [
    ("¿Quién cobró?",     2, "EFF6FF", "1D4ED8"),
    ("¿Dónde vive?",      2, "E1F5EE", "085041"),
    ("¿Cuánto y cuándo?", 4, "FEF9E7", "7D6608"),
    ("¿Alguna nota?",     1, "F4ECF7", "5B21B6"),
]

# ── Columnas (del contrato formato_registro.html) ───────────────────────────
# (nombre, bg_header, txt_header, ancho)
COLUMNAS = [
    ("COBRADOR",        "EFF6FF", "1D4ED8", 20),
    ("FECHA_REGISTRO",  "EFF6FF", "1D4ED8", 16),
    ("MZ",              "E1F5EE", "085041",  8),
    ("LT",              "E1F5EE", "085041",  8),
    ("MONTO",           "FEF9E7", "7D6608", 12),
    ("MONTO_EFECTIVO",  "FEF9E7", "7D6608", 16),
    ("MONTO_YAPE",      "FEF9E7", "7D6608", 14),
    ("FECHA",           "FEF9E7", "7D6608", 14),
    ("COMENTARIO",      "F4ECF7", "5B21B6", 30),
]

EJEMPLO = ["María García", "10/06/2026", "A", "8C", "38.00", "20.00", "18.00", "03/06/2026", ""]

HOJAS   = ["registro_1", "registro_2", "registro_3"]
N_MESAS = 7


# ── Helpers de estilo ───────────────────────────────────────────────────────

def _borde(color="FFFFFF"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _celda_grupo(cell, texto, bg, txt):
    cell.value     = texto
    cell.font      = Font(name="Arial", bold=True, size=9, color=txt)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _borde()

def _celda_col(cell, texto, bg, txt):
    cell.value     = texto
    cell.font      = Font(name="Arial", bold=True, size=10, color=txt)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _borde()

def _celda_ejemplo(cell, valor):
    cell.value     = valor
    cell.font      = Font(name="Arial", size=10, color="9CA3AF", italic=True)
    cell.alignment = Alignment(horizontal="left", vertical="center")


# ── Construir una hoja ──────────────────────────────────────────────────────

def _construir_hoja(ws):
    # Fila 1 — cabeceras de grupo (con merge)
    col = 1
    for texto, n_cols, bg, txt in GRUPOS:
        if n_cols > 1:
            ws.merge_cells(
                start_row=1, start_column=col,
                end_row=1,   end_column=col + n_cols - 1,
            )
        _celda_grupo(ws.cell(row=1, column=col), texto, bg, txt)
        col += n_cols
    ws.row_dimensions[1].height = 20

    # Fila 2 — nombres de columna + anchos
    for idx, (nombre, bg, txt, ancho) in enumerate(COLUMNAS, start=1):
        _celda_col(ws.cell(row=2, column=idx), nombre, bg, txt)
        ws.column_dimensions[get_column_letter(idx)].width = ancho
    ws.row_dimensions[2].height = 22

    # Fila 3 — ejemplo guía en gris itálico
    for idx, valor in enumerate(EJEMPLO, start=1):
        _celda_ejemplo(ws.cell(row=3, column=idx), valor)
    ws.row_dimensions[3].height = 18

    ws.freeze_panes = "A3"


# ── Crear un archivo de mesa ────────────────────────────────────────────────

def crear_mesa(n: int):
    nombre = f"mesa_{n}"
    wb = Workbook()

    for i, hoja in enumerate(HOJAS):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = hoja
        _construir_hoja(ws)

    ruta = INPUTS_DIR / f"{nombre}.xlsx"
    wb.save(ruta)
    print(f"  OK {nombre}.xlsx  ({len(HOJAS)} hojas)")


# ── Template pagos_efectivo_devolucion.xlsx ─────────────────────────────────
# Schema: MZ | LOTE | NOMBRE | MONTO | FECHA | CONCEPTO
# La tesorera llena una fila por cada retorno en efectivo que haga.
# Este archivo va a 5_cobranza/inputs/pagos_efectivo/pagos_efectivo_devolucion.xlsx

_DEV_GRUPOS = [
    ("¿A quién se devolvió?", 3, "E1F5EE", "085041"),
    ("¿Cuánto y cuándo?",     2, "FEF9E7", "7D6608"),
    ("¿Por qué?",             1, "F4ECF7", "5B21B6"),
]

_DEV_COLUMNAS = [
    ("MZ",       "E1F5EE", "085041",  8),
    ("LOTE",     "E1F5EE", "085041",  8),
    ("NOMBRE",   "E1F5EE", "085041", 28),
    ("MONTO",    "FEF9E7", "7D6608", 12),
    ("FECHA",    "FEF9E7", "7D6608", 14),
    ("CONCEPTO", "F4ECF7", "5B21B6", 40),
]

_DEV_EJEMPLO = ["A", "7", "JUAN PEREZ GARCIA", "40.00", "12/06/2026", "Pago fuera de tiempo — retorno acordado con usuario"]


def crear_template_devolucion():
    OUTPUTS_DIR = Path(__file__).parent / "outputs"
    OUTPUTS_DIR.mkdir(exist_ok=True)
    ruta = OUTPUTS_DIR / "pagos_efectivo_devolucion.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "pagos_efectivo_devolucion"
    ws.freeze_panes = "A3"

    col = 1
    for texto, n_cols, bg, txt in _DEV_GRUPOS:
        if n_cols > 1:
            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n_cols - 1)
        _celda_grupo(ws.cell(row=1, column=col), texto, bg, txt)
        col += n_cols
    ws.row_dimensions[1].height = 20

    for idx, (nombre, bg, txt, ancho) in enumerate(_DEV_COLUMNAS, start=1):
        _celda_col(ws.cell(row=2, column=idx), nombre, bg, txt)
        ws.column_dimensions[get_column_letter(idx)].width = ancho
    ws.row_dimensions[2].height = 22

    for idx, valor in enumerate(_DEV_EJEMPLO, start=1):
        _celda_ejemplo(ws.cell(row=3, column=idx), valor)
    ws.row_dimensions[3].height = 18

    wb.save(ruta)
    print(f"  OK pagos_efectivo_devolucion.xlsx  (outputs/)")


# ── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("\nCreando templates de mesa...")
    for n in range(1, N_MESAS + 1):
        crear_mesa(n)
    print("\nCreando template de retornos en efectivo...")
    crear_template_devolucion()
    print(f"\nListo. Abre cada mesa_N.xlsx y llena tus cobros a partir de la fila 4 (la fila 3 es el ejemplo guía, déjala).")
    print("Para retornos: llena pagos_efectivo_devolucion.xlsx desde fila 4 y cópialo a 5_cobranza/inputs/pagos_efectivo/.\n")
