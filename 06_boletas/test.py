# =========================IMPORTS===========================
import re
import unicodedata
import pandas as pd
import shutil
import fitz
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx2pdf import convert
from docx.shared import Mm
from PyPDF2 import PdfMerger

# ========================CONFIGURACION======================
BASE_DIR = Path(".")
INPUT_DIR = BASE_DIR / "Inputs"
OUTPUT_DIR = BASE_DIR / "Outputs"
IMAGES_DIR = OUTPUT_DIR / "Imagenes"

DATA_BOLETAS_PATH = INPUT_DIR / "DATA_boletas.xlsx"
PLANTILLA_BOLETAS_PATH = INPUT_DIR / "PLANTILLA_boletas.docx"
IMG_LOGO_JAAS_PATH = INPUT_DIR / "logo_jaas.png"
IMG_CARITA_TRISTE_PATH = INPUT_DIR / "carita_triste.png"
IMG_CARITA_FELIZ_PATH = INPUT_DIR / "carita_feliz.png"
IMG_QR_PATH = INPUT_DIR / "imagen_qr.png"

NOMBRE_JAAS = "JUNTA ADMINISTRATIVA DE SERVICIOS DE SANEAMIENTO"
SECTOR = "P.J. TUPAC AMARU"
LECTURA_ANTERIOR = "10/12/2O25"
LECTURA_ACTUAL = "10/01/2026"
PERIODO = "11/02/2026 al 10/03/2026"
FECHA_VENCIMIENTO = "07/02/2026"
FECHA_EMISION = "27/03/2026"
FECHA_PAGO = "05-06/06"
HORA_PAGO = "16:00-20:00 y 10:00-13:00 hrs"
LUGAR_PAGO = "LOCAL DEL PUEBLO"
TELEFONO = "948 227 636"

SHEET_DATA_BOLETAS = "Data"

# =====================PATRONES DE TEST======================
# Cada patron busca el primer recibo que cumpla la condicion.
# Si la data cambia cada mes, los patrones siguen siendo validos
# mientras existan registros con esas caracteristicas.
PATRONES = {
    "normal_al_dia":    lambda df: df[
        ~df["Estado"].str.upper().str.contains("NO EST", na=False) &
        (df["M3"] > 0) &
        (df["Multa (faena + reunión)"] == 0) &
        (df["Convenio"] == 0)
    ],
    "sin_consumo":      lambda df: df[df["M3"] == 0],
    "no_al_dia":        lambda df: df[df["Estado"].str.upper().str.contains("NO EST", na=False)],
    "con_multa":        lambda df: df[df["Multa (faena + reunión)"] > 0],
    "con_convenio":     lambda df: df[df["Convenio"] > 0],
    "corte_reconexion": lambda df: df[df["Corte y reconexion"] > 0],
    "mes_anterior":     lambda df: df[df["MES ANTERIOR"] > 0],
}

# ===================SELECCION AUTOMATICA====================
def select_test_receipts(df: pd.DataFrame) -> list:
    seleccionados = {}

    for nombre, filtro in PATRONES.items():
        resultado = filtro(df)
        if resultado.empty:
            print(f"  [AVISO] Patron '{nombre}': sin coincidencias en este mes, se omite.")
            continue
        recibo = resultado.iloc[0]["NUMERO DE RECIBO"]
        seleccionados[nombre] = recibo

    # Deduplicar: un mismo recibo puede cubrir varios patrones
    recibos_unicos = list(dict.fromkeys(seleccionados.values()))

    print(f"\n=== RECIBOS SELECCIONADOS PARA TEST ({len(recibos_unicos)} unicos) ===")
    for nombre, recibo in seleccionados.items():
        print(f"  [{nombre}] -> Recibo {recibo}")
    print()

    return recibos_unicos

# =====================UTILIDADES============================
def sanitize(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^\w]", "_", s).strip("_")

def reset_output_folder(path: Path):
    if path.exists():
        shutil.rmtree(path)
    path.mkdir()

# ====================CARGA DE DATOS=========================
def load_data():
    df = pd.read_excel(DATA_BOLETAS_PATH, sheet_name=SHEET_DATA_BOLETAS)
    return df.dropna(subset=["NUMERO DE RECIBO", "NOMBRES"]).reset_index(drop=True)

# ==================PROCESAMIENTO CON PANDAS==================
def process_data(df: pd.DataFrame):
    numeric_cols = [
        "Marcación anterior", "Marcacion altual", "M3",
        "Total mes actual", "MES ANTERIOR", "Corte y reconexion",
        "Convenio", "Mantenimiento", "Multa (faena + reunión)",
        "Cuota directa", "Importe a pagar",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df

# =====================VALIDACION============================
def validate_data(df: pd.DataFrame):
    required_cols = ["NUMERO DE RECIBO", "NOMBRES", "PERIODO", "Total"]
    missing = df[required_cols].isnull()
    if missing.any().any():
        raise ValueError("Hay datos faltantes en columnas clave")

    duplicated = df.duplicated(subset=["NUMERO DE RECIBO"])
    if duplicated.any():
        raise ValueError("Hay recibos duplicados")

    numeric_cols = ["M3", "Total", "Total mes actual"]
    for col in numeric_cols:
        if col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                raise ValueError(f"La columna {col} debe ser numerica")

    if "LECTURA ANTERIOR" in df.columns and "LECTURA ACTUAL" in df.columns:
        invalid = df[df["LECTURA ACTUAL"] < df["LECTURA ANTERIOR"]]
        if not invalid.empty:
            print("Error en estas filas:")
            print(invalid[["NOMBRES", "LECTURA ANTERIOR", "LECTURA ACTUAL"]])
            raise ValueError("Lecturas inconsistentes detectadas")

    if "M3" in df.columns:
        if (df["M3"] < 0).any():
            raise ValueError("Hay valores negativos en M3")

    if "PERIODO" in df.columns:
        invalid_period = df[df["PERIODO"].astype(str).str.contains("2167|2099|3000")]
        if not invalid_period.empty:
            print("Periodos invalidos detectados:")
            print(invalid_period[["NUMERO DE RECIBO", "PERIODO"]])
            raise ValueError("Hay fechas incorrectas en PERIODO")

    print("Datos validados correctamente")

# =================GENERACION DE DOCUMENTOS==================
def generate_boletas(grouped, recibos_test: list):
    total = len(recibos_test)

    for i, recibo in enumerate(recibos_test):
        try:
            data = grouped.get_group(recibo)
        except KeyError:
            print(f"  [ERROR] Recibo {recibo} no encontrado en los datos.")
            continue

        print(f"[{i+1}/{total}] Generando recibo {recibo}...")

        doc = DocxTemplate(PLANTILLA_BOLETAS_PATH)
        row = data.iloc[0]

        logo_jaas = InlineImage(doc, str(IMG_LOGO_JAAS_PATH), width=Mm(31))
        imagen_qr = InlineImage(doc, str(IMG_QR_PATH), width=Mm(100))

        estado = str(row["Estado"]).strip().upper()
        if "NO EST" in estado:
            carita = InlineImage(doc, str(IMG_CARITA_TRISTE_PATH), width=Mm(26))
        else:
            carita = InlineImage(doc, str(IMG_CARITA_FELIZ_PATH), width=Mm(26))

        def _rv(col, fallback=""):
            v = row.get(col, fallback)
            if str(v).strip() in ("", "nan", "None", "NaT"):
                return fallback
            if hasattr(v, "strftime"):
                return v.strftime("%d/%m/%Y")
            return str(v).strip()

        context = {
            "nu_reci": recibo,
            "nombre_usuario": row["NOMBRES"],
            "direccion_usuario": "Mz." + " " + str(row["MZ"]) + " " + "Lt." + " " + str(row["LT"]),
            "an": row["Marcación anterior"],
            "ac": row["Marcacion altual"],
            "c": row["M3"],
            "tmac": row["Total mes actual"],
            "tman": row["MES ANTERIOR"],
            "c_rec": row["Corte y reconexion"],
            "con": row["Convenio"],
            "man": row["Mantenimiento"],
            "mul": row["Multa (faena + reunión)"],
            "cd": row["Cuota directa"],
            "ip": row["Importe a pagar"],
            "ep": row["Estado"],
            "icono_estado": carita,
            "imagen_qr": imagen_qr,
            "nombre_jaas": NOMBRE_JAAS,
            "sector": SECTOR,
            "lectura_anterior": _rv("LECTURA ANTERIOR", LECTURA_ANTERIOR),
            "lectura_actual":   _rv("LECTURA ACTUAL",   LECTURA_ACTUAL),
            "periodo":          _rv("PERIODO",           PERIODO),
            "fv":               _rv("FECHA DE VENCIMIENTO", FECHA_VENCIMIENTO),
            "fe":               _rv("FECHA DE EMISIÓN",  FECHA_EMISION),
            "fecha_pago":       _rv("FECHA_PAGO",        FECHA_PAGO),
            "hora_pago":        HORA_PAGO,
            "telefono": TELEFONO,
            "logo_jaas": logo_jaas,
        }

        doc.render(context)

        mz = str(row["MZ"]).strip().replace(" ", "")
        lt = str(row["LT"]).strip().replace(" ", "")
        base_name = f"RECIBO_{recibo}_{mz}_{lt}"

        output_docx = OUTPUT_DIR / f"{base_name}.docx"
        doc.save(output_docx)

        output_pdf = OUTPUT_DIR / f"{base_name}.pdf"
        try:
            convert(str(output_docx), str(output_pdf))
        except Exception as e:
            print(f"  [ERROR PDF] Recibo {recibo}: {e}")
            continue

        try:
            img_name = f"{sanitize(mz)}_{sanitize(lt)}_{sanitize(row['NOMBRES'])}.jpg"
            pdf_doc = fitz.open(str(output_pdf))
            pix = pdf_doc[0].get_pixmap(matrix=fitz.Matrix(150/72, 150/72))
            pix.save(str(IMAGES_DIR / img_name))
            pdf_doc.close()
        except Exception as e:
            print(f"  [ERROR IMG] Recibo {recibo}: {e}")

    print("\nBoletas de test generadas.")

# ====================CONSOLIDACION==========================
def merge_pdfs(output_dir: Path, output_name="TEST_CONSOLIDADO.pdf"):
    merger = PdfMerger()
    pdf_files = sorted(output_dir.glob("RECIBO_*.pdf"))
    for pdf in pdf_files:
        merger.append(str(pdf))
    merger.write(str(output_dir / output_name))
    merger.close()
    print("Consolidado de test generado.")

# ======================MAIN=================================
def main():
    reset_output_folder(OUTPUT_DIR)
    IMAGES_DIR.mkdir()

    df = load_data()
    validate_data(df)
    df = process_data(df)

    recibos_test = select_test_receipts(df)

    grouped = df.groupby("NUMERO DE RECIBO")
    generate_boletas(grouped, recibos_test)

    merge_pdfs(OUTPUT_DIR)

if __name__ == "__main__":
    main()
