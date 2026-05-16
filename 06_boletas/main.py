# =========================IMPORTS===========================
import os
import pandas as pd
import shutil
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx2pdf import convert
from docx.shared import Mm
from PyPDF2 import PdfMerger
# ========================CONFIGURACION======================
# Variables de carpeta
BASE_DIR=Path(".")
INPUT_DIR=BASE_DIR/"Inputs"
OUTPUT_DIR=BASE_DIR/"Outputs"

DATA_BOLETAS_PATH=INPUT_DIR/"DATA_boletas.xlsx"
PLANTILLA_BOLETAS_PATH=INPUT_DIR/"PLANTILLA_boletas.docx"
IMG_LOGO_JAAS_PATH=INPUT_DIR/"logo_jaas.png"
IMG_CARITA_TRISTE_PATH=INPUT_DIR/"carita_triste.png"
IMG_CARITA_FELIZ_PATH=INPUT_DIR/"carita_feliz.png"
IMG_QR_PATH=INPUT_DIR/"imagen_qr.png"

# Variables constante del recibo.
NOMBRE_JAAS = "JUNTA ADMINISTRATIVA DE SERVICIOS DE SANEAMIENTO"
SECTOR = "P.J. TUPAC AMARU"
LECTURA_ANTERIOR = "10/12/2O25"
LECTURA_ACTUAL = "10/01/2026"
PERIODO = "11/02/2026 al 10/03/2026"
FECHA_VENCIMIENTO = "07/02/2026"
FECHA_EMISION = "27/03/2026"
FECHA_PAGO = "04/04"
HORA_PAGO = "4-6 pm"
LUGAR_PAGO = "LOCAL DEL PUEBLO"
TELEFONO =  "948 227 636"

SHEET_DATA_BOLETAS="Data"

#=====================UTILIDADES=========================
# Elimina carpeta y crea otra
def reset_output_folder(path: Path):
    if path.exists():
        shutil.rmtree(path)  # elimina todo (carpeta incluida)
    path.mkdir()

#====================CARGA DE DATOS========================
def load_data():
    data_boletas_df = pd.read_excel(DATA_BOLETAS_PATH, sheet_name=SHEET_DATA_BOLETAS)

    return data_boletas_df

#=====================VALIDACION DE DATOS====================
def validate_data(df: pd.DataFrame):
    # 1. Columnas obligatorias
    required_cols = ["NUMERO DE RECIBO", "NOMBRES", "PERIODO", "Total"]

    missing = df[required_cols].isnull()
    if missing.any().any():
        raise ValueError("Hay datos faltantes en columnas clave")

    # 2. Duplicados
    duplicated = df.duplicated(subset=["NUMERO DE RECIBO"])
    if duplicated.any():
        raise ValueError("Hay recibos duplicados")

    # 3. Validar columnas numéricas
    numeric_cols = ["M3", "Total", "Total mes actual"]

    for col in numeric_cols:
        if col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                raise ValueError(f"La columna {col} debe ser numérica")

    # 4. Validaciones lógicas

    # Lecturas coherentes
    if "LECTURA ANTERIOR" in df.columns and "LECTURA ACTUAL" in df.columns:
        invalid = df[df["LECTURA ACTUAL"] < df["LECTURA ANTERIOR"]]

        if not invalid.empty:
            print("Error en estas filas:")
            print(invalid[["NOMBRES", "LECTURA ANTERIOR", "LECTURA ACTUAL"]])
            raise ValueError("Lecturas inconsistentes detectadas")

    # M3 no negativos
    if "M3" in df.columns:
        if (df["M3"] < 0).any():
            raise ValueError("Hay valores negativos en M3")

    # 🔥 NUEVO: Validar años absurdos en PERIODO
    if "PERIODO" in df.columns:
        invalid_period = df[df["PERIODO"].astype(str).str.contains("2167|2099|3000")]

        if not invalid_period.empty:
            print("⚠️ Periodos inválidos detectados:")
            print(invalid_period[["NUMERO DE RECIBO", "PERIODO"]])
            raise ValueError("Hay fechas incorrectas en PERIODO")

    print("✔ Datos validados correctamente")
#==================PROCESAMIENTO CON PANDAS==================
def process_data(df: pd.DataFrame):
    # No se requiere procesamiento adicional
    return df
#======================AGRUPACION==================
def group_data(df: pd.DataFrame):
    return df.groupby("NUMERO DE RECIBO")

#=================GENERACION DE DOCUMENTOS===================
def generate_boletas(grouped, limit=None):
    total = grouped.ngroups  # número total de recibos

    for i, (recibo, data) in enumerate(grouped):

        # 🔥 límite de pruebas
        if limit is not None and i >= limit:
            break

        print(f"[{i+1}/{total}] Generando recibo {recibo}...")

        doc = DocxTemplate(PLANTILLA_BOLETAS_PATH)

        # Tomamos la primera fila del grupo
        row = data.iloc[0]

        # LOGO JAAS
        logo_jaas = InlineImage(doc, str(IMG_LOGO_JAAS_PATH), width=Mm(31))

        # IMAGEN QR
        imagen_qr = InlineImage(doc, str(IMG_QR_PATH), width=Mm(100))

        # ESTADO (carita)
        estado = str(row["Estado"]).strip().upper()

        if "NO ESTÁ AL DÍA" in estado:
            carita = InlineImage(doc, str(IMG_CARITA_TRISTE_PATH), width=Mm(26))
        else:
            carita = InlineImage(doc, str(IMG_CARITA_FELIZ_PATH), width=Mm(26))

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

            # Datos constantes
            "nombre_jaas": NOMBRE_JAAS,
            "sector": SECTOR,
            "lectura_anterior": LECTURA_ANTERIOR,
            "lectura_actual": LECTURA_ACTUAL,
            "periodo": PERIODO,
            "fv": FECHA_VENCIMIENTO,
            "fe": FECHA_EMISION,
            "fecha_pago": FECHA_PAGO,
            "hora_pago": HORA_PAGO,
            "telefono": TELEFONO,
            "logo_jaas": logo_jaas,
        }

        doc.render(context)

        filename = f"RECIBO_{recibo}.docx"
        output_docx = OUTPUT_DIR / filename

        doc.save(output_docx)

        output_pdf = OUTPUT_DIR / f"RECIBO_{recibo}.pdf"

        try:
            convert(str(output_docx), str(output_pdf))
        except Exception as e:
            print(f"❌ Error en recibo {recibo}: {e}")

    print("✔ Boletas generadas en PDF correctamente")
#========================CONSOLIDACION DE DOCUMENTOS==========================
def merge_pdfs(output_dir: Path, output_name="CONSOLIDADO.pdf"):
    merger = PdfMerger()

    pdf_files = sorted(output_dir.glob("RECIBO_*.pdf"))

    for pdf in pdf_files:
        merger.append(str(pdf))

    merger.write(str(output_dir / output_name))
    merger.close()

    print("✔ Consolidado generado")

#======================FUNCION PRINCIPAL====================
def main():
    # 1. Resetear carpeta de salida
    reset_output_folder(OUTPUT_DIR)

    # 2. Cargar datos
    df = load_data()

    # 3. Validar datos
    validate_data(df)

    # 4. Procesar datos (por ahora no hace cambios)
    df = process_data(df)

    # 5. Agrupar por número de recibo
    grouped = group_data(df)

    # 6. Generar boletas (Word + PDF)
    generate_boletas(grouped, limit=None)

    # 7. Consolidar las boletas
    merge_pdfs(OUTPUT_DIR)

if __name__ == "__main__":
    main()