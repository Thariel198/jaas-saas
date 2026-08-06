# =========================IMPORTS===========================
import os
import re
import sys
import unicodedata
import pandas as pd
import shutil
import fitz
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage, RichText
from docx2pdf import convert
from docx.shared import Mm
from PyPDF2 import PdfMerger
# ========================CONFIGURACION======================
# Variables de carpeta
BASE_DIR=Path(".")
INPUT_DIR=BASE_DIR/"Inputs"
OUTPUT_DIR=BASE_DIR/"Outputs"
IMAGES_DIR=OUTPUT_DIR/"Imagenes"

DATA_BOLETAS_PATH=INPUT_DIR/"DATA_boletas.xlsx"
PENDIENTES_CONVENIO_MULTAS_PATH=INPUT_DIR/"pendientes_convenio_multas.xlsx"
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
FECHA_PAGO = "05-06/06"
HORA_PAGO = "16:00-20:00 y 10:00-13:00 hrs"
LUGAR_PAGO = "LOCAL DEL PUEBLO"
TELEFONO =  "948 227 636"

SHEET_DATA_BOLETAS="Data"

#=====================UTILIDADES=========================
def sanitize(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^\w]", "_", s).strip("_")

# Elimina carpeta y crea otra
def reset_output_folder(path: Path):
    if path.exists():
        shutil.rmtree(path)  # elimina todo (carpeta incluida)
    path.mkdir()

#====================CARGA DE DATOS========================
def load_data():
    data_boletas_df = pd.read_excel(DATA_BOLETAS_PATH, sheet_name=SHEET_DATA_BOLETAS)
    data_boletas_df = data_boletas_df.dropna(subset=["NUMERO DE RECIBO", "NOMBRES"]).reset_index(drop=True)
    return data_boletas_df

# CONCEPTO (pendientes_convenio_multas.xlsx) -> campo del contexto de la boleta
CONCEPTO_A_CAMPO_BOLETA = {"CONVENIO": "con", "MULTA": "mul", "ACUERDOS_ASAMBLEA": "cd", "MES_ANTERIOR": "tman"}

def _norm_mz(v):
    return str(v).strip().upper()

def _norm_lt(v):
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.upper()

def load_pendientes_convenio_multas():
    """Predios con CONVENIO/MULTA/ACUERDOS_ASAMBLEA en verificacion (overlay que reemplaza
    el monto por la palabra Verificando al imprimir; no toca DATA_boletas.xlsx)."""
    if not PENDIENTES_CONVENIO_MULTAS_PATH.exists():
        return {}
    df = pd.read_excel(PENDIENTES_CONVENIO_MULTAS_PATH, header=1, dtype=str).fillna("")
    pendientes = {}
    for _, row in df.iterrows():
        mz = _norm_mz(row.get("MZ", ""))
        lt = _norm_lt(row.get("LT", ""))
        concepto = str(row.get("CONCEPTO", "")).strip().upper()
        estado = str(row.get("ESTADO", "")).strip().upper()
        if not mz or not lt or concepto not in CONCEPTO_A_CAMPO_BOLETA:
            continue
        pendientes[(mz, lt, concepto)] = estado
    return pendientes

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
            print("[AVISO] Periodos invalidos detectados:")
            print(invalid_period[["NUMERO DE RECIBO", "PERIODO"]])
            raise ValueError("Hay fechas incorrectas en PERIODO")

    print("[OK] Datos validados correctamente")
#==================PROCESAMIENTO CON PANDAS==================
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
#======================AGRUPACION==================
def group_data(df: pd.DataFrame):
    return df.groupby("NUMERO DE RECIBO")

#=================GENERACION DE DOCUMENTOS===================
def generate_boletas(grouped, limit=None):
    total = grouped.ngroups  # número total de recibos
    pendientes_verificando = load_pendientes_convenio_multas()

    for i, (recibo, data) in enumerate(grouped):

        # 🔥 límite de pruebas
        if limit is not None and i >= limit:
            break

        # Tomamos la primera fila del grupo
        row = data.iloc[0]

        mz = str(row["MZ"]).strip().replace(" ", "")
        lt = str(row["LT"]).strip().replace(" ", "")
        base_name = f"RECIBO_{recibo}_{mz}_{lt}"
        output_docx = OUTPUT_DIR / f"{base_name}.docx"
        output_pdf = OUTPUT_DIR / f"{base_name}.pdf"

        # Reanudación: si el PDF ya existe de una corrida anterior, no se regenera.
        # Para regenerar todo desde cero: py main.py --desde-cero
        if output_pdf.exists():
            print(f"[{i+1}/{total}] Recibo {recibo} ya generado — se saltea")
            continue

        print(f"[{i+1}/{total}] Generando recibo {recibo}...")

        doc = DocxTemplate(PLANTILLA_BOLETAS_PATH)

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

        def _rv(col, fallback=""):
            v = row.get(col, fallback)
            if str(v).strip() in ("", "nan", "None", "NaT"):
                return fallback
            if hasattr(v, "strftime"):
                return v.strftime("%d/%m/%Y")
            return str(v).strip()

        def _num(col):
            # 77.0 → 77 (como salían en junio); 77.5 queda 77.5
            v = row[col]
            if isinstance(v, float) and v == int(v):
                return int(v)
            return v

        context = {
            "nu_reci": recibo,
            "nombre_usuario": row["NOMBRES"],
            "direccion_usuario": "Mz." + " " + str(row["MZ"]) + " " + "Lt." + " " + str(row["LT"]),
            "an": _num("Marcación anterior"),
            "ac": _num("Marcacion altual"),
            "c": _num("M3"),
            "tmac": _num("Total mes actual"),
            "tman": _num("MES ANTERIOR"),
            "c_rec": _num("Corte y reconexion"),
            "con": _num("Convenio"),
            "man": _num("Mantenimiento"),
            "mul": _num("Multa (faena + reunión)"),
            "cd": _num("Cuota directa"),
            "ip": _num("Importe a pagar"),
            "ep": row["Estado"],
            "icono_estado": carita,
            "imagen_qr": imagen_qr,

            # Datos del mes — leídos del row si 06a los puso; fallback a constantes
            "nombre_jaas": NOMBRE_JAAS,
            "sector": SECTOR,
            "lectura_anterior": _rv("LECTURA ANTERIOR", LECTURA_ANTERIOR),
            "lectura_actual":   _rv("LECTURA ACTUAL",   LECTURA_ACTUAL),
            "periodo":          _rv("PERIODO",           PERIODO),
            "fv":               _rv("FECHA DE VENCIMIENTO", FECHA_VENCIMIENTO),
            "fe":               _rv("FECHA DE EMISIÓN",  FECHA_EMISION),
            "fecha_pago":       _rv("FECHA_PAGO",        FECHA_PAGO),
            "hora_pago":        _rv("HORA_PAGO",         HORA_PAGO),
            "telefono": TELEFONO,
            "logo_jaas": logo_jaas,
        }

        # Overlay: CONVENIO/MULTA/ACUERDOS_ASAMBLEA en verificacion -> imprime "Verificando"
        # en vez del monto, y se resta del Total/Importe a pagar (no se cobra este
        # ciclo mientras el reclamo esta en revision). No toca DATA_boletas.
        clave_mz, clave_lt = _norm_mz(mz), _norm_lt(lt)
        for concepto, campo in CONCEPTO_A_CAMPO_BOLETA.items():
            if pendientes_verificando.get((clave_mz, clave_lt, concepto)) == "VERIFICANDO":
                monto_en_revision = context[campo]
                try:
                    context["ip"] = context["ip"] - float(monto_en_revision)
                except (TypeError, ValueError):
                    pass
                # Fuente mas chica: "Verificando" no entra en 1 linea en la columna
                # angosta (~1.23in) al tamano normal (20pt) y desborda la boleta a
                # una 2da pagina invisible. 12pt (sz=24, medio-puntos) si entra.
                context[campo] = RichText("Verificando", size=24)

        if isinstance(context["ip"], float) and context["ip"] == int(context["ip"]):
            context["ip"] = int(context["ip"])

        doc.render(context)
        doc.save(output_docx)

        try:
            convert(str(output_docx), str(output_pdf))
        except Exception as e:
            print(f"[ERROR PDF] Recibo {recibo}: {e}")
            continue

        try:
            img_name = f"{sanitize(mz)}_{sanitize(lt)}_{sanitize(row['NOMBRES'])}.jpg"
            pdf_doc = fitz.open(str(output_pdf))
            pix = pdf_doc[0].get_pixmap(matrix=fitz.Matrix(150/72, 150/72))
            pix.save(str(IMAGES_DIR / img_name))
            pdf_doc.close()
        except Exception as e:
            print(f"[ERROR IMG] Recibo {recibo}: {e}")

    print("[OK] Boletas generadas en PDF correctamente")
#========================CONSOLIDACION DE DOCUMENTOS==========================
def _es_mz_compuesta(pdf: Path) -> bool:
    # RECIBO_{recibo}_{mz}_{lt}.pdf → mz en la posición 2
    mz = pdf.stem.split("_")[2]
    return bool(re.fullmatch(r"[A-Z]+\d+", mz))


def merge_pdfs(output_dir: Path):
    # Dos bloques de impresión: MZ simples (A…Z) y compuestas (A1…H1).
    # Archivos más chicos → abren y se imprimen por tandas sin colgarse.
    pdf_files = sorted(output_dir.glob("RECIBO_*.pdf"))
    bloques = {
        "CONSOLIDADO_SIMPLES.pdf":    [p for p in pdf_files if not _es_mz_compuesta(p)],
        "CONSOLIDADO_COMPUESTAS.pdf": [p for p in pdf_files if _es_mz_compuesta(p)],
    }
    for nombre, pdfs in bloques.items():
        if not pdfs:
            print(f"[AVISO] {nombre}: sin boletas — no se genera")
            continue
        merger = PdfMerger()
        for pdf in pdfs:
            merger.append(str(pdf))
        destino = output_dir / nombre
        merger.write(str(destino))
        merger.close()

        # Word embebe una copia de las fuentes en CADA boleta — al fusionar 500+
        # quedan cientos de copias redundantes (subset_fonts las funde en una sola
        # por fuente, ~90% menos peso). Sin esto el consolidado no entra en WhatsApp.
        tmp = output_dir / f"{nombre}.tmp"
        doc = fitz.open(str(destino))
        doc.subset_fonts()
        doc.save(str(tmp), garbage=4, deflate=True, clean=True)
        doc.close()
        tmp.replace(destino)

        print(f"[OK] {nombre}: {len(pdfs)} boletas ({destino.stat().st_size / 1024 / 1024:.1f} MB)")

#======================FUNCION PRINCIPAL====================
def main():
    # 1. Carpeta de salida — solo se borra con --desde-cero (permite reanudar
    #    una corrida colgada sin perder las boletas ya generadas)
    if "--desde-cero" in sys.argv:
        reset_output_folder(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(exist_ok=True)
    IMAGES_DIR.mkdir(exist_ok=True)

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