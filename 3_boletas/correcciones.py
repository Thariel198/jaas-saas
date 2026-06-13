# =============================================================
# correcciones.py — Reimpresion de boletas por MZ y LT
#
# USO:
#   1. Edita la lista RECIBOS_A_REIMPRIMIR mas abajo con los
#      pares (MZ, LT) que quieras corregir.
#   2. Corre el archivo. Por cada par:
#        - Busca la fila en DATA_boletas.xlsx (matching normalizado).
#        - Elimina en Outputs/  los .docx y .pdf antiguos de ese MZ/LT.
#        - Elimina en Outputs/Imagenes/  el .jpg antiguo de ese MZ/LT.
#        - Genera nuevamente .docx + .pdf en Outputs/ y .jpg en Imagenes/.
#        - Copia el .docx y .pdf nuevo a Outputs/Correcciones/  (ACUMULA).
#   3. Regenera Outputs/CONSOLIDADO.pdf con todos los recibos.
#   4. Regenera Outputs/Correcciones/CONSOLIDADO_CORRECCIONES.pdf
#      con todo lo acumulado del mes.
# =============================================================

import re
import shutil
import unicodedata

import fitz
import pandas as pd
from PyPDF2 import PdfMerger
from docx.shared import Mm
from docx2pdf import convert
from docxtpl import DocxTemplate, InlineImage
from pathlib import Path

# ── paths ─────────────────────────────────────────────────────
BASE_DIR = Path(".")
INPUT_DIR = BASE_DIR / "Inputs"
OUTPUT_DIR = BASE_DIR / "Outputs"
IMAGES_DIR = OUTPUT_DIR / "Imagenes"
OUTPUT_CORRECCIONES_DIR = OUTPUT_DIR / "Correcciones"

DATA_BOLETAS_PATH = INPUT_DIR / "DATA_boletas.xlsx"
PLANTILLA_BOLETAS_PATH = INPUT_DIR / "PLANTILLA_boletas.docx"
IMG_LOGO_JAAS_PATH = INPUT_DIR / "logo_jaas.png"
IMG_CARITA_TRISTE_PATH = INPUT_DIR / "carita_triste.png"
IMG_CARITA_FELIZ_PATH = INPUT_DIR / "carita_feliz.png"
IMG_QR_PATH = INPUT_DIR / "imagen_qr.png"

# ── constantes del recibo ─────────────────────────────────────
NOMBRE_JAAS = "JUNTA ADMINISTRATIVA DE SERVICIOS DE SANEAMIENTO"
SECTOR = "P.J. TUPAC AMARU"
LECTURA_ANTERIOR = "10/12/2O25"
LECTURA_ACTUAL = "10/01/2026"
PERIODO = "11/02/2026 al 10/03/2026"
FECHA_VENCIMIENTO = "07/02/2026"
FECHA_EMISION = "27/03/2026"
FECHA_PAGO = "04/04"
HORA_PAGO = "4-6 pm"
TELEFONO = "948 227 636"

SHEET_DATA_BOLETAS = "Data"

# ══════════════════════════════════════════════════════════════
#  EDITAR AQUI — pares (MZ, LT) a reimprimir
# ══════════════════════════════════════════════════════════════
RECIBOS_A_REIMPRIMIR = [
    ("A1", "12"),
    ("I",  "16"),
    ("C1", "7"),
]
# ══════════════════════════════════════════════════════════════


# ── utilidades ────────────────────────────────────────────────
def sanitize(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^\w]", "_", s).strip("_")


def normalize_filename(s):
    """Igual que main.py: str(s).strip().replace(' ', '')."""
    return str(s).strip().replace(" ", "")


def normalize_key(s):
    """Para matching en DataFrame: sin espacios, sin tildes, mayusculas."""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.strip().replace(" ", "").upper()


# ── filtrado ─────────────────────────────────────────────────
def filtrar_df(df, pares):
    df = df.copy()
    df["_MZ_KEY"] = df["MZ"].apply(normalize_key)
    df["_LT_KEY"] = df["LT"].apply(normalize_key)

    filas = []
    no_encontrados = []
    for mz, lt in pares:
        mz_n, lt_n = normalize_key(mz), normalize_key(lt)
        match = df[(df["_MZ_KEY"] == mz_n) & (df["_LT_KEY"] == lt_n)]
        if match.empty:
            no_encontrados.append((mz, lt))
            continue
        filas.append(match)

    if no_encontrados:
        print("\n[AVISO] No encontrados en DATA_boletas:")
        for mz, lt in no_encontrados:
            print(f"  - MZ={mz}  LT={lt}")

    if not filas:
        return df.iloc[0:0]
    return pd.concat(filas).drop(columns=["_MZ_KEY", "_LT_KEY"])


# ── limpieza de archivos previos ─────────────────────────────
def eliminar_archivos_previos(mz, lt):
    """Borra el .docx, .pdf y .jpg antiguos asociados a este MZ/LT."""
    mz_f = normalize_filename(mz)
    lt_f = normalize_filename(lt)

    borrados = []
    # Outputs/RECIBO_*_{mz}_{lt}.docx y .pdf
    for ext in ("docx", "pdf"):
        for f in OUTPUT_DIR.glob(f"RECIBO_*_{mz_f}_{lt_f}.{ext}"):
            try:
                f.unlink()
                borrados.append(f.name)
            except Exception as e:
                print(f"  [ERROR BORRAR] {f.name}: {e}")

    # Outputs/Imagenes/{san_mz}_{san_lt}_*.jpg
    if IMAGES_DIR.exists():
        mz_s = sanitize(mz_f)
        lt_s = sanitize(lt_f)
        for f in IMAGES_DIR.glob(f"{mz_s}_{lt_s}_*.jpg"):
            try:
                f.unlink()
                borrados.append(f.name)
            except Exception as e:
                print(f"  [ERROR BORRAR] {f.name}: {e}")

    if borrados:
        for n in borrados:
            print(f"    [-] {n}")


# ── generacion de un recibo ──────────────────────────────────
def generar_recibo(row):
    recibo = row["NUMERO DE RECIBO"]
    doc = DocxTemplate(PLANTILLA_BOLETAS_PATH)

    logo_jaas = InlineImage(doc, str(IMG_LOGO_JAAS_PATH), width=Mm(31))
    imagen_qr = InlineImage(doc, str(IMG_QR_PATH), width=Mm(100))

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

    def _rn(col):
        v = row.get(col, 0)
        if pd.isna(v) or str(v).strip() in ("", "nan", "None"):
            return 0
        try:
            return float(v)
        except (ValueError, TypeError):
            return 0

    context = {
        "nu_reci": recibo,
        "nombre_usuario": row["NOMBRES"],
        "direccion_usuario": "Mz." + " " + str(row["MZ"]) + " " + "Lt." + " " + str(row["LT"]),
        "an": _rn("Marcación anterior"),
        "ac": _rn("Marcacion altual"),
        "c": _rn("M3"),
        "tmac": _rn("Total mes actual"),
        "tman": _rn("MES ANTERIOR"),
        "c_rec": _rn("Corte y reconexion"),
        "con": _rn("Convenio"),
        "man": _rn("Mantenimiento"),
        "mul": _rn("Multa (faena + reunión)"),
        "cd": _rn("Cuota directa"),
        "ip": _rn("Importe a pagar"),
        "ep": row["Estado"],
        "icono_estado": carita,
        "imagen_qr": imagen_qr,
        "nombre_jaas": NOMBRE_JAAS,
        "sector": SECTOR,
        "lectura_anterior": _rv("LECTURA ANTERIOR", LECTURA_ANTERIOR),
        "lectura_actual": _rv("LECTURA ACTUAL", LECTURA_ACTUAL),
        "periodo": _rv("PERIODO", PERIODO),
        "fv": _rv("FECHA DE VENCIMIENTO", FECHA_VENCIMIENTO),
        "fe": _rv("FECHA DE EMISIÓN", FECHA_EMISION),
        "fecha_pago": _rv("FECHA_PAGO", FECHA_PAGO),
        "hora_pago": HORA_PAGO,
        "telefono": TELEFONO,
        "logo_jaas": logo_jaas,
    }

    doc.render(context)

    mz_f = normalize_filename(row["MZ"])
    lt_f = normalize_filename(row["LT"])
    base_name = f"RECIBO_{recibo}_{mz_f}_{lt_f}"

    output_docx = OUTPUT_DIR / f"{base_name}.docx"
    output_pdf = OUTPUT_DIR / f"{base_name}.pdf"

    doc.save(output_docx)

    try:
        convert(str(output_docx), str(output_pdf))
    except Exception as e:
        print(f"    [ERROR PDF] Recibo {recibo}: {e}")
        return None, None

    # JPG en Outputs/Imagenes/
    try:
        img_name = f"{sanitize(mz_f)}_{sanitize(lt_f)}_{sanitize(row['NOMBRES'])}.jpg"
        pdf_doc = fitz.open(str(output_pdf))
        pix = pdf_doc[0].get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
        pix.save(str(IMAGES_DIR / img_name))
        pdf_doc.close()
    except Exception as e:
        print(f"    [ERROR JPG] Recibo {recibo}: {e}")

    return output_docx, output_pdf


# ── consolidaciones ──────────────────────────────────────────
def regenerar_consolidado(output_dir: Path, output_name: str):
    consolidado = output_dir / output_name
    if consolidado.exists():
        try:
            consolidado.unlink()
        except Exception as e:
            print(f"[ERROR] No se pudo borrar {consolidado.name}: {e}")
            return

    pdfs = sorted(output_dir.glob("RECIBO_*.pdf"))
    if not pdfs:
        print(f"[INFO] Sin PDFs en {output_dir}, no se genera {output_name}.")
        return

    merger = PdfMerger()
    for p in pdfs:
        merger.append(str(p))
    merger.write(str(consolidado))
    merger.close()
    print(f"[OK] {consolidado}")


# ── main ─────────────────────────────────────────────────────
def main():
    # 1. Tomar pares desde la lista hardcoded
    pares = RECIBOS_A_REIMPRIMIR
    if not pares:
        print("[INFO] RECIBOS_A_REIMPRIMIR esta vacio. Saliendo.")
        return

    print("\n" + "=" * 60)
    print(f"  REIMPRESION DE {len(pares)} BOLETA(S)")
    for mz, lt in pares:
        print(f"    - MZ {mz}  LT {lt}")
    print("=" * 60)

    # 2. Cargar y filtrar
    df = pd.read_excel(DATA_BOLETAS_PATH, sheet_name=SHEET_DATA_BOLETAS)
    df_correc = filtrar_df(df, pares)
    if df_correc.empty:
        print("\n[ERROR] Ningun recibo coincide con los MZ/LT ingresados.")
        return

    # 3. Carpeta de correcciones acumulativa
    OUTPUT_CORRECCIONES_DIR.mkdir(parents=True, exist_ok=True)
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)

    # 4. Regenerar cada recibo
    total = len(df_correc)
    print(f"\nReimprimiendo {total} recibo(s)...\n")
    ok = 0
    for i, (_, row) in enumerate(df_correc.iterrows(), 1):
        recibo = row["NUMERO DE RECIBO"]
        mz = str(row["MZ"]).strip()
        lt = str(row["LT"]).strip()

        print(f"[{i}/{total}] Recibo {recibo}  MZ {mz}  LT {lt}")

        # 4a. Borrar archivos previos
        eliminar_archivos_previos(mz, lt)

        # 4b. Generar nuevo docx + pdf + jpg
        new_docx, new_pdf = generar_recibo(row)
        if new_pdf is None:
            continue

        # 4c. Copiar a Outputs/Correcciones/ (acumula, sobrescribe si misma boleta)
        try:
            shutil.copy2(new_docx, OUTPUT_CORRECCIONES_DIR / new_docx.name)
            shutil.copy2(new_pdf, OUTPUT_CORRECCIONES_DIR / new_pdf.name)
            print(f"    [OK] Copiado a Correcciones/")
        except Exception as e:
            print(f"    [ERROR COPIA] {e}")
            continue

        ok += 1

    # 5. Regenerar consolidados
    print("\nRegenerando consolidados...")
    regenerar_consolidado(OUTPUT_DIR, "CONSOLIDADO.pdf")
    regenerar_consolidado(OUTPUT_CORRECCIONES_DIR, "CONSOLIDADO_CORRECCIONES.pdf")

    print(f"\n[OK] {ok}/{total} correcciones generadas.")


if __name__ == "__main__":
    main()
