"""
validar_boletas_sin_servicio.py
Valida las boletas PDF de Outputs/Sin_servicio contra
data_boletas_sin_servicio.xlsx (registro emitido por boletas_sin_servicio.py).
Compara nombre, MZ, LT y montos clave.
"""
import sys; sys.stdout.reconfigure(encoding="utf-8")
import re
import pandas as pd
from pathlib import Path

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

BASE_DIR   = Path(".")
OUTPUT_DIR = BASE_DIR / "Outputs" / "Sin_servicio"

DATA_PATH  = OUTPUT_DIR / "data_boletas_sin_servicio.xlsx"

CAMPOS_NUM = {
    "AGUA":      "Mes anterior (agua vieja)",
    "CORTE":     "Corte , reconexión",
    "CONVENIO":  "Convenio",
    "MULTA":     "Multa (reu/fae)",
    "ACUERDOS":  "Cuota directa",
    "TOTAL":     "Importe a pagar",
}


def extraer_texto_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join((pg.extract_text() or "") for pg in reader.pages)


def fmt_variants(v) -> list[str]:
    """Devuelve los formatos posibles en que el PDF puede mostrar el numero.
    El template/Word puede mostrar 8.50 como '8.5', 75 como '75.0' o '75'."""
    try:
        f = float(v)
    except Exception:
        s = str(v).strip()
        return [s] if s else []
    if f == 0:
        return []
    if f == int(f):
        return [f"{int(f)}.0", str(int(f))]
    s = f"{f:.10f}".rstrip("0").rstrip(".")
    return [s]


def main():
    df = pd.read_excel(DATA_PATH)
    df = df.dropna(subset=["NUMERO DE RECIBO"]).reset_index(drop=True)

    for col in CAMPOS_NUM:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    total = len(df)
    ok = 0
    errores = []

    print(f"\n{'='*70}")
    print(f"  VALIDACION DE {total} BOLETAS SIN SERVICIO")
    print(f"{'='*70}\n")

    for _, row in df.iterrows():
        recibo = int(row["NUMERO DE RECIBO"])
        mz     = str(row["MZ"]).strip().replace(" ", "")
        lt     = str(row["LT"]).strip().replace(" ", "")
        nombre = str(row["NOMBRES"]).strip() if pd.notna(row["NOMBRES"]) else ""

        pdf_path = OUTPUT_DIR / f"RECIBO_{recibo}_{mz}_{lt}.pdf"
        if not pdf_path.exists():
            errores.append(f"  [FALTA ARCHIVO] Recibo {recibo} — {pdf_path.name}")
            continue

        try:
            texto = extraer_texto_pdf(pdf_path)
        except Exception as e:
            errores.append(f"  [ERROR LECTURA] Recibo {recibo} — {e}")
            continue

        texto_flat     = re.sub(r"\s+", " ", texto).upper()
        texto_noespace = re.sub(r"\s+", "", texto).upper()
        nombre_flat    = re.sub(r"\s+", " ", nombre).upper()
        fallas = []

        if nombre_flat and nombre_flat not in texto_flat:
            fallas.append(f"NOMBRE: esperado '{nombre}'")

        mz_key = mz.replace(" ", "").upper()
        lt_key = lt.replace(" ", "").upper()
        if f"MZ.{mz_key}" not in texto_noespace:
            fallas.append(f"MZ: esperado 'Mz. {mz}'")
        if f"LT.{lt_key}" not in texto_noespace:
            fallas.append(f"LT: esperado 'Lt. {lt}'")

        for col_excel, _label_pdf in CAMPOS_NUM.items():
            if col_excel not in df.columns:
                continue
            variants = fmt_variants(row[col_excel])
            if variants and not any(v in texto for v in variants):
                fallas.append(f"{col_excel}: esperado {variants[0]}")

        if fallas:
            errores.append(f"  [ERROR] Recibo {recibo} {mz}-{lt} {nombre[:30]}")
            for f in fallas:
                errores.append(f"         - {f}")
        else:
            print(f"  [OK]  Recibo {recibo:>6}  {mz}-{lt:<5}  {nombre[:40]}")
            ok += 1

    print(f"\n{'-'*70}")
    if errores:
        print(f"\n  FALLOS DETECTADOS:")
        for e in errores:
            print(e)
    print(f"\n  Resultado: {ok}/{total} correctos")
    print(f"{'='*70}\n")

    sys.exit(0 if ok == total else 1)


if __name__ == "__main__":
    main()
