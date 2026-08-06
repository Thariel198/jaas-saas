# =============================================================
# recibos_medidor_pagado.py — Constancias de CANCELACIÓN de medidor
#
# Papel de constancia para el socio que ya canceló su convenio de
# medidor: deuda original, detalle de pagos mes a mes y saldo en
# cero con sello CANCELADO. No es boleta de cobro (sin agua, sin
# mantenimiento, sin vencimiento).
#
#   - Fuente: shared/vista_seguimiento_pueblo.xlsx · hoja
#     CONVENIO_HISTORIAL (header=1, lectura en vivo).
#   - Universo: SALDO ACTUAL == 0 y DEUDA > 0. Solo medidor —
#     inscripción/multa/acuerdos no entran ni bloquean.
#   - Numeración propia MP-001… en orden de predio (no consume la
#     serie de recibos del ciclo).
#   - PDF directo con PyMuPDF (sin Word/COM). Regenerable siempre.
#
# Contrato: docs/formato_recibo_medidor_pagado.html
# Outputs/Medidor_pagado/: RECIBO_MP_NNN_MZ_LT.pdf,
#     CONSOLIDADO_MEDIDOR_PAGADO.pdf, data_recibos_medidor_pagado.xlsx
#
# Uso:  py recibos_medidor_pagado.py   (desde 3_boletas/)
# =============================================================

import glob
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import fitz
import pandas as pd
from PyPDF2 import PdfMerger

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "Inputs"
OUT_DIR = BASE_DIR / "Outputs" / "Medidor_pagado"
IMG_DIR = OUT_DIR / "Imagenes"
VISTA_PATH = BASE_DIR.parent / "shared" / "vista_seguimiento_pueblo.xlsx"

LOGO_PATH = INPUT_DIR / "logo_jaas.png"
CARITA_PATH = INPUT_DIR / "carita_feliz.png"

# sobrepagos aprobados a mano: la constancia se emite con los pagos
# CAPEADOS a la deuda (el exceso se corrige aparte en el ledger).
# Van al final de la serie para no correr la numeración ya emitida.
EXCESO_APROBADO = {("A", "1")}  # pagó 175 vs 100 — S/75 a favor pendiente

NOMBRE_JAAS = "JUNTA ADMINISTRATIVA DE SERVICIOS DE SANEAMIENTO"
SECTOR = "P.J. TUPAC AMARU"
TELEFONO = "Tel. 948 227 636"

def _mes_anterior(mes: str) -> str:
    y, m = mes.split("-")
    y, m = int(y), int(m)
    m -= 1
    if m == 0:
        m, y = 12, y - 1
    return f"{y:04d}-{m:02d}"

def _detectar_mes_ciclo() -> str:
    """Mes del ciclo actual, detectado desde la planilla_YYYY-MM.xlsx más reciente."""
    matches = sorted(glob.glob(str(BASE_DIR.parent / "2_planilla" / "outputs" / "planilla_*.xlsx")))
    if not matches:
        return "2026-06"  # fallback si no hay ninguna
    return Path(matches[-1]).stem.replace("planilla_", "")

MES_CIERRE = _mes_anterior(_detectar_mes_ciclo())

# A5 vertical (148 × 210 mm)
PAGE_W, PAGE_H = 420.94, 595.28
M = 36  # margen

VERDE = (6 / 255, 95 / 255, 70 / 255)
VERDE_BG = (209 / 255, 250 / 255, 229 / 255)
GRIS = (0.42, 0.45, 0.5)
NEGRO = (0.12, 0.16, 0.22)
ROJO = (0.73, 0.11, 0.11)
BORDE = (209 / 255, 213 / 255, 219 / 255)


def _img_reducida(path, px_max=256):
    """PNG reducido en memoria — los originales son 1024×1024 (~800KB) y en
    el papel se ven a ~2cm: incrustarlos completos en cada página inflaba
    el consolidado a ~450MB. A 256px la calidad de impresión no cambia."""
    pix = fitz.Pixmap(str(path))
    n = 0
    while max(pix.width, pix.height) >> n > px_max:
        n += 1
    if n:
        pix.shrink(n)
    return pix.tobytes("png")


def sanitize(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^\w]", "_", s).strip("_")


def _clave_orden(mz: str, lt: str) -> tuple:
    """Igual que enriquecimiento: MZ simples A-Z primero, compuestas después."""
    m = re.match(r"^([A-Z]+)(\d*)$", mz)
    letras, num = (m.group(1), m.group(2)) if m else (mz, "")
    mz_key = (1 if num else 0, letras, int(num) if num else 0)
    ml = re.match(r"^(\d+)\s*([A-Z]*)$", lt)
    lt_key = (0, int(ml.group(1)), ml.group(2)) if ml else (1, 0, lt)
    return mz_key + lt_key


def _soles(v: float) -> str:
    return f"S/ {v:,.2f}"


def _cargar_pagados():
    """Filas de CONVENIO_HISTORIAL con SALDO ACTUAL == 0, DEUDA > 0 y
    pagos == deuda EXACTO (la vista clampea el saldo a 0, así que un
    sobrepago también muestra saldo 0 — acá se excluye: ni exceso ni
    residuo). Detalle de pagos por mes (solo meses con pago > 0)."""
    df = pd.read_excel(VISTA_PATH, sheet_name="CONVENIO_HISTORIAL", header=1)
    cols_mes = df.columns.tolist()[4:-1]  # entre DEUDA y SALDO ACTUAL

    filas, excluidos = [], []
    for _, r in df.iterrows():
        deuda = pd.to_numeric(r["DEUDA"], errors="coerce")
        saldo = pd.to_numeric(r["SALDO ACTUAL"], errors="coerce")
        if pd.isna(deuda) or deuda <= 0 or pd.isna(saldo) or abs(saldo) > 0.005:
            continue
        pagos = []
        for c in cols_mes:
            v = pd.to_numeric(r[c], errors="coerce")
            if not pd.isna(v) and v > 0:
                pagos.append((c.replace("PAGO ", ""), float(v)))
        fila = {
            "MZ": str(r["MZ"]).strip(), "LT": str(r["LT"]).strip(),
            "NOMBRE": "" if pd.isna(r["NOMBRE"]) else str(r["NOMBRE"]).strip(),
            "DEUDA": float(deuda), "PAGOS": pagos,
            "TOTAL_PAGADO": sum(v for _, v in pagos),
        }
        if abs(fila["TOTAL_PAGADO"] - fila["DEUDA"]) > 0.005:
            excluidos.append(fila)
            continue
        filas.append(fila)

    filas.sort(key=lambda f: _clave_orden(f["MZ"], f["LT"]))

    # sobrepagos aprobados: pagos capeados a la deuda, al final de la serie
    for f in sorted(excluidos, key=lambda f: _clave_orden(f["MZ"], f["LT"])):
        if (f["MZ"], f["LT"]) not in EXCESO_APROBADO:
            print(f"[EXCLUIDO] {f['MZ']}-{f['LT']}  {f['NOMBRE'][:35]}  "
                  f"deuda S/{f['DEUDA']:.2f} vs pagado S/{f['TOTAL_PAGADO']:.2f} "
                  f"— no es saldo 0 exacto")
            continue
        capeados, resta = [], f["DEUDA"]
        for mes, v in f["PAGOS"]:
            if resta <= 0.005:
                break
            capeados.append((mes, min(v, resta)))
            resta -= min(v, resta)
        print(f"[CAPEADO]  {f['MZ']}-{f['LT']}  {f['NOMBRE'][:35]}  "
              f"pagado S/{f['TOTAL_PAGADO']:.2f} → constancia por S/{f['DEUDA']:.2f}")
        f["PAGOS"], f["TOTAL_PAGADO"] = capeados, f["DEUDA"]
        filas.append(f)

    for i, f in enumerate(filas, 1):
        f["NRO"] = f"MP-{i:03d}"
    return filas


def _txt_centrado(page, y, texto, size, color, font="helv"):
    w = fitz.get_text_length(texto, fontname=font, fontsize=size)
    page.insert_text(((PAGE_W - w) / 2, y), texto, fontsize=size,
                     fontname=font, color=color)


def _dibujar_constancia(doc, f, hoy, imgs):
    page = doc.new_page(width=PAGE_W, height=PAGE_H)

    # --- encabezado: logo + JASS + N° ---
    if imgs.get("logo"):
        page.insert_image(fitz.Rect(M, M, M + 52, M + 52), stream=imgs["logo"])
    x_txt = M + 62
    page.insert_text((x_txt, M + 14), NOMBRE_JAAS, fontsize=8, fontname="hebo", color=NEGRO)
    page.insert_text((x_txt, M + 26), SECTOR, fontsize=8, fontname="helv", color=NEGRO)
    page.insert_text((x_txt, M + 38), TELEFONO, fontsize=8, fontname="helv", color=GRIS)
    lbl = "CONSTANCIA"
    w = fitz.get_text_length(lbl, fontname="helv", fontsize=7)
    page.insert_text((PAGE_W - M - w, M + 14), lbl, fontsize=7, fontname="helv", color=GRIS)
    w = fitz.get_text_length(f"N° {f['NRO']}", fontname="hebo", fontsize=11)
    page.insert_text((PAGE_W - M - w, M + 28), f"N° {f['NRO']}", fontsize=11,
                     fontname="hebo", color=ROJO)
    page.draw_line((M, M + 62), (PAGE_W - M, M + 62), color=VERDE, width=1.5)

    # --- título ---
    _txt_centrado(page, M + 88, "CONSTANCIA DE CANCELACIÓN — MEDIDOR", 13, VERDE, "hebo")
    _txt_centrado(page, M + 102, f"Convenio de medidor · deuda original {_soles(f['DEUDA'])}",
                  8, GRIS)

    # --- datos del socio ---
    y = M + 128
    for lbl, val in (("SOCIO:", f["NOMBRE"]), ("PREDIO:", f"Mz. {f['MZ']}  Lt. {f['LT']}"),
                     ("FECHA DE EMISIÓN:", hoy)):
        page.insert_text((M + 8, y), lbl, fontsize=9, fontname="hebo", color=GRIS)
        page.insert_text((M + 110, y), val, fontsize=9, fontname="helv", color=NEGRO)
        y += 15

    # --- tabla de pagos ---
    tw, rh = 240.0, 16.0
    x0 = (PAGE_W - tw) / 2
    y += 8
    filas_tabla = [(mes, _soles(v), False) for mes, v in f["PAGOS"]]
    filas_tabla += [("TOTAL PAGADO", _soles(f["TOTAL_PAGADO"]), True),
                    ("SALDO", _soles(0), True)]

    page.draw_rect(fitz.Rect(x0, y, x0 + tw, y + rh), fill=VERDE, color=VERDE)
    page.insert_text((x0 + 10, y + 11.5), "MES", fontsize=8, fontname="hebo", color=(1, 1, 1))
    w = fitz.get_text_length("PAGO", fontname="hebo", fontsize=8)
    page.insert_text((x0 + tw - 10 - w, y + 11.5), "PAGO", fontsize=8,
                     fontname="hebo", color=(1, 1, 1))
    y += rh
    for mes, monto, es_total in filas_tabla:
        fill = VERDE_BG if es_total else (1, 1, 1)
        font = "hebo" if es_total else "helv"
        page.draw_rect(fitz.Rect(x0, y, x0 + tw, y + rh), fill=fill, color=BORDE, width=0.6)
        page.insert_text((x0 + 10, y + 11.5), mes, fontsize=9, fontname=font, color=NEGRO)
        w = fitz.get_text_length(monto, fontname=font, fontsize=9)
        page.insert_text((x0 + tw - 10 - w, y + 11.5), monto, fontsize=9,
                         fontname=font, color=NEGRO)
        y += rh

    # --- sello CANCELADO + carita ---
    y += 26
    sello = "CANCELADO"
    sw = fitz.get_text_length(sello, fontname="hebo", fontsize=15) + 30
    total_w = sw + 14 + 40
    x0 = (PAGE_W - total_w) / 2
    page.draw_rect(fitz.Rect(x0, y, x0 + sw, y + 30), color=VERDE, width=2.5)
    page.insert_text((x0 + 15, y + 20), sello, fontsize=15, fontname="hebo", color=VERDE)
    if imgs.get("carita"):
        page.insert_image(fitz.Rect(x0 + sw + 14, y - 5, x0 + sw + 14 + 40, y + 35),
                          stream=imgs["carita"])

    # --- pie ---
    page.draw_line((M, PAGE_H - M - 30), (PAGE_W - M, PAGE_H - M - 30),
                   color=BORDE, width=0.8)
    _txt_centrado(page, PAGE_H - M - 18,
                  f"Detalle según registro de seguimiento del pueblo al cierre {MES_CIERRE}.",
                  7, GRIS)
    _txt_centrado(page, PAGE_H - M - 8,
                  "Conserve esta constancia como comprobante de la cancelación de su medidor.",
                  7, GRIS)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    hoy = datetime.now().strftime("%d/%m/%Y")

    filas = _cargar_pagados()
    print(f"Medidor pagado: {len(filas)} constancias ({filas[0]['NRO']} al {filas[-1]['NRO']})\n")

    imgs = {"logo": _img_reducida(LOGO_PATH) if LOGO_PATH.exists() else None,
            "carita": _img_reducida(CARITA_PATH) if CARITA_PATH.exists() else None}

    pdfs = []
    for i, f in enumerate(filas, 1):
        print(f"[{i}/{len(filas)}] {f['NRO']}  {f['MZ']}-{f['LT']}  {f['NOMBRE'][:35]}")
        doc = fitz.open()
        _dibujar_constancia(doc, f, hoy, imgs)
        out = OUT_DIR / f"RECIBO_{f['NRO'].replace('-', '_')}_{f['MZ']}_{f['LT']}.pdf"
        doc.save(str(out), garbage=3, deflate=True)
        # JPG para WhatsApp (mismo patrón que las boletas)
        img = f"{sanitize(f['MZ'])}_{sanitize(f['LT'])}_{sanitize(f['NOMBRE'] or 'SIN_NOMBRE')}.jpg"
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
        pix.save(str(IMG_DIR / img))
        doc.close()
        pdfs.append(out)

    # registro de lo emitido (para auditoría / reimpresión)
    reg = pd.DataFrame([{
        "NRO": f["NRO"], "MZ": f["MZ"], "LT": f["LT"], "NOMBRE": f["NOMBRE"],
        "DEUDA": f["DEUDA"], "TOTAL_PAGADO": f["TOTAL_PAGADO"],
        "PAGOS": " · ".join(f"{m} {_soles(v)}" for m, v in f["PAGOS"]),
    } for f in filas])
    reg.to_excel(OUT_DIR / "data_recibos_medidor_pagado.xlsx", index=False)

    merger = PdfMerger()
    for p in pdfs:
        merger.append(str(p))
    merger.write(str(OUT_DIR / "CONSOLIDADO_MEDIDOR_PAGADO.pdf"))
    merger.close()

    print(f"\n[OK] {len(pdfs)} constancias · CONSOLIDADO_MEDIDOR_PAGADO.pdf")


if __name__ == "__main__":
    main()


