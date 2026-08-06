# =============================================================
# boletas_sin_servicio.py — Boletas de DEUDA para predios sin servicio
#
# Predios con deuda de pueblo (MULTA / ACUERDOS / CONVENIO en
# shared/seguimiento_pueblo.xlsx) que NO reciben boleta del ciclo
# porque no tienen lectura (sin servicio). La deuda es real y la
# mesa necesita papel para cobrarla.
#
#   - Universo: saldo > 0 en el ledger al cierre del mes, fuera de
#     DATA_boletas. Se agrega agua vieja/corte del arrastre_consolidado
#     si existe (ej. E1-6, J-6).
#   - Sin mantenimiento (sin servicio activo no lo genera) y consumo 0.
#   - Recibos correlativos continuando la serie del ciclo (max + 1).
#   - EXCLUIDOS: lotes viejos de reasignaciones pendientes de génesis
#     (B-29→B-20, C-45→C-43) — su deuda se factura cuando se muevan
#     los cargos al lote nuevo.
#
# Outputs/Sin_servicio/: RECIBO_*.docx/.pdf, Imagenes/*.jpg,
#     CONSOLIDADO_SIN_SERVICIO.pdf, data_boletas_sin_servicio.xlsx
#
# Uso:  py boletas_sin_servicio.py   (desde 3_boletas/)
# =============================================================

import glob
import re
import sys
import unicodedata
from pathlib import Path

import fitz
import pandas as pd
from PyPDF2 import PdfMerger
from docx.shared import Mm
from docx2pdf import convert
from docxtpl import DocxTemplate, InlineImage

sys.path.insert(0, str(Path(__file__).parent.parent / "shared"))
import seguimiento_repo as repo  # noqa: E402

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "Inputs"
OUT_DIR = BASE_DIR / "Outputs" / "Sin_servicio"
IMG_DIR = OUT_DIR / "Imagenes"

DATA_BOLETAS_PATH = INPUT_DIR / "DATA_boletas.xlsx"
PLANTILLA_PATH = INPUT_DIR / "PLANTILLA_boletas.docx"

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
ARRASTRE_PATH = BASE_DIR.parent / "5_cobranza" / "outputs" / f"arrastre_consolidado_{MES_CIERRE}.xlsx"
NOMBRE_JAAS = "JUNTA ADMINISTRATIVA DE SERVICIOS DE SANEAMIENTO"
SECTOR = "P.J. TUPAC AMARU"
TELEFONO = "948 227 636"
ESTADO = "NO ESTÁ AL DÍA"

# lotes viejos de reasignaciones — deuda pendiente de mover al lote nuevo
EXCLUIDOS = {("B", "29"), ("C", "45")}


def sanitize(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^\w]", "_", s).strip("_")


def _num(v):
    v = float(v)
    return int(v) if v == int(v) else v


def _clave_orden(mz: str, lt: str) -> tuple:
    """Igual que enriquecimiento: MZ simples A-Z primero, compuestas después."""
    m = re.match(r"^([A-Z]+)(\d*)$", mz)
    letras, num = (m.group(1), m.group(2)) if m else (mz, "")
    mz_key = (1 if num else 0, letras, int(num) if num else 0)
    ml = re.match(r"^(\d+)\s*([A-Z]*)$", lt)
    lt_key = (0, int(ml.group(1)), ml.group(2)) if ml else (1, 0, lt)
    return mz_key + lt_key


def _cargar_universo():
    db = pd.read_excel(DATA_BOLETAS_PATH, sheet_name="Data", dtype=str)
    claves_ciclo = {(repo._norm(r["MZ"]), repo._norm(r["LT"])) for _, r in db.iterrows()}
    recibo_inicio = int(pd.to_numeric(db["NUMERO DE RECIBO"]).max()) + 1
    cfg = db.iloc[0]  # PERIODO/fechas/pago: mismas del ciclo

    deudores = {}
    for con in ("MULTA", "ACUERDOS", "CONVENIO"):
        for k, v in repo.get_saldos_bulk(con, MES_CIERRE).items():
            if v > 0.005 and k not in claves_ciclo and k not in EXCLUIDOS:
                deudores.setdefault(k, {})[con] = v

    # agua vieja / corte del arrastre (solo para los que ya están en el universo)
    def _f(v):
        v = pd.to_numeric(v, errors="coerce")   # celdas "—" del arrastre → NaN → 0
        return 0.0 if pd.isna(v) else float(v)

    if ARRASTRE_PATH.exists():
        arr = pd.read_excel(ARRASTRE_PATH, header=1)
        for _, r in arr.iterrows():
            k = (repo._norm(r["MZ"]), repo._norm(r["LT"]))
            if k in deudores:
                agua, corte = _f(r["DEUDA_AGUA"]), _f(r["CORTE_RECONEXION"])
                if agua > 0:
                    deudores[k]["AGUA"] = agua
                if corte > 0:
                    deudores[k]["CORTE"] = corte

    nombres = repo._lookup_nombres()
    filas = []
    for (mz, lt) in sorted(deudores, key=lambda k: _clave_orden(*k)):
        d = deudores[(mz, lt)]
        total = sum(d.values())
        filas.append({
            "MZ": mz, "LT": lt, "NOMBRES": nombres.get((mz, lt), ""),
            "AGUA": d.get("AGUA", 0.0), "CORTE": d.get("CORTE", 0.0),
            "MULTA": d.get("MULTA", 0.0), "ACUERDOS": d.get("ACUERDOS", 0.0),
            "CONVENIO": d.get("CONVENIO", 0.0), "TOTAL": total,
        })
    for i, f in enumerate(filas):
        f["NUMERO DE RECIBO"] = recibo_inicio + i
    return filas, cfg


def _generar_recibo(fila, cfg, doc_tpl_path):
    doc = DocxTemplate(doc_tpl_path)
    logo = InlineImage(doc, str(INPUT_DIR / "logo_jaas.png"), width=Mm(31))
    qr = InlineImage(doc, str(INPUT_DIR / "imagen_qr.png"), width=Mm(100))
    carita = InlineImage(doc, str(INPUT_DIR / "carita_triste.png"), width=Mm(26))

    context = {
        "nu_reci": fila["NUMERO DE RECIBO"],
        "nombre_usuario": fila["NOMBRES"],
        "direccion_usuario": f"Mz. {fila['MZ']} Lt. {fila['LT']}",
        "an": 0, "ac": 0, "c": 0, "tmac": 0,
        "tman": _num(fila["AGUA"]),
        "c_rec": _num(fila["CORTE"]),
        "con": _num(fila["CONVENIO"]),
        "man": 0,                      # sin servicio: no genera mantenimiento
        "mul": _num(fila["MULTA"]),
        "cd": _num(fila["ACUERDOS"]),
        "ip": _num(fila["TOTAL"]),
        "ep": ESTADO,
        "icono_estado": carita,
        "imagen_qr": qr,
        "nombre_jaas": NOMBRE_JAAS,
        "sector": SECTOR,
        "lectura_anterior": str(cfg["LECTURA ANTERIOR"]).strip(),
        "lectura_actual": str(cfg["LECTURA ACTUAL"]).strip(),
        "periodo": str(cfg["PERIODO"]).strip(),
        "fv": str(cfg["FECHA DE VENCIMIENTO"]).strip(),
        "fe": str(cfg["FECHA DE EMISIÓN"]).strip(),
        "fecha_pago": str(cfg["FECHA_PAGO"]).strip(),
        "hora_pago": str(cfg["HORA_PAGO"]).strip(),
        "telefono": TELEFONO,
        "logo_jaas": logo,
    }
    doc.render(context)

    mz_f = str(fila["MZ"]).strip().replace(" ", "")
    lt_f = str(fila["LT"]).strip().replace(" ", "")
    base = f"RECIBO_{fila['NUMERO DE RECIBO']}_{mz_f}_{lt_f}"
    out_docx = OUT_DIR / f"{base}.docx"
    out_pdf = OUT_DIR / f"{base}.pdf"
    doc.save(out_docx)
    convert(str(out_docx), str(out_pdf))

    img = f"{sanitize(mz_f)}_{sanitize(lt_f)}_{sanitize(fila['NOMBRES'] or 'SIN_NOMBRE')}.jpg"
    pdf_doc = fitz.open(str(out_pdf))
    pix = pdf_doc[0].get_pixmap(matrix=fitz.Matrix(150 / 72, 150 / 72))
    pix.save(str(IMG_DIR / img))
    pdf_doc.close()
    return out_pdf


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    IMG_DIR.mkdir(parents=True, exist_ok=True)

    filas, cfg = _cargar_universo()
    total = sum(f["TOTAL"] for f in filas)
    print(f"Boletas sin servicio: {len(filas)} predios · deuda total S/ {total:,.2f}")
    print(f"Recibos {filas[0]['NUMERO DE RECIBO']} al {filas[-1]['NUMERO DE RECIBO']}\n")

    ok = 0
    for i, fila in enumerate(filas, 1):
        print(f"[{i}/{len(filas)}] Recibo {fila['NUMERO DE RECIBO']}  "
              f"{fila['MZ']}-{fila['LT']}  {fila['NOMBRES'][:35]}")
        try:
            _generar_recibo(fila, cfg, PLANTILLA_PATH)
            ok += 1
        except Exception as e:
            print(f"    [ERROR] {e}")

    # registro de lo emitido (para auditoría / reimpresión)
    pd.DataFrame(filas).to_excel(OUT_DIR / "data_boletas_sin_servicio.xlsx", index=False)

    pdfs = sorted(OUT_DIR.glob("RECIBO_*.pdf"))
    merger = PdfMerger()
    for p in pdfs:
        merger.append(str(p))
    merger.write(str(OUT_DIR / "CONSOLIDADO_SIN_SERVICIO.pdf"))
    merger.close()

    print(f"\n[OK] {ok}/{len(filas)} boletas · CONSOLIDADO_SIN_SERVICIO.pdf ({len(pdfs)} recibos)")


if __name__ == "__main__":
    main()
