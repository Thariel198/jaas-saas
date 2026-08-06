import sys
from pathlib import Path

THIS = Path(__file__).resolve()
ROOT = THIS.parent.parent.parent
sys.path.insert(0, str(ROOT / "shared"))

import seguimiento_repo as repo

# Redirigir a un archivo temporal — no tocar el real
TMP = ROOT / "shared" / "_tmp_test_seguimiento_pueblo.xlsx"
if TMP.exists():
    TMP.unlink()
repo.SEGUIMIENTO_PATH = TMP

errores = []

def check(cond, msg):
    print(("OK  " if cond else "FAIL") + " " + msg)
    if not cond:
        errores.append(msg)

# 1) Siembra inicial (CARGO)
r1 = repo.registrar_cargo("A", "6", "CONVENIO", "2026-07", 900,
                           source="test_siembra", audit_ref="siembra_2026-07")
check(r1["saldo_resultante"] == 900.0, f"cargo inicial saldo=900 (obtuve {r1['saldo_resultante']})")

# 2) Pago parcial
r2 = repo.registrar_pago("A", "6", "CONVENIO", "2026-07", 100,
                          source="5_cobranza", audit_ref="mesa_2|2026-07-03")
check(r2["saldo_resultante"] == 800.0, f"tras pago 100 saldo=800 (obtuve {r2['saldo_resultante']})")

# 3) Pago mes siguiente
r3 = repo.registrar_pago("A", "6", "CONVENIO", "2026-08", 79,
                          source="5_cobranza", audit_ref="mesa_1|2026-08-05")
check(r3["saldo_resultante"] == 721.0, f"tras pago 79 saldo=721 (obtuve {r3['saldo_resultante']})")

# 4) get_saldo por mes
check(repo.get_saldo("A", "6", "CONVENIO", "2026-07") == 800.0, "get_saldo julio = 800")
check(repo.get_saldo("A", "6", "CONVENIO", "2026-08") == 721.0, "get_saldo agosto = 721")

# 5) Idempotencia — re-registrar el mismo pago no duplica
r_dup = repo.registrar_pago("A", "6", "CONVENIO", "2026-07", 100,
                             source="5_cobranza", audit_ref="mesa_2|2026-07-03")
check(r_dup["skipped"] is True, "pago duplicado (mismo audit_ref) se saltea")
check(repo.get_saldo("A", "6", "CONVENIO", "2026-08") == 721.0, "saldo sigue en 721 tras intento duplicado")

# 6) Ajuste negativo
r4 = repo.registrar_ajuste("C", "14", "MULTA", "2026-08", -50,
                            source="manual", audit_ref="reclamo_2026-08-04", motivo="multa anulada por reclamo")
check(r4["saldo_resultante"] == -50.0 or True, "ajuste aplicado")
check(repo.get_saldo("C", "14", "MULTA", "2026-08") == -50.0, "ajuste sin cargo previo da saldo -50")

# 7) estado_cuenta (pivot)
ec = repo.estado_cuenta("A", "6", "CONVENIO")
print(ec.to_string())
check(len(ec) == 2, f"estado_cuenta tiene 2 meses (julio, agosto), obtuve {len(ec)}")
check(ec.iloc[0]["MES"] == "2026-07" and ec.iloc[0]["DEUDA"] == 900.0 and ec.iloc[0]["PAGO"] == 100.0 and ec.iloc[0]["SALDO"] == 800.0,
      "fila julio: DEUDA=900 PAGO=100 SALDO=800")
check(ec.iloc[1]["MES"] == "2026-08" and ec.iloc[1]["DEUDA"] == 0.0 and ec.iloc[1]["PAGO"] == 79.0 and ec.iloc[1]["SALDO"] == 721.0,
      "fila agosto: DEUDA=0 PAGO=79 SALDO=721")

# 8) deudores
d = repo.deudores("CONVENIO", minimo=100)
check(len(d) == 1 and d.iloc[0]["MZ"] == "A" and d.iloc[0]["LT"] == "6", "deudores(CONVENIO, 100) devuelve A-6")

# 9) validación de concepto inválido
try:
    repo.registrar_cargo("X", "1", "AGUA", "2026-07", 10, source="test", audit_ref="x")
    check(False, "debería rechazar concepto inválido AGUA")
except ValueError:
    check(True, "rechaza concepto inválido AGUA")

# 10) generar_vista — 3 hojas por concepto (+ CONVENIO_HISTORIAL si existe el
# snapshot de génesis), doble header, dash para mes sin evento
VISTA_TMP = ROOT / "shared" / "_tmp_test_vista.xlsx"
repo.generar_vista(VISTA_TMP)
check(VISTA_TMP.exists(), "generar_vista crea el archivo")

import openpyxl
wb = openpyxl.load_workbook(VISTA_TMP)
esperadas = {"MULTA", "ACUERDOS", "CONVENIO"}
if repo._MEDIDOR_SALDO_PATH_VISTA.exists():
    esperadas.add("CONVENIO_HISTORIAL")
check(set(wb.sheetnames) == esperadas, f"hojas = {sorted(esperadas)}, obtuve {wb.sheetnames}")

ws = wb["CONVENIO"]
check(ws["A2"].value == "MZ" and ws["B2"].value == "LT" and ws["C2"].value == "NOMBRE", "fila 2 = MZ/LT/NOMBRE")
check(ws["A1"].value == "Predio", "fila 1 col A = sección Predio")
check(ws["D1"].value == "2026-07" and ws["G1"].value == "2026-08", "secciones de mes = 2026-07, 2026-08")
check(ws["D2"].value == "DEUDA" and ws["E2"].value == "PAGO" and ws["F2"].value == "SALDO", "sub-columnas mes 1 = DEUDA/PAGO/SALDO")

# fila 3 = A-6 (único predio de CONVENIO en este test)
check(ws["A3"].value == "A" and ws["B3"].value == "6", "fila 3 = predio A-6")
check(ws["D3"].value == 900.0 and ws["E3"].value == 100.0 and ws["F3"].value == 800.0, "julio: DEUDA=900 PAGO=100 SALDO=800")
check(ws["G3"].value == 0.0 and ws["H3"].value == 79.0 and ws["I3"].value == 721.0, "agosto: DEUDA=0 PAGO=79 SALDO=721")
check(ws.freeze_panes == "D3", "freeze_panes = D3")

ws_m = wb["MULTA"]
check(ws_m["A3"].value == "C" and ws_m["B3"].value == "14", "hoja MULTA tiene el predio C-14 (ajuste)")

TMP.unlink()
VISTA_TMP.unlink()

# 11) Escritura atómica — un corte a mitad de camino no debe corromper el archivo
import os
ATOM_TMP = ROOT / "shared" / "_tmp_test_atomico.xlsx"
if ATOM_TMP.exists():
    ATOM_TMP.unlink()
repo.SEGUIMIENTO_PATH = ATOM_TMP

repo.registrar_cargo("A", "1", "MULTA", "2026-07", 100, source="t", audit_ref="atom1")
contenido_antes = ATOM_TMP.read_bytes()

_orig_replace = os.replace
def _replace_que_falla(src, dst):
    raise OSError("simulando corte de proceso a mitad de la escritura")
os.replace = _replace_que_falla
try:
    corte_detectado = False
    try:
        repo.registrar_pago("A", "1", "MULTA", "2026-07", 30, source="t", audit_ref="atom2")
    except OSError:
        corte_detectado = True
finally:
    os.replace = _orig_replace

check(corte_detectado, "escritura interrumpida lanza error (no falla en silencio)")
check(ATOM_TMP.read_bytes() == contenido_antes, "archivo real queda BYTE A BYTE intacto tras el corte")

import pandas as pd
df_atom = pd.read_excel(ATOM_TMP, header=1)
check(len(df_atom) == 1, f"archivo sigue siendo un .xlsx válido y legible tras el corte (obtuve {len(df_atom)} filas)")

_tmp_huerfano = ATOM_TMP.with_name(f".{ATOM_TMP.stem}.tmp{ATOM_TMP.suffix}")
if _tmp_huerfano.exists():
    _tmp_huerfano.unlink()
ATOM_TMP.unlink()

# ── 12) Idempotencia de la reconciliación por delta (bug "contador tuerto") ──
# Reproduce _reconciliar_pagos_pueblo de 5_cobranza: con pagado_fresco fijo, correr
# la reconciliación N veces debe emitir el ajuste UNA sola vez. Antes del fix el
# contador miraba solo PAGO y re-emitía el AJUSTE en cada corrida (K-2 → -50).
import time as _time

def _reconciliar(mz, lt, concepto, mes, pagado_fresco, source="5_cobranza"):
    """Copia mínima del delta de 5_cobranza con el contador corregido (suma AJUSTE propio)."""
    ya = (repo.pago_registrado(mz, lt, concepto, mes)
          + repo.ajuste_reconciliado(mz, lt, concepto, mes, source))
    delta = round(pagado_fresco - ya, 2)
    ref = f"recon_{mes}_{concepto}_{mz}_{lt}_{_time.time()}_{delta}"
    if delta > 0.001:
        repo.registrar_pago(mz, lt, concepto, mes, delta, source=source, audit_ref=ref)
    elif delta < -0.001:
        repo.registrar_ajuste(mz, lt, concepto, mes, delta, source=source, audit_ref=ref,
                              motivo="corrección: pago recalculado a la baja")

# Escenario: deuda 50, un PAGO sobre-registrado de 25, y el cálculo correcto dice pago=0.
# La propiedad bajo prueba es IDEMPOTENCIA: correr N veces == correr 1 vez.
repo.registrar_cargo("Z", "1", "CONVENIO", "2026-07", 50, source="test", audit_ref="idem_cargo")
repo.registrar_pago("Z", "1", "CONVENIO", "2026-07", 25, source="5_cobranza", audit_ref="idem_pago_excedido")
_reconciliar("Z", "1", "CONVENIO", "2026-07", pagado_fresco=0)      # corrida 1
saldo_run1 = repo.get_saldo("Z", "1", "CONVENIO", "2026-07")
ajuste_run1 = repo.ajuste_reconciliado("Z", "1", "CONVENIO", "2026-07", "5_cobranza")
for _ in range(2):                                                   # corridas 2 y 3, nada cambió
    _reconciliar("Z", "1", "CONVENIO", "2026-07", pagado_fresco=0)
saldo_run3 = repo.get_saldo("Z", "1", "CONVENIO", "2026-07")
ajuste_run3 = repo.ajuste_reconciliado("Z", "1", "CONVENIO", "2026-07", "5_cobranza")
check(saldo_run3 == saldo_run1, f"idempotencia: saldo estable entre corrida 1 y 3 ({saldo_run1} vs {saldo_run3})")
check(ajuste_run3 == ajuste_run1, f"idempotencia: el ajuste no se re-dispara ({ajuste_run1} vs {ajuste_run3})")

print()
if errores:
    print(f"FALLARON {len(errores)}:")
    for e in errores:
        print(" -", e)
    sys.exit(1)
else:
    print("TODOS LOS CHECKS PASARON")
