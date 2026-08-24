# RETOMAR — Limpieza del ledger por bug de signo + mini-pipeline "simulado" (2026-08-19)

Documento de auditoría. Cubre TODO lo hecho en la sesión del 2026-08-19 sobre
`shared/seguimiento_pueblo.xlsx` (real y su copia `jass_system - simulado`):
qué se eliminó, por qué, con qué evidencia, y dónde está el respaldo de cada
paso para poder reconstruir el estado anterior si hace falta.

---

## 1. Causa raíz — el bug que originó todo

`5_cobranza/main.py::_reconciliar_pagos_pueblo()` recalcula cada corrida el
total de cascada por (predio, concepto, mes) y compara contra lo ya
registrado en el ledger. Si el nuevo total es **menor** al ya reconocido,
debe registrar un `AJUSTE` que **restaure** la deuda (signo positivo).

Hasta el **2026-08-12 (commit `bda176d`)** el código llamaba
`registrar_ajuste(..., delta, ...)` con el `delta` crudo (negativo) en vez de
`-delta`. Efecto: en vez de restaurar la deuda, la **descontaba una segunda
vez** — un "pago fantasma" que el propio sistema generó al recalcular, más un
ajuste que en vez de anularlo la duplicó en sentido contrario.

El fix de código (12/08) **no fue retroactivo** — todos los eventos con esa
firma, escritos ANTES del fix, quedaron con el error permanentemente
grabado en el ledger histórico. Esta sesión fue encontrar y corregir esos
eventos, uno por uno, con evidencia antes de tocar nada.

**Firma para detectar un evento roto:**
`TIPO_EVENTO=AJUSTE`, `CLASE=CORRECCION_SISTEMA`, `SOURCE=5_cobranza`,
`AJUSTE<0`, `TIMESTAMP < 2026-08-12 15:57:27`, precedido en el mismo
predio+concepto+mes por un `PAGO` de **magnitud exactamente igual**
(`|PAGO| == |AJUSTE|`). Ese PAGO es el "pago fantasma": no es plata real,
es el propio recálculo de 5_cobranza.

---

## 2. Mini-pipeline "jass_system - simulado"

Copia completa del repo en `C:\Users\wilde\PycharmProjects\jass_system - simulado`,
usada para correr `5_cobranza --force` en segundos (en vez de ~30 min) y
comparar resultados contra el ledger real sin tocar el repo real.

**Filtrado a los 39 predios** (el universo de "deuda escondida" que salió de
`auditar_pago_vs_ledger.deuda_escondida_ledger()`):
- `shared/seguimiento_pueblo.xlsx`, `shared/planilla_mes/planilla_2026-08.xlsx`
- `4_pagos/efectivo/outputs/pagos_efectivo.xlsx` y `pagos_efectivo_2026-08.xlsx`
- `4_pagos/yape/motor_matching/outputs/pagos_yape_tepago_2026-08.xlsx` y `pagos_yape_pagaste_2026-08.xlsx`
- `shared/blancos_acumulados.xlsx`, `6_corte/outputs/audit_penalidad.xlsx`, `5_cobranza/inputs/correcciones_lote.xlsx`
- Los 9 precursores de `shared/`: `abonos_rezagados.xlsx` (+ manifest JSON), `ajustes_cargo.xlsx`,
  `genesis_tardia.xlsx`, `reidentificacion.xlsx`, `deuda_correcciones.xlsx`, `blancos_efectivo.xlsx`,
  `aportes_tanque_manuales.xlsx`, `reasignaciones_aplicacion.xlsx`, `devoluciones_aplicadas.xlsx`

Cada archivo tiene su backup `.bak_pre_filtro39` (o similar) al lado, dentro
de la copia `- simulado`, con el contenido ANTES de filtrar.

**Validación:** con todo filtrado, `5_cobranza --force` corrió en **28
segundos** (antes 23 min con solo ledger/pagos filtrados, planilla completa).
Comparado el `SALDO` final de las 68 combinaciones (MZ,LT,CONCEPTO) de los 39
predios entre real y simulado → **0 diferencias**. Confirma que la
metodología de copia+filtro reproduce el pipeline real fielmente.

**Cambio de código (solo en la copia, NO en el repo real):**
`shared/seguimiento_repo.py` → `VISTA_PATH` y `VISTA_PDF_PATH` apuntan a
`vista_seguimiento_pueblo_simulado.xlsx/.pdf` en vez del nombre real, para
no pisar ni confundir con la vista del repo real.

---

## 3. Limpieza del bug de signo, por concepto

Para cada caso: se verificó que el `PAGO` "fantasma" no vino de ningún
archivo real de pagos (`pagos_efectivo*.xlsx`, `pagos_yape*.xlsx`,
`trazabilidad_cobranza.xlsx`) antes de borrar. Se excluyó todo predio que
ya tuviera una corrección manual humana posterior (para no duplicar ni
pisar una decisión de negocio ya tomada).

### CONVENIO — 9 lotes limpiados (24 filas eliminadas en real y en simulado)

| Predio | Saldo con bug | Saldo corregido |
|---|---|---|
| A-8 | -50 | **50** |
| B-5 | -50 | **50** |
| C-1 | -50 | **50** |
| C-7 | -25 | **25** |
| E-12 | -16 | **26** |
| I-11 | -25 | **25** |
| J-3 | -30 | **50** |
| K-17 | -25 | **25** |
| K-2 | -25 | **25** |

**Excluidos y por qué:**
- **A-4** — ya resuelto a mano el 27/07 (`fix_race_condition_yape_20260713_A_4_CONVENIO`, +225), un bug distinto (yape duplicado por race condition). No se tocó.
- **G-18, T-7** — el "pago fantasma" ahí era en realidad un pago declarado por la secretaria (evidencia de negocio real), no un artefacto del bug. Instrucción explícita del usuario de excluirlos.

### ACUERDOS — 6 lotes limpiados (12 filas eliminadas en real y en simulado)

| Predio | Saldo con bug | Saldo corregido |
|---|---|---|
| B-5 | -25 | **25** |
| C-1 | -25 | **25** |
| D-16 | 0 | **50** |
| H-16 | 47 | **75** |
| I-16 | 47 | **75** |
| W-5 | 37 | **47** |

D-16 es el caso que arrancó toda la investigación (el usuario detectó que
la deuda de junio bajó de 50 a 25 sin que hubiera pago real).

**Excluidos** — 12 predios ya con condonación manual humana (30-31/07,
`...-condonacion-estable-30072026` o `...-revertir-condonacion-fallida`):
B1-12, F1-5, G1-14, J-6, L-16, Q-11, Q-4, R-4, S-1, S-5, S-9, T-7.

### MULTA — 2 lotes limpiados (4 filas eliminadas en real y en simulado)

| Predio | Saldo con bug | Saldo corregido |
|---|---|---|
| F-4 | 48 | **50** |
| I-16 | -18 | **18** |

**Excluidos** — C-21 y J-6 ya resueltos a mano el 30/07 con un ajuste
"revertir-condonación-fallida" (+100) que los devolvió al valor correcto (50).

### Total de deuda restaurada por el bug (no era "ruido", era plata real que faltaba cobrar)

CONVENIO 622 + ACUERDOS 216 + MULTA 38 = **S/876** de deuda que estaba
oculta por el bug y ahora vuelve a aparecer correctamente en el ledger.

---

## 4. Lotes fantasma — eliminados por completo (no solo el bug, todo el rastro)

A diferencia de la sección anterior (donde el predio es real y solo se
corrigió un evento puntual), estos 4 códigos de lote **no existen** como
predios reales — se eliminó su historial entero, todos los conceptos.

| Lote | Evidencia | Filas eliminadas (real / simulado) |
|---|---|---|
| **C1-17** | `5_cobranza/inputs/correcciones_lote.xlsx` fila 5: mapeo explícito y documentado C1-17→C1-9 ("la regla auto-recuperada tenía el sentido al revés"), con guarda en `main.py:1348` que bloquea la reversión. | 2 / 0 (no estaba en los 39) |
| **S-14** | Duplicado exacto de **S-13** (Ana Garro Rojas, real, confirmada en `DATA_boletas.xlsx`). Sembrados con el MISMO monto (50) un segundo aparte (18:19:07 vs 18:19:08) en la misma corrida de siembra. S-13 se pagó igual ese mismo día; S-14 quedó flotando hasta que un humano lo marcó "duplicado" el 31/07. | 4 (MULTA+ACUERDOS) / 0 |
| **C-29A** | No existe en `padron_reconciliado.xlsx` (el padrón consolidado). El dueño registrado, "Eddy Santiago Garcilazo Trujillo", no aparece en `DATA_boletas.xlsx` bajo ningún MZ/LT. Sin gemelo identificado, pero sin sustento como predio real. | 4 (MULTA) / 0 |
| **Q-16** | Igual que C-29A: no existe en `padron_reconciliado.xlsx`. Sin gemelo identificado. | 4 (ACUERDOS) / 0 |

**No se tocó E-14A** (investigado por el mismo motivo, pero es un caso
distinto): tiene una reidentificación completa y documentada hacia **E-14B**
(`reidentificacion_cargo|E-14A-a-E-14B|ACUERDOS|origen/destino`, ambos lados
presentes, mismo día). E-14 y E-14B SÍ existen en el padrón consolidado;
E-14A no — confirma que la reidentificación fue correcta y no requiere
ninguna acción.

---

## 5. J-1 y J-6 — predios comunitarios (con una corrección a mitad de camino)

J-1 ("Comedor Popular Club de Madres") y J-6 son predios reales pero
comunitarios. El usuario pidió eliminar su deuda porque MULTA/ACUERDOS ya
estaban exonerados por la directiva (J-1) o por una condonación revertida
que no debía existir (J-6), y ninguno de los dos tuvo nunca un pago real
para esos 2 conceptos.

**Primer paso (excesivo):** se eliminó el historial COMPLETO de J-1 y J-6
— los 3 conceptos (MULTA, ACUERDOS, **y CONVENIO**). 14 filas en el real
(J-1: 5, J-6: 9), 9 filas en el simulado (solo J-6, J-1 no era parte de los
39).

**Error detectado:** CONVENIO sí era deuda real y activa (J-1: 75, J-6: 100)
— nunca estuvo exonerada ni condonada, no debía borrarse.

**Corrección aplicada:** se restauraron las filas de CONVENIO desde el
backup, con el CARGO y TIMESTAMP originales exactos:
- Real: J-1/CONVENIO (75, siembra 2026-06-siembra) y J-6/CONVENIO (100) — ambas restauradas.
- Simulado: solo J-6/CONVENIO (100) restaurada (J-1 no pertenece al set de 39).

**Verificación del pago real de J-1** (con `reporte_historico.py`): hubo
un único yape real de S/44 (21/06/2026, "Comedor Popular Club de Madres")
que se consumió ÍNTEGRO en el agua de julio (consumo+mantenimiento+mes
anterior = 44 exacto, pago completo). Nunca tocó MULTA, ACUERDOS ni
CONVENIO — confirma que la deuda eliminada en esos 2 conceptos (MULTA,
ACUERDOS) nunca tuvo plata real detrás.

**Estado final:**
- J-1: MULTA=0 (exonerado, sin tocar), ACUERDOS=0 (exonerado, sin tocar), **CONVENIO=75 (restaurado)**.
- J-6: MULTA=0 (eliminado, sin pago real), ACUERDOS=0 (eliminado, sin pago real), **CONVENIO=100 (restaurado)**.

---

## 6. Todos los backups creados (para reconstruir cualquier paso)

Todos en `shared/` de cada repo (real y `- simulado`), junto al archivo
original, con el contenido de `seguimiento_pueblo.xlsx` **inmediatamente
antes** de cada operación:

1. `seguimiento_pueblo.xlsx.bak_pre_filtro39` — antes de filtrar el simulado a los 39 predios (solo simulado).
2. `seguimiento_pueblo.xlsx.bak_pre_limpieza_bug_convenio_<timestamp>` — antes de limpiar los 9 lotes de CONVENIO.
3. `seguimiento_pueblo.xlsx.bak_pre_limpieza_bug_acuerdos_<timestamp>` — antes de limpiar los 6 lotes de ACUERDOS.
4. `seguimiento_pueblo.xlsx.bak_pre_limpieza_bug_multa_<timestamp>` — antes de limpiar F-4/I-16 en MULTA.
5. `seguimiento_pueblo.xlsx.bak_pre_eliminar_lotes_fantasma_<timestamp>` — antes de eliminar C1-17/S-14.
6. `seguimiento_pueblo.xlsx.bak_pre_eliminar_c29a_q16_20260819.xlsx` — antes de eliminar C-29A/Q-16.
7. `seguimiento_pueblo.xlsx.bak_pre_eliminar_j1_j6_20260819.xlsx` — antes de eliminar J-1/J-6 completo (**este backup es el que se usó para restaurar CONVENIO**).
8. `seguimiento_pueblo.xlsx.bak_pre_restaurar_convenio_j1_j6_20260819.xlsx` (real) / `..._j6_...` (simulado) — justo antes de reinsertar las filas de CONVENIO.

Cada archivo `.xlsx` filtrado del entorno simulado (pagos, planilla,
precursores) tiene también su propio `.bak_pre_filtro39` al lado.

---

## 7. Herramientas usadas (todas ya existentes, ninguna nueva creada esta sesión)

- `4b_reclamos/herramienta/auditar_pago_vs_ledger.py` — detección inicial de
  "deuda escondida" (48 eventos / 39 predios). Su clasificación
  `YA_ESTABILIZADO` tiene una limitación conocida: no filtra por MES al
  buscar un ajuste posterior que compense uno roto, por lo que puede dar
  falsos positivos si otro mes tiene, por coincidencia, un ajuste de igual
  magnitud (encontrado en A-8/B-5/E-12 — **no corregido en el código esta
  sesión**, quedó pendiente).
- `4b_reclamos/herramienta/reporte_historico.py` — verificación de
  cascada agua+deuda por predio (usado para confirmar el destino del yape de J-1).
- `shared/seguimiento_repo.py::generar_vista()` / `exportar_vista_pdf()` —
  regeneración de `vista_seguimiento_pueblo.xlsx/.pdf` después de cada
  cambio al ledger.

---

## 8. Pendientes / no resuelto en esta sesión

- **C-29A**: eliminado por no existir en el padrón consolidado, pero sin
  identificar a qué predio real (si alguno) pertenecía originalmente su
  deuda de multa (S/20). A diferencia de S-14, no se encontró el "gemelo".
- **Heurística `YA_ESTABILIZADO`** de `auditar_pago_vs_ledger.py` no quedó
  corregida — sigue sin filtrar por MES. Si se vuelve a correr
  `deuda_escondida_ledger()`, puede reportar como "resuelto" un caso que
  en realidad no lo está (como pasó con A-8/B-5/E-12 antes de investigarlos
  a mano).
- No se revisó el concepto **CONVENIO** en busca del mismo patrón de lotes
  fantasma (C1-17/S-14 style) fuera de los ya encontrados — la búsqueda de
  "gemelos" se hizo solo para los casos que el usuario señaló puntualmente.
