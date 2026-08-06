# RETOMAR — Reclamos julio + bug idempotencia de ajustes (2026-07-09)

Sesión Opus. Se trabajó reclamos `2026-07` y se encontró/arregló un bug en la
reconciliación de `5_cobranza` hacia `seguimiento_pueblo`. Este doc deja todo el
contexto para continuar mañana.

---

## 1. BUG DE IDEMPOTENCIA EN LA RECONCILIACIÓN — RESUELTO HOY

### Qué pasaba
`5_cobranza/main.py · _reconciliar_pagos_pueblo()` reconcilia por delta hacia el
ledger append-only. El contador de "lo ya hecho" era `repo.pago_registrado()`, que
**suma solo eventos PAGO, no AJUSTE**. Cuando el branch `delta<0` emitía un AJUSTE,
la corrida siguiente no lo veía → re-emitía el mismo AJUSTE. Cada re-corrida restaba
otro −25.

- Disparado por: la corrección manual de convenio (−25, `source=correccion_genesis_formula`)
  que el usuario metió como evento en el ledger (génesis NUNCA se tocó — solo se
  agregó un AJUSTE). Eso hizo que 5_cobranza generara ajustes automáticos.
- Se corrió 2 veces (Jul-08 16:22 y Jul-09 05:31) → **2 tandas gemelas** →
  18 predios con saldo hundido (K-2 CONVENIO llegó a −50).

### Fix aplicado (código) — testeado
- `shared/seguimiento_repo.py`: **nueva función** `ajuste_reconciliado(mz,lt,concepto,mes,source)`
  — espejo de `pago_registrado` pero suma AJUSTE filtrando por `source`.
- `5_cobranza/main.py` (~línea 1789): contador ahora
  `ya = pago_registrado(...) + ajuste_reconciliado(..., "5_cobranza")`.
  (Fix mínimo, solo idempotencia. NO se tocó el signo del ajuste.)
- Test: `shared/tests/test_seguimiento_repo.py` — bloque nuevo "12) Idempotencia"
  prueba que 3 corridas = 1 corrida. **Todos los checks pasan.**

### "Bug 2" (signo del ajuste / pago fantasma) — DESCARTADO, NO EXISTE
Durante la sesión propuse un segundo bug (signo del ajuste invertido). Se verificó
con datos: `saldo_SIN_aj5cob ≥ 0` en TODOS los predios → nadie pagó de más → no hubo
pagos fantasma → ese branch nunca disparó legítimamente. Era sobreingeniería.
Razón de fondo (usuario): el arrastre fija la deuda ANTES de cobrar y el período de
cuotas ya venció → el escenario del re-reparto no ocurre. Se revirtió al fix mínimo.

### Remediación de datos — HECHA
- Backup: `shared/backups/seguimiento_pueblo_2026-07-09_163232_pre_remediacion_ajustes.xlsx`
- Se **borraron las 36 filas** `AJUSTE / source=5_cobranza` (artefactos de bug, todas
  eventos finales, verificado que no rompen cadena). Precedente: B-20 en el doc de
  decisión seguimiento_pueblo.
- Resultado: todos los saldos vuelven a ≥ 0 (K-2 CONVENIO = 0). Vista regenerada
  (`shared/vista_seguimiento_pueblo.xlsx`).
- ⚠ Se re-verificó que el repo lee correcto (K-2=0, A-8=0, J-3=10).

### ⚠ PENDIENTE de este bug
- **NO se re-corrió `5_cobranza`** (queda para cuando sigas el ciclo, ya con el fix).
  Ahora es seguro correrlo: ledger limpio + código arreglado.
- Regresión: falta correr `2_planilla` (lee `get_saldo`) para confirmar que consume
  bien los saldos corregidos.
- Aprendizaje: `docs/aprendizaje/Aprendizaje html/contador_tuerto_idempotencia_ajuste_20260709.html`
  (ya corregido a solo bug 1).

---

## 2. RECLAMOS 2026-07 — CLASIFICACIÓN Y RECHAZOS

Archivo: `4b_reclamos/outputs/reclamos_2026-07.xlsx` (54 reclamos).

### Regla de rechazo para reclamos de CONVENIO/MEDIDOR
```
RECHAZAR  =  (1) es reclamo de CONVENIO  → el predio estaba en las 109 correcciones
                                            (lote 3_boletas/Outputs/Correcciones)
         AND (2) su convenio en el ledger (seguimiento_pueblo) = 0
```
- El error del viernes fue **convenio sobrecobrado** (fórmula Saldo de génesis omitía
  pagos de abril). Se reimprimieron 109 boletas (`correcciones.py`).
- Convenio = deuda MEDIDOR + deuda INSTALACIÓN. Si el residual es medidor real, NO se
  rechaza (el reclamo puede ser válido por pago no acreditado).

### Estado actual: 17 RECHAZADO
- Sin-tipo con convenio ledger=0 (estaban en las 109): 10 + H-11 + M-19 + **K-2**
  (K-2 se cerró tras la remediación, su −50 era el bug).
- Tipo=convenio con ledger=0: **M-21, G-21, L-2, B-14**.

### ⚠ M-15 quedó SIN marcar (el .xlsx estaba abierto en Excel al intentar)
- M-15: exceso −20, sin deuda, NO era convenio (nunca tuvo medidor/instalación).
  Regla nueva confirmada por el usuario: **exceso + sin deuda → RECHAZADO**
  (el exceso NO se pierde, va a `arrastre_devolucion`). **Marcar M-15 = RECHAZADO.**

### Exceso — se maneja en arrastre_devolucion (no en reclamos)
`5_cobranza/outputs/arrastre_devolucion_2026-07.xlsx` lleva los excesos. 5 reclamos
con exceso: M-15(20), F-8(25), H-11(25), I-17(25), O-25(33) — los 4 últimos ya
RECHAZADO. **Ese archivo tiene la columna COMENTARIO con el motivo real** del reclamo
(ej. M-15 "Cambiar la multa de G1-M15").

---

## 3. PENDIENTE PARA MAÑANA

### A) La hoja de reclamos está "ciega" (sin comentario) — PLOMERÍA
- El campo RECLAMO se auto-puebla del COMENTARIO de origen
  (`4b_reclamos/main.py:214`), leyendo `4_pagos/efectivo/outputs/pagos_efectivo.xlsx`.
- Ese origen SÍ tiene comentarios ricos para muchos predios (ej. O-22 "Ya pague
  convenio, pago solo consumo"; C1-9 "Cambiar mz C1-17 a C1-9"; A1-13; M-22 "Error
  en medidor"). Pero **20 de los 25 sin-tipo salen terse ("reclamo")** o no aparecen
  en esa fuente → vinieron por otra vía (arrastre/manual).
- TAREA: rastrear de dónde salieron esos 25 sin comentario y por qué no heredaron el
  texto. Cruzar con `pagos_efectivo` (comentarios ricos) y con `arrastre_devolucion`
  (comentarios de los con exceso) para poblar el motivo.
- NOTA usuario: pensó que estaban vacíos porque vio el comentario del mes pasado.
  El comentario SÍ existe en origen para muchos.

### B) Reclamos sin-tipo pendientes (24, tras marcar M-15)
Sin texto no se pueden clasificar objetivo. Con deuda de MEDIDOR real (revisar pago
no acreditado): **B-6, D-1, D-6, G-12, G-17, L-3, L-6, N-3, W-5**. El resto son
disputas de faena/acuerdos/agua o casi-al-día — necesitan el motivo (ver punto A).

### C) Reclamos con tipo pendientes (EN_REVISION)
- **C-20**: debe convenio 10, reclamo "ya pagué convenio pero me figura 10" — disputa
  ese saldo, verificar pago.
- **K-9** (2 filas): debe convenio 75, "ya cancelé medidor" — verificar.
- **B1-13 y B1-14**: piden no ser cobrados campo/faena por estar el lote FUERA del
  pueblo. NO pagaron este mes (siguen PENDIENTE). Es **decisión de la directiva**:
  ¿exentos de faena (multa 30 c/u) por estar fuera? Si sí → anular MULTA (AJUSTE en
  ledger) + ACEPTADO; si no → RECHAZADO.

### D) Arrastres — OK, no es error
16 reclamos tienen `MES_ANO_ORIGEN=2026-06` (arrastres de junio sin resolver).
Es intencional (`_cargar_arrastres`). Cuando se les pone ESTADO resuelto, dejan de
arrastrarse.

---

## Archivos tocados hoy
- `shared/seguimiento_repo.py` (nueva func `ajuste_reconciliado`)
- `5_cobranza/main.py` (contador de reconciliación)
- `shared/tests/test_seguimiento_repo.py` (test idempotencia)
- `shared/seguimiento_pueblo.xlsx` (remediado: −36 filas) + backup
- `shared/vista_seguimiento_pueblo.xlsx` (regenerada)
- `4b_reclamos/outputs/reclamos_2026-07.xlsx` (17 RECHAZADO; falta M-15)
- `docs/aprendizaje/Aprendizaje html/contador_tuerto_idempotencia_ajuste_20260709.html` (nuevo)
