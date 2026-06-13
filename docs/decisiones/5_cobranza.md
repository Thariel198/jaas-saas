# Decisión — Módulo 5_cobranza: diseño de carga de pagos + lista de corte

**Fecha:** 2026-06-10
**Módulo afectado:** `5_cobranza`

---

## Problema

Después de correr `4_pagos`, los pagos identificados (yape + efectivo) necesitan cargarse
en la planilla del mes para producir el estado final de cobro. El módulo debe re-correr
sin perder auditoría (llegan pagos nuevos vía yape durante el periodo de gracia), generar
la lista de usuarios a cortar, y producir el arrastre de deuda para el ciclo siguiente.

## Decisiones adoptadas

### D1 — Modelo B: 5_cobranza lee planilla + pagos separados

`4_pagos` genera dos archivos independientes (`pagos_yape_tepago.xlsx`,
`pagos_efectivo.xlsx`). `5_cobranza` los lee junto con `planilla_YYYY-MM.xlsx`
de `2_planilla` y genera `planilla_cobrado.xlsx` como copia enriquecida.

La planilla original de `2_planilla` **nunca se modifica**.

### D2 — planilla_cobrado.xlsx = planilla 21 cols + CICLO_COBRANZA

`planilla_cobrado.xlsx` tiene exactamente las mismas 21 columnas de la planilla
(`MZ…FECHA_PAGO`) más `CICLO_COBRANZA` como columna 22, agregada por `5_cobranza`.

`5_cobranza` rellena: `MONTO_YAPE`, `MONTO_EFECTIVO`, `ESTADO`, `FECHA_PAGO`,
y agrega `CICLO_COBRANZA`.

### D3 — 5_cobranza genera arrastre_deuda_YYYY-MM.xlsx

`5_cobranza` es quien tiene el `SALDO` en memoria, por lo que genera el arrastre
directamente. `5b_validacion` es un observador puro (solo compara totales) y no
genera datos.

Columnas de `arrastre_deuda_YYYY-MM.xlsx`: `MZ`, `LT`, `NOMBRE`, `monto` (= SALDO),
`MES_ANO_ORIGEN`.

Solo se incluyen usuarios con `SALDO > 0`. Usuarios con `SALDO ≤ 0` no aparecen.

### D4 — Condición de lista_corte: SALDO > 0 AND ARRASTRE ≥ 8

Un usuario va a `lista_corte.xlsx` si cumple ambas condiciones:
- `SALDO > 0` — tiene deuda pendiente este mes
- `ARRASTRE ≥ 8` — tenía al menos S/8 de arrastre del mes anterior

> S/8 = mínimo mensual (mantenimiento S/3 + mínimo agua S/5). Arrastre ≥ 8 confirma
> que no hubo ningún pago el mes anterior. Arrastre menor puede ser saldo residual.

Usuarios con deuda nueva (primer mes sin pagar, ARRASTRE = 0) aparecen como
PARCIAL/PENDIENTE en planilla_cobrado pero **no** van a lista_corte.

Penalidad por reconexión: S/20. `TOTAL_A_PAGAR = SALDO + 20`.

### D5 — Estados CANCELADO / EXCESO / PARCIAL / PENDIENTE

| Estado | Condición |
|---|---|
| CANCELADO | saldo = 0 (tolerancia ±0.005) |
| EXCESO | saldo < 0 — pagó de más; se registra, no se devuelve hasta reclamo |
| PARCIAL | saldo > 0 y pagó algo |
| PENDIENTE | saldo > 0 y no pagó nada |

### D6 — Estructura plana: un solo main.py sin sub-módulos

No hay sub-módulos (`carga_de_pago/`, `lista_de_corte/`). Un único `main.py`
genera los cinco outputs en una sola corrida.

Outputs generados por `5_cobranza/main.py`:
1. `planilla_cobrado.xlsx`
2. `lista_corte.xlsx`
3. `trazabilidad_cobranza.xlsx`
4. `resumen_recaudacion.xlsx`
5. `arrastre_deuda_YYYY-MM.xlsx`

### D7 — Dos contadores independientes: CICLO_CORRECCION y CICLO_COBRANZA

| Contador | Módulo | Mide |
|---|---|---|
| `CICLO_CORRECCION` | yape, efectivo | Rondas de corrección manual en 4_pagos |
| `CICLO_COBRANZA` | 5_cobranza | Rondas de carga en planilla_cobrado |

No están acoplados. El ciclo nuevo de cobranza se detecta comparando los pagos
entrantes contra `trazabilidad_cobranza.xlsx` (mismo patrón `max(CICLO)+1`).

`CICLO_COBRANZA` se escribe en:
- `planilla_cobrado.xlsx` col 22
- `trazabilidad_cobranza.xlsx` por fila
- `pagos_yape_tepago.xlsx` y `pagos_efectivo.xlsx` (retroescritura en inputs/)

`CICLO_CORRECCION_ORIGEN` en trazabilidad registra el ciclo de corrección en que
`4_pagos` identificó ese pago.

### D8 — 5b_validacion es observador puro

`5b_validacion` solo compara totales: suma de `MONTO_YAPE` vs total transferencias
Yape del banco; suma de `MONTO_EFECTIVO` vs total efectivo registrado. No genera
archivos de arrastre ni de deuda.

## Alternativas descartadas

| Alternativa | Por qué se descartó |
|---|---|
| Sub-módulos separados (carga_de_pago/, lista_de_corte/) | Complejidad innecesaria — el volumen de datos no lo justifica; un main.py es suficiente |
| Condición lista_corte: "2 meses consecutivos sin pagar" | Requiere historial mensual; ARRASTRE ≥ 8 es equivalente y se deriva directamente de la planilla sin estado adicional |
| 5b_validacion genera arrastre_deuda | 5b_validacion es observador puro; el SALDO ya está disponible en 5_cobranza, que es el módulo correcto para generarlo |
| Orquestador con estado compartido (cycle_state.json) | Introduce componente externo, rompe con el patrón del proyecto |
