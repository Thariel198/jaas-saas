# README — 5_cobranza

## Qué hace

Recibe la planilla del mes (desde `2_planilla/outputs/`) y los pagos identificados
de Yape y efectivo, los cruza contra las deudas, aplica blancos automáticos desde
`shared/`, calcula el saldo de cada usuario y genera:

- **planilla_cobrado.xlsx** — estado de pago de todos los usuarios (todos los conceptos)
- **lista_corte.xlsx** — usuarios con ARRASTRE ≥ 8 y saldo pendiente (penalidad S/20)
- **trazabilidad_cobranza.xlsx** — un registro por pago cargado en este ciclo
- **resumen_recaudacion.xlsx** — totales del mes
- **arrastre_deuda_YYYY-MM.xlsx** — deuda pendiente para alimentar 2_planilla el ciclo siguiente
- **arrastre_devolucion_YYYY-MM.xlsx** — excesos (SALDO < 0) pendientes de reclamo, para 7_cierre
- **run.log** — log de ejecución con conteos y errores

## Cuándo se corre

Después de que `4_pagos` está completo:
- `4_pagos/yape/motor_matching/outputs/pagos_yape_tepago.xlsx` existe y no tiene pendientes
- `4_pagos/efectivo/outputs/pagos_efectivo.xlsx` existe y no tiene discrepancias

## Estructura de carpetas

```
5_cobranza/
├── inputs/
│   ├── planilla/             ← copia de 2_planilla/outputs/planilla_YYYY-MM.xlsx
│   ├── pagos_yape/           ← pagos_yape_tepago.xlsx (copia de 4_pagos)
│   └── pagos_efectivo/       ← pagos_efectivo.xlsx (copia de 4_pagos)
├── outputs/
│   ├── planilla_cobrado.xlsx
│   ├── lista_corte.xlsx
│   ├── trazabilidad_cobranza.xlsx
│   ├── resumen_recaudacion.xlsx
│   ├── arrastre_deuda_YYYY-MM.xlsx
│   ├── arrastre_devolucion_YYYY-MM.xlsx
│   ├── run.log
│   ├── run_validacion.log              ← lo genera validacion_planilla_cobrado.py
│   └── validacion_errores.xlsx         ← solo si hay errores (lo genera validacion)
├── docs/
│   ├── README_cobranza.md                 ← este archivo
│   ├── diagrama_5_cobranza.html           ← diagrama de flujo del módulo
│   ├── arquitectura_5_cobranza.html       ← arquitectura detallada
│   ├── planilla_cobrado_diseno.html       ← diseño visual de planilla_cobrado
│   ├── lista_corte_diseno.html            ← diseño visual de lista de corte
│   ├── trazabilidad_cobranza.html         ← diseño visual del log de carga
│   ├── arrastre_deuda_diseno.html         ← diseño visual del arrastre de deuda
│   └── arrastre_devolucion_diseno.html    ← diseño visual del arrastre de devolución
├── tests/
│   └── test_cobranza.py
├── main.py
└── validacion_planilla_cobrado.py      ← script de validación post-cobranza
```

## Inputs — qué necesita antes de correr

### 1. planilla_YYYY-MM.xlsx (desde 2_planilla)

Copia de `2_planilla/outputs/planilla_YYYY-MM.xlsx` — generada por el módulo de
planilla con lecturas, consumos y conceptos de deuda ya calculados. No se modifica:
`main.py` la lee y genera `planilla_cobrado.xlsx` como copia enriquecida con pagos.

**Columnas que trae la planilla (las genera 2_planilla · 21 columnas):**

| Columna | Tipo | Descripción |
|---|---|---|
| MZ | texto | Manzana (ej: A, B, G1) |
| LT | texto | Lote (ej: 1, 7, 11A) |
| NOMBRE | texto | Nombre completo del usuario |
| MES_ANO | texto | Periodo (ej: 2026-06) |
| MARC_ANT | número | Marcación anterior |
| MARC_ACT | número | Marcación actual |
| M3 | número | Consumo calculado |
| MES_ACTUAL | número | Cobro por consumo del mes |
| MANTENIMIENTO | número | Cuota de mantenimiento mensual fija |
| MES_ANTERIOR | número | Deuda pendiente del mes anterior |
| CORTE_RECONEXION | número | Cargo por corte/reconexión si aplica |
| CONVENIO | número | Cuota de convenio de pago activo |
| MULTA | número | Multa por reunión/faena no asistida |
| ACUERDOS_ASAMBLEA | número | Cuotas aprobadas en asamblea |
| BLANCO | número | Descuento por blanco (negativo · sale en 0 · 5_cobranza lo actualiza) |
| DEVOLUCION | número | Devolución de exceso (negativo · sale en 0 · 5_cobranza lo actualiza) |
| TOTAL_A_PAGAR | fórmula | Suma de todos los conceptos (fórmula Excel) |
| MONTO_YAPE | vacío | Lo llena 5_cobranza desde pagos_yape_tepago |
| MONTO_EFECTIVO | vacío | Lo llena 5_cobranza desde pagos_efectivo |
| ESTADO | vacío | Lo llena 5_cobranza (CANCELADO / EXCESO / PARCIAL / PENDIENTE) |
| FECHA_PAGO | vacío | Lo llena 5_cobranza — fecha del pago más reciente |

> `CICLO_COBRANZA` **no** viene en la planilla — 5_cobranza lo agrega como columna 22 al crear `planilla_cobrado.xlsx`.

### 2. pagos_yape_tepago.xlsx

Output de `4_pagos/yape/motor_matching`. Doble cabecera (fila 1 = grupos, fila 2 =
columnas). Se usa `MONTO_ASIGNA` para pagos multi-lote. Solo se procesan filas con
`TIPO = TE PAGÓ`, `MZ + LOTE` identificados y sin `CONCEPTO`.
La columna `CICLO_CORRECCION` indica el ciclo de corrección en que fue identificado.

### 3. pagos_efectivo.xlsx

Output de `4_pagos/efectivo`. Doble cabecera. Se usa la columna `MONTO`.
La columna `CICLO_CORRECCION` indica el ciclo de corrección en que fue registrado.

### 4. shared/blancos_acumulados.xlsx (automático)

Archivo compartido en `shared/`. El módulo lo lee automáticamente — no requiere
copia manual. Lee los registros donde `MZ + LOTE` están identificados y
`ESTADO ≠ aplicado`. Al finalizar el cálculo marca esos registros como
`ESTADO = aplicado` y escribe el mes en `MES_APLICADO`.

Si el archivo no existe, el módulo continúa sin blancos y lo registra en el log.

## Reglas de negocio

### Estados de planilla_cobrado

| Estado | Condición |
|---|---|
| CANCELADO | saldo = 0 (tolerancia ±0.005) |
| EXCESO | saldo < 0 — pagó de más; se registra, no se devuelve hasta que el usuario reclame |
| PARCIAL | saldo > 0 y pagó algo |
| PENDIENTE | saldo > 0 y no pagó nada |

### Regla de lista_corte

Un usuario va a lista_corte **solo si** cumple ambas condiciones:

- `SALDO > 0` — tiene deuda pendiente este mes
- `ARRASTRE ≥ 8` — tenía al menos S/8 de arrastre del mes anterior

> El umbral de S/8 es el mínimo mensual (mantenimiento S/3 + mínimo agua ≈ S/5 = S/8).
> Si el arrastre es menor (ej: S/4) puede significar que el mes anterior pagó parcialmente.
> Con arrastre ≥ 8 se confirma que no hubo ningún pago el mes anterior.

Los usuarios con deuda nueva (primer mes sin pagar, `ARRASTRE = 0`) aparecen como
PARCIAL o PENDIENTE en planilla_cobrado pero **no** van a lista_corte.

**Filtrado de ya-cortados:** lista_corte incluye a todos los que cumplen la condición.
Si un usuario ya fue cortado en un mes anterior, eliminarlo manualmente de la lista
antes de pasarla a `6_corte`.

### Penalidad

Los usuarios en lista_corte tienen penalidad de **S/20** por reconexión.
`TOTAL_A_PAGAR = SALDO + 20`

### CICLO_COBRANZA

Al cargar pagos, `main.py` detecta si ya existen cargos previos en este mes
(leyendo el valor máximo de `CICLO_COBRANZA` en trazabilidad_cobranza) e incrementa.
El valor se escribe en:
- `planilla_cobrado.xlsx` columna `CICLO_COBRANZA` — **agregada** por 5_cobranza (no existe en la planilla original)
- `trazabilidad_cobranza.xlsx` columna `CICLO_COBRANZA` (en cada fila del log)
- `pagos_yape_tepago.xlsx` y `pagos_efectivo.xlsx` columna `CICLO_COBRANZA` (retroescritura)

## Outputs

### planilla_cobrado.xlsx

21 columnas de la planilla original + `CICLO_COBRANZA` (col 22). Ver `docs/planilla_cobrado_diseno.html`.

Las columnas de pago (`MONTO_YAPE`, `MONTO_EFECTIVO`, `ESTADO`, `FECHA_PAGO`) llegan vacías
de 2_planilla y son completadas por este módulo. `BLANCO` y `DEVOLUCION` se escriben como
negativos cuando aplican (reducen `TOTAL_A_PAGAR`).

Incluye **todos los usuarios** (pagaron o no). Es la fuente de verdad para
`5b_validacion`, `3_boletas` y `7_cierre`.

### lista_corte.xlsx

Solo usuarios con `SALDO > 0 AND MES_ANTERIOR ≥ 8`. Ver `docs/lista_corte_diseno.html`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ, LT, NOMBRE |
| ¿Por qué va a corte? | DEUDA_ARRASTRE, SALDO |
| ¿Qué debe pagar? | PENALIDAD (S/20), TOTAL_A_PAGAR |

### trazabilidad_cobranza.xlsx

Un registro por pago cargado. Acumulada — cada corrida agrega filas sin borrar las anteriores.
Ver `docs/trazabilidad_cobranza.html`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ, LT, NOMBRE |
| ¿Qué se cargó? | MONTO, FUENTE |
| ¿Cuándo y de dónde? | CICLO_CORRECCION_ORIGEN, CICLO_COBRANZA, FECHA_CARGA |

### resumen_recaudacion.xlsx

Totales del mes: deuda total, recaudado por Yape/efectivo/combinado, saldo pendiente,
conteo por estado (CANCELADO/EXCESO/PARCIAL/PENDIENTE), y N° de usuarios en lista_corte.

### arrastre_deuda_YYYY-MM.xlsx

Solo usuarios con `SALDO > 0`. Ver `docs/arrastre_deuda_diseno.html`.

| Columna | Descripción |
|---|---|
| `MZ` | Manzana |
| `LT` | Lote |
| `NOMBRE` | Nombre del usuario |
| `monto` | SALDO pendiente del ciclo actual (S/) |
| `MES_ANO_ORIGEN` | Mes en que se originó la deuda (ej: 2026-06) |

Este archivo se copia manualmente a `2_planilla/inputs/deuda_anterior/` al iniciar
el ciclo siguiente. Alimenta la columna `MES_ANTERIOR` de la planilla del mes nuevo.

### arrastre_devolucion_YYYY-MM.xlsx

Solo usuarios con `SALDO < 0` (EXCESO — pagaron de más). Ver `docs/arrastre_devolucion_diseno.html`.

| Columna | Descripción |
|---|---|
| `MZ` | Manzana |
| `LT` | Lote |
| `NOMBRE` | Nombre del usuario |
| `monto` | Exceso a devolver — valor absoluto del SALDO (S/) |
| `MES_ANO_ORIGEN` | Mes en que se originó el exceso (ej: 2026-06) |

La JASS no devuelve automáticamente — espera a que el usuario reclame.
`7_cierre` acumula estos excesos en `shared/` para el seguimiento inter-ciclo.

## Flujo mensual

```
1. Copiar planilla_YYYY-MM.xlsx       → inputs/planilla/       (desde 2_planilla/outputs/)
2. Copiar pagos_yape_tepago.xlsx      → inputs/pagos_yape/     (desde 4_pagos/yape/outputs/)
3. Copiar pagos_efectivo.xlsx         → inputs/pagos_efectivo/ (desde 4_pagos/efectivo/outputs/)
4. python main.py
5. python validacion_planilla_cobrado.py
   → si hay errores: revisar outputs/validacion_errores.xlsx y corregir antes de continuar
6. Revisar outputs/planilla_cobrado.xlsx
7. Filtrar manualmente ya-cortados de lista_corte.xlsx
8. Pasar lista_corte.xlsx             → 6_corte
9. Copiar arrastre_deuda_YYYY-MM.xlsx → 2_planilla/inputs/deuda_anterior/  (para el ciclo siguiente)
```

## Validación post-cobranza

`validacion_planilla_cobrado.py` — script independiente que verifica la coherencia
interna de los 6 outputs generados por `main.py`. Se corre **después** de `main.py`.

```
python validacion_planilla_cobrado.py
```

| Bloque | Qué verifica |
|---|---|
| 1 — Re-cálculo SALDO | Recalcula `TOTAL_A_PAGAR` y `SALDO` desde los 9 conceptos · verifica `ESTADO` consistente |
| 2 — Trazabilidad ↔ planilla | `sum(MONTO por FUENTE)` en trazabilidad == `MONTO_YAPE` / `MONTO_EFECTIVO` en planilla · huérfanos |
| 3 — arrastre_deuda ↔ planilla | Cada fila tiene `SALDO > 0` en planilla · `monto == SALDO` · faltantes y huérfanos |
| 3 — arrastre_devolucion ↔ planilla | Espejo del anterior: `SALDO < 0` · `monto == |SALDO|` · faltantes y huérfanos |

**Outputs del script:**
- `outputs/run_validacion.log` — siempre se genera
- `outputs/validacion_errores.xlsx` — solo si se encontraron errores (una hoja por bloque)

**Exit code:** `0` = todo OK · `1` = se encontraron errores

> Los errores de Bloque 2 tipo `HUERFANO` indican pagos en mesa con MZ-LT que no
> existen en la planilla. Deben corregirse en `4_pagos/efectivo/` antes de volver a correr.

## Lo que NO hace este módulo

- No genera boletas (eso es `3_boletas`)
- No envía notificaciones a usuarios
- No corta el servicio físicamente (eso es `6_corte`)
- No modifica la planilla Excel original
- No resuelve pagos sin identificar (eso queda en `4_pagos`)

## Errores comunes

| Error | Causa | Solución |
|---|---|---|
| `FileNotFoundError: Inputs faltantes` | Falta alguno de los archivos de input | Ver mensaje en log — indica cuál falta y dónde copiarlo |
| `Columnas faltantes: {'MARC_ANT', ...}` | Planilla con schema incorrecto | Verificar que viene de `2_planilla/outputs/` del mes correcto |
| `PermissionError al guardar blancos` | blancos_acumulados.xlsx está abierto en Excel | Cerrarlo y volver a correr |
| Saldo negativo inesperado (EXCESO) | Usuario pagó de más | Normal — registrar y esperar reclamo del usuario |
| Lista corte vacía | Nadie cumple la condición arrastre ≥ 8 | Normal si es el primer mes con deudas nuevas |
