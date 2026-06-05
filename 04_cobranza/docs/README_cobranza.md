# README — 04_cobranza

## Qué hace

Recibe los pagos identificados de Yape y efectivo, los cruza contra las deudas
de la planilla del mes, aplica blancos automáticos desde `shared/`, calcula el
saldo de cada usuario y genera:

- **cobranza_final.xlsx** — estado de pago de todos los usuarios (todos los conceptos)
- **lista_corte.xlsx** — usuarios con 2+ meses de deuda acumulada (penalidad S/20)
- **resumen_recaudacion.xlsx** — totales del mes
- **run.log** — log de ejecución con conteos y errores

## Cuándo se corre

Después de que `03_pagos` está completo:
- `03_pagos/yape/motor_matching/outputs/pagos_yape_tepago.xlsx` existe y no tiene pendientes
- `03_pagos/efectivo/outputs/pagos_efectivo.xlsx` existe y no tiene discrepancias

## Estructura de carpetas

```
04_cobranza/
├── inputs/
│   ├── planilla_base/        ← planilla_base.xlsx (llenar manualmente)
│   ├── pagos_yape/           ← pagos_yape_tepago.xlsx (copia de 03_pagos)
│   └── pagos_efectivo/       ← pagos_efectivo.xlsx (copia de 03_pagos)
├── outputs/
│   ├── cobranza_final.xlsx
│   ├── lista_corte.xlsx
│   ├── resumen_recaudacion.xlsx
│   └── run.log
├── docs/
│   ├── README_cobranza.md         ← este archivo
│   ├── cobranza_final_diseno.html ← diseño visual del output principal
│   └── lista_corte_diseno.html    ← diseño visual de lista de corte
├── tests/
│   └── test_cobranza.py
├── crear_template.py   ← crea planilla_base.xlsx vacía con columnas correctas
└── main.py
```

## Inputs — qué necesita antes de correr

### 1. planilla_base.xlsx

Se llena manualmente cada mes con las lecturas y los conceptos de deuda de cada
usuario. Generar el archivo vacío con `python crear_template.py`.

**Columnas requeridas:**

| Columna | Tipo | Descripción |
|---|---|---|
| MZ | texto | Manzana (ej: A, B, G1) |
| LT | texto | Lote (ej: 1, 7, 11A) |
| NOMBRE | texto | Nombre completo del usuario |
| MARC_ANT | número | Marcación anterior (lectura del mes pasado) |
| MARC_ACT | número | Marcación actual (lectura de este mes) |
| ARRASTRE | número | Deuda pendiente del mes anterior (0 si al día) |
| CONVENIO | número | Cuota de convenio de pago activo (0 si no tiene) |
| MANT | número | Mantenimiento mensual fijo (ej: 3) |
| REUNION_FAENA | número | Multa por reunión/faena no asistida (0 si no aplica) |
| TECHADO | número | Cuota techado si aplica (0 si no) |
| DEVOLUCION | número | Exceso anterior reconocido — reduce la deuda (0 si no aplica) |
| AJUSTE | número | Ajuste manual positivo o negativo (0 si no aplica) |

> `main.py` calcula automáticamente:
> - `m3 = MARC_ACT − MARC_ANT`
> - `TOTAL_MES = m3 × COSTO_M3` (tarifa configurada en `main.py`, actualmente S/1 por m3)
> - `BLANCOS` desde `shared/blancos_acumulados.xlsx` (ver sección Blancos)
> - `TOTAL = TOTAL_MES + ARRASTRE + CONVENIO + MANT + REUNION_FAENA + TECHADO − DEVOLUCION − BLANCOS + AJUSTE`

### 2. pagos_yape_tepago.xlsx

Output de `03_pagos/yape/motor_matching`. Tiene doble cabecera (fila 1 = grupos,
fila 2 = nombres de columnas). Se usa `MONTO_ASIGNA` para pagos multi-lote.
Solo se procesan filas con `TIPO = TE PAGÓ`, `MZ + LOTE` identificados y sin `CONCEPTO`
(los conceptos son gastos comunitarios, no pagos de lote).

### 3. pagos_efectivo.xlsx

Output de `03_pagos/efectivo`. También con doble cabecera. Se usa la columna `MONTO`.

### 4. shared/blancos_acumulados.xlsx (automático)

Archivo compartido en `shared/`. El módulo lo lee automáticamente — no requiere
copia manual. Lee los registros donde `MZ + LOTE` están identificados y
`ESTADO ≠ aplicado`. Al finalizar el cálculo marca esos registros como
`ESTADO = aplicado` y escribe el mes en `MES_APLICADO`.

Si el archivo no existe, el módulo continúa sin blancos y lo registra en el log.

## Reglas de negocio

### Estados de cobranza_final

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

> El umbral de S/8 es el mínimo mensual (mantenimiento). Si el arrastre es menor
> (ej: S/4) puede significar que el mes anterior pagó parcialmente — no se puede
> confirmar que no pagó nada. Con arrastre ≥ 8 se confirma que no hubo ningún pago.

Los usuarios con deuda nueva (primer mes sin pagar, `ARRASTRE = 0`) aparecen como
PARCIAL o PENDIENTE en cobranza_final pero **no** van a lista_corte.

**Filtrado de ya-cortados:** lista_corte incluye a todos los que cumplen la condición.
Si un usuario ya fue cortado en un mes anterior, eliminarlo manualmente de la lista
antes de pasarla a `05_corte`.

### Penalidad

Los usuarios en lista_corte tienen penalidad de **S/20** por reconexión.
`TOTAL_A_PAGAR = SALDO + 20`

## Outputs

### cobranza_final.xlsx

Doble cabecera — fila 1 grupos, fila 2 columnas. Ver `docs/cobranza_final_diseno.html`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ, LT, NOMBRE |
| ¿Cuánto consumió? | MARC_ANT, MARC_ACT, m3, COSTO_m3, TOTAL_MES |
| ¿Cuánto debía? | ARRASTRE, CONVENIO, MANT, REUNION_FAENA, TECHADO, DEVOLUCION, BLANCOS, AJUSTE, TOTAL |
| ¿Cómo pagó? | YAPE, EFECTIVO, TOTAL_PAGADO |
| ¿Qué queda? | SALDO, ESTADO |

Incluye **todos los usuarios** (pagaron o no). Es la fuente de verdad para
`04b_validacion`, `06_boletas` y `07_nueva_planilla`.

### lista_corte.xlsx

Solo usuarios con `SALDO > 0 AND ARRASTRE ≥ 8`. Ver `docs/lista_corte_diseno.html`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ, LT, NOMBRE |
| ¿Por qué va a corte? | DEUDA_ARRASTRE, SALDO |
| ¿Qué debe pagar? | PENALIDAD, TOTAL_A_PAGAR |

### resumen_recaudacion.xlsx

Totales del mes: deuda total, recaudado por Yape/efectivo/combinado, saldo pendiente,
conteo por estado (CANCELADO/EXCESO/PARCIAL/PENDIENTE), y N° de usuarios en lista_corte.

## Flujo mensual

```
1. python crear_template.py           ← solo si no existe el template aún
2. Llenar planilla_base.xlsx          → inputs/planilla_base/
3. Copiar pagos_yape_tepago.xlsx      → inputs/pagos_yape/
4. Copiar pagos_efectivo.xlsx         → inputs/pagos_efectivo/
5. python main.py
6. Revisar outputs/cobranza_final.xlsx
7. Filtrar manualmente ya-cortados de lista_corte.xlsx
8. Pasar lista_corte.xlsx             → 05_corte
```

## Lo que NO hace este módulo

- No genera boletas (eso es `06_boletas`)
- No envía notificaciones a usuarios
- No corta el servicio físicamente (eso es `05_corte`)
- No modifica la planilla Excel original
- No resuelve pagos sin identificar (eso queda en `03_pagos`)

## Errores comunes

| Error | Causa | Solución |
|---|---|---|
| `FileNotFoundError: Inputs faltantes` | Falta alguno de los 3 archivos de input | Ver mensaje en log — indica cuál falta y dónde copiarlo |
| `Columnas faltantes: {'MARC_ANT', ...}` | planilla_base tiene columnas del schema anterior | Regenerar con `crear_template.py` y rellenar |
| `PermissionError al guardar blancos` | blancos_acumulados.xlsx está abierto en Excel | Cerrarlo y volver a correr |
| Saldo negativo inesperado (EXCESO) | Usuario pagó de más | Normal — registrar y esperar reclamo del usuario |
| Lista corte vacía | Nadie cumple la condición arrastre ≥ 8 | Normal si es el primer mes con deudas nuevas |
