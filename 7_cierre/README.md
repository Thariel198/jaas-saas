# 7_cierre — Consolidador de cierre de ciclo

Consolida las decisiones del mes (reclamos resueltos + cortes ejecutados) y produce
los arrastres **finales** listos para que `2_planilla` del próximo mes los pre-cargue.

> **Estado:** diseñado · pendiente de implementación.
> Los arrastres intermedios que este módulo consume aún no los genera `6_corte` — eso
> va en la siguiente fase de desarrollo.

---

## Cuándo correr

Al **cierre del período de gracia**, después de que:
- `6_corte` ejecutó los cortes físicos y generó `arrastre_corte_intermedio`
- `4b_reclamos` resolvió todos los reclamos posibles del mes

```
1_lecturas → 2_planilla → 3_boletas → 4_pagos → 5_cobranza
→ 6_corte → [período de gracia + resolución de reclamos] → 7_cierre → (próximo mes)
```

Es idempotente: puede correrse varias veces mientras lleguen nuevas resoluciones
de reclamos. El output final siempre refleja el estado más reciente.

---

## Estructura de carpetas

```
7_cierre/
├── README.md
├── config.py
├── consolidar_cierre.py        ← script principal
├── docs/
│   ├── diagrama_flujo_7_cierre.html     ← vista rápida del flujo
│   └── diagrama_consolidador_cierre.html ← diseño detallado
├── inputs/                     ← 7_cierre lee de aquí (copiados por el módulo fuente)
│   ├── arrastre_corte_intermedio_YYYY-MM.xlsx   ← de 6_corte/outputs/
│   ├── resolucion_reclamos_YYYY-MM.xlsx         ← de 4b_reclamos/outputs/
│   └── planilla_cobrado_YYYY-MM.xlsx            ← de 5_cobranza/outputs/
└── outputs/
    ├── arrastre_corte_final_YYYY-MM.xlsx
    ├── arrastre_deuda_final_YYYY-MM.xlsx
    └── run.log
```

---

## Inputs

| Archivo | Origen | Qué aporta |
|---------|--------|------------|
| `arrastre_corte_intermedio_YYYY-MM.xlsx` | `6_corte/outputs/` | Todos los usuarios elegibles para corte con ESTADO_RECLAMO |
| `resolucion_reclamos_YYYY-MM.xlsx` | `4b_reclamos/outputs/` | Decisiones del supervisor: FUNDADO / RECHAZADO / EN_REVISION |
| `planilla_cobrado_YYYY-MM.xlsx` | `5_cobranza/outputs/` | Saldos finales del mes (fuente de verdad de deuda) |

### Columnas requeridas — arrastre_corte_intermedio

| Columna | Tipo | Descripción |
|---------|------|-------------|
| MZ | str | Manzana |
| LT | str | Lote |
| NOMBRE | str | Nombre del usuario |
| SALDO | float | Deuda total al momento del corte |
| MES_ANTERIOR | float | Monto de mes anterior que originó el corte |
| PENALIDAD | float | S/20 aplicado por corte |
| ESTADO_RECLAMO | str | SIN_RECLAMO / EN_REVISION / PENDIENTE |
| MES_ORIGEN | str | YYYY-MM del ciclo que generó el corte |

### Columnas requeridas — resolucion_reclamos

| Columna | Tipo | Valores válidos |
|---------|------|-----------------|
| MZ | str | |
| LT | str | |
| ESTADO | str | `FUNDADO` · `RECHAZADO` · `EN_REVISION` |

---

## Outputs

### arrastre_corte_final_YYYY-MM.xlsx

Usuarios que deben pagar corte y reconexión en el próximo mes.
`2_planilla` lo lee y pre-carga la columna `CORTE_RECONEXION`.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| MZ | str | |
| LT | str | |
| NOMBRE | str | |
| PENALIDAD | float | S/20 por corte |
| RECONEXION | float | S/20 por reconexión |
| TOTAL_CORTE | float | S/40 (penalidad + reconexión) |
| ESTADO_DECISION | str | CORTADO / RECHAZADO / DIFERIDO |
| MES_ORIGEN | str | YYYY-MM del ciclo de corte |

### arrastre_deuda_final_YYYY-MM.xlsx

Usuarios con saldo pendiente que pasa como `MES_ANTERIOR` al próximo ciclo.
`2_planilla` lo lee y pre-carga la columna `MES_ANTERIOR`.

| Columna | Tipo | Descripción |
|---------|------|-------------|
| MZ | str | |
| LT | str | |
| NOMBRE | str | |
| SALDO | float | Saldo total pendiente al cierre |
| MES_ANTERIOR | float | Monto que va en MES_ANTERIOR del próximo mes |
| MES_ORIGEN | str | YYYY-MM del ciclo de origen |

---

## Reglas de negocio

### Cómo se procesa cada usuario de arrastre_corte_intermedio

| ESTADO_RECLAMO en intermedio | Resolución en resolucion_reclamos | Acción en 7_cierre | Va en arrastre_corte_final |
|------------------------------|-----------------------------------|-------------------|---------------------------|
| SIN_RECLAMO | — | Corte confirmado | Sí · TOTAL_CORTE = S/40 |
| EN_REVISION o PENDIENTE | RECHAZADO | Corte confirmado retroactivo | Sí · TOTAL_CORTE = S/40 |
| EN_REVISION o PENDIENTE | FUNDADO | Corte anulado | No · sin cargo |
| EN_REVISION o PENDIENTE | EN_REVISION (sin resolver) | Diferido al próximo mes | No este mes (ver regla especial) |

### Cálculo de TOTAL_CORTE

```
PENALIDAD  = S/20  (corte aplicado)
RECONEXION = S/20  (reconexión al pagar)
TOTAL_CORTE = S/40
```

### Regla especial — Ciclo junio 2026

Los reclamos con ESTADO `EN_REVISION` al cierre **no reciben corte físico ni penalidad**
en este ciclo. Su deuda (`SALDO`) pasa como `MES_ANTERIOR` en `arrastre_deuda_final`
para que el próximo mes el sistema los detecte nuevamente como elegibles.

> Esta regla es temporal: el sistema está en rodaje y hay muchos reclamos sin resolver.
> En ciclos futuros evaluar si el corte procede de todos modos al vencer el período de gracia.

### Devolución por reclamo FUNDADO

Si el usuario ya pagó S/20 de penalidad antes de que se resolviera el reclamo:
- 7_cierre registra el caso en la columna `DEVOLUCION_PENDIENTE` del arrastre_corte_final
- `4b_reclamos` gestiona la devolución efectiva (fuera del alcance de 7_cierre)

---

## Flujo paso a paso

```
python 7_cierre/consolidar_cierre.py --mes 2026-06
```

1. Validar que existen los 3 inputs del mes
2. Leer `arrastre_corte_intermedio` — todos los elegibles para corte
3. Leer `resolucion_reclamos` — cruzar por MZ+LT con los elegibles
4. Aplicar reglas de negocio → clasificar cada usuario
5. Escribir `arrastre_corte_final_YYYY-MM.xlsx`
6. Leer `planilla_cobrado` — extraer saldos pendientes (SALDO > 0)
7. Escribir `arrastre_deuda_final_YYYY-MM.xlsx`
8. Copiar ambos finales a `2_planilla/inputs/arrastres/` para el próximo mes
9. Escribir `run.log`

---

## Lo que NO hace este módulo

- No calcula quién tiene saldo (lo hace `5_cobranza`)
- No decide quién es elegible para corte (lo hace `6_corte`)
- No aplica correcciones a `DATA_boletas` (lo hace `4b_reclamos`)
- No genera la planilla del próximo mes (lo hace `2_planilla`)
- No gestiona devoluciones de dinero (las registra pero las ejecuta `4b_reclamos`)

---

## Errores comunes

| Error | Causa | Solución |
|-------|-------|----------|
| `arrastre_corte_intermedio no encontrado` | `6_corte` aún no genera este archivo | Pendiente de implementar en `6_corte` |
| Usuario en intermedio sin match en resolucion | Reclamo no registrado en `4b_reclamos` | Verificar que el reclamo se cargó en `reclamos_YYYY-MM.xlsx` |
| ESTADO desconocido en resolucion_reclamos | Valor distinto a FUNDADO/RECHAZADO/EN_REVISION | Corregir el valor en el archivo y re-correr |

---

## Dependencias de implementación

Antes de implementar `7_cierre`, se requiere que `6_corte` genere:
- `lista_corte.xlsx` — auditoría (todos los elegibles + ESTADO_RECLAMO + EJECUTAR_CORTE)
- `corte_fisico_YYYY-MM.xlsx` — vista operacional (solo EJECUTAR_CORTE = SI)
- `arrastre_corte_intermedio_YYYY-MM.xlsx` — todos los cortados sin decisión final de penalidad

Ver `6_corte/README.md` para el estado de implementación de estos outputs.
