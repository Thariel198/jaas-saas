# Decisión — Arranque digital ciclo Junio 2026

**Fecha:** 2026-06-08
**Módulos afectados:** `2_planilla`, `4_pagos/yape/motor_matching`

---

## Problema

Primer ciclo digital del sistema. Las boletas de junio 2026 fueron generadas manualmente
y ya se enviaron a los usuarios. No existe `planilla_2026-06.xlsx` porque `2_planilla`
aún no había corrido. El módulo `4_pagos` necesita una planilla con deudas para hacer
el matching de pagos Yape.

## Criterios de éxito

1. Cero discrepancia entre el Total que el usuario recibió en su boleta y el Total
   que el sistema usa para el matching.
2. Auditable: cualquier fila puede explicarse fila por fila.
3. Julio arranca 100% digital sin excepciones.

## Enfoque elegido — Técnica del cangrejo

Usar `DATA_boletas.xlsx` (output del proceso manual, fuente de verdad del mes) como
punto de entrada para reconstruir la planilla en sentido inverso:

```
DATA_boletas.xlsx → cangrejo_jun2026.py → planilla_2026-06.xlsx → 4_pagos
```

El script `2_planilla/cangrejo_jun2026.py` mapea las columnas de DATA_boletas al
formato de 21 columnas que esperan `4_pagos` y `5_cobranza`. Las columnas que
`4_pagos` llenará (`MONTO_YAPE`, `MONTO_EFECTIVO`, `ESTADO`, `FECHA_PAGO`) salen
vacías. `BLANCO` y `DEVOLUCION` se inicializan en 0.

## Alternativas descartadas

| Alternativa | Por qué se descartó |
|---|---|
| Correr `2_planilla` + maquillaje | Riesgo de diferencias residuales con la boleta enviada. Viola criterio 1. |
| Aceptar diferencias y documentarlas | El sistema diría un Total distinto al de la boleta. Genera reclamos. |

## Fix relacionado

`4_pagos/yape/motor_matching/main.py` — `ALIAS_PLANILLA` actualizado para aceptar
los nombres canónicos del nuevo formato de planilla:
- `"total_a_pagar"` como alias de `"total"`
- `"mes_anterior"` (con guión bajo) como alias de `"mes anterior"` (con espacio)

Esto deja el motor compatible con la planilla que genera `2_planilla/main.py`
a partir de julio.

## Señal de alerta

Si al correr `cangrejo_jun2026.py` aparece `ValueError: Columnas faltantes en DATA_boletas`,
significa que el archivo manual tiene columnas con nombres distintos a los esperados.
Revisar el MAPEO en el script contra las columnas reales del archivo.

## Cierre

Al terminar el ciclo de junio:
- Mover `2_planilla/cangrejo_jun2026.py` → `backup/`
- Desde julio: correr `2_planilla/main.py` normalmente
