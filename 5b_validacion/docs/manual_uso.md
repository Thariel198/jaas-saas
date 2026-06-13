# Manual de uso — 5b_validacion

**Para quién es este manual:** la persona que cierra la cobranza cada mes y
necesita verificar que el dinero registrado en el sistema cuadra con lo que
realmente entró.

---

## Qué hace este módulo

Compara tres fuentes de información y detecta si hay diferencias:

```
Reporte banco (crudo)
       ↓
pagos_yape_tepago.xlsx  ←→  cobranza_final.xlsx
       ↓
pagos_efectivo.xlsx     ←→  cobranza_final.xlsx
```

Si todo cuadra → imprime `VALIDACION OK`.
Si hay diferencias → genera `validacion_diferencias.xlsx` con el detalle exacto.

---

## Antes de correr este módulo

Verificar que estos pasos ya estén completos:

- [ ] `5_cobranza/main.py` ya corrió y generó `cobranza_final.xlsx`
- [ ] El reporte del banco (crudo) del mes está en `shared/reporte_mes_crudo/`
- [ ] `pagos_yape_tepago.xlsx` está en `5_cobranza/inputs/pagos_yape/`
- [ ] `pagos_efectivo.xlsx` está en `5_cobranza/inputs/pagos_efectivo/`

Este módulo **no necesita que copies ningún archivo** — lee todo desde donde
ya está.

---

## Cómo correr el módulo

Abrir una terminal en la carpeta `5b_validacion/` y ejecutar:

```
python main.py
```

El sistema mostrará el progreso:

```
═══════════════════════════════════════════════════════
  5b_validacion — Validación de cobranza
═══════════════════════════════════════════════════════

[1/4] Validando inputs...
[2/4] Cargando datos...
[3/4] Calculando diferencias...
[4/4] Exportando reporte...

═══════════════════════════════════════════════════════
  VALIDACION OK — todos los montos cuadran
═══════════════════════════════════════════════════════
```

O si hay diferencias:

```
  ALERTA  Yape por MZ: 1 MZ(s) con diferencia
  ALERTA  Efectivo por MZ: 3 MZ(s) con diferencia
  Revisar outputs/validacion_diferencias.xlsx
```

---

## Cómo se determina el período validado

El módulo detecta automáticamente el período del ciclo leyendo la columna
`FECHA` de `pagos_yape_tepago.xlsx`. Usa la fecha más antigua y la más reciente
como límites. Con eso filtra el reporte crudo del banco para comparar solo el
mismo tramo de tiempo.

No hay que configurar fechas manualmente.

---

## Revisar validacion_diferencias.xlsx

El archivo tiene tres hojas:

### Hoja "resumen"

Vista rápida de los totales del mes:

| MEDIO | CRUDO BANCO | PROCESADO | PLANILLA | DIFERENCIA |
|---|---|---|---|---|
| Yape | S/ del banco | S/ en yape_tepago | S/ en cobranza | ¿cuadra? |
| Efectivo | — n/a | S/ en pagos_efectivo | S/ en cobranza | ¿cuadra? |

- **Verde** → cuadra exacto
- **Rojo** → hay diferencia — investigar

> Para Efectivo no hay columna "CRUDO BANCO" porque el archivo
> `pagos_efectivo.xlsx` ya es la fuente directa (no hay un reporte bancario
> separado para efectivo).

### Hoja "yape_por_mz"

Compara `pagos_yape_tepago` con `cobranza_final` manzana por manzana:

| PERIODO | MZ | REPORTE_YAPE | PLANILLA_YAPE | DIFERENCIA | ESTADO |
|---|---|---|---|---|---|

- **DIFERENCIA = PLANILLA − REPORTE**
- Positivo (+): la planilla registró más de lo que entró — puede ser una devolución reconocida
- Negativo (−): la planilla registró menos de lo que entró — puede ser un pago mal asignado de medio (estaba en efectivo pero el reporte lo tiene como Yape)
- **OK** → diferencia = 0
- **ALERTA** → diferencia ≠ 0

### Hoja "efectivo_por_mz"

Misma estructura que la anterior pero para efectivo.

---

## Qué hacer con cada tipo de alerta

### Diferencia en el resumen (totales)

El total global no cuadra. Causas frecuentes:

| Causa | Qué revisar |
|---|---|
| Pago sin identificar en crudo que sí entró a cobranza | Revisar `pagos_yape_tepago` — ¿hay filas sin MZ/LOTE cuyo monto explica la diferencia? |
| Devolución reconocida este mes | Normal — en la planilla puede aparecer más porque se devolvió un exceso del mes anterior |
| Error de monto al ingresar efectivo | Revisar `pagos_efectivo.xlsx` contra el papel de cobro |

### Alerta en una MZ específica (por MZ)

Solo una manzana tiene diferencia. Causas frecuentes:

| Causa | Qué pasa | Ejemplo del mundo real |
|---|---|---|
| Pago registrado en el medio equivocado | Un pago de efectivo quedó contabilizado como Yape en cobranza (o al revés) | MZ C: efectivo dice 410 pero planilla dice 385 — los 25 están en Yape |
| Devolución de exceso anterior | La planilla registró más que el reporte porque se devolvió un exceso | MZ K: reporte=462, planilla=466, dif=+4 |
| Pago de convenio sin identificar | Un pago entró al crudo pero no se asignó correctamente | Buscar en `pagos_yape_tepago` esa MZ |

### Cuando la diferencia es cruzada entre Yape y Efectivo

Si Yape tiene −25 en MZ C y Efectivo tiene +25 en MZ C, el pago simplemente
se registró en el medio equivocado. El dinero está, solo en la columna
incorrecta. Eso se corrige en `6b_override`.

---

## Lo que NO hace este módulo

- No modifica ningún archivo — solo lee y reporta
- No bloquea el pipeline — la decisión de continuar o corregir es tuya
- No explica por qué hay una diferencia — solo la muestra con el número exacto
- No valida conceptos de deuda (MANT, CONVENIO, etc.) — solo los montos pagados

---

## Referencia rápida

```
INPUTS (no copies nada — el módulo los lee directo)
  shared/reporte_mes_crudo/*.xlsx         ← banco crudo del mes
  5_cobranza/inputs/pagos_yape/
    pagos_yape_tepago.xlsx                ← procesado de 4_pagos
  5_cobranza/inputs/pagos_efectivo/
    pagos_efectivo.xlsx                   ← registro de cobro en mano
  5_cobranza/outputs/
    cobranza_final.xlsx                   ← output de 5_cobranza

OUTPUT
  outputs/validacion_diferencias.xlsx     ← siempre se genera
  outputs/run.log                         ← log de ejecución

SIGUIENTE PASO
  Si VALIDACION OK     → continuar a 6_corte
  Si hay ALERTAS       → investigar diferencias · corregir en 6b_override si aplica
```
