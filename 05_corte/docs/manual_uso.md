# Manual de uso — 05_corte

**Para quién es este manual:** la persona que gestiona el proceso de corte cada
mes — verifica quién pagó la penalidad, entrega la lista al operario y mantiene
el registro de usuarios cortados.

---

## Qué hace este módulo

Después de que se genera la lista de corte, los usuarios tienen una ventana de
2 días para pagar la penalidad de S/20 y evitar el corte físico. Este módulo:

1. Detecta quién pagó la penalidad durante esa ventana
2. Genera la lista para el operario (quiénes cortar físicamente)
3. Mantiene el registro acumulado de todos los usuarios cortados activos

```
lista_corte.xlsx  ←  04_cobranza
cobranza_final.xlsx  ←  04_cobranza        ↘
pagos_yape_tepago.xlsx  ← 03_pagos (re-corrido)  →  05_corte  →  pagaron_penalidad.xlsx
pagos_efectivo.xlsx  ← actualizado                            →  corte_fisico.xlsx
                                                              →  corte_acumulado.xlsx
```

---

## Antes de correr este módulo

Estos pasos deben estar completos en orden:

- [ ] `04_cobranza/main.py` ya corrió y generó `lista_corte.xlsx` y `cobranza_final.xlsx`
- [ ] Se descargó el nuevo reporte del banco — cubre el período extendido incluyendo los días de ventana de penalidad
- [ ] El nuevo reporte está en `shared/reporte_mes_crudo/` (reemplaza el anterior)
- [ ] `03_pagos` fue re-corrido con el nuevo reporte → `pagos_yape_tepago.xlsx` actualizado
- [ ] `pagos_efectivo.xlsx` actualizado si hubo cobros en mano durante la ventana

> El módulo compara los pagos actuales contra lo que ya estaba registrado en
> `cobranza_final.xlsx`. La diferencia es el pago nuevo de la ventana de penalidad.

---

## Cómo correr el módulo

Abrir una terminal en la carpeta `05_corte/` y ejecutar:

```
python main.py
```

El sistema mostrará el progreso:

```
═══════════════════════════════════════════════════════
  05_corte — Clasificación penalidad y corte físico
═══════════════════════════════════════════════════════

[1/5] Validando inputs...
[2/5] Cargando datos...
[3/5] Clasificando...
[4/5] Exportando resultados...
[5/5] Actualizando corte_acumulado...

═══════════════════════════════════════════════════════
  12 pagaron la penalidad
   3 van a corte físico → corte_fisico.xlsx
  Pasar corte_fisico.xlsx al operario
═══════════════════════════════════════════════════════
```

O si todos pagaron:

```
  15 pagaron la penalidad
  Sin corte físico este mes
```

---

## Cómo se detecta quién pagó

El módulo calcula el **pago nuevo** para cada usuario en lista de corte:

```
PAGO_NUEVO = (Yape actualizado + Efectivo actualizado) − lo que ya pagó en cobranza_final
```

- Si `PAGO_NUEVO ≥ S/20` → pagó la penalidad → va a `pagaron_penalidad.xlsx`
- Si `PAGO_NUEVO < S/20` → no pagó → va a `corte_fisico.xlsx`

---

## Revisar los outputs

### pagaron_penalidad.xlsx

Usuarios que pagaron la penalidad durante la ventana. Tres estados posibles:

| ESTADO | Qué significa | Qué hacer |
|---|---|---|
| **CANCELADO** | Pagó penalidad + saldo completo — queda en cero | Nada |
| **PARCIAL** | Pagó los S/20 pero todavía debe el SALDO | Registrar SALDO_FINAL como ARRASTRE en planilla del mes siguiente |
| **EXCESO** | Pagó más de lo que debía | Registrar exceso en DEVOLUCION del mes siguiente |

### corte_fisico.xlsx

Lista para el operario. Contiene solo los usuarios que **no pagaron** la penalidad.

| Columna | Descripción |
|---|---|
| MZ / LT / NOMBRE | Identificación del usuario |
| SALDO | Lo que debía al momento del corte |
| PAGO_EN_VENTANA | Lo que pagó durante la ventana (puede ser 0 o un pago parcial insuficiente) |
| RECONEXION | S/40 fijo — costo de reconexión (el doble porque perdió la ventana) |
| TOTAL_RECONEXION | SALDO + S/40 — lo que debe pagar para que el operario reconecte |

> Entregar este archivo al operario para que ejecute los cortes físicos.

### corte_acumulado.xlsx

Registro persistente de todos los usuarios que han sido cortados. Se actualiza
cada mes automáticamente:

- **CORTADO** → tiene el agua cortada, no ha pagado la reconexión
- **RECONECTADO** → pagó la reconexión y se le restableció el servicio

El módulo detecta automáticamente cuándo alguien se reconectó: si en
`cobranza_final.xlsx` del mes actual aparece con `CORTE_RECONEXION > 0` y
estado `CANCELADO`, se marca como RECONECTADO con el mes actual.

---

## Qué hacer después del corte físico

### Con los usuarios en corte_fisico.xlsx

1. El operario va y corta el suministro físicamente
2. En la planilla del **mes siguiente**, agregar en la columna `CORTE_RECONEXION`
   el valor de `TOTAL_RECONEXION` para cada usuario cortado

   > Ejemplo: MZ C LT 3 tiene TOTAL_RECONEXION = S/82 → en planilla del mes
   > siguiente, columna CORTE_RECONEXION = 82

3. Cuando ese usuario pague ese mes (incluyendo los S/40 de reconexión), el
   operario reconecta y el módulo registrará automáticamente el RECONECTADO en
   corte_acumulado

### Con los usuarios en pagaron_penalidad.xlsx estado PARCIAL

Todavía deben el SALDO del agua. Registrar en planilla del mes siguiente:
- `ARRASTRE` = valor de `SALDO_FINAL`

---

## Referencia rápida

```
PRE-CONDICIONES (en orden)
  04_cobranza corrió → lista_corte.xlsx + cobranza_final.xlsx
  Descargar nuevo reporte banco (período extendido)
  Re-correr 03_pagos → pagos_yape_tepago.xlsx actualizado
  Actualizar pagos_efectivo.xlsx si hubo cobros en mano

INPUTS (no copies nada — el módulo los lee directo)
  04_cobranza/outputs/lista_corte.xlsx
  04_cobranza/outputs/cobranza_final.xlsx
  04_cobranza/inputs/pagos_yape/pagos_yape_tepago.xlsx   ← actualizado
  04_cobranza/inputs/pagos_efectivo/pagos_efectivo.xlsx  ← actualizado

OUTPUTS
  outputs/pagaron_penalidad.xlsx   ← quiénes se salvaron y qué queda
  outputs/corte_fisico.xlsx        ← lista para el operario
  outputs/corte_acumulado.xlsx     ← registro persistente (se actualiza)
  outputs/run.log                  ← log de ejecución

SIGUIENTE PASO
  corte_fisico.xlsx   → operario ejecuta cortes físicos
  PARCIAL             → registrar SALDO_FINAL como ARRASTRE en planilla siguiente
  EXCESO              → registrar en DEVOLUCION en planilla siguiente
  Cortados            → agregar TOTAL_RECONEXION en columna CORTE_RECONEXION planilla siguiente
```
