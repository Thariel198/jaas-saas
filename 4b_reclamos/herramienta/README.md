# 4b_reclamos/herramienta

Utilidades **on-demand** que corren fuera del ciclo mensual automático (`main.py` →
`resolucion.py` → `aplicar_correcciones.py` → `validacion_resolucion_correcciones.py`).
El supervisor las invoca manualmente cuando las necesita — no forman parte del pipeline
que corre todos los meses.

## Qué hace cada herramienta

| Script | Qué hace | Cuándo se corre |
|---|---|---|
| `clasificar_tipo_reclamo.py` | Clasifica `TIPO_RECLAMO` por palabra clave en el texto de `RECLAMO`, duplicando la fila si el reclamo mezcla más de un concepto | Cada mes, después de que `4b_reclamos/main.py` generó `reclamos_YYYY-MM.xlsx` y antes de que el supervisor termine de clasificar a mano |
| `buscar_pago.py` | Para los reclamos `TIPO_RECLAMO=mes_anterior`, busca evidencia del pago en todo el historial del predio (3 repos por ciclo), en los precursores de `shared/` y en los pools de pagos sin identificar | On-demand, cuando hay un lote de reclamos "ya pagué mes anterior" para verificar |
| `reporte_historico.py` | Genera PDF y Excel desde eventos activos del último ledger comprometido; separa DEUDA/PAGO/AJUSTE/SALDO y no proyecta ciclos abiertos | On-demand para resolver reclamos y auditar el estado mensual real |
| `aplicar_reimputacion_ca1.py` | Simula y aplica la migracion auditada de pagos junio/julio al orden `CONVENIO -> ACUERDOS -> MULTA`; excluye instalaciones/reactivaciones y abonos rezagados | Uso unico para la reimputacion CA1 cerrada el 24/08/2026; una segunda aplicacion se bloquea |

## clasificar_tipo_reclamo.py

Reglas de mapeo (palabra clave en `RECLAMO` → `TIPO_RECLAMO`):

| Palabra clave | TIPO_RECLAMO |
|---|---|
| mes anterior / mes pasado | `mes_anterior` |
| faena / reunion / "multa" | `multa` |
| medidor / "convenio" | `convenio` |
| techado / campo / "cuota" | `cuota` |

Si el texto matchea 2+ categorías, la fila se **duplica**: una fila por categoría, con el
resto de columnas idéntico. Si matchea 1 sola y no coincide con lo que ya tenía el
supervisor, **corrige** el valor existente. Si no matchea ninguna, no toca la fila (queda
para clasificar a mano).

**Idempotente por diseño** — agrupa por identidad (todo menos `TIPO_RECLAMO`) antes de
clasificar, así correr el script 2 veces sobre filas ya divididas no las vuelve a
triplicar (bug real encontrado y corregido el 2026-08-12, ver commit de esta sesión).

```
py clasificar_tipo_reclamo.py                # mes = ciclo activo
py clasificar_tipo_reclamo.py --mes 2026-08
```

## buscar_pago.py

Diagrama de flujo: `4b_reclamos/docs/diagrama_flujo_buscar_pago.html`
Decisión de diseño completa: `docs/decisiones/buscar_pago.md`

```
py buscar_pago.py                # mes = ciclo activo
py buscar_pago.py --mes 2026-08
→ outputs/busqueda_pago_mes_anterior_YYYY-MM.xlsx
```

### Qué hace

Para cada reclamo `TIPO_RECLAMO=mes_anterior`, responde: **¿el pago existe? ¿un bug lo
ocultó? ¿está acreditado a otro predio? ¿nunca existió?** — con evidencia, nunca con un
cierre automático del reclamo.

### Solo mes_anterior — por qué

`convenio` ("ya pagué mi medidor") y `cuota` ("ya pagué techado y campo") **no se auditan
acá**: su causa es anterior y distinta — el orden de la cascada.

```
hoy reparte    consumo · mantenimiento · mes anterior · multa · acuerdos · convenio
va a repartir  consumo · mantenimiento · mes anterior · convenio · acuerdos · multa
               (la multa al final: es lo único que se cubre con faena o se exonera)
```

Antes de ese reorden, buscar el pago de un convenio devuelve "pagó parte de sus cuotas"
— cierto y sin valor. Se probó contra los 43 reclamos de convenio y 29 caían ahí. Ese
reorden es otro trabajo, ya simulado en `4b_reclamos/reporte_reimputacion_cascada.py`.

### El embudo — primer veredicto que matchea gana

```
GATE   DATA_boletas["MES ANTERIOR"] del ciclo activo == 0 ?
       → SÍ: RESUELTO_YA, cierra sin buscar más

BLOQUE A — EXPLICAR (¿la plata ya está adentro?)
  A1a  el pago de ese mes era aporte al tanque y no hubo pago de agua
                                                  → PAGO_PERO_NO_ERA_AGUA
  A1b  abono rezagado con desfase real (pagó en un ciclo, se aplicó en otro)
                                                  → PAGO_ANTES_APLICADO_DESPUES
  A0   el mes disputado  ─┐  misma regla para los dos, en ese orden de fuerza:
  A2   el ciclo del       │   mes anterior recibió algo → PAGO_PARCIAL
       propio reclamo   ──┘   quedó en 0 y la plata fue a multa/acuerdos/convenio
                                                       → CASCADA_FUERA_DE_ORDEN
                              quedó en 0 y la plata fue a consumo+mantenimiento
                                                       → PAGO_SOLO_EL_MES
  A1c  otro precursor de shared/ toca el predio    → EXPLICADO_POR_PRECURSOR

BLOQUE B — BUSCAR (la plata no está, ¿dónde se fue?)
  B2b  el MENSAJE de un pago nombra este lote      → CANDIDATO_MULTILOTE
  B2a  lote confundible + monto cubre el cargo     → CANDIDATO_TIPEO
  B1   blanco sin reclamar, monto y ventana ok     → CANDIDATO_BLANCO
  B3   exceso sin resolver de un lote confundible  → CANDIDATO_EXCESO

SIN_EVIDENCIA → pedir recibo o captura de yape
```

**Qué es anomalía y qué no.** Que el pago del mes se vaya a consumo+mantenimiento es la
cascada **correcta** (P1: agua del mes primero) — no es un hallazgo. La anomalía real es
que se cobre **multa / acuerdos / convenio** dejando el arrastre sin pagar, porque esos
tres van *después* del mes anterior. Medido en 2026-08: **0 casos** de fuera de orden.

**`PAGO_SOLO_EL_MES` es el resultado dominante — 14 de 29.** El vecino paga el mes y no
paga el arrastre, precisamente porque su reclamo sostiene que no lo debe (ya lo pagó en
meses anteriores). Entonces lo que hay que verificar es el **origen del arrastre**, no
re-imputar ese pago.

Si `consumo+mant` suma lo mismo que el cargo de mes anterior (6 de los 14: 5+3=8 y
anterior=8), el monto **no dice** cuál de los dos quiso pagar — la herramienta lo declara
en vez de afirmar una lectura.

### La columna que decide a quién mirar primero

`PAGO_ARRASTRES_ANTES` parte los 14 en dos grupos con acción distinta:

```
 8  "nunca"                 nunca pagó un arrastre → el reclamo no tiene respaldo
 6  "S/xx en N meses"       venía pagando arrastres y aun así le sigue apareciendo
       Q-9   S/76.00 en 4 meses     Z-17  S/101.00 en 4 meses
       W-5   S/23.00 en 3 meses     H-13  S/96.00  en 3 meses
       O-21  S/33.00 en 2 meses     L-2   S/20.00  en 1 mes
```

Los 6 del segundo grupo son donde puede haber un problema real de reconciliación. El
resumen en consola los lista solo.

**Cuánto discrimina:** contra un grupo de control de 60 predios que también deben arrastre
y no reclamaron, "nunca pagó un arrastre" aparece en 18%, contra 41% entre los reclamantes
que hoy deben arrastre. Discrimina ≈2.3× pero no es determinante — por eso va como columna
para priorizar, no como parte del veredicto.

**Regla de propuesta:** un candidato del Bloque B solo se propone si la lista queda en
exactamente 1. Con 2+: `"N candidatos"`, sin elegir (mismo criterio que
`4_pagos/efectivo/verificar_lotes.py` — un falso OK es peor que un no-sé).

**Ventana temporal** (aplica solo al Bloque B):
```
distancia = ciclo_reclamo − mes_del_candidato
≤ 1 mes    → PLAUSIBLE siempre (en agosto reclaman julio: evitar el corte de 2 meses)
≥ 2 meses  → PLAUSIBLE solo si SE DEBÍA en TODOS los meses intermedios
             (si en algún mes quedó en 0, la deuda se cerró y el candidato se descarta)
```

### Reusa, no reconstruye

| Pieza reusada | De dónde |
|---|---|
| `tabla_predio()` (historial, 3 repos por ciclo) | `4b_reclamos/reporte_historico.py` |
| `_cargar_pagos_yape_crudo()` · `_cargar_pagos_efectivo_crudo()` | `4b_reclamos/reporte_referencias_pago.py` |
| `confundible()` · `subconjuntos()` · `leer_boletas()` | `4_pagos/efectivo/verificar_lotes.py` |
| `_cargar_planilla_correcta()` (lo que se DEBÍA cada mes) | `4b_reclamos/reporte_historico.py` |

### Resultado de la primera corrida (29 reclamos, ciclo 2026-08)

```
14  PAGO_SOLO_EL_MES               pagó el mes, no el arrastre → auditar su origen
10  RESUELTO_YA                    la boleta vigente ya no cobra mes anterior
 1  PAGO_PARCIAL                   E-8, cubrió 8 de 16
 1  PAGO_PERO_NO_ERA_AGUA          Q-12, aporte al tanque
 1  PAGO_ANTES_APLICADO_DESPUES    I-9, pagó 05/06 y se aplicó en julio
 1  SIN_BOLETA                     S-16 no existe en DATA_boletas — revisar MZ/LT
 1  SIN_EVIDENCIA                  F1-8, pedir recibo
 0  CASCADA_FUERA_DE_ORDEN         ← el cero es el hallazgo: la cascada está bien
```

Tests: `4b_reclamos/tests/test_buscar_pago.py` — 34/34, los 6 primeros bloques son
**contrafactuales** de falsos positivos reales que la 1a versión emitió (incluido el
encuadre equivocado de "consumo+mant primero" como si fuera una anomalía).

### Lo que NO hace

- No cierra el reclamo — el supervisor decide en `resolucion_reclamos_YYYY-MM.xlsx`.
- No escribe en `DATA_boletas.xlsx` ni en ningún archivo de otro módulo.
- No inventa un candidato único cuando hay ambigüedad real (2+ candidatos → se listan, no se elige).

## reporte_historico.py

```text
seguimiento_pueblo.xlsx + anulaciones_ledger.json
                    |
                    v
eventos activos hasta el ultimo ciclo comprometido
                    |
                    +-- PDF por usuario o lote
                    +-- Excel: Resumen / Mensual / Ajustes / Referencias
```

- Junio/julio muestran cobertura parcial: `MULTA`, `ACUERDOS` y `CONVENIO`.
- La cuenta completa comienza en agosto de 2026.
- Las referencias Yape/efectivo se muestran aparte como evidencia externa; nunca cambian el saldo.
- No lee `abonos_rezagados.xlsx`, precursores ni snapshots abiertos para calcular importes.

```powershell
py -u -X utf8 4b_reclamos/herramienta/reporte_historico.py A 4
py -u -X utf8 4b_reclamos/herramienta/reporte_historico.py --con-deuda 2026-08
py -u -X utf8 4b_reclamos/herramienta/reporte_historico.py --todos 2026-08
```

## Estructura

```
4b_reclamos/herramienta/
├── README.md                      # este archivo
├── aplicar_reimputacion_ca1.py    # migracion CA1: dry run por defecto, --aplicar escribe
├── clasificar_tipo_reclamo.py     # clasifica TIPO_RECLAMO, corre cada mes
├── buscar_pago.py                 # busca el pago de los reclamos mes_anterior
├── reporte_historico.py           # tabla_predio(): historial de un lote (lee ciclo_activo.json)
└── verificar_yape.py
```
