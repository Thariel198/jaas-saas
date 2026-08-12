# 4b_reclamos/herramienta

Utilidades **on-demand** que corren fuera del ciclo mensual automático (`main.py` →
`resolucion.py` → `aplicar_correcciones.py` → `validacion_resolucion_correcciones.py`).
El supervisor las invoca manualmente cuando las necesita — no forman parte del pipeline
que corre todos los meses.

## Qué hace cada herramienta

| Script | Qué hace | Cuándo se corre |
|---|---|---|
| `clasificar_tipo_reclamo.py` | Clasifica `TIPO_RECLAMO` por palabra clave en el texto de `RECLAMO`, duplicando la fila si el reclamo mezcla más de un concepto | Cada mes, después de que `4b_reclamos/main.py` generó `reclamos_YYYY-MM.xlsx` y antes de que el supervisor termine de clasificar a mano |
| `buscar_pago.py` *(diseñado, sin codificar aún)* | Para reclamos de tipo "ya pagué X", busca evidencia del pago en todo el historial del predio (3 repos por ciclo) y en los pools de pagos sin identificar | On-demand, cuando hay un lote de reclamos "mes_anterior"/"convenio"/etc. para verificar |

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

## buscar_pago.py — diseño (Fase 2, aún sin código)

Diagrama de flujo: `4b_reclamos/docs/diagrama_flujo_buscar_pago.html`
Decisión de diseño completa: `docs/decisiones/buscar_pago.md`

### Qué hace

Para cada reclamo de un `TIPO_RECLAMO` dado, responde: **¿el pago existe? ¿un bug lo
ocultó? ¿está acreditado a otro predio? ¿nunca existió?** — con evidencia y nivel de
confianza, nunca con un cierre automático del reclamo.

### Reglas de negocio — el embudo

```
GATE   DATA_boletas["MES ANTERIOR"] del ciclo activo == 0 ?
       → SÍ: RESUELTO_YA, cierra sin buscar más
       → NO: reclamo vivo, sigue

BLOQUE A — EXPLICAR (¿la plata ya está adentro?)
  A0  historial completo del predio (3 repos: ago activo · Julio · Junio)
        entró y MES_ANT quedó en 0   → INFUNDADO
        entró y MES_ANT siguió > 0   → MAL_IMPUTADO (bug real)
  A1  ¿algún precursor de shared/ ya cuenta la historia?
        evento en shared/                    → EXPLICADO_POR_PRECURSOR
        aporte al tanque, no era agua         → PAGÓ_PERO_NO_ERA_AGUA
        pagó antes, se aplicó después         → PAGÓ_ANTES_APLICADO_DESPUÉS

BLOQUE B — BUSCAR (la plata no está, ¿dónde se fue?)
  filtros previos: precursores apagan candidatos con dueño · ventana temporal ·
                   cobrador nombrado en el texto reduce el pool casi a 0
  B1  pool de blancos sin reclamar               → CANDIDATO_BLANCO
  B2  registrado en otro predio del mismo ciclo
        origen distinto (tipeo/OCR)              → CANDIDATO_TIPEO
        mismo origen, 1 pago cubrió 2 lotes       → CANDIDATO_MULTILOTE
  B3  exceso no resuelto de un predio confundible → CANDIDATO_EXCESO

SIN_EVIDENCIA → pedir recibo o captura de yape
```

**Regla de propuesta:** un candidato del Bloque B solo se propone si la lista queda en
exactamente 1. Con 2+: `"N candidatos"`, sin elegir (mismo criterio que
`4_pagos/efectivo/verificar_lotes.py`).

**Ventana temporal:**
```
distancia = ciclo_reclamo − mes_del_candidato
≤ 1 mes    → PLAUSIBLE siempre
≥ 2 meses  → PLAUSIBLE solo si MES_ANT > 0 en TODOS los meses intermedios
             (si en algún mes quedó en 0, el candidato viejo se descarta)
```

### Genérico por TIPO_RECLAMO

El mismo embudo sirve para cualquier tipo — "ya pagué mi medidor" (`convenio`) usa la
misma búsqueda que "ya pagué mes anterior" (`mes_anterior`).

### Reusa, no reconstruye

| Pieza reusada | De dónde |
|---|---|
| `tabla_predio()` | `4b_reclamos/reporte_historico.py` |
| `referencias_pago()` | `4b_reclamos/reporte_referencias_pago.py` |
| `confundible()` + `subconjuntos()` | `4_pagos/efectivo/verificar_lotes.py` |

### Blocker a resolver antes de codificar

`reporte_referencias_pago.py` tiene `"2026-07"` apuntando a `BASE_DIR.parent` (el repo
activo). Desde que `shared/ciclo_activo.json` rodó a `2026-08`, eso lee el repo de
**agosto** y lo rotula julio — ya roto hoy. Fix: julio vive en
`C:\Users\wilde\PycharmProjects\Julio\jass_system - Julio` (nombre de archivo distinto:
`planilla_cobrado_julio.xlsx`, necesita alias en `ciclo.resolver`).

### Lo que NO hace

- No cierra el reclamo — el supervisor decide en `resolucion_reclamos_YYYY-MM.xlsx`.
- No escribe en `DATA_boletas.xlsx` ni en ningún archivo de otro módulo.
- No inventa un candidato único cuando hay ambigüedad real (2+ candidatos → se listan, no se elige).

## Estructura

```
4b_reclamos/herramienta/
├── README.md                      # este archivo
├── clasificar_tipo_reclamo.py     # codificado, corre cada mes
└── buscar_pago.py                 # diseñado, pendiente de codificar (Fase 3)
```
