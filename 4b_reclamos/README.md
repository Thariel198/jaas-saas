# 4b_reclamos

Sub-módulo independiente que gestiona el ciclo de vida completo de los reclamos detectados en cobros de efectivo: detección, clasificación, preparación de corrección y cierre con auditoría.

## Qué hace

1. **Detecta** reclamos en `pagos_efectivo.xlsx` (COMENTARIO contiene "reclamo", caso insensible).
2. **Preserva** trabajo manual del supervisor en `reclamos_YYYY-MM.xlsx` entre re-corridas.
3. **Arrastra** reclamos PENDIENTE/EN_REVISION del mes anterior cuando no tienen match en el mes actual.
4. **Prepara correcciones**: `resolucion.py` cruza los reclamos clasificados con `DATA_boletas.xlsx` y genera `resolucion_reclamos_YYYY-MM.xlsx` con CAMPO y VALOR_ACTUAL auto-poblados.
5. **Cierra** reclamos RESUELTO/RECHAZADO a `trazabilidad_reclamos.xlsx` (auditoría permanente) con los datos de corrección mergeados desde `resolucion_reclamos_YYYY-MM.xlsx`.

## Cuándo se corre

| Momento | Script | Condición |
|---|---|---|
| Después de `4_pagos/efectivo/main.py` | `main.py` | Al inicio del ciclo mensual, o en cualquier re-corrida |
| Después de que el supervisor clasifique TIPO_RECLAMO | `resolucion.py` | Cuando `reclamos_YYYY-MM.xlsx` tiene ≥1 fila con TIPO_RECLAMO lleno |
| Después de que el supervisor complete VALOR_A_CORREGIR y ESTADO | `main.py` | Segunda corrida del mes, para cerrar reclamos a trazabilidad |

## Estructura

```
4b_reclamos/
├── main.py                               # Detecta, preserva, arrastra, cierra
├── resolucion.py                         # Cruza reclamos clasificados con DATA_boletas
├── inputs/                               # Vacío — lee directo de 4_pagos y 3_boletas
├── outputs/
│   ├── reclamos_YYYY-MM.xlsx             # Vista operacional mensual (PENDIENTE + EN_REVISION)
│   └── resolucion_reclamos_YYYY-MM.xlsx  # Hoja de corrección mensual
├── trazabilidad/
│   └── trazabilidad_reclamos.xlsx        # Auditoría permanente (nunca se borra)
├── backup/
│   └── reclamos/                         # Backups de reclamos_YYYY-MM previos
├── tests/
└── docs/
    ├── diagrama_4b_reclamos.html
    ├── arquitectura_4b_reclamos.html
    ├── formato_reclamos.html
    ├── formato_resolucion_reclamos.html
    └── formato_trazabilidad_reclamos.html
```

## Inputs externos (no se copian a `inputs/`)

| Archivo | Ruta externa | Quién lo genera |
|---|---|---|
| `pagos_efectivo.xlsx` | `4_pagos/efectivo/outputs/` | `4_pagos/efectivo/main.py` |
| `DATA_boletas.xlsx` | `3_boletas/inputs/` | Manual / pipeline `3_boletas` |

`DATA_boletas.xlsx` se lee directo desde `3_boletas/inputs/` sin copiar, para garantizar que VALOR_ACTUAL sea el dato vigente al momento del cruce, no un snapshot del inicio del mes.

## Reglas clave

- **Preservación manual:** clave `(MESA, MZ, LT, FECHA_COBRO)`. Nunca sobreescribir RECLAMO, TIPO_RECLAMO, RESOLUCION, ESTADO, FECHA_RESOLUCION que el supervisor llenó.
- **Un solo archivo para el cierre:** ESTADO + FECHA_RESOLUCION + RESOLUCION se editan únicamente en `resolucion_reclamos.xlsx`. `main.py` los sincroniza hacia `reclamos.xlsx` automáticamente — no hay que mantener el mismo dato en dos lugares.
- **Idempotencia:** re-correr sin cambios en inputs produce el mismo output.
- **Una historia por reclamo:** todo el ciclo (detección + corrección + cierre) queda en una sola fila de `trazabilidad_reclamos.xlsx`. Sin JOINs para responder "¿qué pasó con este reclamo?".
- **Corrección null para RECHAZADO:** si ESTADO_FINAL=RECHAZADO, las columnas CAMPO/VALOR_ANTERIOR/VALOR_APLICADO quedan vacías en trazabilidad.
- **DATA_boletas como fuente de verdad:** el módulo no escribe en `DATA_boletas.xlsx`. La corrección registrada en trazabilidad es la decisión acordada, no necesariamente el cambio aplicado en el origen.

## Flujo mensual

```
1. main.py (primera corrida)
   ← pagos_efectivo.xlsx
   ← reclamos_YYYY-MM.xlsx (previo, si existe)
   → outputs/reclamos_YYYY-MM.xlsx      [PENDIENTE + EN_REVISION + arrastres]
   → trazabilidad_reclamos.xlsx         [filas RESUELTO/RECHAZADO pendientes del mes previo]

2. [Supervisor clasifica TIPO_RECLAMO en reclamos_YYYY-MM.xlsx]

3. resolucion.py
   ← outputs/reclamos_YYYY-MM.xlsx
   ← 3_boletas/inputs/DATA_boletas.xlsx
   → outputs/resolucion_reclamos_YYYY-MM.xlsx  [CAMPO + VALOR_ACTUAL auto-llenados]

4. [Supervisor llena VALOR_A_CORREGIR + RESOLUCION + ESTADO + FECHA_RESOLUCION
    en resolucion_reclamos.xlsx — UN SOLO archivo, no hay que volver a reclamos.xlsx]

5. main.py (segunda corrida)
   ← outputs/reclamos_YYYY-MM.xlsx
   ← outputs/resolucion_reclamos_YYYY-MM.xlsx  [ESTADO/FECHA/RESOLUCION sincronizan a reclamos]
   → trazabilidad_reclamos.xlsx         [append — fila con corrección mergeada]
   → outputs/reclamos_YYYY-MM.xlsx      [solo quedan PENDIENTE + EN_REVISION]
```

## Lifecycle de outputs

| Archivo | Lifecycle |
|---|---|
| `reclamos_YYYY-MM.xlsx` | Mensual — se regenera; al inicio del mes siguiente los PENDIENTE/EN_REVISION se arrastran |
| `resolucion_reclamos_YYYY-MM.xlsx` | Mensual — se regenera con cada corrida de `resolucion.py`; no persiste entre meses |
| `trazabilidad_reclamos.xlsx` | Permanente — nunca se borra; crece mes a mes |

## Lo que este módulo NO hace

- No modifica `DATA_boletas.xlsx` — solo lo lee.
- No modifica `pagos_efectivo.xlsx` — solo lo lee.
- No duplica datos de `3_boletas` ni de `4_pagos`.
- No aplica correcciones directamente en el origen de datos — registra la corrección acordada para auditoría.

## Señales de alerta

| Señal | Diagnóstico |
|---|---|
| `trazabilidad_reclamos.xlsx` crece pero `reclamos_YYYY-MM.xlsx` no baja | El supervisor no está marcando RESUELTO/RECHAZADO |
| `resolucion_reclamos_YYYY-MM.xlsx` tiene CAMPO vacío | `resolucion.py` no encontró el predio en DATA_boletas — revisar MZ/LT |
| Mismo (MESA, MZ, LT, FECHA_COBRO) en trazabilidad dos veces | Error de duplicado — investigar antes de continuar |
| Más del 50% de trazabilidad son RECHAZADO | El filtro COMENTARIO="reclamo" puede estar capturando falsos positivos |
