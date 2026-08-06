# 4b_reclamos

Sub-módulo independiente que gestiona el ciclo de vida completo de los reclamos detectados en cobros de efectivo: detección, clasificación, preparación de corrección, cierre con auditoría, aplicación a `DATA_boletas` vía repo, y verificación.

## Qué hace

1. **Detecta** reclamos en `pagos_efectivo.xlsx` (COMENTARIO contiene "reclamo", caso insensible) **y** en `inputs/reclamos_manuales.xlsx` — reclamos que llegan por fuera del sistema (secretaria, verbal, WhatsApp), sin cobro asociado. Ambas fuentes se fusionan en un solo checklist.
2. **Preserva** trabajo manual del supervisor en `reclamos_YYYY-MM.xlsx` entre re-corridas.
3. **Arrastra** reclamos PENDIENTE/EN_REVISION del mes anterior cuando no tienen match en el mes actual.
4. **Prepara correcciones**: `resolucion.py` cruza los reclamos clasificados con `DATA_boletas.xlsx` (vía repo) y genera `resolucion_reclamos_YYYY-MM.xlsx` con CAMPO y VALOR_ACTUAL auto-poblados.
5. **Cierra** reclamos RESUELTO/RECHAZADO a `trazabilidad_reclamos.xlsx` (auditoría permanente) con los datos de corrección mergeados desde `resolucion_reclamos_YYYY-MM.xlsx`.
6. **Aplica correcciones** a `DATA_boletas.xlsx`: `aplicar_correcciones.py` lee las filas RESUELTO + CAMPO + VALOR_A_CORREGIR de `resolucion_reclamos_YYYY-MM.xlsx` y llama `shared/data_boletas_repo.apply_correction(...)` por cada una. El repo encapsula backup + write + audit.
7. **Verifica** correcciones: `validacion_resolucion_correcciones.py` cruza las filas RESUELTO de `trazabilidad_reclamos.xlsx` con `DATA_boletas.xlsx` vigente (vía repo) y genera `validacion_correcciones_YYYY-MM.xlsx`. Después de `aplicar_correcciones.py`, lo esperado es 100% OK.

## Cuándo se corre

| Momento | Script | Condición |
|---|---|---|
| Después de `4_pagos/efectivo/main.py` | `main.py` | Al inicio del ciclo mensual, o en cualquier re-corrida |
| Después de que el supervisor clasifique TIPO_RECLAMO | `resolucion.py` | Cuando `reclamos_YYYY-MM.xlsx` tiene ≥1 fila con TIPO_RECLAMO lleno |
| Después de que el supervisor complete VALOR_A_CORREGIR y ESTADO | `main.py` | Segunda corrida del mes, para cerrar reclamos a trazabilidad |
| Después de cerrar RESUELTO a trazabilidad | `aplicar_correcciones.py` | Aplica los VALOR_A_CORREGIR a `DATA_boletas` vía repo |
| Después de aplicar | `validacion_resolucion_correcciones.py` | Confirma que la aplicación funcionó (esperado: 100% OK) |

## Estructura

```
4b_reclamos/
├── main.py                                # Detecta, preserva, arrastra, cierra
├── resolucion.py                          # Cruza reclamos clasificados con DATA_boletas (vía repo)
├── aplicar_correcciones.py                # ★ Baja VALOR_A_CORREGIR a DATA_boletas vía repo
├── validacion_resolucion_correcciones.py  # Verifica RESUELTO → DATA_boletas coinciden
├── inputs/
│   └── reclamos_manuales.xlsx             # Writer único humano — reclamos que llegan por fuera
│                                             (secretaria/WhatsApp/verbal), sin cobro asociado
├── outputs/
│   ├── reclamos_YYYY-MM.xlsx              # Vista operacional mensual (PENDIENTE + EN_REVISION)
│   ├── resolucion_reclamos_YYYY-MM.xlsx   # Hoja de corrección mensual (input doble: main + aplicar)
│   └── validacion_correcciones_YYYY-MM.xlsx  # Verificación de correcciones aplicadas
├── trazabilidad/
│   └── trazabilidad_reclamos.xlsx         # Auditoría permanente (nunca se borra)
├── backup/
│   └── reclamos/                          # Backups de reclamos_YYYY-MM previos
├── tests/
└── docs/
    ├── diagrama_4b_reclamos.html
    ├── arquitectura_4b_reclamos.html
    ├── formato_reclamos.html
    ├── formato_resolucion_reclamos.html
    ├── formato_trazabilidad_reclamos.html
    └── formato_validacion_correcciones.html
```

## Dependencias externas

| Recurso | Tipo | Quién lo gobierna |
|---|---|---|
| `pagos_efectivo.xlsx` | archivo (lectura) | `4_pagos/efectivo/main.py` |
| `inputs/reclamos_manuales.xlsx` | archivo (lectura) | Writer único humano — nadie lo regenera |
| `shared/data_boletas_repo` | módulo (API) | `shared/` — único acceso a `DATA_boletas.xlsx` |
| `shared/utils_lote` | módulo (utilidad pura) | `shared/` — lee `correcciones_lote.xlsx` de 5_cobranza |

**`DATA_boletas.xlsx` ya no se lee ni se escribe directamente.** Toda interacción pasa por `shared/data_boletas_repo`:
- `repo.get_predio(mz, lt)` — lectura para `resolucion.py` y `validacion_*.py`
- `repo.apply_correction(...)` — escritura para `aplicar_correcciones.py`

## Resolución de reclamos de dinero → ledger (Fase 2)

Un reclamo puede ser de **dos naturalezas**, y cada una tiene su destino:

| Naturaleza del reclamo | Ejemplo | Dónde se corrige |
|---|---|---|
| **Datos del predio** | nombre mal escrito, lote equivocado | `DATA_boletas.xlsx` vía `shared/data_boletas_repo` (hoy) |
| **Dinero / deuda** | "ya pagué mayo" (blanco), pago mal atribuido, me cobraron de más | `libro_mayor` — hechos al ledger (Fase 2) |

Para los reclamos de **dinero**, 4b **no escribe la deuda ni la aplicación**. Al
autorizar la resolución, invoca una tool del ledger (writer único intacto); el
**motor de aplicación** de `libro_mayor/estado_cuenta` re-corre y deriva la
aplicación que baja el saldo. Esto es la **decisión ⑨** del contrato del ledger
(ver `libro_mayor/estado_cuenta/README.md`).

| Reclamo resuelto | Tool que invoca 4b | Escribe en |
|---|---|---|
| Blanco reclamado (el pago sí es suyo) | `identificar_abono(abono_id, mz, lt, reclamo_id)` | caja (evento identificación) |
| Pago mal atribuido (fue a otro predio) | `reasignar_abono(abono_id, mz, lt, reclamo_id)` | caja (re-identificación) |
| Cargo incorrecto (le cobraron de más) | `registrar_ajuste(mz, lt, concepto, monto, reclamo_id)` | estado_cuenta (ajuste) |

- Toda tool recibe `reclamo_id` → la aplicación queda linkeada al pago (`ABONO_ID`)
  **y** al reclamo que la autorizó.
- El descuento por blanco **deja de escribirse como columna `BLANCO`/`DEVOLUCION`
  en la planilla** (esa vía se retira en Fase 2) — pasa a ser una aplicación
  auditable en el ledger.

## Reglas clave

- **Reclamos manuales (fuera del sistema):** `inputs/reclamos_manuales.xlsx` (FECHA · MZ · LT · RECLAMO · QUIEN_REPORTA, tipeado por la secretaria) no nace de un cobro, así que no tiene MESA/COBRADOR/MONTO reales. `_cargar_manuales()` los sintetiza (`MESA="MANUAL"`, `FECHA_COBRO=FECHA` del reporte) para producir el mismo shape que `_cargar_detectados()` y reusar sin cambios toda la preservación/arrastre/dedup/trazabilidad de abajo — la clave de identidad es la misma.
- **Preservación manual:** clave `(MESA, MZ, LT, FECHA_COBRO)`. Nunca sobreescribir RECLAMO, TIPO_RECLAMO, RESOLUCION, ESTADO, FECHA_RESOLUCION que el supervisor llenó.
- **Un solo archivo para el cierre:** ESTADO + FECHA_RESOLUCION + RESOLUCION se editan únicamente en `resolucion_reclamos.xlsx`. `main.py` los sincroniza hacia `reclamos.xlsx` automáticamente.
- **Idempotencia:** re-correr sin cambios en inputs produce el mismo output. El repo además garantiza que aplicar el mismo `audit_ref` dos veces es un solo write efectivo.
- **Una historia por reclamo:** todo el ciclo (detección + corrección + cierre + aplicación) queda trazable: trazabilidad (acuerdo) + audit del repo (write real).
- **Corrección null para RECHAZADO:** si ESTADO_FINAL=RECHAZADO, CAMPO/VALOR_ANTERIOR/VALOR_APLICADO quedan vacíos en trazabilidad y nunca se llama al repo.
- **DATA_boletas solo vía repo:** ningún script de este módulo abre `DATA_boletas.xlsx` directamente. Sustituye la antigua invariante "no modifica orígenes" — ahora sí modifica, pero de forma centralizada y auditada.
- **Audit obligatorio en cada write:** cada llamada a `repo.apply_correction` desde aquí pasa `source="4b_reclamos"`, `audit_ref="MESA|FECHA_COBRO"`, `motivo=RESOLUCION`.
- **Verificación no bloquea el flujo:** `validacion_resolucion_correcciones.py` es solo lectura — reporta discrepancias pero no modifica nada.

## Flujo mensual

```
1. main.py (primera corrida)
   ← pagos_efectivo.xlsx
   ← reclamos_YYYY-MM.xlsx (previo, si existe)
   → outputs/reclamos_YYYY-MM.xlsx      [PENDIENTE + EN_REVISION + arrastres]

2. [Supervisor clasifica TIPO_RECLAMO en reclamos_YYYY-MM.xlsx]

3. resolucion.py
   ← outputs/reclamos_YYYY-MM.xlsx
   ← repo.get_predio(mz, lt)            [valor vigente — sin snapshot]
   → outputs/resolucion_reclamos_YYYY-MM.xlsx  [CAMPO + VALOR_ACTUAL auto]

4. [Supervisor llena VALOR_A_CORREGIR + RESOLUCION + ESTADO + FECHA_RESOLUCION]

5. main.py (segunda corrida)
   ← outputs/reclamos_YYYY-MM.xlsx
   ← outputs/resolucion_reclamos_YYYY-MM.xlsx
   → trazabilidad_reclamos.xlsx         [append RESUELTO/RECHAZADO]
   → outputs/reclamos_YYYY-MM.xlsx      [solo activos]

6. aplicar_correcciones.py  ★ NUEVO
   ← outputs/resolucion_reclamos_YYYY-MM.xlsx  [RESUELTO + CAMPO + VALOR]
   → repo.apply_correction(...) por cada fila
        internamente:
          backup → 3_boletas/backup/DATA_boletas/...
          write  → 3_boletas/inputs/DATA_boletas.xlsx
          append → shared/data_boletas_audit.xlsx

7. validacion_resolucion_correcciones.py
   ← trazabilidad_reclamos.xlsx (RESUELTO del mes)
   ← repo.read_padron()
   → outputs/validacion_correcciones_YYYY-MM.xlsx  [esperado: 100% OK]
```

## Lifecycle de outputs

| Archivo | Lifecycle |
|---|---|
| `reclamos_YYYY-MM.xlsx` | Mensual — se regenera; al inicio del mes siguiente los PENDIENTE/EN_REVISION se arrastran |
| `resolucion_reclamos_YYYY-MM.xlsx` | Mensual — input de `main.py` y `aplicar_correcciones.py`; no persiste entre meses |
| `validacion_correcciones_YYYY-MM.xlsx` | Mensual opcional — se regenera con cada corrida de `validacion_resolucion_correcciones.py` |
| `trazabilidad_reclamos.xlsx` | Permanente — nunca se borra; crece mes a mes |

> **Nota:** `aplicar_correcciones.py` no produce archivo de output propio. Su efecto persiste en: `DATA_boletas.xlsx` (update), `shared/data_boletas_audit.xlsx` (append) y `3_boletas/backup/DATA_boletas/` (snapshot).

## Lo que este módulo NO hace

- No abre `DATA_boletas.xlsx` directamente — siempre vía `shared/data_boletas_repo`.
- No modifica `pagos_efectivo.xlsx` — solo lo lee.
- No duplica datos de `3_boletas` ni de `4_pagos`.
- No bypassea el audit log — toda mutación queda registrada por el repo.
- **(Fase 2) No escribe aplicaciones ni el saldo del ledger** — invoca las tools
  de `libro_mayor`; el motor deriva la aplicación. 4b nunca es writer del ledger.

## Señales de alerta

| Señal | Diagnóstico |
|---|---|
| `trazabilidad_reclamos.xlsx` crece pero `reclamos_YYYY-MM.xlsx` no baja | El supervisor no está marcando RESUELTO/RECHAZADO |
| `resolucion_reclamos_YYYY-MM.xlsx` tiene CAMPO vacío | `resolucion.py` no encontró el predio vía repo — revisar MZ/LT |
| Mismo (MESA, MZ, LT, FECHA_COBRO) en trazabilidad dos veces | Error de duplicado — investigar antes de continuar |
| Más del 50% de trazabilidad son RECHAZADO | El filtro COMENTARIO="reclamo" puede estar capturando falsos positivos |
| `validacion_correcciones_YYYY-MM.xlsx` tiene filas ERR después de aplicar | `aplicar_correcciones.py` falló o se saltó filas — revisar consola y audit log |
| `validacion_correcciones_YYYY-MM.xlsx` tiene filas AUSENTE | (MZ,LT) no encontrado en DATA_boletas — el predio puede haber cambiado de clave |
| `aplicar_correcciones.py` reporta "0 aplicados" pero hay RESUELTO | El filtro de aplicar no encuentra filas — revisar que ESTADO=RESUELTO y CAMPO+VALOR no estén vacíos |
