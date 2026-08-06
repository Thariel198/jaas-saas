# 4_pagos — Registro e identificación de pagos del mes

Toma los pagos que entraron por los dos canales —**Yape** (reporte del banco) y
**efectivo** (mesas de cobro)— y los identifica: a qué predio (MZ-LT) corresponde cada
uno. Es la fuente de la que `5_cobranza` calcula el saldo y, en Fase 2, la fuente de
**abonos** del ledger `libro_mayor/caja`.

> Orquestador del módulo: `main.py` corre los 7 pasos en orden. Cada sub-módulo tiene su
> propio README con el detalle. Manual de orquestación en `docs/manual_orquestacion.md`.

---

## Qué hace

```
python main.py     ← corre el pipeline completo (7 pasos)

[1/7] Yape — construir orígenes      yape/construir_maestro/crear_origenes
[2/7] Yape — construir maestro       yape/construir_maestro/crear_maestro
[3/7] Yape — motor de matching       yape/motor_matching        ← identifica MZ-LT de cada yape
[4/7] Yape — validación              yape/validacion
[5/7] Efectivo — procesar cobros     efectivo/                  ← cuadra mesas, detecta discrepancias
[6/7] Tanque — consolidar aportes    consolidar_tanque.py       → outputs/aportes_tanque.xlsx
[7/7] Deuda directiva — append       consolidar_deuda_directiva.py → shared/deuda_directiva.xlsx
```

## Cuándo se corre

Una vez por mes, después de `3_boletas` (el socio ya tiene su boleta) y antes de
`5_cobranza`. Los sub-módulos se pueden correr sueltos para re-procesar un canal.

---

## Estructura

```
4_pagos/
├── main.py                          ← orquestador de los 7 pasos
├── yape/                            ← canal Yape (reporte del banco)
│   ├── construir_maestro/              crear_origenes + crear_maestro
│   ├── motor_matching/                 identifica MZ-LT de cada pago (ciclos hasta pendientes=0)
│   ├── validacion/
│   └── readme_03_pagos_yape.md         README del canal Yape
├── efectivo/                        ← canal efectivo (mesas de cobro)
│   ├── main.py                         cuadra registros por mesa, detecta discrepancias
│   └── README.md                       README del canal efectivo
├── consolidar_tanque.py             ← ledger canal-agnóstico de aportes al tanque
├── consolidar_deuda_directiva.py    ← ledger append-only de deuda de la directiva
├── outputs/aportes_tanque.xlsx
└── docs/
    ├── formato_deuda_directiva.html
    └── manual_orquestacion.md
```

---

## Los dos canales

| Canal | Sub-módulo | Cómo identifica el predio | Output clave |
|---|---|---|---|
| **Yape** | `yape/motor_matching` | motor de matching sobre el mensaje del banco; ciclos hasta `pendientes=0`. `Sin_identificar` = intervención manual. `CONCEPTO=comunitario/multiple` → hoja `Segregacion` (desglose por lote) | `pagos_yape_tepago.xlsx` (+ `_retorno`, `_devolucion`, `_pagaste`, `*_procesado`) |
| **Efectivo** | `efectivo/main.py` | los cobradores registran MZ-LT en campo; el módulo cuadra registros por mesa y detecta discrepancias | `pagos_efectivo.xlsx` (+ `discrepancias.xlsx` si hay) |

## Los dos consolidadores por concepto

- **`consolidar_tanque.py`** → `outputs/aportes_tanque.xlsx`. Cosecha `CONCEPTO=tanque`
  de las vistas de yape + efectivo (canal-agnóstico). **REGENERA** desde el mes actual —
  al cerrar el ciclo pierde meses viejos (el "Gap" del tanque, aceptado: es aporte, no deuda).
- **`consolidar_deuda_directiva.py`** → `shared/deuda_directiva.xlsx`. Cosecha
  `CONCEPTO=deuda_directiva`, hace **APPEND con dedup** (writer único). A diferencia del
  tanque, es **append-only** con columna `CICLO`: sobrevive el cierre.

---

## Reglas de negocio

- **1 depósito = 1 abono.** Un yape que paga varios lotes es **un solo pago** con
  `MONTO = MONTO_PAGO` (depósito completo); el reparto por lote no vive acá.
- **Motor de matching corre en ciclos** hasta que no queden pendientes; los ambiguos se
  resuelven solos salvo `comunitario`/`multiple` (van a segregación por lote).
- **Efectivo se re-corre** hasta que `discrepancias.xlsx` desaparece.
- **Writer único por ledger de concepto:** solo `consolidar_deuda_directiva.py` escribe
  `shared/deuda_directiva.xlsx`; solo `consolidar_tanque.py` escribe `aportes_tanque.xlsx`.

---

## Alimenta el ledger `libro_mayor/caja` (Fase 2)

Los outputs de este módulo son la **fuente de ABONOS** del ledger:

| Fuente | Importador (Fase 2) | Clave natural del `ABONO_ID` |
|---|---|---|
| `pagos_efectivo.xlsx` | `libro_mayor/caja/importar_efectivo.py` | `(jass, MESA, COBRADOR, FECHA, MONTO, MZ, LT)` |
| `*_procesado.xlsx` (yape) | `libro_mayor/caja/importar_yape.py` | `(jass, ORIGEN, TIMESTAMP)` — **predio-agnóstica** |
| `aportes_tanque.xlsx` | importador de caja | balde `tanque` (INGRESO, no cruza a deuda) |
| `shared/deuda_directiva.xlsx` | importador de caja | balde `deuda_directiva` (INGRESO, caja-only) |
| `pagos_yape_pagaste.xlsx` | `importar_egresos.py` | EGRESO: devolución / retorno / gasto |

Esas columnas **no deben cambiar de nombre** — son la clave del id determinista. El
reparto por lote de un yape compartido son **aplicaciones** en `estado_cuenta`, no
abonos separados. Ver el contrato en `libro_mayor/caja/README.md` (①⑩).

---

## Lo que NO hace

- **No calcula saldos ni aplica pagos a deudas** — eso es `5_cobranza` (hoy) / el motor
  de aplicación del ledger (Fase 2). 4_pagos solo dice "entró este pago, es de este predio".
- **No valida el cuadre de caja** — eso es `5b_validacion`.

## Errores comunes

- Pasar a `5_cobranza` con pendientes de Yape sin identificar o discrepancias de efectivo
  sin resolver → saldos incompletos.
- Correr `consolidar_tanque` esperando historial multi-mes: regenera, no acumula (por diseño).
