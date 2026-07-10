# 7b_historial_pagos

Ledger **append-only de pagos crudos por canal**, acumulado en todos los meses y
consultable por predio. Responde: *"¿qué pagó este predio, cuándo y por qué canal?"*

> Estado: **Fase 1 — diseño cerrado, sin implementar.** Este README es la fuente de
> verdad del módulo. El código lo implementa; si algo del código contradice este
> documento, manda el README.

---

## Cuándo corre — después de 7_cierre, no en caliente

`7b_historial_pagos` carga el mes al histórico **recién después de que `7_cierre` lo
consolida**, nunca durante el mes en curso.

**Por qué:** mientras el mes está abierto, los reclamos generan reatribuciones entre
predios (ej. exceso de M-15 cubre la multa de G-1), rechazos y correcciones de lote —
el pago "verdadero" del predio todavía se está asentando. `7_cierre` es el momento en
que esas decisiones quedan consolidadas y el mes se vuelve inmutable. Cargar antes
fotografiaría un mes a medio corregir, el mismo problema que ya se resolvió para
`mesa_N` (verdad que cambia bajo los pies) — el histórico terminaría con datos que
después habría que volver a tocar, rompiendo el invariante append-only.

```
mes en curso → reclamos, reatribuciones, correcciones de lote (el pago se asienta)
     │
7_cierre     → consolida decisiones, produce arrastres finales, el mes queda inmutable
     │
7b_historial_pagos  → recién ahí importa el mes al ledger — 1 sola vez, definitivo
```

Consecuencia de diseño: `importar_efectivo.py` e `importar_yape.py` toman como fuente
el estado **posterior** a `7_cierre` (no `pagos_efectivo.xlsx` en caliente), y las
correcciones de reclamos que ya se aplicaron durante el mes (reatribuciones, rechazos)
entran como parte del pago consolidado, no como eventos de corrección separados —
la corrección post-cierre (si aparece un reclamo tardío) sí sigue usando
`registrar_correccion`.

---

## Por qué existe

Los libros de Excel viejos tenían una **hoja de pagos completa por mes**: para verificar
un "te pagué tal mes" se abría la hoja y se buscaba al usuario. Al modularizar el sistema
esa capacidad **se perdió para el efectivo**:

| Canal | ¿Responde "¿pagó en el mes M?" hoy? |
|---|---|
| **Yape** | Sí — `shared/reporte_acumulado_procesado/YYYY-MM_procesado.xlsx` (1 archivo por mes) |
| **Efectivo** | **No** — `pagos_efectivo.xlsx` es solo el mes actual (se regenera); la trazabilidad de efectivo solo guarda discrepancias entre cobradores, no todos los pagos |

Este módulo **recupera esa capacidad y la unifica** (efectivo + yape), y la expone como
herramientas que un agente puede invocar.

**Justificación (Regla del Tres, ya cumplida):** en una sola sesión se necesitó a mano
4 veces — convenio SALDO=0, reclamos vagos, la pareja G-1/M-15 (cruce por exceso), y
"¿te pagué efectivo en mayo?". El driver es un dolor real y recurrente: **resolver
reclamos**. No es "muchos módulos lo piden".

---

## Qué NO hace (para no confundir con módulos vecinos)

- **No** lleva la deuda por concepto (convenio/multa/acuerdos). Eso es `seguimiento_pueblo`.
  Son complementarios: un pago crudo de 25 en efectivo es **un** evento acá; cómo ese pago
  salda varios conceptos es trabajo de `seguimiento_pueblo`.
- **No** decide cobranza ni saldos del mes. Eso es `5_cobranza`.
- **No** reemplaza el motor de yape. Lo **lee** como fuente (y a futuro lo migra al repo).

---

## Arquitectura — separar escribir de leer, un solo writer

La pieza que hace que esto escale a Postgres y a agentes no es dónde viven los archivos:
es que **haya un repo como writer único**. Todo lo que escribe el store pasa por él.

```
  FUENTES (productores thin · 1 responsabilidad c/u)
  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────────┐
  │ libros viejos     │  │ pagos_efectivo    │  │ yape procesado    │
  │ (Drive, Oct-25→)  │  │ mesa_N (mensual)  │  │ (ya existe)       │
  └────────┬─────────┘  └────────┬─────────┘  └────────┬─────────┘
     importar_libros.py    importar_efectivo.py    importar_yape.py
           └──────────┬──────────┴──────────┬───────────┘
                      ▼                      ▼
              ┌───────────────────────────────────┐
              │  historial_repo.py  (WRITER ÚNICO) │
              │  registrar_pago() · registrar_     │
              │  correccion() · identificar_blanco()│
              └────────────────┬──────────────────┘
                               ▼ append-only
              ┌───────────────────────────────────┐
              │  STORE  shared/reporte_acumulado_  │
              │  procesado/  YYYY-MM_*.xlsx         │
              └───────────────────────────────────┘
                               ▲ read-only
              ┌───────────────────────────────────┐
              │  consultar.py  (LECTORES · TOOLS)  │
              │  analizar_reclamo(predio)          │
              │  reporte_pagos(predio)             │
              └───────────────────────────────────┘
```

| Pieza | Responsabilidad | Por qué escala a agentic SaaS |
|---|---|---|
| `historial_repo.py` | Único que escribe el store | Swap a Postgres = reescribir SOLO este archivo |
| `importar_*.py` | Un productor por fuente | Cada uno es una capability self-documenting |
| `consultar.py` | Solo lee | Tools read-only-safe que el agente invoca sin riesgo |
| STORE | Append-only, nunca se pisa | Auditable y reversible |

---

## El evento — 1 fila = 1 pago crudo

Sin concepto (el reparto por concepto es de `seguimiento_pueblo`).

| Sección | Columnas |
|---|---|
| **¿Quién es?** | `MZ` · `LT` · `NOMBRE` |
| **¿Qué pagó?** | `CANAL` (efectivo/yape) · `MONTO` · `FECHA` (día para efectivo; día+hora:min:seg para yape) · `MES_CICLO` (a qué ciclo se atribuye) |
| **Referencia** (1 columna compacta, nunca vacía) | `REFERENCIA` — efectivo: `"MESA-COBRADOR"` · yape: `ORIGEN` (junto con FECHA completa, identifica el pago de forma única) |
| **Auditoría** | `ESTADO` (identificado/blanco) · `FUENTE` (pago/correccion) · `MOTIVO` · `AUDIT_REF` · `CICLO_CORRECCION` · `ORIGEN_ARCHIVO` (libro+hoja+fila \| run del importador) · `TIMESTAMP` |

### Reglas de negocio

1. **Crudo, sin concepto.** Un pago es un evento; su reparto por concepto no vive acá.
2. **FECHA exacta + MES_CICLO.** Los libros viejos se titulan por rango de fechas que no
   calza un mes calendario (ej. "mayo … 2026-03-11 a 2026-04-10"). La `FECHA` es la verdad;
   `MES_CICLO` es la atribución al ciclo de boleta. Para `yape` la `FECHA` se guarda con
   hora:min:seg — junto con `REFERENCIA` (el `ORIGEN`), es lo que identifica el pago de
   forma única (mismo criterio que ya usa motor_matching).
3. **Yape en blanco.** Pago recibido sin predio → `PREDIO` vacío, `ESTADO=blanco`. Cuando
   se identifica, entra un **evento de identificación** que lo enlaza (no se edita el viejo).
4. **Correcciones = eventos de primera clase.** Reatribución entre predios e identificación
   de blancos se registran con `FUENTE=correccion` + `CICLO_CORRECCION` + `AUDIT_REF`,
   nunca pisando el evento original. Reversibles borrando la fila.
5. **El store nunca se pisa.** Sólo crece. `mesa_N` y demás fuentes quedan intactas: son
   la observación física; la corrección es una decisión posterior que va en su propio evento.

---

## Flujo

### Ingesta (cómo entra un pago)
```
fuente → importar_*.py → historial_repo.registrar_pago(...) → STORE (append)
```

### Consulta (cómo se responde un reclamo)
```
"¿te pagué efectivo en mayo?"  →  consultar.analizar_reclamo(predio)
     → LEE todo el store, filtra por predio
     → respuesta directa  +  reporte_pagos(predio) (mes a mes, para el vecino)
```

### Corrección (mismo writer)
```
reclamo pareja G-1/M-15
  → historial_repo.registrar_correccion(orig=M-15, dest=G-1, 20, motivo, audit_ref)
  → nueva fila FUENTE=correccion · las viejas no se tocan
```

---

## Estructura de carpetas

```
7b_historial_pagos/
├── README.md                 ← fuente de verdad (este archivo)
├── historial_repo.py         ← writer único del store
├── importar_libros.py        ← productor: libros viejos (Drive)
├── importar_efectivo.py      ← productor: pagos_efectivo / mesa_N
├── importar_yape.py          ← productor: yape procesado existente
├── consultar.py              ← lectores read-only (tools del agente)
├── inputs/
│   └── libros/               ← libro_YYYY-MM.xlsx (bajados de Drive)
├── docs/
│   ├── diagrama_flujo_7b_historial_pagos.html
│   ├── arquitectura_7b_historial_pagos.html
│   └── formato_evento_pago.html
└── tests/

STORE (compartido, no se duplica):
shared/reporte_acumulado_procesado/YYYY-MM_*.xlsx
```

---

## Relación con otros módulos

| Módulo | Relación |
|---|---|
| `seguimiento_pueblo` | Complementario. Acá: pago crudo por canal. Allá: deuda por concepto. |
| `4b_reclamos` | Consumidor principal. `analizar_reclamo(predio)` resuelve reclamos. |
| `5_cobranza` | Fuente (pagos del mes) y consumidor futuro (validaciones cruzadas). |
| `4_pagos/yape` | Fuente. Su salida procesada se lee; a futuro escribe vía el repo. |

---

## Visión — local hoy, Postgres mañana, agentic SaaS

```
local hoy       → los eventos viven en Excel append-only (el folder shared)
Postgres mañana → eventos = una tabla; se reescribe SOLO historial_repo.py, los que llaman no cambian
agentic SaaS    → analizar_reclamo(predio), reporte_pagos(predio), reatribuir_pago(...) son tools
                  que un agente invoca cada mes (validaciones proactivas antes del cierre)
```

Guardrail de alcance: se arranca con **1 mes validado** (parser contra dato conocido — el
SALDO convenio ya cruzado), luego batch Oct-2025→hoy con reporte de desviaciones por
archivo. Las validaciones cruzadas (planilla vs medidores) se suman después como queries,
no se diseñan ahora.
