# backfill_ledger — carga inicial del ledger desde el sistema viejo

> **Tarea de desarrollo — Agosto 2026. Fase 2, no empezada (cero código).**
> No es un módulo del pipeline mensual ni parte de `libro_mayor/` (eso es el
> *diseño* del ledger). Esto es la **tarea puntual** de sembrar el ledger nuevo
> con los 9+ meses de historia que ocurrieron antes de que existiera.

---

## Qué es el backfill

El ledger (`libro_mayor/`), cuando se construya, **nace vacío**. Pero desde
Oct-2025 ya hay meses de pagos, deudas y cortes registrados en el sistema viejo
(mesas de `4_pagos`, `seguimiento_pueblo`, etc.). El backfill es el paso de
**carga inicial** que mete esa historia al ledger.

No fotocopia los resultados ya calculados — **siembra los HECHOS crudos**
(abonos + cargos) y deja que el **motor** re-derive todas las aplicaciones con
su link. Corre **una sola vez**; después el ledger corre en vivo mes a mes tras
`7_cierre`.

```
sistema viejo (mesas, seguimiento)  ──►  siembra HECHOS  ──►  motor re-deriva  ──►  ledger poblado
   (con tachones y correcciones)         (abonos, cargos)     (aplicaciones)        y consistente
```

---

## El terreno que ya estamos preparando (julio 2026)

Antes de que exista el backfill, cada corrección manual que hacemos hoy se
guarda en un **precursor durable** en `shared/`. Hoy esos archivos hacen un
trabajo real (los aplica el overlay de `5_cobranza`); mañana son la fuente
limpia que el backfill lee. Un archivo, dos vidas.

| precursor (`shared/`) | qué es | evento de ledger mañana |
|---|---|---|
| `abonos_rezagados.xlsx` | un abono que **faltaba** (nunca entró a caja) | `registrar_movimiento` (abono nuevo) |
| `reidentificacion.xlsx` | un abono en el **lote equivocado** | `reasignar_abono` |
| `devoluciones_aplicadas.xlsx` | un **exceso** dirigido a otro concepto | `SALDO_A_FAVOR` → motor lo imputa |
| `aportes_tanque_manuales.xlsx` | un abono **voluntario** mal ubicado *(pendiente crear + cablear)* | `registrar_movimiento` BALDE=tanque |

Detalle visual y ejemplos → `docs/cuaderno_backfill.html`.

---

## Casos que el backfill tendrá que resolver (descubiertos en julio)

- **Bloque mixto sin partir** (C1-17): un depósito de junio de S/218.50 anotado
  en C1-9 sin concepto = 18.5 agua + 200 tanque. El backfill lo **parte en 2
  abonos por balde**, validado por chequeo de suma (`18.5 + 200 == 218.50`).
  Requiere agregar el `BALDE` a la llave del `ABONO_ID` (extensión chica del
  contrato, aún no decidida). Ver cuaderno.
- **Fecha real vs. fecha de caja** (abonos rezagados): el vecino pagó en junio,
  la plata entró a la JASS en julio (efectivo). El abono se ancla a junio para
  saldar la deuda de junio; el canal/fecha de caja es julio. Ver `docs`.

---

## Estado

- **Precursores:** 3 de 4 codificados y funcionando (falta `aportes_tanque_manuales`).
- **Backfill:** sin empezar. Es la tarea de agosto.
- **Depende de:** que `libro_mayor/` Fase 2 exista primero (caja_repo, motor,
  importadores) — el backfill los usa.

---

## Artefactos

```
backfill_ledger/
├── README.md                        ← este archivo
└── docs/
    └── cuaderno_backfill.html        ← cuaderno de aprendizaje (diagramas de cajas)
```
