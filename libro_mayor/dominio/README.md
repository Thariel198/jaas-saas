# libro_mayor/dominio — las reglas puras del ledger

> Sub-contexto de `libro_mayor/` (ver `libro_mayor/README.md`). **Fase 1 — spec
> CERRADO** (2026-07-17): las 6 firmas de esta carpeta están decididas. Sin código
> todavía — Fase 2 (implementación) es la siguiente sesión.
>
> **Verdad única de cada decisión:** `docs/decisiones/ledger_fase1_decisiones.md`
> (Bloque A). Este README es el spec consolidado y detallado — coincide con esa
> verdad, no la reemplaza. Material didáctico (diagramas de cajas, opciones A/B,
> el porqué enseñado) → `docs/cuaderno/libro_mayor.html`.

---

## Qué es

La capa que define **qué conceptos existen, cómo se reparte un pago y quién se
corta** — sin tocar ningún archivo. Reglas puras: entra por parámetros, sale por
retorno, nunca abre un `.xlsx` ni sabe de `JASS_ID` como ruta. Se testea con dicts.

```
┌──────────────┬───────────────────────┬───────────────┬─────────────┐
│  QUÉ ENTRA   │  QUÉ HACE (6 firmas)  │   QUÉ SALE    │  QUIÉN LEE  │
├──────────────┼───────────────────────┼───────────────┼─────────────┤
│ nada de disco│ 1 taxonomia           │ valores puros:│ caja_repo   │
│              │ 2 entidades           │  · Reparto    │ cuenta_repo │
│ en memoria:  │ 3 cascada             │  · saldo(int) │ motor       │
│  · cargos    │ 4 politica_corte      │  · ¿corta?    │ tools       │
│  · un abono  │ 5 saldo               │  · CARGO_ID   │             │
│              │ 6 identidad           │ nunca archivo │             │
└──────────────┴───────────────────────┴───────────────┴─────────────┘
```

**Por qué una capa pura:** el `.xlsx` de hoy es un adapter detrás del patrón repo
(lente de escala, `docs/lente_escala.md`) — se reemplaza por Postgres sin que
`dominio/` cambie una línea. La lógica de negocio no se cablea a `openpyxl`.

## Orden de dependencia (por qué se cerraron en este orden)

```
   taxonomia ◄── entidades ◄── cascada ◄── saldo
    ▲                                │
    └──────── politica_corte ◄───────┘ ◄── identidad (transversal)
```

Todo cuelga de `taxonomia` (el diccionario de conceptos) — se cierra primero.
`cascada` necesita las `entidades` (qué es un Cargo, una Aplicación). `saldo`
necesita `cascada` (lee las Aplicaciones que ya produjo). `politica_corte` necesita
`saldo` (su precondición: `meses_impagos`). `identidad` es transversal — la usan
todas las demás para nombrar lo que crean.

---

## Estructura (destino Fase 2)

```
dominio/
├── taxonomia.py       ← el diccionario: conceptos, sub-conceptos, cascada P1-P6
├── entidades.py        ← los tipos: MovimientoCaja, Cargo, Aplicacion, Ajuste,
│                          IdentificacionAbono (frozen, validan al nacer)
├── cascada.py           ← aplicar(cargos_abiertos, abono) → Reparto
├── politica_corte.py    ← evaluar(saldo, meses_impagos, ..., cfg) → (motivo, penalidad, salvado)
├── saldo.py             ← saldo(cargos, aplicaciones) → int · meses_impagos(...) → int
├── identidad.py         ← cargo_id() · abono_id() · canon() — nombres deterministas
└── README.md            ← este archivo
```

---

## 1 · `taxonomia.py` — el diccionario del ledger

Define qué conceptos existen y en qué orden se reparte un pago cuando no alcanza.
Rediseño de `5_cobranza/main.py::_descomponer_saldo` (L1741) — la fuente es el
código, el destino es un rediseño, no una transcripción literal.

**7 conceptos:** `AGUA` · `MANTENIMIENTO` · `CORTE_RECONEXION` · `MULTA` ·
`ACUERDOS` · `CONVENIO` · `OTROS`.

**Cascada de prioridad (6 niveles) — REABIERTA y corregida 2026-07-23, ver CA1:**

```
P1  AGUA (consumo, FIFO por MES_CARGO — incluye lo arrastrado de meses previos)
    MANTENIMIENTO (cuota fija, mismo MES_CARGO)              ambos sin sub, misma prioridad
P2  CORTE_RECONEXION                                          sin sub
P3  CONVENIO       → sub: MEDIDOR primero, luego INSTALACION
P4  ACUERDOS       → sub: TECHADO primero, luego CAMPO
P5  MULTA          → sub: REUNION primero, luego FAENA
P6  OTROS                                                      sin sub · slot residual, sin emisor hoy
```

> Orden viejo (hasta 2026-07-22): P3 MULTA · P4 ACUERDOS · P5 CONVENIO. Se invirtió
> el tramo pueblo — CONVENIO y ACUERDOS pasan antes que MULTA. Ver CA1.

| # | Decisión | Verdad única |
|---|---|---|
| T1 | AGUA vs MANTENIMIENTO | **SEPARADOS** — `AGUA` = consumo actual (mes_actual) · `MANTENIMIENTO` = concepto propio. Cobros distintos (variable vs cuota fija); comparten prioridad P1 porque nacen de la misma boleta mensual |
| T2 | "arrastre" (mes_anterior) | **NO es concepto** — es AGUA de meses previos, resuelta por FIFO (`MES_CARGO`). No hay un 8º balde "arrastre" |
| T3 | P6 OTROS | **SE AGREGA** (aunque hoy no tenga emisor) — slot de escape para una JASS futura con un cobro raro, sin tocar el motor |
| T4 | sub_concepto | **SE AGREGA** — multa: reunión/faena · acuerdos: techado/campo · convenio: medidor/instalación/reactivación (⑬, 24/07/2026). Las fuentes crudas ya traen el split; no hay "sub genérico" permanente |
| T5 | dinero céntimos int, sin TOL | **SÍ** — plata en céntimos enteros, comparación exacta, `TOL` eliminado |

`comunitario` (segregación de `motor_matching`) y `deuda_directiva` (balde
caja-only) **no son conceptos de esta taxonomía** — no generan CARGO, no entran a
la cascada (ver `libro_mayor/caja/README.md` §⑩).

---

## 2 · `entidades.py` — las cosas (sustantivos)

5 entidades `frozen`, validan al construirse (`__post_init__` contra la
taxonomía). Frozen = append-only real (nada se muta después de nacer).

| Entidad | Qué es |
|---|---|
| `MovimientoCaja` | un ingreso o egreso de dinero — lo emite `caja` |
| `Cargo` | una obligación que nace — lo emiten varias fuentes hacia `estado_cuenta` |
| `Aplicacion` | una porción de un abono imputada a un cargo — la deriva el MOTOR |
| `Ajuste` | corrección manual de un cargo — apunta por `CARGO_ID` |
| `IdentificacionAbono` | corrección de identidad del DUEÑO de un abono (no mueve plata) |

| # | Decisión | Verdad única |
|---|---|---|
| E1 | Ajuste apunta al cargo por | **`CARGO_ID`** (único/determinista). La tool recibe la llave humana `(mz, lt, concepto, sub, mes)` y calcula el id adentro |
| E2 | frozen + validan al construirse | **SÍ** — `__post_init__` valida vs taxonomía (concepto/sub) + `monto_centimos > 0` en Cargo y MovimientoCaja (Ajuste/devolución van con signo) |
| E3 | IdentificacionAbono | **entidad propia** (no mueve plata; es corrección de identidad `abono_id → mz·lt·reclamo_id`) |

Todas llevan `recorded_at`. Céntimos `int` en todo monto (T5).

---

## 3 · `cascada.py` — cómo un pago cubre las deudas

**Firma:** `aplicar(cargos_abiertos_de_1_predio, abono) → Reparto(aplicaciones, saldo_a_favor_centimos)`
· `clave_orden(c) = (prioridad, indice_sub, mes_cargo)`.

Rediseño de `5_cobranza/main.py::_descomponer_saldo` (L1741): 5 comps floats →
6 prioridades con sub-orden y FIFO.

| # | Decisión | Verdad única |
|---|---|---|
| CA1 | orden del reparto | **P1-P6 → sub-orden → FIFO por `MES_CARGO`**. `clave_orden=(prioridad, indice_sub, mes_cargo)`. **Reabierta 2026-07-23**: dentro del tramo pueblo el orden pasa a CONVENIO→ACUERDOS→MULTA (antes: MULTA→ACUERDOS→CONVENIO). **Principio** (mismo del sub-orden, subido un nivel): *el dinero cubre primero lo que SOLO el dinero salda* — la MULTA es la deuda más "no-monetaria" (faena se paga con trabajo, reunión se exonera por asistencia — pasa muy seguido); convenio/acuerdos solo se saldan con plata → van antes, multa al final. Origen: caso C-1 — el orden viejo declaraba "pagada" una multa real (reunión+faena, sin ningún abono en la fuente) solo porque el waterfall llegaba a ella primero, tapando que el pago real había sido para convenio/acuerdos. Ver la consecuencia sobre corte en PC6. `5_cobranza/main.py::_descomponer_saldo` (código vivo, pre-ledger) **sigue con el orden viejo** — no se tocó código, solo se corrigió el dominio para el backfill de agosto. Corrección puntual de C-1 vive en `shared/reidentificacion.xlsx` mientras tanto |
| CA2 | sobrante del pago | **`SALDO_A_FAVOR`** (concepto explícito, céntimos). El código viejo lo tiraba (`restante` descartado) |
| CA3 | granularidad | **1 abono, función pura**; el **motor** ordena por fecha e itera N abonos. La cascada no sabe de tiempo ni I/O |
| CA4 | devolución | **fuera de la cascada** — es egreso de caja vs `SALDO_A_FAVOR`, lo maneja el motor. El código viejo la sumaba dentro de P1 agua |

---

## 4 · `politica_corte.py` — ¿este predio se corta?

**Firma:** `evaluar(saldo, meses_impagos, ya_cortado, en_revision, pago_ventana, cfg) → (motivo, penalidad, salvado)`
· `cfg` = manifiesto del tenant por trigger: `{umbral_meses, penalidad_base, permite_salvarse}`.

Rediseño de `6_corte`: `generar_lista` (elegibilidad Día 0) + `seguimiento`
(salvado Día 2) colapsan en **1 función pura**; el motor la llama en ambos días.
Decidida con el lente de escala: los valores no se hardcodean, entran por config;
agua y multa son dos filas de manifiesto, no dos ramas de código.

| # | Decisión | Verdad única |
|---|---|---|
| PC1 | umbral de corte | **conductual "N meses impagos seguidos"**, N=config(tenant, trigger) — agua N=2. El `MES_ANTERIOR≥8` viejo era un proxy (planilla independiente solo veía 1 mes); el ledger da el conteo real. **Precondición:** `meses_impagos` lo deriva `saldo.py` |
| PC1b | qué es "mes impago" | **pagó S/0 ese mes**. Un pago parcial NO cuenta como impago (regla universal) |
| PC2 | qué salva del corte | **cubrir la penalidad** (`pago_ventana ≥ penalidad`). Un parcial menor NO salva. Se retira el gate Día-0 "cualquier pago salva" |
| PC3 | granularidad | **función pura de 1 instante**; el MOTOR la llama Día 0 y Día 2 (misma fn, otro input). No sabe de tiempo ni I/O — análogo a CA3 |
| PC4 | penalidad base vs escalada | **base = config(trigger)**; el motor la escala tras la ventana de gracia. `politica_corte` no sabe de días |
| PC5 | ya_cortado / exoneración | **3 estados**: `activo` \| `cortado` \| `exonerado` (mensual, caduca, motivo obligatorio · o permanente, decisión de junta). Guarda `{tipo, motivo, periodo}` + `JASS_ID` |
| PC6 | multa | **mismo motor, otra cfg**: `umbral_meses=0` (cualquier multa impaga corta), `permite_salvarse=no`, `penalidad_base=20→40` (igual que agua). `6b_corte_multas` = cut-trigger `multa` en el manifiesto, NO módulo aparte. **Consecuencia querida del reorden CA1 (2026-07-23):** con la multa en P5 (última), un pago parcial ya no la cubre → muchos más predios quedan con multa impaga → para un tenant con `multa` como cut-trigger, esto **saca a la luz las multas realmente impagas**. Es intencional, no un efecto a mitigar: muchos vecinos reclaman "ya pagué convenio/acuerdos" (y es cierto) mientras su multa seguía sin pagarse — el orden nuevo deja esa multa visible y elegible para corte |
| PC7 | forma de salida | **motivo = solo el trigger** ∈ {`agua`, `multa`, `""`}; `salvado`=bool. El "por qué se salvó" lo reconstruye el motor de `en_revision`/`pago_ventana`, no viaja en el retorno |

---

## 5 · `saldo.py` — ¿cuánto debe? ¿cuántos meses sin pagar?

**Firma:** `saldo(cargos, aplicaciones) → int céntimos` · `meses_impagos(cargos_agua, aplicaciones) → int`
(la **precondición de PC1**).

Rediseño de `6_corte`, que hoy adivina la conducta con el proxy `MES_ANTERIOR ≥ 8`.
No re-corre la cascada — lee las `Aplicacion` que el motor ya creó y las suma por
cargo.

| # | Decisión | Verdad única |
|---|---|---|
| S1 | qué es "un mes" para el conteo | **AGUA + MANTENIMIENTO** (la boleta del `MES_CARGO`, no un concepto suelto). Impago ⟺ S/0 imputado a ambos cargos del mes |
| S2 | qué es "mes impago" bajo FIFO | **el mes que recibió S/0 imputado**; el mes-frontera parcial NO cuenta (puso algo → enganchado, PC1b exacto) |
| S3 | ¿guardar el conteo o re-derivar? | **re-derivar** (función pura, cero estado). En el pago solo se APILA el evento; el número se calcula al LEER. Igual que el saldo = deuda − Σpagos |
| S4 | `saldo()` total o desglose | **solo el total** (int céntimos = Σcargos − Σaplicaciones, piso 0 por cargo). El desglose por concepto/mes lo arma `estado_cuenta` de los mismos datos |
| S5 | "N meses seguidos" | **gratis por FIFO** — imputar al más viejo primero hace imposible un mes pagado en medio de impagos; los S/0 son la cola contigua. Sin código de consecutividad |

---

## 6 · `identidad.py` — el nombre único de cada hecho

**Firma:** `cargo_id(cargo) → str` · `abono_id(mov) → str` ·
`canon(mz, lt, concepto, sub, mes) → tuple` (normaliza antes de generar el id).

Consolida en la capa pura las reglas fijadas en el contrato del ledger (①
`ABONO_ID`, `CARGO_ID`) y resuelve sus contradicciones. Solo las COSAS que nacen
(`Cargo`, `MovimientoCaja`) tienen id propio; los vínculos (`Aplicacion`,
`Ajuste`, `IdentificacionAbono`) se identifican por el PAR que conectan.

| # | Decisión | Verdad única |
|---|---|---|
| I1 | formato del id | **prefijo legible + shorthash** (`{jass}-{mz}-{lt}-{concepto}-#hash` cargo · `{jass}-{mes}-{canal}-#hash` abono). El `[:8]` es del sufijo hash, no del id entero |
| I2 | canonicalización | **`identidad.py` dueño de un solo `canon()`** (MZ→upper · LT→int · concepto canónico · mes→YYYY-MM). Unifica los `_norm_mz/_norm_lt` hoy repetidos por módulo |
| I3 | ¿el abono lleva el predio? | **no — abono predio-agnóstico** (yape: clave `(jass, origen, timestamp)`; 1 depósito=1 abono aunque pague N predios, el reparto vive en las APLICACIONES). Efectivo SÍ lleva mz/lt (capturado en mesa) |
| I4 | ¿re-identificar cambia el id? | **no** — el `ABONO_ID` sale de canal+ref, nunca del predio. La identidad del DUEÑO se corrige con un HECHO append-only (`identificar_abono`/`reasignar_abono`); el motor lo lee y re-aplica |

**Hallazgo de la sesión que cerró I1-I4 (2026-07-17):** la clave de efectivo
original (`jass, mesa, cobrador, fecha, monto, mz, lt`) es frágil — incluía
`fecha`, campo editable por un humano; un typo re-corrido duplicaba el pago aunque
hubiera hash. Fix: ancla a la **procedencia** (`origen_archivo + fila`) + gatillo
de casi-duplicado con autorización humana al importar. Ver
`libro_mayor/caja/README.md` §① (corrección 2026-07-17).

---

## Qué NO hace `dominio/`

- No abre archivos, no sabe de `.xlsx`, no conoce rutas.
- No decide *cuándo* correr (eso es el motor / los feeders).
- No persiste nada — cada firma es pura, se re-invoca desde el estado que el
  llamador le pasa.
- No sabe de `JASS_ID` como infraestructura (routing, conexión) — lo recibe como
  parámetro, como cualquier otro dato.

## Deuda pendiente (no bloquea el spec, sí bloquea Fase 2 en esos puntos)

- Re-sembrar histórico de MULTA/ACUERDOS/CONVENIO con `SUB_CONCEPTO` real (hoy
  suma faena+reunión, techado+campo, medidor+instalación).
- `formato_aplicacion.html` sin columna `SUB_CONCEPTO` — HTML se actualiza después
  de cerrar todos los READMEs (decisión del usuario), no antes.
- `6_corte/README.md` tiene una sección "diseño destino" (2026-07-16) con una
  firma de `evaluar()` **anterior** a PC1-PC7 (umbral por monto, sin
  `meses_impagos`, sin los 3 estados de PC5) — pendiente de reconciliar contra
  este spec.
