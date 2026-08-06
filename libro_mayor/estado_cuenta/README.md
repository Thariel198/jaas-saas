# libro_mayor/estado_cuenta — Cuenta corriente del predio (cargos + aplicaciones + motor)

> Agregado del bounded context `libro_mayor/` (no es un módulo del pipeline). Ver `libro_mayor/README.md`.
>
> **Estado:** Fase 1 — spec cerrado con `libro_mayor/caja`.
> **Migración:** rediseña y expande `shared/seguimiento_pueblo.xlsx` /
> `seguimiento_repo.py` (hoy 3 conceptos, hecho en horas sin diseño).

---

## Qué hace

Registro **append-only** de la deuda de cada predio y cómo se cubre:

- **CARGO:** una obligación nace (consumo del mes, corte, multa, acuerdo, cuota de
  convenio, otros). Lo emiten varias fuentes; todas escriben acá.
- **APLICACIÓN:** una porción de un abono de `caja` se imputa a un cargo, siguiendo
  la escalera de prioridad. **La deriva el MOTOR DE APLICACIÓN** — referencia el
  `ABONO_ID`, no recopia la plata.
- **AJUSTE:** corrección manual (nunca edita un evento pasado).
- **SALDO** = Σcargos − Σaplicaciones ± Σajustes — siempre derivado, nunca celda a mano.

Responde la historia que un reclamo necesita: *"tu pago fue a consumo y convenio
(medidor), por eso aún debes la multa"* — el orden nuevo (2026-07-23, dominio CA1)
cubre convenio/acuerdos antes que la multa, así el sistema no declara "pagada" una
multa que el vecino nunca cubrió.

## Por qué existe (y por qué se rediseña)

`seguimiento_pueblo` nació horas antes de un cobro, sin diseño: solo 3 conceptos
(MULTA/ACUERDOS/CONVENIO), sin consumo/corte, con las aplicaciones **sin link al
pago** de caja, y con el histórico pre-junio colapsado en un CARGO de génesis. Para
que un **agente** resuelva reclamos solo, la cuenta corriente tiene que estar completa,
linkeada a la caja, y separar el HECHO (cargo) de la INTERPRETACIÓN (aplicación).

## Cuándo corre

- **CARGOS:** los emiten `2_planilla` (consumo/corte) y `obligaciones` (multa/acuerdo/
  convenio, incl. convenio instalación). `deuda_directiva` NO es cargo — es balde caja-only (⑩).
- **APLICACIONES:** las produce el **motor de aplicación** on-demand / post-cierre,
  leyendo cargos abiertos + abonos de `caja`.
- **AJUSTES / resolución de reclamos:** `4b_reclamos` invoca `registrar_ajuste`
  (cargo incorrecto) o las tools de caja `identificar_abono` / `reasignar_abono`
  (blanco reclamado / pago mal atribuido) al autorizar; el motor re-corre y
  re-deriva las aplicaciones (decisión ⑨).
- **Tools de lectura:** on-demand, cuando el agente investiga un reclamo.

## Conceptos y prioridad

```
P1 AGUA·MANTENIMIENTO·arrastre → P2 CORTE_RECONEXION → P3 CONVENIO(medidor·instalación·reactivación)
   → P4 ACUERDOS(techado·campo) → P5 MULTA(reunión·faena)
```
Orden corregido 2026-07-23 (ver dominio CA1): el tramo pueblo pasó de MULTA→ACUERDOS→CONVENIO
a CONVENIO→ACUERDOS→MULTA. El vecino que pagó sin especificar concepto: su pago cae primero
en agua/corte, después convenio/acuerdos, y la MULTA queda al final — así un pago que en
realidad fue para convenio/acuerdos no "paga" por error una multa que nadie cubrió (caso C-1).
El motor lo resuelve fila por fila, incluido el sub-orden dentro de cada concepto (ver
"Taxonomía de conceptos y cascada de prioridad ⑪" en el contrato, abajo).

## Estructura

```
estado_cuenta/
├── cuenta_repo.py                 ← WRITER ÚNICO de cargos/aplicaciones/ajustes
│                                     (evoluciona seguimiento_repo.py)
│                                     registrar_cargo · registrar_aplicacion · registrar_ajuste
├── motor_aplicacion.py            ← EL MOTOR: aplicar(cargos, abonos) → aplicaciones
│                                     (única pieza que llena ABONO_ID + MES_CARGO)
├── tools/                         ← lo que el AGENTE invoca (read-only)
│   ├── estado_cuenta.py           ← panorama completo de un predio
│   ├── explicar_reclamo.py        ← la historia de un pago: a dónde fue cada sol
│   ├── extracto_predio.py         ← ⑫ filas de N meses (deuda por concepto/sub +
│   │                                 distribución de cada pago) — datos, no PDF
│   └── auditoria_conservacion.py  ← ¿entró == registrado? por predio y global
├── vista/                         ← vista ancha + PDF imprimible para la mesa
└── docs/
    └── formato_extracto.html      ← ⑫ contrato visual del extracto (3 secciones)
```

El ledger-repo y el motor quedan accesibles como recurso compartido (los cargos los
emiten `2_planilla` y `5_cobranza`); las **tools** y el reporte viven acá, que es lo
que el agente usa.

## Reglas de negocio

- **Writer único:** solo `cuenta_repo.py` escribe el store.
- **Append-only** · **SALDO siempre derivado**, nunca mutable.
- **Solo el motor aplica:** los feeders emiten cargos; el motor deriva aplicaciones.
- **Toda APLICACIÓN referencia un `ABONO_ID` de `caja`** — sin ese link no hay
  conservación auditable.
- **`JASS_ID` en cada evento** (ver decisión ⑦).
- **Idempotencia** del motor por `(JASS_ID, ABONO_ID, CARGO)`.

## Lo que NO hace

- **No guarda el hecho de caja** (cuánto entró, por qué canal). Eso es `caja` —
  estado_cuenta solo lo referencia por `ABONO_ID`.
- No decide cuánta plata entró; parte de que el abono ya existe en caja.

---

<!-- ══════════════════════════════════════════════════════════════════════════ -->
<!-- CONTRATO DE INTERFAZ — IDÉNTICO en caja/README.md y estado_cuenta/README.md -->
<!-- Si cambia acá, cambia allá en la MISMA pasada. Es el punto de cero-discrepancia. -->
<!-- ══════════════════════════════════════════════════════════════════════════ -->

## CONTRATO DE INTERFAZ · el ledger (caja · estado_cuenta · motor de aplicación)

> Esta sección es **byte-idéntica** en los README de ambos módulos.

### Principio rector: HECHO vs INTERPRETACIÓN

```
   HECHOS (inmutables, cada fuente escribe el suyo)      INTERPRETACIÓN (derivada, un solo motor)
   ┌─────────────────────────────┐
   │ ABONOS  (entró plata)       │ ── caja            ┌──────────────────────────┐
   │  quién, cuánto, canal, ref  │                      │  MOTOR DE APLICACIÓN      │
   └─────────────────────────────┘                      │  aplicar(cargos, abonos)  │
                                          ──────────►    │  camina prioridad · FIFO  │ ── estado_cuenta
   ┌─────────────────────────────┐                      │  → APLICACIONES           │
   │ CARGOS  (nació deuda)       │ ── 2_planilla        │    (abono_id → cargo_id,  │
   │  concepto, mes, monto       │    obligaciones      │     monto_aplicado)       │
   └─────────────────────────────┘    6_corte           └──────────────────────────┘
```

**Regla de oro:** cada módulo emite solo su HECHO (abono o cargo). **Nadie "aplica"
salvo un único motor** que ve los dos lados. Esto traduce 1:1 a un ledger en Postgres
(tablas `abonos`, `cargos`, `aplicaciones`) y es el substrato correcto para tools de
agente idempotentes.

### Multi-tenant · `JASS_ID` en todo (decisión ⑦)

Cada evento del ledger — abono, cargo, aplicación — lleva `JASS_ID`. Hoy siempre
`tupac_amaru`; mañana es la clave que separa 25.000 JASS. El núcleo es
**tenant-agnóstico**: nada de rutas fijas, "la JASS" ni globals; lo específico de cada
JASS (manzanas, tarifa, conceptos activos) entra por **config**, no por fork de código.
Toda clave natural, todo id y toda query incluyen `JASS_ID`.

### Representación de dinero y campos comunes (2026-07-14)

- **Todo monto se almacena como `int` de céntimos** (`monto_centimos`), no `float`.
  Los `100.00` / `40.00` de las tablas de abajo son **display**; internamente son
  `10000` / `4000`. Suma y comparación exactas → **desaparece el `TOL=0.005`** que el
  código de hoy arrastra. Los feeders convierten a céntimos en el borde de entrada.
- **Todo evento del ledger lleva `recorded_at`** (cuándo se asentó, ≠ `FECHA` / `MES_CARGO`,
  que es cuándo ocurrió el hecho) — para orden estable y depuración de idempotencia en el
  backfill. No se repite en cada tabla de abajo; aplica a movimiento, cargo, aplicación,
  ajuste e identificación.

### Entidad 1 · MOVIMIENTO DE CAJA — lo emite `caja`

Un evento de dinero: **ingreso** (abono) o **egreso** (devolución / gasto). "Abono"
es el caso `DIRECCION=INGRESO`; la caja los registra a todos con el mismo esquema
para responder tesorería del mes (cuánto entró, cuánto salió) y reproducir los
baldes de `5b_validacion` desde una sola fuente.

| Campo | Ejemplo | Nota |
|---|---|---|
| `JASS_ID` | `tupac_amaru` | tenant |
| `ABONO_ID` · `DEVOLUCION_ID` · `GASTO_ID` | `tupac_amaru-2026-05-YA-a3f9c1` | id **determinista** del movimiento — `ABONO_ID` (ingreso) · `DEVOLUCION_ID` / `GASTO_ID` (egreso); mismo esquema |
| `CANAL` | `efectivo` / `yape` | |
| `DIRECCION` | `INGRESO` / `EGRESO` | el **signo** de tesorería (antes `TIPO`) |
| `BALDE` | `agua` · `tanque` · `deuda_directiva` · `devolucion` · `retorno` · `honorario` · `gasto` | naturaleza del movimiento — **≠ `CONCEPTO` de deuda** (ver abajo) |
| `DESTINO` | `PREDIO` / `CONCEPTO` / `PENDIENTE` | a qué se dirige |
| `MONTO` | `100.00` | **el depósito completo** (yape: `MONTO_PAGO`, no lo asignado); el signo lo da `DIRECCION` |
| `FECHA` | `05/05/2026` (yape: + hh:mm:ss) | |
| `REFERENCIA` | `MESA-Wilder` / origen yape | 1 columna, nunca vacía |
| `MES_CICLO` | `2026-05` | |
| `MZ` · `LT` | `C` · `43` / vacío | solo si `DESTINO=PREDIO`. Blanco = `DESTINO=PENDIENTE` (sin predio aún) |

**① `ABONO_ID` = `{JASS_ID}-{MES_CICLO}-{CANAL}-{shorthash(clave_natural)}`** —
determinista, NO secuencial, NO aleatorio (los ids de egreso siguen el mismo esquema). Clave natural:
- efectivo: `(jass, origen_archivo, fila)` — **procedencia**, no contenido (ver corrección abajo)
- yape: `(jass, origen, timestamp)` — **predio-agnóstico** (la identidad que ya usa `motor_matching`)

**Hueco 1 resuelto:** un yape de S/100 que paga dos predios es **un** abono de S/100
(no dos con id colisionado). El predio autoritativo y el reparto viven en las
APLICACIONES, no en la caja. `MZ/LT` en el abono es informativo (predio declarado).
Determinista → re-sembrar un mes en orden libre mapea al mismo id → import idempotente.

> **CORRECCIÓN (2026-07-17) — la clave de efectivo original era frágil.** Incluía
> `fecha`, un campo que tipea un humano: un typo en la fecha, re-corrido, generaba un
> `ABONO_ID` distinto para el **mismo** pago → duplicado en el ledger (bug real
> observado, no hipotético). El hash no protege si la identidad misma incluye un campo
> editable. **Fix:** ancla a la **procedencia** (`origen_archivo + fila`) — estable
> aunque se corrija la fecha en la fuente, porque sigue siendo la misma fila.
> `mesa/cobrador/fecha/monto/mz/lt` pasan a ser **atributos** del abono, no parte de
> su identidad. Yape no tiene este problema: su clave (`origen, timestamp`) la fija
> el banco, no un humano.
>
> **Gatillo de casi-duplicado (autorización humana):** al importar efectivo, si un
> pago nuevo coincide en predio + monto + mesa con uno ya existente y la fecha cae
> dentro de una ventana de pocos días, el importador **no lo inserta solo** — lo
> marca "¿posible duplicado?" para que un humano confirme o descarte antes de
> asentarlo. Detalle → `docs/decisiones/ledger_fase1_decisiones.md` (Bloque A,
> `identidad.py`, decisión I-contexto).

### Baldes y auditoría de tesorería (⑩)

Cada movimiento cae en un **balde** según su `DIRECCION` y su `DESTINO`. Es lo que
hoy `5b_validacion` re-deriva con filtros dispersos (`main.py:552`, Nivel 1a/1b) —
la caja lo lleva **explícito**, así la validación pasa a ser una query por balde en
vez de 8 lecturas con filtros, y desaparecen los falsos descuadres por plata sin balde.

```
                A PREDIO (MZ-LT)          POR CONCEPTO (balde nombrado, sin predio)
INGRESO (+)     agua                       tanque · deuda_directiva
  (TE PAGÓ)     blancos → DESTINO=PENDIENTE (concepto libre futuro)
EGRESO  (−)     devolucion · retorno       honorario · gasto  (GASTO institucional)
  (PAGASTE)
```

> **CORRECCIÓN (2026-07-13) — dos errores de dominio del contrato anterior, verificados
> contra código antes de escribir esta versión:**
>
> 1. **`comunitario` NO es un balde.** Colisión de nombres: "tanque comunitario"
>    (adjetivo — el tanque es propiedad de la comunidad, `5b_validacion/main.py:560`)
>    se confundió con `CONCEPTO=comunitario` de `motor_matching` (readme_motor
>    líneas 312-315): un **mecanismo de segregación** — un cobrador presta su Yape,
>    agrega el cobro de varios vecinos y envía un depósito único; el sistema no sabe
>    a qué lotes va, se marca `comunitario` y se **desgloza por lote** en la hoja
>    `Segregacion` (`PADRE_SEGREGADO` → N × `HIJO_SEGREGADO`). Un depósito
>    `comunitario` sigue siendo un INGRESO normal, `BALDE=agua` (mayormente) — ya lo
>    cubre la decisión ① (`ABONO_ID` sin mz/lt) + el motor lo reparte en N
>    aplicaciones. **Efecto en `buscar_abono`:** debe buscar también en los
>    `HIJO_SEGREGADO` de `motor_matching`, no solo en depósitos de primer nivel —
>    el pago de agua de un vecino puede estar enterrado dentro del depósito de
>    su cobrador.
> 2. **`deuda_directiva` NO cruza a `estado_cuenta`.** Es un caso específico: dos
>    miembros de la directiva anterior repagando un faltante de caja detectado —
>    no es deuda de ningún predio. Mismo tratamiento que `tanque`: balde de
>    INGRESO caja-only, ya reconciliado en `5b_validacion` como "otros conceptos"
>    (Nivel 1a, `_cargar_otros_conceptos`). Nunca genera CARGO, nunca entra a la
>    cascada de prioridad (ver ⑪, abajo).

**Qué balde cruza a `estado_cuenta` (toca la deuda del predio) y qué se queda solo en caja:**

| Balde | Dir. | ¿Cruza a estado_cuenta? | Cómo |
|---|---|---|---|
| `agua` | + | **Sí** | el motor lo aplica a cargos por concepto/mes (cascada P1-P6, ⑪) |
| `blancos` | + | tras identificarse (⑨) | queda PENDIENTE hasta que un reclamo le asigna predio |
| `tanque` · `deuda_directiva` | + | **No** | **aporte/reposición voluntaria** — no genera cargo, vive solo en caja |
| `devolucion` · `retorno` | − | **Sí** | baja `SALDO_A_FAVOR` del predio (④) |
| `honorario` · `gasto` | − | **No** | egreso institucional — solo tesorería, no toca ningún predio |

**Vocabulario — no confundir:**
- **`BALDE`** (caja) = naturaleza del *movimiento de dinero* (agua, tanque, honorario…).
- **`CONCEPTO`** (estado_cuenta) = tipo de *cargo/deuda* (AGUA, MULTA, CONVENIO…).
  Un movimiento de balde `agua` se aplica a cargos de concepto `AGUA`; son
  vocabularios distintos que el motor conecta. `comunitario` no es vocabulario de
  ninguno de los dos — es un mecanismo de segregación aguas arriba, en `motor_matching`.

**Auditoría de tesorería del mes:**
```
ENTRÓ = Σ MONTO donde DIRECCION=INGRESO     SALIÓ = Σ MONTO donde DIRECCION=EGRESO
NETO  = ENTRÓ − SALIÓ            desglose:  GROUP BY BALDE
```
Reproduce la validación de `5b` (agua + blancos + tanque + otros = crudo TE PAGÓ ·
devolución + retorno + gastos = crudo PAGASTE) desde el ledger, no desde 8 archivos.

### Entidad 2 · CARGO — lo emiten varias fuentes hacia `estado_cuenta`

| Campo | Ejemplo | Nota |
|---|---|---|
| `JASS_ID` | `tupac_amaru` | tenant |
| `CARGO_ID` | `tupac_amaru-C-43-multa-a1b2c3d4` | **determinista siempre** = `sha256[:8]` de la clave natural canónica `(JASS_ID, MZ, LT, CONCEPTO, SUB_CONCEPTO, MES_CARGO)` — misma regla que `ABONO_ID`, sin identidad condicional |
| `MZ` · `LT` | `C` · `43` | predio |
| `CONCEPTO` | `MULTA` | agua·mantenimiento·corte_reconexion·multa·acuerdos·convenio·**otros** |
| `SUB_CONCEPTO` | `faena` | **⑪** desglose dentro del concepto — ver taxonomía completa abajo; vacío si el concepto no tiene sub (agua/mantenimiento/corte_reconexion/otros) |
| `MES_CARGO` | `2026-03` | **② el mes en que nació la deuda** |
| `MONTO` | `40.00` | monto de la obligación |
| `SOURCE` | `obligaciones` | quién lo emitió (`2_planilla`·`6_corte`·`obligaciones`); `deuda_directiva` NO es cargo (caja-only, ⑩) |

**Fuentes de cargo (Hueco 4 resuelto — una sola cuenta corriente):**

| SOURCE | Emite concepto(s) | Fase |
|---|---|---|
| `2_planilla` | agua · mantenimiento | 2 |
| `6_corte` | corte_reconexion (penalidad, evento de corte) | 2 |
| `obligaciones` | multa (reunión/faena) · acuerdos (techado/campo) · convenio (medidor/instalación/reactivación, unifica el viejo `arrastre_consolidado`) | 2 |

Toda deuda de un predio es un CARGO en `estado_cuenta`, venga de donde venga. El hack
`PREDIOS_INSTALACION_EXCLUIDOS` (hoy en la siembra) desaparece: el cargo de instalación
ahora existe, el pago se aplica normal. **`deuda_directiva` NO figura acá** — no es
deuda de predio, es balde caja-only (ver corrección arriba).

### Entidad 3 · APLICACIÓN — la deriva el MOTOR DE APLICACIÓN (dentro de `estado_cuenta`)

| Campo | Ejemplo | Nota |
|---|---|---|
| `JASS_ID` | `tupac_amaru` | tenant |
| `ABONO_ID` | `tupac_amaru-2026-05-YA-a3f9c1` | FK al abono de caja (o `DEVOLUCION_ID`) |
| `CARGO` | `(MULTA, faena, 2026-03)` | el cargo que salda `(CONCEPTO, SUB_CONCEPTO, MES_CARGO)` o `CARGO_ID` |
| `MONTO_APLICADO` | `40.00` | porción del abono imputada a este cargo |

**⑥ El MOTOR DE APLICACIÓN es responsabilidad propia (Hueco 2 resuelto).** Es una
función pura / tool: `aplicar(cargos_abiertos, abonos) → aplicaciones`, que camina la
prioridad y reparte FIFO por `MES_CARGO`. **Ve los dos lados**, por eso es el único que
puede llenar `ABONO_ID` + `MES_CARGO`. Los emisores (`2_planilla`, `obligaciones`) **solo
emiten cargos**; nadie más aplica. `5_cobranza` **se disuelve** — nunca aplicó ni emitió
cargos (el código confirma: solo `registrar_pago`/`registrar_ajuste`). Idempotente por
`(ABONO_ID, CARGO)` → re-corrible sin duplicar.

**② La aplicación referencia el CARGO concreto** vía `(CONCEPTO, SUB_CONCEPTO, MES_CARGO)`
— un abono de mayo puede saldar la multa de marzo; sin `MES_CARGO` la tool no sabe qué
deuda cubrió, y sin `SUB_CONCEPTO` no distingue faena de reunión dentro de la misma multa.

**③ `SALDO_A_FAVOR` es un CONCEPTO explícito**, no un residual. Cuando un abono supera
todos los cargos abiertos, el sobrante se imputa a `CONCEPTO=SALDO_A_FAVOR` (crédito del
predio). El dinero no aplicado es una FILA que se consulta y se suma.

**④ La DEVOLUCION baja el `SALDO_A_FAVOR`** (balance corrido, FIFO — la plata parqueada
es fungible). caja registra la salida (`DEVOLUCION_ID`); el motor anota una aplicación
`(DEVOLUCION_ID, SALDO_A_FAVOR, −monto)`. Si excede el saldo a favor, la auditoría lo marca.

### Resolución de reclamos → hechos al ledger (⑨ · 4b_reclamos)

Un reclamo de dinero ("ya pagué mayo", "ese pago no era mío", "me cobraron de
más") lo gestiona `4b_reclamos`, pero **4b no escribe aplicaciones ni toca el
saldo**. Al autorizar la resolución emite el HECHO que faltaba e invoca una tool
del ledger; el MOTOR re-corre y deriva la aplicación. La deuda baja como
consecuencia, nunca por edición directa de una celda.

| Reclamo resuelto | Tool que invoca 4b | Escribe en | Efecto tras re-correr el motor |
|---|---|---|---|
| Blanco reclamado (el pago sí es suyo) | `identificar_abono(abono_id, mz, lt, reclamo_id)` | caja — evento IDENTIFICACIÓN (append-only) | el abono deja de ser blanco → se aplica a los cargos del predio |
| Pago mal atribuido (fue a otro predio) | `reasignar_abono(abono_id, mz, lt, reclamo_id)` | caja — re-identificación (append-only) | revierte la aplicación vieja y aplica al predio correcto |
| Cargo incorrecto (le cobraron de más) | `registrar_ajuste(mz, lt, concepto, monto, reclamo_id)` | estado_cuenta — AJUSTE (append-only) | el cargo baja → el saldo se corrige |

**⚠ Pendiente — `reasignar_abono` sobre abono PARTIDO (detectado 25/07/2026, caso
M-15/G-1).** El contrato de arriba describe `reasignar_abono` a nivel de abono
completo ("revierte la aplicación vieja y aplica al predio correcto"). No cubre el
caso donde solo una PARTE del abono está mal atribuida: M-15 pagó S/28 (mesa_1,
04/07) — S/8 es deuda real suya (consumo+mantenimiento), S/20 le correspondían a
G-1/MULTA (mismo cobrador/día, error de mesa, comentario "Cambiar la multa de
G1-M15"). El abono NO se puede reasignar completo — M-15 sí debía los S/8.
Hipótesis de diseño (no cerrada): como ③ ya hace explícito `SALDO_A_FAVOR` como
CONCEPTO propio, el motor aplicaría normal (S/8→cargo real de M-15, S/20→
`SALDO_A_FAVOR` de M-15) y `reasignar_abono` apuntaría a ESA fila de
`SALDO_A_FAVOR`, no al abono original — evita partir el abono a mano. Mismo tipo
de hueco que el caso C1-17 (bloque partido, ver
`backfill_ledger/docs/cuaderno_backfill.html`, lámina 3). Precursor manual de hoy:
`shared/reidentificacion.xlsx` (columna `CONCEPTO_DESTINO`) — solo sabe sumar
crédito al destino (G-1); el origen (M-15) no se corrige solo, queda como
`resuelto` a mano en `arrastre_devolucion` con nota explicando a dónde fue la plata.

**⑭ Primer precursor real de `registrar_ajuste` (24/07/2026):** `shared/ajustes_cargo.xlsx`
— un CARGO nació (penalidad de corte, F1-4, junio) y se determinó que no
correspondía (reclamo confirmado, revertido en el audit del módulo que lo generó).
Distinto de `SALDO_A_FAVOR` (③): acá no hubo plata de más, el cargo mismo se anula.
`REF_AUDIT` ancla el hecho original (fila APLICADO en `6_corte/outputs/audit_penalidad.xlsx`)
para que el backfill cuente los dos hechos reales — nace en `MES_ANO_ORIGEN`, se anula
en `MES_ANO_APLICA` — sin perder cuál módulo generó el cargo ni cuándo.

**⑮ Primer precursor real de `registrar_cargo` tardío (24/07/2026):** `shared/genesis_tardia.xlsx`
— 6 predios pagaron "techado y campo" (cuota de asamblea, S/75 c/u) cobrado en mesa,
pero la fuente en `obligaciones/inputs/DEUDORES Y PAGOS DEL TECHADO Y CAMPO.xlsx`
nunca los tenía — génesis nunca sembrada. Distinto de `ajustes_cargo` (⑭): acá el
CARGO es **legítimo**, solo llegó tarde a un ciclo ya congelado. La fuente real ya
se corrigió (misma fecha) — el precursor solo tapa el ciclo que quedó frozen antes
de la corrección; el ciclo siguiente lee la fuente arreglada directo, sin necesitar
esta fila. `REF_FUENTE` ancla dónde quedó corregida la fuente.

**⑯ `ajustes_cargo` se extiende a AGUA (24/07/2026):** caso A1-13 — su consumo de
julio nació como 38 m³ pero el predio paga tarifa mínima siempre; la secretaria dio
la orden directa de cobrarle 5 (no es disputa del usuario, es corrección
administrativa). Mismo mecanismo que ⑭ (F1-4/corte), pero primera vez que el
CONCEPTO es `AGUA` en vez de `CORTE_RECONEXION` — se agregó `AGUA→mes_actual` y
`MANTENIMIENTO→mantenimiento` a `_CONCEPTO_DEVOLUCION_A_CAMPO` (antes solo cubría
convenio/multa/acuerdos/corte). A1-13 combina ⑮+⑯: techado/campo (75) + consumo
corregido (38→5) + mantenimiento (3) = 83, exacto lo que pagó → `SALDO=0`.

- **Un blanco es un abono con `DESTINO=PENDIENTE`** (entró plata, sin predio).
  Identificarlo NO edita el abono (append-only) — agrega un evento que le asigna
  predio; recién ahí el motor puede aplicarlo. El abono viejo (mayo) se aplica a
  los cargos abiertos por prioridad FIFO, no necesariamente al cargo de su mes.
- **4b no es writer del ledger:** solo invoca las tools (writer único intacto).
  Toda aplicación resultante linkea `ABONO_ID` **y** `reclamo_id` — dice de qué
  pago salió y qué reclamo la autorizó. Idempotente: re-correr el motor con la
  misma identificación produce la misma aplicación.
- **La columna `BLANCO`/`DEVOLUCION` de la planilla se retira** (Fase 2): el
  descuento deja de ser una celda manual regenerable (que se pisa al regenerar y
  nunca cuadra en 5b) y pasa a ser una aplicación auditable. Cierra el balde:
  `Σ blancos = Σ aplicados + Σ pendientes`.
- **La boleta refleja la corrección leyendo estado_cuenta** (saldo derivado) + una
  línea "pago reconocido", no una columna de descuento (pendiente `3_boletas`, Fase 2).
- **No confundir con la DEVOLUCION real** (decisión ④, sale plata por yape/efectivo):
  identificar un blanco NO mueve dinero — el pago ya había entrado.

### Backfill histórico (Hueco 3 resuelto — no se recupera, se re-deriva)

```
  backfill        ──► caja: ABONOS históricos (id determinista) ─┐
  migrar segui.   ──► estado_cuenta: CARGOS históricos                        │─► motor aplicar()
  transcribir     ──► DEVOLUCIONES del libro (no se re-derivan)    ─┘   → aplicaciones linkeadas ✓
```

No hay que inventarle `ABONO_ID` a un pago viejo: se siembran los HECHOS (abonos +
cargos) y el **mismo motor** re-deriva todas las aplicaciones con su link. Las
aplicaciones viejas de `seguimiento` (solo "total por concepto") se descartan.
Excepción: las devoluciones se transcriben del libro (cambio de régimen — antes se
devolvía, ahora se acumula; re-derivarlas mentiría sobre la resolución del exceso).

### Invariante de conservación

```
  Por cada (JASS_ID, ABONO_ID):
     MONTO(abono) = Σ MONTO_APLICADO a conceptos  +  SALDO_A_FAVOR generado

  Global por JASS_ID:
     Σ abonos = Σ aplicado + Σ saldo_a_favor_vigente + Σ devoluciones

  Se expone como tool: auditoria_conservacion() — el agente la corre antes y
  después de cualquier corrección para auto-verificarse.
```

### Taxonomía de conceptos y cascada de prioridad (⑪)

Cuando un pago no alcanza a cubrir toda la deuda, el motor reparte por una escalera
de prioridad. El código viejo (`5_cobranza/main.py::_descomponer_saldo()`, P1→P5)
usaba MULTA→ACUERDOS→CONVENIO; el orden se **corrigió 2026-07-23** (ver dominio CA1) —
el motor de aplicación replica el orden nuevo **agregando P6 OTROS**, con `SUB_CONCEPTO`:

```
P1  AGUA (consumo, FIFO por MES_CARGO — incluye lo arrastrado de meses previos)
    MANTENIMIENTO (cuota fija, mismo MES_CARGO)              ambos sin sub, misma prioridad
P2  CORTE_RECONEXION                                          sin sub
P3  CONVENIO       → sub: MEDIDOR primero, luego INSTALACION
P4  ACUERDOS       → sub: TECHADO primero, luego CAMPO
P5  MULTA          → sub: REUNION primero, luego FAENA
P6  OTROS                                                      sin sub · slot residual, sin emisor hoy (especulativo)
```

> Orden viejo (hasta 2026-07-22): P3 MULTA · P4 ACUERDOS · P5 CONVENIO. Se invirtió el
> tramo pueblo para que un pago sin concepto declarado cubra primero convenio/acuerdos
> y deje la MULTA impaga al final (evita marcar como "pagada" una multa que nadie pagó —
> caso C-1). El sub-orden dentro de cada concepto NO cambió. `_descomponer_saldo` (código
> vivo, pre-ledger) sigue con el orden viejo; el orden nuevo rige para el backfill.

**`arrastre` no es un concepto propio** (taxonomía T2 del dominio): es AGUA con
`MES_CARGO` de meses anteriores — el mismo concepto, distinto mes, resuelto por FIFO.
No hay un 8º balde "arrastre"; el nombre solo describe que ese `MES_CARGO` quedó
viejo. AGUA y MANTENIMIENTO son dos **conceptos separados** (T1) que comparten
prioridad P1 porque nacen de la misma boleta mensual.

**Por qué el tramo pueblo va CONVENIO → ACUERDOS → MULTA** (nivel concepto, corregido
2026-07-23): el mismo principio del sub-orden, subido un nivel — *el dinero cubre primero
lo que SOLO el dinero puede saldar*. La **MULTA es la deuda más "no-monetaria"**: la faena
se paga con trabajo y la reunión se exonera por asistencia (pasa muy seguido). CONVENIO
(medidor/instalación/reactivación) y ACUERDOS (techado/campo) **solo se saldan con plata**. Por eso un
pago escaso cubre primero convenio/acuerdos y deja la multa al final — que es además lo más
probable de terminar trabajado o perdonado. Efecto de negocio: un pago sin concepto
declarado ya no marca como "pagada" una multa que el vecino nunca cubrió (caso C-1).

**Por qué ese orden dentro de cada concepto** (regla de negocio, no técnica):

| Concepto | 1º | 2º | 3º | Por qué |
|---|---|---|---|---|
| CONVENIO | medidor | instalación | reactivación | medidor es deuda chica; instalación es grande y se paga de a pocos; reactivación (⑬) es la más vieja e incierta en monto — va última, mismo principio que instalación llevado al extremo |
| ACUERDOS | techado | campo | — | techado es el monto más bajo — se salda primero |
| MULTA | reunión | faena | — | la faena se puede pagar con trabajo (doblar a 8h en vez de 4h); la reunión **nunca** se paga con trabajo — el dinero debe cubrir primero lo que solo el dinero puede cubrir |

**`DEUDA_DIRECTIVA` NO entra a esta cascada.** No es deuda de predio — es el caso
específico de dos miembros de la directiva anterior repagando un faltante de caja.
Mismo tratamiento que `tanque`: balde de caja, nunca genera CARGO (ver corrección
en "Baldes y auditoría de tesorería", arriba).

**El desglose `SUB_CONCEPTO` está disponible en las fuentes crudas de la secretaria**
(hojas separadas por sub: reunión/faena · techado/campo · medidor/instalación — ver
`obligaciones/README`). `obligaciones` emite el sub **real desde el inicio**; no hay
fallback genérico. (La siembra de MULTA espera reconciliar el gap de asistencia bruta vs
residual; ACUERDOS/CONVENIO ya son sembrables.)

**⑤ Los conceptos entran al contrato desde hoy**, pero `2_planilla` escribe los cargos
de AGUA/MANTENIMIENTO/CORTE_RECONEXION en **Fase 2** (tocar 2_planilla se aísla del build de
caja/cuenta para no multiplicar riesgo). En Fase 1 salen vacíos; la forma de la tool
no cambia bajo el agente.

**⑬ Decisión (24/07/2026) — `reactivación` como 3er `SUB_CONCEPTO` de CONVENIO.**
Naturaleza: deuda que **nace** para el usuario (no es arrastre de AGUA re-etiquetado)
cuando un predio queda años sin uso y acumula deuda que nunca se generó como CARGO
mes a mes (el predio no tenía lecturas activas). El día que el usuario reactiva y
paga, paga el TOTAL acumulado — usualmente en cuotas, igual que instalación. Encaja
en CONVENIO porque es **negociado con el usuario** (no automático) y se paga en
cuotas con TOTAL fijo — misma forma que medidor/instalación, distinto origen del
monto. **Guardarraíl aplicado antes de sembrar:** verificar que el predio NO tenía
`MES_ANTERIOR` corriendo mes a mes en ese período — si lo tenía, la deuda ya está
contada en AGUA (P1) y meterla también acá sería doble conteo. Caso semilla: M-12
(Iglesia Evangélica Bautista → Ramon Requez Mendoza, override en `0_padron`), deuda
"2019 hasta octubre 2025" = S/266, confirmado sin arrastre mes a mes en ese período.
Fuente: nueva hoja `REACTIVACION` en
`obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx` (mismo
esqueleto que `MEDIDORES 2026`: NOMBRES·MZ·LT·TOTAL·meses·PAGADO·SALDO).

### Extracto de cuenta — vista cross-agregado (⑫)

Capacidad **nueva** (pedido directo del usuario, 2026-07-13): *"quiero mi reporte de
seguimiento de todos los meses / de los últimos N, con todos mis pagos, en PDF o
impreso."* Distinta de la boleta:

```
BOLETA      = 1 mes, "esto debes ahora"          → cobrar
EXTRACTO    = N meses, "esto pasó en tu cuenta"   → auditar / resolver reclamo
```

**Las 5 decisiones cerradas:**

| # | Decisión | Razón |
|---|---|---|
| 1 | Clave = por **PREDIO** (MZ-LT), no por persona | sistema es predio-céntrico; identidad-persona es problema aparte, no se construye por caso raro. Persona con 2 predios → 2 extractos |
| 2 | **Tanque** en sección aparte, informativo — **no** afecta `SALDO` de deuda | es aporte voluntario, no obligación del predio (igual que en la cascada, ⑪) |
| 3 | Rango default = **TODO el histórico**; `desde/hasta` opcional | uso real = reclamos ("te pagué en marzo") — no se sabe el mes, el default útil es todo. Con `desde` a mitad → fila "saldo inicial" (arrastre) |
| 4 | Trigger = **usuario on-demand**, 1 predio, con o sin reclamo | no es batch de asamblea |
| 5 | Template **nuevo** — no reusa `PLANTILLA_boletas.docx` | el extracto es tabla tipo libreta (saldo corriente), la boleta es otro layout |

**Arquitectura:** `extracto_predio(mz, lt, desde, hasta)` es una **tool de solo
lectura** en `estado_cuenta/tools/` — no es un writer. Cruza los dos agregados:
cargos+aplicaciones de `estado_cuenta`, abonos de `caja`. Devuelve filas
estructuradas (no un PDF); el render (ver abajo) las convierte en documento.

**Layout — 3 secciones** (ver `estado_cuenta/docs/formato_extracto.html`):

```
① DEUDA POR CONCEPTO→SUB (estado actual, con prioridad P1-P6)
   Pri | CONCEPTO | SUB | CARGADO | PAGADO | DEBE

② PAGOS RECIBIDOS Y CÓMO SE APLICARON (traza la cascada, por pago)
   FECHA | MONTO | "se aplicó a: agua mar, agua abr, mant, corte,
                    multa reunión, multa faena (parcial, faltó X)"

③ APORTES VOLUNTARIOS (tanque) — informativo, no resta de la deuda
   SALDO ACTUAL DE DEUDA: <derivado>
```

La sección ② es la que responde "cómo se distribuyó mi pago" — muestra, en el orden
de la cascada ⑪, qué llenó primero y dónde se cortó si el pago fue parcial.

### Arquitectura de render — data-prep en cada dueño, render como servicio stateless

Decisión de arquitectura (2026-07-13): **ningún módulo de negocio "imprime".** Cada
bounded context arma sus propias filas (sabe qué significa el documento); un
servicio compartido y **sin estado** convierte `(plantilla, filas) → PDF`.

```
estado_cuenta              3_boletas                (cada uno arma SUS filas —
extracto_predio()          data_boletas()            sabe su propio negocio)
      │                         │
      └───────────┬─────────────┘
                   ▼
   RENDER  (hoy: shared/utils_render.py)
   NO sabe qué es boleta ni extracto — solo recibe (plantilla, filas) → PDF
                   │
                   ▼
                  PDF
```

**Por qué NO un módulo `3_impresor` único:** mezclaría dos mundos de datos
(planilla + ledger) en un solo dueño, y probablemente los motores de render sean
distintos (boleta = sustitución de campos en `PLANTILLA_boletas.docx`; extracto =
tabla con saldo corriente). El render se separa **porque es infraestructura
cross-cutting** (como una DB o un servicio de auth), no por código duplicado — no
aplica la Regla del Tres, es un concern de escala/despliegue: generar PDF es
CPU-pesado y candidato a escalar en **su propio contenedor** mañana, sin que
`3_boletas` ni `estado_cuenta` cambien cómo lo invocan (`render(plantilla, filas)`
es el borde estable). `3_boletas` se queda donde está; el extracto vive con su
dueño de datos (`estado_cuenta`).

### Resumen de decisiones del contrato

| # | Decisión | Cierra |
|---|---|---|
| ① | `ABONO_ID` determinista, sin `mz/lt`; 1 depósito = 1 abono. **Corregido 2026-07-17:** clave de efectivo = procedencia (`origen_archivo+fila`), no `fecha` tipeada (bug de duplicado) + gatillo de casi-duplicado con autorización | Hueco 1 (colisión yape) |
| ② | Aplicación referencia `(CONCEPTO, SUB_CONCEPTO, MES_CARGO)` | trazabilidad del cargo |
| ③ | `SALDO_A_FAVOR` = concepto explícito | auditoría en 1 query |
| ④ | DEVOLUCION baja `SALDO_A_FAVOR` FIFO | plata parqueada fungible |
| ⑤ | Los conceptos en el contrato hoy, AGUA/MANTENIMIENTO/CORTE_RECONEXION escritos en Fase 2 | schema estable |
| ⑥ | **MOTOR DE APLICACIÓN separado**; los writers solo emiten hechos | Hueco 2 (5_cobranza sin abono_id) |
| ⑦ | **`JASS_ID` en cada evento** + núcleo tenant-agnóstico | escala a 25k JASS |
| ⑧ | **Todo cargo va a estado_cuenta** (varias fuentes); histórico re-derivado por el motor | Huecos 3 y 4 |
| ⑨ | **4b_reclamos emite la resolución como HECHO** (identificar/reasignar abono · ajuste de cargo); el motor deriva la aplicación, 4b nunca escribe el saldo | reclamos trazables (abono_id + reclamo_id), sin columna manual |
| ⑩ | **El movimiento de caja lleva `DIRECCION`+`BALDE`+`DESTINO`**: captura ingresos y egresos, por predio o por concepto (incl. `GASTO` institucional); `BALDE`≠`CONCEPTO` de deuda. **Corregido 2026-07-13:** `comunitario` no es balde (es segregación de motor_matching); `deuda_directiva` no cruza a estado_cuenta (caja-only, como tanque) | tesorería "cuánto entró/salió" + baldes de 5b desde una sola fuente |
| ⑪ | **`SUB_CONCEPTO` en el CARGO** + cascada P1→P6 con sub-orden (multa: reunión→faena · acuerdos: techado→campo · convenio: medidor→instalación); `DEUDA_DIRECTIVA` fuera de la cascada | el reporte de seguimiento puede mostrar deuda y distribución de pago desglosada |
| ⑫ | **Extracto de cuenta** — tool read-only cross-agregado (`extracto_predio`), por predio, rango=todo por default, template propio; render separado como servicio stateless | responde "mi historial de N meses en PDF" sin acoplar 3_boletas al ledger |
