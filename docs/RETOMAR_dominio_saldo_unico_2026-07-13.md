# RETOMAR — decisión de arquitectura: `libro_mayor` como ÚNICO ledger + arranque de `dominio/` · 2026-07-13 (2ª sesión Opus del día)

Amplía `docs/RETOMAR_comunitario_y_extracto_2026-07-13.md` (misma fecha, sesión
Opus siguiente — empezó diseñando `riesgo_corte` y terminó destapando un problema
de arquitectura más de fondo).

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **Decisión tomada: vamos con la Opción B.** Antes de diseñar más capacidades
   nuevas (`riesgo_corte`, `morosidad_total`, etc.) hay que resolver **quién es el
   dueño del saldo** — hoy son 3 dueños distintos (`2_planilla`, `5_cobranza`,
   `seguimiento_pueblo`) y eso es la raíz de los bugs B4/B5/B7 del ciclo anterior.
2. **`riesgo_corte` quedó PAUSADO a la mitad** — la regla de negocio ya se extrajo
   tal cual de `6_corte` (ver abajo), pero no se cerró de dónde lee `SALDO`/
   `MES_ANTERIOR` porque esa pregunta abrió el problema de fondo. Se retoma
   *después* de B1-B2, no antes — con el ledger único ya no hay ambigüedad de dónde leer.
3. **Aclaración de nombres (esto costó confusión esta sesión, no repetir el error):**
   **`libro_mayor` es EL ledger** — el bounded context completo. `caja` y
   `estado_cuenta` son sus **dos agregados**, no dos ledgers distintos.
   `seguimiento_pueblo` es el prototipo viejo que se **absorbe y desaparece**
   (su README ya lo dice: "rediseña y expande seguimiento_repo.py").
4. **Siguiente paso concreto:** confirmar 2 preguntas cerradas abajo (sección
   "B1 — decisiones a confirmar") y detallar las firmas (I/O) de cada archivo de
   `libro_mayor/dominio/`. Carpeta existe pero vacía (`.gitkeep`).
5. **`docs/pendientes_plan.md` fue descartado por el usuario esta sesión**
   ("Olvidate del plan no sirve"). No usarlo como fuente de SIGUIENTE_ACCION —
   esta sesión y la de comunitario/extracto (07-13) son la fuente de verdad hoy.

---

## Cómo se llegó a esto

Se retomó el proyecto queriendo diseñar las capacidades nuevas del backlog del
ledger (`docs/RETOMAR_comunitario_y_extracto_2026-07-13.md`, sección 7). Se eligió
empezar por `riesgo_corte` (recomendado por esa misma RETOMAR, por el bug de "2
meses" descubierto en producción). Se extrajo la regla tal cual de
`6_corte/generar_lista.py` (ver sección propia abajo) y al preguntar de dónde iba a
leer `SALDO`/`MES_ANTERIOR`, el usuario frenó: **no sabía si esos dos campos hoy
"pertenecen" a `2_planilla` o a `5_cobranza`, ni si el `arrastre_consolidado` de
`5_cobranza` seguía vigente o ya lo reemplazaba el ledger.** Pidió leer los módulos
`4_pagos` → `7_cierre` antes de seguir, y una opinión de arquitecto sobre si el
diseño actual (archivos Excel regenerados + arrastres por archivo) es profesional
para el lente agentic SaaS + Postgres + Docker que el proyecto ya declaró como norte.

---

## 1 · Inventario de "ledgers" — qué existe en CÓDIGO vs qué es solo DISEÑO

El usuario notó (correctamente) que "ledger" se estaba usando para varias cosas
distintas. Esto es lo que hay, verificado contra código real (no memoria):

### Ledgers append-only REALES hoy (saldo derivado, nunca celda a mano)

| Archivo | Writer | Cubre | Nota |
|---|---|---|---|
| `shared/seguimiento_pueblo.xlsx` | `shared/seguimiento_repo.py` | CARGO/PAGO/AJUSTE — **solo** MULTA/ACUERDOS/CONVENIO | El prototipo bueno. Es el que más se parece a `estado_cuenta` — de hecho es su antecesor directo |
| `shared/deuda_directiva.xlsx` | `4_pagos/consolidar_deuda_directiva.py` | reposición de caja de 2 ex-directivos | Balde caja-only (ver corrección ⑩ del contrato, no cruza a deuda de predio) |
| `4_pagos/efectivo/entregas.xlsx` | `4_pagos/efectivo/entregas_repo.py` | arqueo diario: qué entregó cada cobrador | Mismo patrón repo que seguimiento_pueblo |

### Maestro con audit (NO es ledger de dinero)

| Archivo | Writer |
|---|---|
| `DATA_boletas.xlsx` | `shared/data_boletas_repo.py` — padrón + correcciones auditadas |

### Audit logs por-proceso (NO son ledgers de saldo — trazan una corrida, no derivan estado)

`trazabilidad_cobranza.xlsx`, trazabilidad de `motor_matching`, `4_pagos/efectivo`,
`1_lecturas`, `4b_reclamos`, `0_padron` — responden "qué pasó en esta corrida", no
"cuánto se debe".

### Estado persistente simple (acumulado, no event-sourced)

`6_corte/inputs/registro_cortes.xlsx` (CORTADO/EXONERADO),
`1_lecturas/inputs/registro_operario_acumulado.xlsx`.

### Lo que NO tiene ledger — el punto ciego

**La deuda de agua/mantenimiento/corte no vive en ningún ledger.** Se re-computa
cada corrida de `5_cobranza` a partir de `2_planilla` (planilla del mes) +
`4_pagos` (pagos identificados). El "arrastre" al mes siguiente es un **archivo
Excel que se regenera y se vuelve a leer** — memoria-por-archivo, no un registro.
Los bugs B4 (NaN), B5 (columna manual pisada), B7 (dual-writer, el −20) son todos
síntomas de este mismo paradigma: estado derivado que se regenera ciegamente.

### Lo que solo existe en DISEÑO (0 código)

`libro_mayor/` completo — `caja/`, `estado_cuenta/`, `dominio/`, `stores/`,
`tools/` son carpetas vacías (`.gitkeep`). Contrato cerrado en
`caja/README.md` + `estado_cuenta/README.md` (12 decisiones, byte-idénticas).

### Lo que se eliminó

`7b_historial_pagos/` — otro intento de ledger de pagos, arquitectura
pre-contrato. Se descartó en vez de migrarlo (recuperable en git history).

---

## 2 · Nomenclatura — de dónde en adelante, un solo nombre

```
libro_mayor            ← EL ledger (system of record). En Postgres = 1 esquema.
├── caja               agregado 1: MOVIMIENTOS de dinero (abonos, egresos)
│                        tabla futura: abonos
└── estado_cuenta      agregado 2: DEUDA del predio (cargos + aplicaciones)
                         tablas futuras: cargos, aplicaciones
   + motor_aplicacion  la pieza que ve las dos mitades juntas
```

Vocabulario a partir de ahora:
- **`libro_mayor`** = "el ledger" a secas.
- **`libro_mayor/caja`** = mitad "entró/salió plata".
- **`libro_mayor/estado_cuenta`** = mitad "se debe".
- **`seguimiento_pueblo`** = nombre que **muere** en la migración; su contenido
  renace como `libro_mayor/estado_cuenta` (generalizado a todos los conceptos).
- **`trazabilidad_*`** = audit logs, **nunca** llamarlos "el ledger".

---

## 3 · Mapeo hoy → mañana (lo que migra a dónde)

```
① seguimiento_pueblo (3 conceptos)   ─┐
   + agua/mant/corte (re-cómputo      ─┼──►  libro_mayor/estado_cuenta
     en planilla_cobrado, sin ledger) ─┘      (deuda, TODOS los conceptos, saldo derivado)

② deuda_directiva.xlsx  ─┐
③ entregas.xlsx          ┤
   pagos_efectivo.xlsx   ┼──────────────►  libro_mayor/caja
   pagos_yape_*.xlsx     ┘                  (movimientos: abonos/egresos, baldes)

trazabilidad_cobranza.xlsx  ──►  desaparece — las APLICACIONES del motor la reemplazan
DATA_boletas.xlsx            ──►  se queda igual (es padrón, no dinero)
registro_cortes.xlsx         ──►  se queda por ahora (estado operativo de corte físico)
```

---

## 4 · Veredicto de arquitectura (pedido explícito del usuario, "como ingeniero de 15 años")

**No, el diseño actual (archivos regenerados + arrastres por archivo) no es
profesional para el lente agentic SaaS + Postgres + Docker.** Funciona para una
JASS en una laptop; no sobrevive a multi-tenant porque el estado de verdad vive en
archivos derivados regenerables, no en un registro de hechos. Cada bug de este
ciclo (B4/B5/B7) es esta deuda cobrando intereses.

**Pero el camino correcto ya está construido a medias:** `seguimiento_pueblo` es
event-sourcing bien hecho — el prototipo correcto, solo que cubre 3 de N
conceptos. `libro_mayor` es su generalización ya diseñada (contrato ⑥⑧). El
trabajo no es reinventar: es terminar de mover agua/mant/corte al mismo ledger y
que los módulos dejen de ser dueños del saldo.

**Bajo el ledger único, el saldo deja de tener 3 dueños** — pasa a ser siempre una
query derivada (`SALDO = Σcargos − Σaplicaciones ± Σajustes`), nunca una columna
que un `main.py` escribe. Esto es lo que hace posible `riesgo_corte` y las demás
tools sin ambigüedad de "de dónde leo".

---

## 5 · Opción elegida — B (cerrar dueño único del saldo) sobre A (catálogo de tools primero)

Se presentaron 2 caminos:
- **A** — diseñar las 9 capacidades asumiendo el ledger completo, quedan "vivas"
  cuando Fase 2 migre agua/mant/corte.
- **B** — primero migrar agua/mant/corte al ledger + poblar `dominio/` con las
  reglas puras (cascada, política de corte, taxonomía); DESPUÉS retomar el
  catálogo de tools, que en ese punto caen casi solas (son queries sobre una
  tabla, no diseño nuevo).

**El usuario eligió B.**

### Roadmap de B

```
B1 · dominio/          reglas de negocio PURAS, sin I/O, tenant-agnósticas   ← EMPEZAMOS ACÁ
       ↓
B2 · estado_cuenta     migrar deuda agua/mant/corte al ledger
       (quién emite el cargo, cuándo, idempotencia, arrastre→query en vez de archivo)
       ↓
B3 · caja              unificar pagos efectivo/yape + deuda_directiva + entregas
       ↓
B4 · cutover           coexistencia transitoria + backfill de junio/julio
       (Hueco 3 del contrato: NO se migra el histórico, se RE-DERIVA con el motor)
```

---

## 6 · B1 en curso — propuesta de `libro_mayor/dominio/`

Reglas de negocio puras (sin I/O, testeables, independientes de Excel/Postgres).
Cada una ya existe hoy, solo dispersa dentro de `main.py` de distintos módulos —
poblar `dominio/` es **extraer y nombrar, no inventar**:

```
libro_mayor/dominio/
├── entidades.py       Abono · Cargo · Aplicacion (dataclasses)
│                        hoy: no existen como tipo — son dicts sueltos en main.py
├── taxonomia.py       árbol concepto→sub_concepto + orden P1-P5 (⑪)
│                      + mapa BALDE(caja) → CONCEPTO(deuda)
│                        hoy: hardcodeado en 5_cobranza/main.py + CONCEPTOS_VALIDOS
│                        de seguimiento_repo.py
├── cascada.py         aplicar(cargos_abiertos, monto) → aplicaciones
│                      [FIFO por MES_CARGO + sub-orden reunión→faena,
│                       techado→campo, medidor→instalación]
│                        hoy: 5_cobranza/main.py::_descomponer_saldo() (línea 1741)
├── politica_corte.py  evaluar(saldo, agua_atrasada, protecciones) → veredicto
│                      [umbral configurable (JASS actual: 8), TOL=0.005,
│                       protección reclamo EN_REVISION, protección pago parcial
│                       cualquier canal]
│                        hoy: 6_corte/generar_lista.py::_filtrar_corte()
├── saldo.py           derivar(cargos, aplicaciones, ajustes) =
│                      Σcargo − Σaplicaciones ± Σajustes
│                        hoy: seguimiento_repo.get_saldo() (parcial, 3 conceptos)
│                        + re-cómputo disperso en 5_cobranza para agua/mant/corte
└── identidad.py       abono_id(...) determinista · clave_cargo(...)
                         hoy: motor_matching (id yape) + contrato ① del README
```

**Separación clave que define B1 — regla pura vs. orquestación con I/O:**

```
dominio/cascada.py            →  LA LÓGICA de repartir (función pura, sin leer nada)
estado_cuenta/motor_aplicacion.py →  lee cargos+abonos reales, LLAMA a
                                   cascada.aplicar(), escribe las aplicaciones
```

Mismo patrón para corte: `dominio/politica_corte.py` decide (función pura); la
tool `riesgo_corte()` (en `estado_cuenta/tools/`) le da de comer datos ya
derivados del ledger. Así `riesgo_corte` no lee Excel ni Postgres dentro de su
regla de negocio — eso es trabajo del adapter.

**Tenant-agnóstico (decisión ⑦ del contrato):** `dominio/` no conoce "JASS
tupac_amaru", ni el umbral 8, ni la tarifa. Todo eso entra por `config` específico
de cada JASS — `dominio/` es el mismo código para las 25.000 JASS futuras.

---

## 7 · B1 — decisiones a confirmar ANTES de detallar firmas (Opus, esto es debate de diseño)

Dos preguntas, recomendación dada pero **no confirmada explícitamente por el
usuario** — cerrar esto es lo primero de la próxima sesión:

**Pregunta 1 — ¿`dominio/` es un único núcleo compartido por `caja` Y
`estado_cuenta`, o cada agregado tiene el suyo?**
Recomendación: **uno solo compartido.** La taxonomía y el mapa balde→concepto los
necesitan las dos mitades (el motor cruza ambos lados); partirlo en dos duplicaría
el vocabulario y reabriría el riesgo de desincronización que ya pasó con el
contrato (dos README que tenían que mantenerse byte-idénticos).

**Pregunta 2 — ¿la lógica pura de la cascada vive en `dominio/cascada.py` y el
I/O (leer cargos/abonos reales, escribir aplicaciones) vive aparte en
`estado_cuenta/motor_aplicacion.py`?**
Recomendación: **sí**, es la separación pura/adapter de la sección 6. La
alternativa (todo el motor en `estado_cuenta/`, sin separar) mezcla regla con
lectura y hace la cascada no-testeable sin montar archivos/DB de prueba.

**Una vez confirmadas ambas:** detallar la firma exacta (parámetros de entrada,
qué devuelve, qué excepción lanza) de cada archivo listado en la sección 6, un
archivo a la vez, empezando por `taxonomia.py` (todo lo demás depende de sus
constantes) → `entidades.py` → `cascada.py` → `politica_corte.py` → `saldo.py` →
`identidad.py`.

---

## 8 · `riesgo_corte` — estado parcial, PAUSADO (no perder este trabajo)

Antes de abrir la pregunta de arquitectura mayor, se llegó a esto — **no
descartar, retomar después de B1-B2**:

**Regla de negocio extraída tal cual de `6_corte/generar_lista.py::_filtrar_corte`
+ `config.py`** (verbatim, no reinterpretada):

```
riesgo_corte(mz, lt)  →  ¿este predio va a corte?

① excluir     (mz,lt) en CORTADO/EXONERADO (registro_cortes.xlsx)
② elegible    SALDO > 0.005 (TOL)  AND  MES_ANTERIOR (agua atrasada) >= 8
③ protección  reclamo EN_REVISION (4b_reclamos) → NO corta
④ protección  pagó algo este mes, CUALQUIER canal (mesa O monto_yape>TOL) → NO corta
⑤ si no aplica ninguna protección → VA_A_CORTE = SI
```

Constantes exactas a preservar: `MES_ANTERIOR_MIN = 8`, `TOL = 0.005`,
`PENALIDAD = 20.0` (inicial) / `PENALIDAD_FINAL = 40.0` (tras corte físico).

**Lo que quedó SIN cerrar (por eso se pausó):** de dónde lee `SALDO` y
`MES_ANTERIOR` mientras el ledger no tiene todavía los cargos de AGUA
(`2_planilla` los escribe recién en Fase 2, decisión ⑤ del contrato). Se
plantearon 2 opciones (ledger vacío hasta Fase 2, vs. bridge temporal leyendo
`planilla_cobrado.xlsx`) pero el usuario prefirió resolver primero **quién es el
dueño real del saldo** (esta sesión) antes de decidir el bridge. **Con B1+B2
resueltos, esta pregunta se disuelve sola** — ya no hay bridge que decidir, el
ledger es la única fuente.

**Salida propuesta (no cerrada, retomar cuando vuelva el tema):**
```
riesgo_corte(mz=None, lt=None) → filas de:
  MZ · LT · NOMBRE · SALDO · MES_ANTERIOR ·
  ELEGIBLE(bool) · ESTADO(SIN_RIESGO|EN_RIESGO|YA_CORTADO) ·
  PROTEGIDO_POR("" | reclamo_en_revision | pago_parcial) ·
  VA_A_CORTE(SI/NO)          ← idéntico semánticamente a EJECUTAR_CORTE de 6_corte
```

---

## 9 · Qué pasó con `docs/pendientes_plan.md`

Al retomar la sesión se detectó que `pendientes_plan.md` está desactualizado
(última edición 2026-07-02, 11 días de trabajo real no reflejados — contrato
ledger 07-11, libro mayor 07-12, comunitario+extracto 07-13). Se recomendó
`/inventario` para reconciliar. **El usuario cortó por lo sano: "Olvidate del
plan no sirve."** No usar ese archivo como fuente de siguiente acción por ahora.
Si se quiere retomar el tracking formal de pendientes en algún momento, correr
`/inventario` de nuevo — pero no es lo que se pidió para esta sesión ni la
siguiente.

---

## Orden sugerido al retomar (próxima sesión, Opus — esto es debate de diseño)

1. Confirmar las 2 preguntas de la sección 7 (dominio compartido sí/no · cascada
   pura separada del motor sí/no).
2. Detallar firma I/O de cada archivo de `dominio/`, en el orden: `taxonomia.py` →
   `entidades.py` → `cascada.py` → `politica_corte.py` → `saldo.py` →
   `identidad.py`.
3. Con `dominio/` cerrado (spec, no código todavía — sigue Fase 1), decidir si se
   escribe el código de `dominio/` ya (es lógica pura, bajo riesgo) o se sigue
   con B2 (migración de estado_cuenta) primero en spec.
4. Recién ahí retomar `riesgo_corte` — a esa altura ya no tiene la ambigüedad de
   dónde leer.

---

## 10 · SPEC CERRADO de `dominio/` — continuación 2026-07-14 (Opus)

Se confirmaron las 2 preguntas de la sección 7 y se detalló la firma (I/O) de las 6
piezas, una por una, cada una **verificada contra código real** (no inventada).
**`dominio/` queda cerrado en spec — falta escribir el código (avisar antes).**

### Preguntas 7 — confirmadas
- **P1 = uno solo compartido.** `dominio/` es un único núcleo que importan `caja` y
  `estado_cuenta` (la taxonomía + mapa balde→concepto los usa el motor, que cruza
  ambos lados; partirlo reabriría el riesgo de desincronización del contrato).
- **P2 = sí, cascada pura separada del motor I/O.** `dominio/cascada.py` = lógica
  pura; `estado_cuenta/motor_aplicacion.py` = adapter (lee cargos/abonos reales,
  llama a `cascada.aplicar`, escribe aplicaciones).

### Recalibración de escala (el usuario corrigió el lente)
25k JASS máx en Perú × ≤10k predios c/u = **un Postgres modesto, NO hiperescala.**
Nada de sharding ni FK optimizados para miles de millones de filas. Esto cambió qué
recomendaciones sobreviven:
- **Sobrevive por corrección/limpieza (no escala):** dinero en **céntimos int** →
  elimina todo el `TOL=0.005`, comparación exacta `==`.
- **Sobrevive por borde transaccional (no RAM):** unidad de trabajo = **1 predio**
  (el abono de un predio salda cargos de ese predio; el motor ya corre así).
- **Ya no forzado por escala, se mantiene por consistencia:** `cargo_id` surrogate
  determinista igual que `abono_id` (un FK compuesto de 6 cols andaría fino a esta
  escala; lo que se evita es la *identidad condicional* "id sintético solo si colisión").

### Las 6 firmas cerradas

```
taxonomia.py     CONCEPTOS (tupla ordenada = prioridad P1..P6; OTROS al final = slot
                   especulativo sin emisor hoy) · SUB_CONCEPTOS {MULTA:(REUNION,FAENA),
                   ACUERDOS:(TECHADO,CAMPO), CONVENIO:(MEDIDOR,INSTALACION)} ·
                   SALDO_A_FAVOR (fuera de cascada) · BALDES_INGRESO/EGRESO ·
                   BALDE_CRUZA {agua:AGUA, tanque:None, deuda_directiva:None,
                   blancos:None, devolucion/retorno:SALDO_A_FAVOR, honorario/gasto:None}
                 fn: prioridad() · sub_orden() · concepto_de_balde() · es_valido()
                 · MANTENIMIENTO en P1 junto con AGUA (verificado: se suman en
                   _descomponer_saldo) · nombres canónicos del contrato (ACUERDOS, no
                   ACUERDOS_ASAMBLEA); los feeders traducen su nombre viejo al emitir

entidades.py     frozen dataclasses, céntimos int, sin I/O, __post_init__ valida vs taxonomia:
                 MovimientoCaja  (UNA clase; direccion INGRESO/EGRESO distingue abono/egreso;
                    mov_id = ABONO_ID|DEVOLUCION_ID|GASTO_ID mismo esquema; balde, destino,
                    monto_centimos>0, mz/lt solo si destino=PREDIO)
                 Cargo           (cargo_id determinista SIEMPRE; concepto, sub_concepto,
                    mes_cargo, monto_centimos, source)
                 Aplicacion      (abono_id FK, cargo_id FK, monto_aplicado_centimos)
                 Ajuste, IdentificacionAbono (⑨) también entidades frozen
                 + recorded_at en cada evento (orden + debug idempotencia en backfill)

cascada.py       CargoAbierto(cargo_id, concepto, sub_concepto, mes_cargo,
                    saldo_pendiente_centimos)
                 Reparto(aplicaciones: list[Aplicacion], saldo_a_favor_centimos: int)
                 clave_orden(c) -> (prioridad, indice_sub, mes_cargo)   # P1..P6, sub, FIFO
                 aplicar(cargos_abiertos_de_1_predio, abono: MovimientoCaja) -> Reparto
                   · UN abono (no la suma) → cada Aplicacion sabe su abono_id (traza extracto ⑫)
                   · varios abonos: el MOTOR los itera por FECHA y actualiza saldos; la
                     cascada procesa siempre uno
                   · la DEVOLUCIÓN (④) NO vive acá — es egreso vs SALDO_A_FAVOR, la maneja el motor

politica_corte.py Veredicto(va_a_corte: bool, estado: SIN_RIESGO|EN_RIESGO|YA_CORTADO,
                    protegido_por: ""|reclamo_en_revision|pago_parcial)
                 evaluar(saldo_centimos, arrastre_centimos, umbral_centimos,
                         ya_cortado, en_revision, pago_parcial) -> Veredicto
                   · replica _filtrar_corte exacto: cortado→elegibilidad(saldo>0 AND
                     arrastre>=umbral)→reclamo(prioridad)→pago_parcial→corta
                   · umbral por CONFIG del tenant (agnóstico); penalidad 20/40 = config del
                     caller, NO en evaluar (es consecuencia total=saldo+penalidad, no veredicto)
                 PARQUEADO: "8 soles ¿= 2 meses?" es decisión de negocio; la función es
                   agnóstica (config pasa el número) → NO bloquea cerrar dominio; se decide
                   al retomar riesgo_corte (post B1-B2)

saldo.py         saldo_cargo(cargo, aplicaciones, ajustes) -> int   # → CargoAbierto del motor
                 saldo_por_concepto(cargos, aplic, ajustes) -> dict # → sección ① extracto ⑫
                 saldo_predio(cargos, aplic, ajustes) -> int        # deuda neta, EXCLUYE saldo_a_favor
                   · SALDO = Σcargo − Σaplicado + Σajuste(signed) — verificado vs get_saldo
                   · deuda y SALDO_A_FAVOR SEPARADOS hasta el borde de presentación (contrato ③,
                     extracto ⑫); netear temprano oculta info que el reclamo necesita

identidad.py     abono_id_efectivo(jass, mes_ciclo, mesa, cobrador, fecha, monto_centimos, mz, lt)
                 abono_id_yape(jass, mes_ciclo, origen, timestamp)  # (origen,ts) predio-agnóstico
                 egreso_id(jass, mes_ciclo, canal, balde, fecha, monto_centimos, referencia)
                 cargo_id(jass, mz, lt, concepto, sub_concepto, mes_cargo)
                 formato: {JASS_ID}-{MES_CICLO}-{CANAL}-{shorthash}
                   · shorthash = sha256(clave_canónica).hexdigest()[:8] — NUNCA hash() de
                     Python (salteado por proceso → no reproducible → import no idempotente)
                   · CANONICALIZACIÓN obligatoria antes de hashear (reusar _norm_mz/_norm_lt,
                     concepto al canon, monto en céntimos int, fecha formato fijo): "C"/"c",
                     "43"/"43 ", 40.0/40.00 deben dar el MISMO id o se duplican abonos
                   · clave natural verificada: efectivo (jass,mesa,cobrador,fecha,monto,mz,lt);
                     yape (jass,origen,timestamp) = ORIGEN|FECHA de motor_matching (main.py:1022)

Transversal (lente SaaS): tenant-agnóstico (umbral/tarifa/conceptos por config, nunca
  hardcode) · cero I/O · sin TOL (céntimos exactos) · todo id determinista/idempotente.
```

### Siguiente paso (avisar antes de codificar — pedido del usuario)
1. **[HECHO]** persistir este spec (esta sección).
2. **Codificar `dominio/` con tests unitarios** — orden: `taxonomia.py` (todo depende de
   sus constantes) → `entidades.py` → `cascada.py` → `politica_corte.py` → `saldo.py` →
   `identidad.py`. Pura, testeable con dicts, no toca ningún módulo existente = bajo riesgo.
   Es Opus (lógica no trivial: cascada, saldo).
3. **B2 después** (migración `estado_cuenta`: quién emite el cargo de agua, cuándo,
   idempotencia, arrastre→query en vez de archivo), ya con `dominio/` funcionando.

---

## 11 · Cierre de diseño de los módulos que el ledger remodela — 2026-07-14 (Opus)

Antes de codificar `dominio/` había que cerrar el diseño destino de los 3 módulos que
el ledger cambia (Regla 6: Fase 1 cerrada antes de Fase 2). El usuario frenó bien el
sesgo a codificar: `dominio/` no estaba realmente cerrado porque `politica_corte`
depende de decisiones parqueadas y la forma de estos módulos no estaba decidida. Ver
[[feedback_no_codificar_diseno_no_cerrado]].

### 5_cobranza post-ledger → SE DISUELVE (verificado contra código)

Cada responsabilidad tiene otro dueño bajo el ledger:
```
aplicar pagos (cascada)  → motor de aplicación
saldo / arrastre         → query derivada (arrastre deja de ser archivo)
reconciliar pagos pueblo → motor deriva aplicaciones desde abonos
reportes                 → tools de lectura del ledger
cargos no-agua           → obligaciones/  (NUEVO)
```
**Prueba de código:** `5_cobranza` solo llama `registrar_pago`/`registrar_ajuste`,
**nunca** `registrar_cargo`. El que CREA los cargos multa/acuerdos/convenio es
`shared/sembrar_seguimiento_pueblo.py` (génesis). → El contrato ⑧ estaba MAL
("SOURCE: 5_cobranza") y se corrigió a `obligaciones` en ambos README (byte-idéntico
verificado con diff).

### obligaciones/ → NUEVO módulo (creado, solo README, sin código)

Emisor de cargos no-agua (multa/acuerdos/convenio), **event-driven**, **cliente** del
ledger, **sin número** (no es paso del pipeline, igual que `libro_mayor`). No es `2b`:
el `b` implica subordinación a su principal + paso mensual, y no es ninguna de las dos.
```
obligaciones/  registrar_multa · registrar_acuerdo · registrar_convenio  → registrar_cargo
               backfill()  = modo migración (corre en lote sobre histórico, parte de B4)
```
Relación con la migración: `seguimiento_pueblo` es el store viejo que DESAPARECE;
`obligaciones` es el emisor PERMANENTE que puebla `estado_cuenta` (la migración es una
de sus corridas, no su razón de ser). NO es desechable. Diseño detallado (triggers
reales, anti-doble-conteo carry vs nueva, convenio=medidor+inscripción) marcado como
Fase 1 propia pendiente en `obligaciones/README.md`.

### 5b_validacion post-ledger → SE DISUELVE en 2 tools

A diferencia de 5_cobranza, `5b` tenía una función SIN otro dueño: comparar contra el
banco (fuente externa). Se parte:
```
arqueo_caja(fecha)      Σ MONTO por BALDE  → query PURA sobre caja → tool en libro_mayor
                        (el ledger es su dueño; ⑩ ya lo decidió: "reproduce 5b desde el
                         ledger, no desde 8 archivos")
conciliar_caja(crudo)   arqueo vs reporte_mes_crudo (banco) → cuadra/discrepancias
                        cliente del ledger (conoce lo externo) — el control que sobrevive
```

**El evento que refina el diseño (aporte del usuario): 5b se corre 2 VECES por ciclo** —
1ª para generar lista de corte (mitad de ciclo), 2ª para validar todo (cierre). Esto
NO rompe el ledger, lo confirma:
```
HOY:    re-correr el batch entero × 2 (regenera estado desde 8 archivos, caro, buggy)
LEDGER: conciliar_caja + saldo son QUERIES point-in-time sobre un ledger append-only
        T1 (corte):  conciliar_caja(crudo)@T1 → riesgo_corte lee saldo@T1 → lista_corte
        T2 (cierre): conciliar_caja(crudo)@T2 → asiento
        entre T1 y T2 solo se AGREGAN abonos; la query los ve sola, sin regenerar
```
→ Corrige mi framing "b1 vs b2 gate único": `conciliar_caja` NO es un gate fijo, es una
**tool invocada en N momentos** = la capacidad "estado a-una-fecha" del backlog. El
`estado_ciclo.validado` deja de gatear a `5_cobranza` (disuelto) → gatea el asiento.

### Estado y qué falta

```
✓ 5_cobranza  cerrado (se disuelve)          ✓ obligaciones/ creado (README)
✓ 5b          cerrado (se disuelve → 2 tools) ✓ contrato ⑧ corregido (byte-idéntico)
✓ 6_corte     CERRADO — SOBREVIVE reshaped (NO se disuelve: tiene estado operativo que
              no es ledger). Tres decisiones:
              (a) regla ACLARADA (2026-07-14): "MES_ANTERIOR>=8" NO es "2 meses" (marco
                  equivocado). El 8 = DEUDA MÍNIMA MENSUAL; >=8 = filtro de PAGO PARCIAL
                  (MES_ANTERIOR=7 = pagó parcial → protegido). Amount-based, confirma
                  politica_corte sin cambios (umbral = deuda_mínima de cada JASS, config).
              (b) el CARGO corte_reconexion lo emite 6_corte (evento de corte), NO 2_planilla
                  (⑧ re-corregido: 2_planilla = agua+mantenimiento; 6_corte = corte_reconexion).
                  obligaciones es para obligaciones standalone; la penalidad es cohesiva con
                  el workflow de corte → 6_corte la emite directo (registrar_cargo).
              (c) phase gate SOBREVIVE como estado de workflow, NO candado de archivo: el
                  commit registra T1 (ancla) + emite cargos; la inmutabilidad viene del ledger
                  append-only (lista = riesgo_corte@T1, re-derivable). auto-release por MES_ANO.
              Qué queda de 6_corte: workflow de corte + estado operativo (registro_cortes,
              phase gate) + emite corte_reconexion. Ya NO: calcula saldo · overlay-hack de
              penalidad · truco algebraico de seguimiento (todo → queries del ledger).
```
**Fase 1 de los tres módulos CERRADA (2026-07-14).** `dominio/` deja de moverse → se
puede codificar (B1) sin reescribir. Pendiente, en orden sugerido:
1. Actualizar los README de `5_cobranza` / `5b_validacion` / `6_corte` al nuevo diseño
   (ahora sí sin inventar — el destino está cerrado).
2. Diseño DETALLADO de `obligaciones/` (Fase 1 propia: triggers reales, anti-doble-conteo,
   convenio=medidor+inscripción).
3. Codificar `dominio/` + repos (`caja_repo`, `cuenta_repo`) + motor de aplicación.
