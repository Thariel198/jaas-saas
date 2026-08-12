# Decisión de diseño — buscar_pago (herramienta nueva de 4b_reclamos)

Fecha: 2026-08-12
Estado: Aprobado en conversación · Fase 2.0.6

---

**Problema:**
Un vecino reclama "ya pagué mes anterior" (29 casos abiertos en agosto, `TIPO_RECLAMO=mes_anterior`
en `reclamos_2026-08.xlsx`). El supervisor hoy busca a mano en planillas viejas o le pregunta a la
secretaria. No hay forma sistemática de responder: **¿el pago existe? ¿algún bug lo ocultó? ¿está
acreditado a otro predio por error de tipeo/OCR? ¿nunca existió?**

El texto libre del reclamo es la única pista (`RECLAMO`): a veces trae monto ("pagó 16"), a veces
cobrador ("con Maximo"), a veces nada ("Reclamo mes anterior 9.00"). No es un formulario, es lo que
el cobrador anotó en la mesa.

**Criterios:**
- Reusar lo que ya existe y está probado (`reporte_referencias_pago.py`,
  `4_pagos/efectivo/verificar_lotes.py`) en vez de reconstruirlo.
- Un candidato propuesto con evidencia débil es peor que no proponer nada — mismo criterio que
  `verificar_lotes.py` (ver `docs/decisiones/verificacion_lotes_efectivo.md`): "un falso OK es peor
  que un no-sé".
- Auditable fila por fila por una persona no técnica.
- **Solo `mes_anterior`.** Revisado durante la implementación (2026-08-12): la intención inicial era
  hacerla genérica por `TIPO_RECLAMO`, y se descartó al correrla contra los 43 reclamos de
  `convenio` — devolvía "pagó parte de sus cuotas" para 29 de ellos, que es cierto y no explica
  nada. La causa de "ya pagué mi medidor" y "ya pagué techado y campo" es **anterior y distinta**:
  el orden de la cascada. Hoy reparte `consumo · mantenimiento · mes anterior · multa · acuerdos ·
  convenio` y va a repartir `consumo · mantenimiento · mes anterior · convenio · acuerdos · multa`
  (la multa al final: es lo único que se puede cubrir con faena o exonerar). Buscar el pago de un
  convenio antes de ese reorden es buscar en el lugar equivocado. Ese reorden es otro trabajo, ya
  simulado en `4b_reclamos/reporte_reimputacion_cascada.py`.

**Enfoque elegido — embudo de 2 bloques, con un gate antes:**

```
GATE   DATA_boletas["MES ANTERIOR"] del ciclo activo == 0 ?
       → SÍ: RESUELTO_YA (la deuda ya no le vino), cierra sin buscar nada más.
       → NO: reclamo vivo, entra al embudo.

BLOQUE A — EXPLICAR (¿la plata ya está adentro del sistema?)
  A0  historial completo del predio, 3 repos (jun·jul·ago) vía tabla_predio() +
      referencias_pago() → entró y MES_ANT quedó en 0: INFUNDADO (mostrar el recibo)
                          → entró y MES_ANT siguió > 0: MAL_IMPUTADO (bug real, la
                            plata fue a otro concepto)
  A1  ¿algún precursor de shared/ ya cuenta la historia? (ver "Usos de los
      precursores" abajo) → EXPLICADO_POR_PRECURSOR / PAGÓ_PERO_NO_ERA_AGUA /
      PAGÓ_ANTES_APLICADO_DESPUÉS

BLOQUE B — BUSCAR (la plata no está, ¿dónde se fue?)
  filtros previos: precursores apagan candidatos con dueño · ventana temporal ·
  cobrador nombrado en el texto (si lo hay, reduce el pool a casi 0)
  B1  pool de blancos sin reclamar (blancos_acumulados · blancos_efectivo ·
      hoja "Reporte" pre-mayo con mz=blanco) → CANDIDATO_BLANCO
  B2  registrado en OTRO predio del mismo ciclo (~300 pagos, no una lista curada)
        b2a tipeo/OCR (origen del pago distinto)   → CANDIDATO_TIPEO
        b2b multi-lote (mismo origen, 1 transacción
            cubre 2 lotes y el sistema solo acreditó
            uno — caso real K-3/K-4, 2026-08-10)    → CANDIDATO_MULTILOTE
  B3  exceso no resuelto de un predio confundible (arrastre_devolucion_{06,07,08},
      ESTADO≠resuelto, y solo si el precursor no lo tiene ya explicado)
                                                     → CANDIDATO_EXCESO

SIN_EVIDENCIA → pedir recibo o captura de yape
```

**Regla de propuesta (heredada de `verificar_lotes.py`, ver decisión hermana):** un candidato del
Bloque B solo se propone cuando la lista queda en exactamente 1 (error simple, con dueño impago).
Con 2+ candidatos: `"N candidatos"`, sin elegir.

**Ventana temporal (regla del negocio, no del código):**

```
distancia = ciclo_reclamo − mes_del_candidato

distancia ≤ 1 mes   → PLAUSIBLE siempre (julio reclama junio = evitar el corte de 2 meses,
                       comportamiento normal del vecino)
distancia ≥ 2 meses → PLAUSIBLE SOLO si MES_ANT > 0 en TODOS los meses intermedios
                       (arrastró la deuda, se niega a pagar, re-reclama)
                     → si en algún mes intermedio MES_ANT quedó en 0, el candidato viejo
                       se descarta — no tiene sentido
```

`tabla_predio()` ya devuelve `MES_ANT` mes a mes — no requiere dato nuevo.

**Usos de los precursores (`shared/*.xlsx`) — 5, no solo "explicar":**

1. **Explicar** — el uso obvio: un evento en `abonos_rezagados` / `ajustes_cargo` / etc. ya
   cuenta por qué la deuda quedó como está.
2. **Filtrar el pool de candidatos** — `reidentificacion(_cargo).xlsx` dice qué blanco YA se
   asignó a otro predio. Sin este filtro el embudo puede proponer un blanco con dueño (doble
   crédito). Mismo patrón que la regla ya usada en R-7/M-7 (`feedback_pivot...`, ver memoria).
3. **Causa del reclamo, no su cura** — `aportes_tanque_manuales`: la plata entró de verdad pero
   no era agua. El vecino tiene razón en caja y no en deuda → veredicto propio
   `PAGÓ_PERO_NO_ERA_AGUA` (casos A-4, C1-2/K-2).
4. **El desfase de fechas es la explicación** — `abonos_rezagados` trae `FECHA_REAL` /
   `MES_CICLO` / `MES_ANO_APLICA` como 3 fechas distintas. Un vecino que pagó el 28/06 y se le
   aplicó en julio dice "ya pagué mes anterior" y tiene razón → `PAGÓ_ANTES_APLICADO_DESPUÉS`.
5. **Exceso ya explicado ≠ exceso disponible** — `arrastre_devolucion` con `ESTADO=resuelto`
   trae en `REVISION` el motivo ya cerrado. Ese exceso no es candidato para B3; sin este filtro
   se re-propondría plata que ya tiene dueño.

**Blockers reales encontrados en `reporte_referencias_pago.py` (a arreglar antes de codificar
esta herramienta, porque la reusa):**

1. `_PLANILLAS_RECIENTES`, `_PAGOS_EFECTIVO_CRUDO`, `_PAGOS_YAPE_CRUDO` tienen `"2026-07"` apuntando
   a `BASE_DIR.parent` (el repo activo). Cuando julio era el ciclo activo esto era correcto; desde
   que `shared/ciclo_activo.json` rodó a `2026-08`, `"2026-07"` lee el repo de **agosto** y lo rotula
   julio — ya roto hoy, no es un problema futuro de septiembre.
   Fix: julio vive en `C:\Users\wilde\PycharmProjects\Julio\jass_system - Julio`.
2. Los nombres de archivo no son uniformes entre repos cerrados:
   `planilla_cobrado_julio.xlsx` (Julio, nombre de mes en español) vs
   `planilla_cobrado_2026-06.xlsx` (Junio, `ciclo.resolver` estándar). El resolver estándar no
   encuentra el de julio sin un alias explícito.

**Corregido durante la implementación (medido contra los 29 reclamos reales de 2026-08):**

La 1a versión emitió 5 `CANDIDATO_TIPEO`, de los cuales 4 eran falsos. Los defectos y su arreglo,
todos con test contrafactual en `4b_reclamos/tests/test_buscar_pago.py`:

1. **Faltaba la capa 4 de `verificar_lotes`** (el candidato debe estar impago). C-28, O-25, O-23 y
   O-24 tenían un pago que cuadraba EXACTO con su propia boleta — era su pago, no una confusión.
   Estaba implementado como una nota ("ojo: también cuadra…") en vez de como filtro.
2. **El monto matcheaba subconjuntos que no incluyen el concepto disputado.** Q-9 proponía un pago
   de S/3 (= su `mant`) para una disputa de S/20 de mes anterior; Z-17 uno de S/34 (`consumo+mant`)
   para S/36. Se reemplazó `vl.subconjuntos()` por `_montos_que_cubren()`, que solo genera las
   combinaciones que contienen el cargo reclamado.
3. **`PAGO_ANTES_APLICADO_DESPUES` sin desfase real.** E-8 tenía `MES_CICLO == MES_ANO_APLICA` y
   `FECHA_REAL` vacía — no hay "pagó antes". Ahora se exige `_dist_meses(aplica, ciclo) >= 1`, y
   cuando un predio tiene varias filas se sigue buscando la que sí lo tiene (caso I-9).
4. **El aporte al tanque cortaba el embudo sin explicar el reclamo.** Q-12 (tanque S/50, cuyo propio
   `MOTIVO` dice "sin relación con este aporte") frente a un reclamo de S/15. Ahora es veredicto
   propio solo si el tanque fue la única plata del mes; si además hubo pago de agua, queda como nota
   y la búsqueda sigue.
5. **`MAL_IMPUTADO` sin verificar que el cargo se debía ese mes.** Si no se debía nada, recibir 0 es
   correcto y llamarlo bug es un falso positivo. Se agregó el chequeo contra `shared/planilla_mes/`.

**Capa que faltaba en el diseño, encontrada al correr:** 11 de los 13 casos que quedaban en
`SIN_EVIDENCIA` tenían un pago **del propio ciclo del reclamo** que no fue al arrastre. El diseño en
papel no lo cubría porque asumía que el pago reclamado era siempre de un mes anterior.

**Encuadre corregido por el usuario (importante — la primera lectura era equivocada):** ese pago fue
a **consumo + mantenimiento**, y eso es la cascada **funcionando bien** (P1: agua del mes primero).
No es un hallazgo y no hay nada que re-imputar. Lo que realmente pasó es que el vecino pagó el mes y
**se negó a pagar el arrastre**, precisamente porque su reclamo dice que no lo debe (ya lo pagó en
meses anteriores). Entonces lo que hay que verificar es el **origen del arrastre**, no el destino de
este pago.

De ahí salen dos correcciones concretas:

6. **`PAGO_SOLO_EL_MES` reemplaza a `PAGO_FUE_A_OTRO_CARGO`**, y se agrega
   `CASCADA_FUERA_DE_ORDEN` para la anomalía que sí importa: que se cobre **multa / acuerdos /
   convenio** dejando el arrastre sin pagar — esos tres van *después* del mes anterior. Medido:
   **0 casos** en 2026-08, los 14 pagaron exactamente consumo+mantenimiento. Ese cero es el
   hallazgo: la cascada está bien en lo que respecta al mes anterior.
7. **El monto no siempre discrimina, y ahora se dice.** En 6 de los 14 (O-28, U-2, Y-3, T-15, O-21,
   F1-1) `consumo+mant` suma exactamente lo mismo que el cargo de mes anterior (5+3=8 y anterior=8),
   así que el monto pagado **no permite** saber cuál de los dos quiso pagar. La versión anterior
   afirmaba "S/8 es EXACTO el cargo que reclama, así que pagó justo eso" — falso OK. Ahora la
   ambigüedad se declara explícitamente.
   Se agrega también el contexto histórico (`_historico_mes_anterior`) como **columna propia**
   (`PAGO_ARRASTRES_ANTES`), porque parte los casos en dos grupos con acción distinta — ver abajo.

**El corte que decide a quién mirar primero (medido, no supuesto):**

De los 14 `PAGO_SOLO_EL_MES`:

```
 8  nunca pagaron un arrastre en todo el historial → el reclamo no tiene respaldo
 6  SÍ venían pagando arrastres y aun así les sigue apareciendo → mirar acá
      Q-9   S/76.00 en 4 meses     Z-17  S/101.00 en 4 meses
      W-5   S/23.00 en 3 meses     H-13  S/96.00  en 3 meses
      O-21  S/33.00 en 2 meses     L-2   S/20.00  en 1 mes
```

Grupo de control para saber si la señal discrimina: 60 predios que también deben arrastre y **no**
reclamaron → 18% nunca pagó un arrastre, contra 41% entre los reclamantes que hoy deben arrastre.
La señal discrimina (≈2.3×) pero **no es determinante** — de ahí que se exponga como columna para
que el supervisor priorice, y no como parte del veredicto.

Corrección de método sobre esto: en la primera lectura se afirmó "los 14 nunca pagaron un arrastre"
extrapolando de 3 filas impresas. Era falso (son 8 de 14). El número salió al contarlo sobre el
archivo completo.

Resultado final: 28 de 29 reclamos con explicación (antes 12 sin evidencia y 4 candidatos falsos).

**Alternativas descartadas:**
- *Genérica por `TIPO_RECLAMO`* — se implementó y se retiró el mismo día, ver Criterios arriba: para
  `convenio`/`cuota` la respuesta correcta está bloqueada por el reorden de la cascada, así que la
  herramienta contestaba con una verdad irrelevante. Toda la maquinaria de "veredicto débil +
  fallback" que ese caso exigía se borró en vez de dejarla sin uso.
- *Reconstruir la búsqueda de pagos desde cero* — `reporte_referencias_pago.py` ya resuelve
  correctamente pre-mayo (hojas Cobranza/Reporte/Efectivo) y post-mayo (planilla_cobrado + crudos +
  overlays), con media docena de bugs ya encontrados y corregidos documentados en sus propios
  comentarios. Reescribirlo tira ese conocimiento.
- *Veredicto automático que cierra el reclamo* — el script decide y marca RESUELTO — descartado:
  el supervisor es quien cierra, el script solo entrega evidencia con nivel de confianza (mismo
  criterio que `verificar_lotes.py`: nunca escribe el archivo que el humano resuelve).
- *No filtrar por precursores antes del Bloque B* — sin el filtro, un blanco ya reidentificado o
  un exceso ya resuelto se re-propone como candidato nuevo: sería repetir investigación ya cerrada.
- *Código de billete (B2495) como identificador de pago* — es la serie del billete físico (control
  de falsedad), no un identificador de transacción. No aporta al embudo.
- *Monto declarado por el vecino como llave dura* — descartado a llave, mantenido como pista suave:
  el reclamo nace muchas veces de la emoción (no quiere el corte), no siempre de una confusión real
  de monto.

**Señal de alerta:**
Si el Bloque B empieza a proponer candidatos con 2+ opciones en la mayoría de los casos, el pool de
"~300 pagos por ciclo" dejó de discriminar — señal de que faltan más filtros duros (cobrador
nombrado, ventana temporal), no de que hay que forzar una elección.

Segunda señal, la que importa: **cualquier caso de `CASCADA_FUERA_DE_ORDEN`**. Hoy son 0. Si aparecen,
la cascada está cobrando multa / acuerdos / convenio antes del arrastre y eso sí es un bug de
aplicación, no una disputa con el vecino.

Tercera: **14 de 29 salen `PAGO_SOLO_EL_MES`**, o sea que la mitad de los reclamos de mes anterior no
son un pago perdido ni un error del sistema — son vecinos que pagan el mes y se niegan a pagar el
arrastre porque sostienen que ya lo pagaron antes. Eso no se resuelve buscando mejor ni re-imputando:
hay que **auditar el origen del arrastre** (¿de qué mes impago nació ese saldo?). Si el número no
baja mes a mes, la deuda arrastrada del pueblo no está reconciliada y ese es el trabajo de fondo.

**Pendiente conocido (no implementado):** el filtro por **cobrador nombrado** en el texto del
reclamo ("con Maximo", "yape a Janet", "Garcilazo y Maximo"). Está en el diseño como filtro fuerte
del Bloque B y reduciría el pool casi a cero, pero el Bloque B casi no se usa hoy — el Bloque A
explica 28 de 29. Se implementa cuando el Bloque B empiece a devolver listas largas, no antes.

---

## Escala (lente de tenant)

El embudo (GATE → Bloque A → Bloque B) y la regla de ventana temporal son universales — cualquier
JASS reclama "ya pagué mes anterior" por el mismo miedo al corte. Lo que es específico de esta JASS
y de este momento del proyecto: la ruta a los repos cerrados por mes (`PycharmProjects\Julio\...`,
`PycharmProjects\Junio\...`) es un artefacto de que el pipeline todavía no vive en Postgres — a
escala SaaS eso es una tabla con partición por `mes_ano`, no una carpeta distinta por ciclo.
