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
- Genérica por `TIPO_RECLAMO` desde el día 1 — el mismo embudo sirve para "ya pagué mi medidor"
  (`convenio`), no solo `mes_anterior`.

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

**Alternativas descartadas:**
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
nombrado, ventana temporal), no de que hay que forzar una elección. Segunda señal: si
`GATE`+`Bloque A` resuelven menos del 50% de los reclamos "mes_anterior", el problema real no es de
búsqueda — es que la aplicación de pagos (motor_matching / 5_cobranza) tiene un bug sistemático
imputando mal `MES_ANTERIOR`, y el foco debe moverse ahí.

---

## Escala (lente de tenant)

El embudo (GATE → Bloque A → Bloque B) y la regla de ventana temporal son universales — cualquier
JASS reclama "ya pagué mes anterior" por el mismo miedo al corte. Lo que es específico de esta JASS
y de este momento del proyecto: la ruta a los repos cerrados por mes (`PycharmProjects\Julio\...`,
`PycharmProjects\Junio\...`) es un artefacto de que el pipeline todavía no vive en Postgres — a
escala SaaS eso es una tabla con partición por `mes_ano`, no una carpeta distinta por ciclo.
