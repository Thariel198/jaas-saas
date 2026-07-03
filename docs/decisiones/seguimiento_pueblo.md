# Decisión de diseño — seguimiento_pueblo

Fecha: 2026-07-02
Estado: Diseño aprobado en conversación · Fase 2 y Fase 3 completas · código corrido y verificado contra datos reales

## Fase 3 (código) — HECHA (2026-07-02)

`shared/seguimiento_repo.py` (API + escritura atómica), `shared/sembrar_seguimiento_pueblo.py` (siembra:
712 CARGO, MULTA S/12,710 · ACUERDOS S/14,505 · CONVENIO S/10,065), `5_cobranza/main.py`
(`_reconciliar_pagos_pueblo`, reconciliación por delta), `2_planilla/main.py` (lee `get_saldos_bulk()` en
vez del consolidado para MULTA/ACUERDOS/CONVENIO). 4 suites de test, todas OK (salvo el FAIL pre-existente
de `test_cobranza.py` ajeno a este cambio).

**Bugs encontrados y corregidos durante la verificación con datos reales:**
1. **Escritura no atómica** — `wb.save()` directo podía corromper el archivo si el proceso se cortaba a
   mitad de camino (pasó una vez en una corrida real). Fix: `_save_atomic()` (temp file + `os.replace()`,
   con retry ante `PermissionError` transitorio de Windows — antivirus/indexador). Verificado con test que
   simula el corte + con el fallo real que ocurrió en producción.
2. **`MES_SIEMBRA` mal fechado** — la siembra usaba `2026-07`, pero 5_cobranza sigue procesando el ciclo
   `2026-06` (julio bloqueado por D3). Como `get_saldo`/`get_saldos_bulk` filtran `MES <= mes_consultado`,
   el CARGO de julio quedaba invisible al preguntar el saldo a fin de junio → saldo salía negativo. Fix:
   `MES_SIEMBRA = "2026-06"`. Corregido re-sembrando desde cero (append-only no permite editar el MES de
   un evento ya escrito) — 100% re-derivable, sin pérdida de información real.
3. **Reconciliación sin exclusión de los 6 predios de instalación** — `_reconciliar_pagos_pueblo` no
   conocía `PREDIOS_INSTALACION_EXCLUIDOS` (que sí respeta la siembra) → le creó un evento PAGO de CONVENIO
   a B-20 sin CARGO correspondiente, saldo=-40 falso. Fix: mover la lista a `seguimiento_repo.py` (fuente
   única) e importarla desde ambos lados. La 1 fila mala ya escrita se corrigió a mano (backup en
   `shared/backups/`, no con re-run completo — precedente: mismo patrón que el fix de B7) porque es un
   artefacto de bug, no un evento de negocio real que el append-only deba proteger.

**Estado final verificado:** `shared/seguimiento_pueblo.xlsx` = 1094 filas (712 CARGO + 382 PAGO), 0
duplicados, 0 predios de instalación con filas de CONVENIO. `shared/vista_seguimiento_pueblo.xlsx`
regenerada. MULTA/ACUERDOS calzan exacto contra el consolidado ya validado antes; CONVENIO corregido de
S/3,201 (parcial, bug viejo) a la deuda real completa.

---

**Problema:**
`arrastre_consolidado` (5_cobranza, DE5) ya lleva mes a mes el saldo de MULTA, ACUERDOS_ASAMBLEA y CONVENIO,
pero solo como una **foto del mes** — no queda historial de cuánto se pagó cada mes, ni de cuándo nació cada
deuda. Hoy esa información vive a mano, dispersa en 3 archivos distintos (`padron_secundario/reunion`,
`padron_secundario/faena`, `mayo-planilla/Cobro medidores`, `SEGUIMIENTO/NUEVAS INSTALACIONES`), cada uno
con columnas que crecen mes a mes (`MARZO`, `MAYO`, `JUNIO`...) y sin trazabilidad de quién pagó qué mes.
Cuando un usuario pregunta "¿cuánto pagué en julio? ¿y en agosto?", hoy no hay una respuesta directa —
hay que reconstruirla a mano de varios archivos.

Además, el diseño original de esta sesión (un "ledger de génesis" separado del arrastre, para sembrar
obligaciones nuevas de julio) resultó estar resolviendo el problema equivocado: génesis es solo una foto
de arranque, no dá seguimiento continuo. El problema real es que no existe ningún lugar que registre el
historial de pagos por concepto — génesis y seguimiento son cosas distintas y hay que resolver la segunda.

**Criterios:**
- Responder "¿cuánto debía, cuánto pagó cada mes, cuánto debe ahora?" para MULTA, ACUERDOS y CONVENIO,
  por predio, sin reconstruir nada a mano.
- Sin doble contabilidad — un solo dueño del dato (nunca dos archivos con el mismo saldo, la clase de bug de B7).
- El código imita el proceso manual real: *"del arrastre se sacaban los pagos, se ponían en seguimiento,
  salían los saldos"* (palabras del usuario) — no una abstracción inventada desde cero.
- Preparado para agentic SaaS: cada escritura es una función nombrada con audit (`source`/`audit_ref`),
  no un script que abre el Excel directo.
- Simple — Regla del Tres: no anticipar infraestructura para casos que no ocurrieron 3+ veces.

**Enfoque elegido — registro event-sourced, lote como llave:**

```
Eventos (append-only, nunca se edita una fila)
  CARGO   → deuda nueva (siembra inicial o cargo nuevo del mes)
  PAGO    → 5_cobranza registra la porción de un pago que va a ese concepto
  AJUSTE  → corrección manual, con motivo obligatorio

Estado derivado (nunca se guarda mutable)
  SALDO(mz, lt, concepto, mes) = Σcargos − Σpagos ± Σajustes

Repo — único writer
  shared/seguimiento_repo.py → registrar_cargo() / registrar_pago() / registrar_ajuste()
                                get_saldo() / estado_cuenta()

Persistencia
  shared/seguimiento_pueblo.xlsx — registro largo, 1 fila = 1 evento
```

- **La llave es `(MZ, LT)`, no la persona.** Igual que el resto del sistema. Reasignación de lote es un
  evento puntual (ocurrió una vez, 16 casos, ya resuelto con `reasignaciones_lotes_2026-06.xlsx`) — no un
  patrón recurrente que justifique acuñar una identidad de persona nueva.
- **Génesis = el primer evento CARGO de cada predio.** No hay un ledger de génesis separado — sembrar la
  deuda inicial de julio es exactamente el mismo mecanismo que registrar una multa nueva en agosto.
- **Guardar largo, mostrar ancho.** El registro (fuente de verdad) es una tabla con schema fijo que no
  cambia nunca. El "estado de cuenta" ancho (DEUDA·PAGO·SALDO por mes, el formato que el usuario usa hoy
  a mano) es un **pivot regenerable** desde el registro — nunca la fuente.
- **`arrastre_consolidado` se achica.** Deja de cargar MULTA/ACUERDOS/CONVENIO por dentro — esos 3 pasan
  a ser propiedad de `seguimiento_repo`. El consolidado sigue siendo dueño de agua+corte (sin cambios ahí).
  `2_planilla` pasa a leer dos fuentes en vez de una: el consolidado (agua+corte) + `get_saldo()` (pueblo).

**Alternativas descartadas:**
- *Ledger de génesis aparte, separado del seguimiento* (diseño inicial de esta misma sesión) — resolvía
  solo la siembra de julio, no el problema real (falta de historial mes a mes). Se descartó cuando el
  usuario aclaró: "génesis no hace seguimiento continuo, solo es un saldo inicial de un momento dado".
- *Columnas por mes creciendo indefinidamente* (`JULIO`, `AGOSTO`, `SEPTIEMBRE`...) — es el formato manual
  actual. Funciona para lectura humana pero el schema cambia cada mes, imposible de consumir por un agente
  sin romper cada vez que se agrega una columna. Se conserva como **vista** (pivot), no como fuente.
- *Deuda de pueblo sigue a la persona, no al lote* — propuesto por el diseñador (Opus) razonando sobre el
  caso de reasignación de lote. El usuario corrigió: es sobreingeniería para un caso que pasó una vez en
  el pasado; si el usuario no lo menciona como problema recurrente, no se diseña infraestructura para eso.
  Ver memoria `feedback_no_sobreingenieria_edge_case_raro`.
- *5_cobranza reparte el pago pero el usuario marca a mano la columna del mes en la hoja de seguimiento* —
  descartado: el usuario confirmó que 5_cobranza ya tiene los datos del pago y es lo que hacía manualmente
  ("sacaba los pagos del arrastre, los ponía en seguimiento") — automatizar ese paso es la mejora real,
  no agregar trabajo manual nuevo.
- *Saldo como celda mutable, recalculada cada mes* — mismo patrón que causó el bug B7 (dual-writer,
  lost-update) en `shared/planilla_mes`. Un saldo derivado de eventos nunca tiene ese riesgo porque no hay
  celda que dos escrituras puedan pisarse.

**Señal de alerta:**
- Si `get_saldo()` de un predio no coincide con lo que muestra `estado_cuenta()` para el mismo mes → el
  pivot tiene un bug, revisar que sume CARGO−PAGO±AJUSTE en el orden correcto de eventos.
- Si un evento aparece con `SOURCE` vacío o `AUDIT_REF` vacío → el repo debería haberlo rechazado; revisar
  validación de inputs en `registrar_*`.
- Si el saldo de un predio queda negativo sin que sea un AJUSTE explícito → señal de pago duplicado o mal
  repartido en el waterfall de 5_cobranza — revisar `audit_ref` del pago para encontrar el origen.
- Si `arrastre_consolidado` y `seguimiento_pueblo` alguna vez muestran valores distintos para el mismo
  predio/concepto/mes → el doble-writer que este diseño evita se reintrodujo en algún lado — revisar que
  ningún script siga escribiendo MULTA/ACUERDOS/CONVENIO directo en `shared/planilla_mes` ni en el
  consolidado.

---

## Referencias

- Arquitectura visual (5 capas): `shared/docs/diagrama_seguimiento_pueblo.html`
- Flujo de 5 segundos (LEE/GENERA por paso): `shared/docs/diagrama_flujo_seguimiento_pueblo.html`
- Contrato de columnas del registro: `shared/docs/formato_seguimiento_pueblo.html`
- API y invariantes: `shared/README.md` (sección "Patrón Event-Sourced — seguimiento_repo")
- Concepto de API explicado (aprendizaje): `docs/aprendizaje/api_concepto_20260702.html`
- Bug que este diseño evita por construcción: `docs/aprendizaje/writer_unico_desincronizacion_20260701.html` (B7)
- Memoria de la corrección "no sobreingeniería": `feedback_no_sobreingenieria_edge_case_raro.md`

## Artefactos superados por este diseño (borrados 2026-07-02)

Del intento anterior en esta misma sesión ("ledger de génesis" separado, antes de aclarar que el problema
real era seguimiento continuo): `2_planilla/docs/genesis_obligaciones_diseno.html` y
`shared/docs/formato_obligaciones_genesis.html` — reemplazados por el diseño event-sourced de este documento.

## Fase 3 — HECHA, ver sección arriba. Deuda pendiente (menor, no bloquea):

- `6_corte/seguimiento.py` y `6b/seguimiento_multas.py`: no se re-corrieron tras este cambio (mismo criterio
  que B7 — no mutar estado cerrado de junio innecesariamente). Revisar cuando se reactive el ciclo de corte.
- Performance: cada `registrar_*` hace un read+write completo del archivo (O(n) por evento) — la siembra +
  reconciliación completa tardó ~20 min con ~1700 eventos. Aceptable para volumen actual (mensual, ~700-1000
  eventos/mes), pero si el archivo crece mucho más, considerar batch-write o cache en memoria durante una
  corrida completa.
