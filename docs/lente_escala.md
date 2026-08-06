# Lente de escala — la vara con la que se decide toda arquitectura de jass_system

> Doc de referencia único. Cuando una decisión de diseño tenga dimensión de arquitectura
> (¿un módulo o dos? ¿tabla o archivo? ¿acoplar o separar?), se decide **con este lente**,
> no con lo que es cómodo hoy para un solo cliente.

---

## El lente en una línea

**jass_system es un SaaS agentic multi-tenant en construcción**, no una herramienta para
una JASS. Escala objetivo: **~25 000 JASS**, tope **≤10 000 predios por JASS**. Destino
técnico: **PostgreSQL + Docker**, operado por **agentes** vía un set de tools idempotentes.

```
HOY                                    DESTINO (la vara)
─────────────────────────────────────────────────────────────────────
1 JASS (tupac_amaru)              →    25 000 JASS · ≤10k predios c/u
archivos .xlsx                     →    PostgreSQL (1 esquema, JASS_ID en todo)
scripts main.py corridos a mano   →    tools de agente idempotentes
carpeta del repo                   →    1 servicio Docker por bounded context
lógica hardcodeada a "la JASS"    →    núcleo tenant-agnóstico + config por JASS
```

Escala **modesta, no hiperescala**: 25k × 10k = 250M predios-fila tope teórico. Postgres
lo sostiene sin sharding exótico. El diseño no debe sobre-ingenierizarse para Google-scale;
debe ser **limpio, multi-tenant y config-driven**.

---

## Qué implica el lente al decidir (los criterios)

1. **Tenant-agnóstico por defecto.** Nada de rutas fijas, "la JASS", ni globals. Todo lo
   específico de un cliente (manzanas, tarifas, umbrales de corte, conceptos activos, montos
   de penalidad) entra por **config**, no por fork de código. Toda clave natural, todo id y
   toda query llevan `JASS_ID`.

2. **Diferencia de política ≠ diferencia de arquitectura.** Si dos cosas hacen las mismas
   operaciones y solo cambian parámetros (umbral, monto, condición), son **un motor
   config-driven**, no dos módulos. Dos copias = doble superficie que mantener × 25k tenants
   = riesgo de divergencia. La multi-tenancy se logra parametrizando, no duplicando.

3. **1:1 con Postgres.** Cada bounded context = 1 esquema. Cada "cosa que pasa" = una fila
   append-only con `JASS_ID`. Si en el diseño dos módulos escriben la misma tabla, son **un
   agregado**, no dos. Las dimensiones (motivo, canal, tipo) son **columnas**, no módulos.

4. **Superficie de tools para el agente, mínima y componible.** El agente opera mejor con
   pocas tools coherentes y parametrizadas (`riesgo_corte(jass, motivo?)`) que con muchas
   casi-idénticas (`riesgo_corte_agua` + `riesgo_corte_multa`). Menos tools, más parámetros.

5. **Idempotencia desde el día 1.** Re-correr = mismo resultado. Ids deterministas
   (`sha256[:8]` de la clave natural canónica), no secuenciales ni aleatorios. Es lo que hace
   seguros el backfill y el retry de un agente.

6. **El .xlsx de hoy es un adapter detrás del patrón repo.** Se reemplaza por Postgres sin
   que el resto del contexto cambie. No cablear lógica de negocio a openpyxl; la lógica pura
   va en `dominio/` (cero I/O, céntimos int, testeable con dicts).

---

## De lo general a lo específico — composición por config, no por módulo

Onboarding de una JASS nueva = **agregar o quitar artefactos**, pero el artefacto correcto
es el **CONCEPTO en un manifiesto de tenant**, no un módulo que se enchufa.

```
GENERAL (byte-idéntico en 25k JASS)          ESPECÍFICO (lo único que cambia)
─────────────────────────────────────────────────────────────────────────────
dominio/ · ledger · cascada P1-P6 completa   config/jass/<id>:
catálogo completo de conceptos + emisores      conceptos_activos: {agua, multa, convenio…}
motor de corte único                           cortes.triggers: [{motivo, cuando, penalidad}]
                                               umbrales/tarifas/canales
```

- **Quitar artefacto** = borrar una línea del manifiesto. JASS sin multas: no activa
  `multa` → nadie emite ese CARGO → el motor nunca aplica ahí → la cascada lo saltea sola.
  Cero código muerto, cero branch.
- **Agregar artefacto** = agregar una línea (concepto nuevo → su slot en la cascada, su
  emisor, opcionalmente su trigger de corte).

**La regla:** el motor (corte, ledger, cascada) es **uno e idéntico** para todos; la
variación de capacidad es 100% **data** (manifiesto), nunca una rama de código ni un módulo
toggle. Este es el "de lo general a lo específico" correcto: el general está completo, el
específico es un **subconjunto declarado**.

**Por qué NO un microkernel/plugins:** las capacidades de una JASS no son independientes —
todas comparten el mismo ledger/cascada/motor, son filas de una taxonomía compartida, no
plugins autónomos. Un plugin re-introduciría la duplicación (cada uno tendría que conocer la
cascada) y es sobre-ingeniería para escala "modesta, no hiperescala".

> Corolario que ya usamos: **`6b_corte_multas` no es un módulo agregable/quitable** — es el
> concepto `multa` marcado como *cut-trigger* en el manifiesto. JASS que no corta por multas
> no pone ese trigger. Un solo motor de corte; la config decide.

## Cómo se aplicó ya (precedentes)

- **`libro_mayor/` como bounded context único** (caja + estado_cuenta) = 1 esquema, 1
  servicio, 1 set de tools. El motor imputa atómicamente → 1 frontera transaccional.
- **`5_cobranza` se disuelve** en motor + queries + `obligaciones` — no era un módulo, era
  lógica que el lente reparte entre sus dueños reales.
- **Dinero en céntimos `int`, sin `TOL`** — exactitud a escala, no floats.
- **`JASS_ID` en cada evento** del ledger desde Fase 1, aunque hoy siempre sea `tupac_amaru`.

Ver el contrato en `libro_mayor/estado_cuenta/README.md` (§ "Multi-tenant", "Hacia dónde
escala") y `docs/arquitectura_pipeline_futuro.html` (§08).
