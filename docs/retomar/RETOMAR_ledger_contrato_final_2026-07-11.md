# RETOMAR — Ledger 8_caja + 8b_estado_cuenta: CONTRATO FINAL cerrado · 2026-07-11

Sesión Opus. Amplía y reemplaza `docs/retomar/RETOMAR_8_caja_8b_estado_cuenta_2026-07-11.md`
(ese archivo queda como historia — el contrato de acá lo superó).

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **El CONTRATO DE INTERFAZ del ledger está CERRADO — 8 decisiones, no 5.**
   Vive **byte-idéntico** (verificado con `diff`, 156 líneas) en
   `8_caja/README.md` y `8b_estado_cuenta/README.md`. No re-debatir salvo duda
   de fondo nueva.
2. **Nada de código todavía.** Sigue Fase 1 mecánica (diagramas + formato HTML +
   `git mv`) y recién ahí Fase 2 (implementar).
3. **Próxima sesión: Sonnet** para lo mecánico. Opus solo para implementar
   `caja_repo.py` / `cuenta_repo.py` / `motor_aplicacion.py` (lógica no trivial)
   o si aparece una duda de diseño nueva.
4. Sigue pendiente el **sembrado histórico** (backfill) — sin cambios respecto
   al handoff anterior, ver esa sección abajo.

---

## Qué pasó esta sesión — de 5 decisiones a 8

La sesión anterior (mismo día) cerró el contrato 8↔8b con 5 decisiones (①-⑤),
pero solo a nivel del par 8/8b — **sin cruzarlo contra el resto del pipeline**.
El usuario pidió explícitamente auditar esa costura ("¿son 100% compatibles con
los otros módulos?"). Se encontraron **4 huecos reales**, verificados contra
código (no solo contra README):

| # | Hueco | Evidencia en código |
|---|---|---|
| 1 | `ABONO_ID` de yape colisiona cuando un depósito se reparte en 2+ predios | `motor_matching` tiene `MONTO_ASIGNADO` por fila con `ORIGEN+FECHA` repetido → mismo id, 2 predios |
| 2 | `5_cobranza` no puede emitir la aplicación que pide el contrato (le falta `ABONO_ID` y `MES_CARGO`) | `5_cobranza/main.py:_reconciliar_pagos_pueblo` llama `repo.registrar_pago(mz,lt,concepto,mes,monto)` — nunca vio el abono, aplica por delta de ciclo, no por cargo concreto |
| 3 | Migrar `seguimiento_pueblo` deja aplicaciones sin `ABONO_ID` → rompe `auditoria_conservacion()` | histórico no tiene el link |
| 4 | Deuda dispersa en 3 fuentes sin cruzar (`seguimiento`, `arrastre_consolidado`, `deuda_directiva.xlsx`) | `PREDIOS_INSTALACION_EXCLUIDOS` en `seguimiento_repo.py` es el parche que existe *porque* falta esta unificación; `deuda_directiva.xlsx` confirmado como pago-por-persona suelto |

Antes de resolver, el usuario amplió el lente (importante — quedó registrado
como principio, no como dato suelto):

**Lente de dominio (si cambia, cambia el código) vs lente de tecnología (no
cambia el código, es adaptador):**
- Multi-tenant real: **sí**, meta es dar esto a otras JASS del Perú (~25.000),
  "de lo general a lo específico por artefactos/config, no por fork de código".
- Escala: tope duro **10.000 predios/JASS** (cinturón de seguridad; nunca hubo
  más de 1.000 en la práctica). Esto es "muchas orgs chicas", NO big data —
  simplifica, no complica.
- Multiagente: el usuario no tiene el modelo claro. Se le corrigió: **no se
  diseña para multiagente ahora** — se diseñan tools idempotentes; pasar de 1
  agente a N es gratis después si el substrato (append-only + idempotente) es
  correcto.
- El agente **sí escribe** (más rápido), pero solo vía tools controladas, nunca
  SQL/mutación directa.
- Consistencia: plata = **siempre fuerte** (no eventual). A esta escala es gratis.
- Retención: **append-only para siempre** (ya decidido en la sesión previa).
- Postgres / Docker: confirmado como destino, pero es **adaptador** — no cambia
  el núcleo si el borde `*_repo.py` (patrón writer único) ya existe, que existe.

**Conclusión que se le dio al usuario (y que hay que sostener):** de todo esto,
solo 2 cosas se bakean en el código de HOY — el resto es adaptador y se resuelve
cuando le toque:
1. `JASS_ID` en cada evento del ledger (columna barata hoy, carísima de meter después).
2. Núcleo tenant-agnóstico: nada de rutas fijas / "la JASS" / globals — lo
   específico de cada JASS entra por config.

---

## El contrato final — 8 decisiones (resumen; el texto completo vive en los README)

| # | Decisión | Resuelve |
|---|---|---|
| ① | `ABONO_ID` determinista **sin `mz/lt`**; 1 depósito = 1 abono (el reparto a N predios vive en las APLICACIONES, no en la caja) | Hueco 1 |
| ② | Aplicación referencia el cargo vía `(CONCEPTO, MES_CARGO)` | trazabilidad — ya estaba |
| ③ | `SALDO_A_FAVOR` = concepto explícito, no residual | auditoría en 1 query — ya estaba |
| ④ | `DEVOLUCION` baja `SALDO_A_FAVOR` FIFO | plata parqueada fungible — ya estaba |
| ⑤ | Los 5(+1) conceptos entran al contrato hoy; `2_planilla` escribe CONSUMO/CORTE en Fase 2 | schema estable — ya estaba |
| ⑥ | **MOTOR DE APLICACIÓN es pieza propia** (`aplicar(cargos, abonos) → aplicaciones`), dentro de `8b`. `5_cobranza` y `2_planilla` **dejan de aplicar**, solo emiten CARGOS. Es la única pieza que ve caja+deuda juntas, por eso es la única que puede llenar `ABONO_ID`+`MES_CARGO` | **Hueco 2 — la decisión de más peso, cambia comportamiento real de `5_cobranza`** |
| ⑦ | `JASS_ID` en TODO evento (abono, cargo, aplicación); núcleo tenant-agnóstico, JASS-específico = config | escala a 25k JASS |
| ⑧ | Toda deuda de un predio es CARGO en `8b`, venga de `2_planilla` / `5_cobranza` / `arrastre_consolidado` / `deuda_directiva`. El histórico **no se recupera, se re-deriva**: se siembran los HECHOS (abonos+cargos) y el motor produce las aplicaciones — nunca se le inventa `ABONO_ID` a un pago viejo. Excepción: devoluciones se transcriben (cambio de régimen) | Huecos 3 y 4 |

**Principio rector que sostiene las 8 (repetirlo si se reabre la duda):**
> Cada módulo emite solo su HECHO (abono o cargo, inmutable, append-only). Nadie
> "aplica" salvo el motor único, que ve los dos lados y deriva la interpretación.
> Esto traduce 1:1 a tablas Postgres (`abonos`, `cargos`, `aplicaciones`) y es el
> substrato correcto para tools de agente idempotentes.

Autorización del usuario: las 3 decisiones nuevas (⑥⑦⑧) las tomé yo como
ingeniero senior a partir del contexto que dio el usuario (no las redactó él
línea por línea) — confirmado explícitamente ("autorizo") en esta sesión.

---

## READMEs tocados esta sesión (todos ya escritos, no releer contenido — solo confirmar que siguen así)

```
✅ 8_caja/README.md              — reescrito, CONTRATO con 8 decisiones
✅ 8b_estado_cuenta/README.md    — reescrito, CONTRATO byte-idéntico (verificado)
✅ 2_planilla/README.md          — + sección "Alimentación del ledger 8b (Fase 2)"
✅ 4_pagos/efectivo/README.md    — + pointer a 8_caja (fuente de ABONOS, clave natural)
✅ 4_pagos/yape/.../readme_motor_README.md — + pointer (1 yape = 1 abono, predio-agnóstico)
✅ README.md raíz                — pipeline: 7_cierre → [8_caja + 8b], tabla actualizada
⚠️ 5_cobranza/                   — NO TIENE README. Su rol nuevo (emite cargos,
                                    ya no aplica) está en el contrato de 8b pero
                                    su propio README se escribe recién en Fase 2,
                                    junto con el cambio de código real.
```

**Regla si se retoma y algo no calza:** el contrato en `8_caja/README.md` /
`8b_estado_cuenta/README.md` manda. Si algún módulo feeder (2_planilla, 4_pagos)
parece decir algo distinto, el README del feeder está desactualizado — corregirlo
a él, no al contrato.

---

## Qué falta — orden sugerido (sin cambios de fondo respecto al handoff anterior)

```
FASE 1 (terminar) — Sonnet, mecánico:
  [ ] diagrama_flujo_8_caja.html + diagrama_8_caja.html
  [ ] diagrama_flujo_8b_estado_cuenta.html + diagrama_8b_estado_cuenta.html
      — OJO: el diagrama debe mostrar el MOTOR DE APLICACIÓN como caja propia
        (no meterlo dentro de "cuenta_repo" ni de "5_cobranza")
  [ ] formato_*.html por cada output:
        - store de caja (8_caja) — entidad ABONO
        - store de cargos (8b) — entidad CARGO, con columna SOURCE
        - store de aplicaciones (8b) — entidad APLICACIÓN
        - vista/PDF de estado_cuenta
  [ ] git mv 7b_historial_pagos/ → 8_caja/
        (historial_repo.py → caja_repo.py; consultar.py se queda;
         importar_libros.py DEPRECADO — decidir borrar o dejar documentado)
  [ ] borrar dato sucio shared/reporte_acumulado_procesado/2026-05_historial.xlsx
  [ ] confirmar que README.md raíz sigue sincronizado (ya actualizado esta sesión)

FASE 2 (implementar) — Opus para lógica no trivial, Sonnet para lo mecánico:
  [ ] caja_repo.py (writer único, ABONO_ID determinista SIN mz/lt, JASS_ID en todo)
  [ ] importar_efectivo.py + importar_yape.py + importar_devoluciones.py
  [ ] cuenta_repo.py (registrar_cargo · registrar_aplicacion · registrar_ajuste,
      JASS_ID en todo, SOURCE por cargo)
  [ ] motor_aplicacion.py — PIEZA NUEVA, no existía en el handoff anterior.
      aplicar(cargos_abiertos, abonos) → aplicaciones, camina prioridad FIFO
      por MES_CARGO. Único lugar que llena ABONO_ID + MES_CARGO en una aplicación.
  [ ] 5_cobranza: QUITAR la aplicación directa (_reconciliar_pagos_pueblo hoy
      llama repo.registrar_pago con delta por ciclo) — pasa a emitir SOLO
      registrar_cargo(multa/acuerdos/convenio). El PREDIOS_INSTALACION_EXCLUIDOS
      se puede retirar una vez que arrastre_consolidado emita su cargo a 8b.
  [ ] tools/: estado_cuenta.py, explicar_reclamo.py, auditoria_conservacion.py
  [ ] migrar datos de seguimiento_pueblo.xlsx: sembrar como CARGOS históricos,
      dejar que el motor re-derive las aplicaciones (NO migrar aplicaciones viejas)
  [ ] escribir 5_cobranza/README.md recién acá, junto con el código

FASE 2 SEGUNDA ETAPA (aislada, tocar módulo vivo):
  [ ] 2_planilla emite CARGO consumo/corte al ledger (sección ya escrita en su README)
```

---

## Backfill histórico (Oct-25 → Jun-26) — sin cambios de fondo

Mismo plan que el handoff anterior (repo-copia, PARTE 0/1/2). Único ajuste: al
sembrar cargos históricos, NO traer las aplicaciones viejas de `seguimiento` —
se re-derivan con el motor nuevo (ver decisión ⑧). Detalle completo en
`docs/retomar/RETOMAR_8_caja_8b_estado_cuenta_2026-07-11.md` sección "Backfill histórico"
(esa parte del archivo viejo sigue vigente, no se reescribió).

Trampas ya identificadas (sin cambios): (A) replay independiente por mes; (B)
usuarios_id/maestro de la copia serán los de hoy. NO se necesita la planilla
para producir pagos de efectivo. NO correr downstream de 4_pagos para sembrar caja.

**Set de meses:** sigue pendiente confirmar cuáles hay con libro.

---

## Sub-decisiones diferidas (heredadas + nuevas de esta sesión)

- Consumo/corte histórico: sin cambios, decidir en Fase 2.
- `CARGO_ID` sintético: sin cambios, no antes de que haga falta.
- Colisión de nombres `registrar_pago`: renombrar la de 8b a `registrar_aplicacion`.
- `importar_libros.py`: deprecado, decidir borrar o documentar como descartado.
- **NUEVA — `deuda_directiva.xlsx` como concepto OTROS:** confirmado que hoy es
  un pago por `USER_ID` (no por `MZ/LT` directo). Falta decidir en Fase 2 si
  mapea 1:1 a predio o si `deuda_directiva` es una cuenta de otra naturaleza
  (persona vs predio) antes de emitirla como CARGO a `8b`.
- **NUEVA — 5_cobranza sin README:** cuando se escriba en Fase 2, debe incluir
  explícito el cambio de modelo (antes aplicaba, ahora solo emite cargo) para
  que quede como decisión documentada, no solo como diff de código.

---

## Pendientes sueltos (heredados, sin relación directa — no se tocaron esta sesión)

- mesa_4 A-1/D-6/S-8 (reclasificados en el input, siguen como reclamo activo).
- El ciclo julio en vivo (5_cobranza/5b desactualizados, 6_corte, 4b_reclamos) —
  ver `docs/retomar/RETOMAR_ciclo_julio_2026-07-09.md`.

---

## Modelo por tarea (recordatorio)

| Tarea | Modelo |
|---|---|
| Terminar Fase 1: diagramas, formato HTML, git mv | Sonnet |
| Implementar caja_repo / cuenta_repo / motor_aplicacion / tools (lógica no trivial) | Opus |
| Reescribir 5_cobranza para dejar de aplicar (lógica, cambia comportamiento) | Opus |
| Correr el backfill, validar outputs | Sonnet |
| Cualquier duda de diseño que reabra el contrato | Opus |
