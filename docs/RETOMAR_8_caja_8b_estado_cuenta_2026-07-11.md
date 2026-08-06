# RETOMAR — 8_caja + 8b_estado_cuenta: diseño Fase 1 cerrado · Handoff 2026-07-11

> ⚠️ **SUPERADO** por `docs/RETOMAR_ledger_contrato_final_2026-07-11.md` (misma
> fecha, sesión posterior). Este archivo queda como historia: el contrato de
> interfaz de acá (5 decisiones ①-⑤) fue ampliado a 8 decisiones tras stress-test
> contra el resto del pipeline (huecos 1-4) y el lente agentic SaaS / multi-tenant.
> Leer el archivo nuevo primero.

Diseño cerrado esta sesión (Opus). Leer de arriba a abajo antes de tocar nada.
Reemplaza como handoff activo a `docs/RETOMAR_historial_pagos_2026-07-10.md`
(ese queda como historia: explica por qué `importar_libros.py` se deprecó y de
dónde salió 7b — que ahora se renombra a 8_caja).

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **Fase 1 (diseño) de `8_caja` + `8b_estado_cuenta` está CERRADA.** Los dos
   README están escritos y su `CONTRATO DE INTERFAZ` es **byte-idéntico**
   (verificado con `diff`). No re-debatir el diseño salvo que aparezca una duda
   nueva de fondo.
2. **Nada de código todavía.** Lo que sigue es terminar Fase 1 (diagramas +
   formato HTML + migración + README raíz) y recién ahí Fase 2 (implementar).
3. **Próxima sesión: Sonnet** para lo mecánico que queda de Fase 1 (transcribir
   diagramas/HTML, git mv, actualizar README raíz). Volver a Opus solo cuando se
   empiece a **implementar los repos** (`caja_repo.py`, `cuenta_repo.py` — lógica
   no trivial) o si aparece una duda de diseño.
4. Sigue pendiente el **sembrado histórico** (Oct-25 → Jun-26) por repo-copia —
   ver "Backfill" abajo. Ese es el trabajo operativo del usuario, no bloquea el
   build de los módulos.

---

## Contexto — de dónde salió esto

Veníamos de `7b_historial_pagos` (ledger de pagos para responder reclamos, diseño
Fase 1 cerrado el 10/07). Esta sesión, stress-testeando el diseño contra reclamos
reales (excesos, devoluciones, "pagué medidor pero consumió multa"), se descubrió
que **un solo ledger de pagos no alcanza** — el reclamo necesita también la deuda
por concepto y cómo se aplicó cada pago. Eso partió el diseño en **dos módulos
acoplados**:

```
  7b_historial_pagos  →  se renombra a  →  8_caja        (libro de caja)
  seguimiento_pueblo  →  se rediseña como →  8b_estado_cuenta (cuenta corriente)
```

El `b` marca el par acoplado (igual que 6/6b). El norte de todo esto: que evolucione
a **agentic SaaS** — que el usuario diga "el cliente reclamó X" y el agente encuentre
la respuesta o el error usando **tools** sobre ledgers limpios.

---

## La arquitectura — 3 capas

```
┌─ CAPA 1 · CAJA  (8_caja) ──────────────────────────────────┐
│  abono (+) · devolución (−) · por CANAL · TODO concepto     │
│  hecho de dinero, inmutable, regime-independent             │
│  writer: importar_efectivo / importar_yape / import_devol   │
│  responde "¿pagué? ¿cuándo? ¿canal? ¿cuánto?"               │
└───────────────────────────┬─────────────────────────────────┘
                            │ ABONO_ID (FK)
                            ▼
┌─ CAPA 2 · CUENTA CORRIENTE  (8b_estado_cuenta) ────────────┐
│  cargo (deuda) · aplicación (abono→concepto x prioridad)    │
│  · saldo derivado · 5 conceptos                             │
│  writer: 2_planilla (consumo/corte) + 5_cobranza (resto)    │
│  responde "tu pago fue a multa, aún debes medidor"          │
└───────────────────────────┬─────────────────────────────────┘
                            ▼
┌─ CAPA 3 · TOOLS DEL AGENTE  (dentro de 8b) ────────────────┐
│  estado_cuenta(predio) · explicar_reclamo(predio)           │
│  · auditoria_conservacion()  → read-only, junta capa 1+2    │
└─────────────────────────────────────────────────────────────┘
```

**Por qué 2 módulos y no 1** (decisión sostenida): caja y deuda tienen writers
distintos (import de canales vs reconciliación) y lifecycles distintos (hecho
inmutable vs interpretación con reglas de negocio). Separados, cada uno es una tool
limpia; la capa 3 hace el join. Meter conceptos/deuda dentro de la caja rompería su
invariante "sin concepto", que es lo que la hace simple y auditable.

**Complementarios, ninguno sobra.** `seguimiento_pueblo` hoy solo cubre 3 conceptos
(MULTA/ACUERDOS/CONVENIO), sin consumo/corte, con las aplicaciones sin link al pago,
y colapsó el histórico pre-junio en un CARGO de génesis → no puede responder
"¿pagué efectivo en mayo?". 8_caja llena exactamente ese hueco.

---

## El contrato de interfaz — las 5 decisiones (detalle en los README)

El `CONTRATO DE INTERFAZ` vive idéntico en `8_caja/README.md` y
`8b_estado_cuenta/README.md`. Resumen de las 5 decisiones tomadas (con su porqué
completo en el README):

| # | Decisión | En una línea |
|---|---|---|
| ① | **ABONO_ID = `{mes}-{canal}-{shorthash(clave_natural)}`** | determinista, no secuencial → re-sembrar en orden libre da el mismo id (idempotente) |
| ② | **Aplicación referencia `(CONCEPTO, MES_CARGO)`** | un abono de mayo puede saldar multa de marzo — sin el mes del cargo la tool adivina |
| ③ | **`SALDO_A_FAVOR` = concepto explícito** | el dinero no aplicado es una FILA que se suma, no un residual invisible → auditoría en 1 query |
| ④ | **DEVOLUCION baja el SALDO_A_FAVOR (FIFO, balance corrido)** | plata parqueada es fungible, no se matchea a una moneda; si excede, la auditoría lo marca |
| ⑤ | **CONSUMO/CORTE en el contrato desde hoy, escritos en Fase 2** | schema estable de 5 conceptos; tocar 2_planilla se aísla para no multiplicar riesgo |

**Invariante que atraviesa todo** (y que se expone como tool `auditoria_conservacion()`):
```
  Σ abonos = Σ aplicado a conceptos + Σ saldo_a_favor_vigente + Σ devoluciones
```
El agente la corre antes y después de cualquier corrección para auto-verificarse —
eso es lo que permite decirle "arreglá el reclamo de C-43" y confiar.

**El cambio de régimen (clave para el backfill):** antes los excesos se DEVOLVÍAN
(mes siguiente), ahora se ACUMULAN. El motor nuevo (5_cobranza) aplicaría la regla
de HOY a meses viejos → mentiría sobre la resolución del exceso. Por eso las
**devoluciones históricas se transcriben del libro**, no se re-derivan corriendo el
pipeline. Los pagos sí son hechos regime-independent y sí salen del motor.

---

## Backfill histórico (Oct-25 → Jun-26) — repo-copia

Cómo se consiguen los archivos que 8_caja importa, sin tipear a mano (datos ciegos)
y sin tocar producción. **Elegido: repo-copia** (correr el pipeline real en
`jass_system_copia` desechable, cosechar, sembrar en el repo real). Detalle completo
del mapa operativo en la conversación de esta sesión; resumen:

```
PARTE 0 (una vez, repo real):
  - borrar shared/reporte_acumulado_procesado/2026-05_historial.xlsx (dato sucio)
  - escribir importar_efectivo.py + importar_yape.py (leen de 8_caja/inputs/historico/)

PARTE 1 (por mes, en la copia):
  - resetear estado acumulado de la copia (vaciar reporte_acumulado_procesado,
    trazabilidades, reporte_mes_crudo — dejar SOLO el crudo del mes)
  - EFECTIVO: mesa_1.xlsx sintética (hoja1=Wilder, hoja2=Janet) → correr
    4_pagos/efectivo/main.py → pagos_efectivo.xlsx (0 discrep = 2 bloques cuadran)
  - YAPE: reporte crudo del mes (blanco Gmail) → motor_matching → _procesado.xlsx
    + llenar pendientes a mano
  - transcribir DEVOLUCIONES del libro (el motor no las reproduce bien)
  - cosechar los archivos → 8_caja/inputs/historico/

PARTE 2 (por mes, repo real):
  - importar_efectivo.py + importar_yape.py → siembran el mes
  - validar con la tool de consulta contra un predio conocido
  - commit por mes (store append-only idempotente → orden libre)
```

Trampas ya identificadas: (A) replay independiente por mes — no arrastrar estado
acumulado; (B) usuarios_id/maestro de la copia serán los de hoy → algún pago cae
"blanco", se llena a mano (mapea a identidad actual, mejor). NO se necesita la
planilla para producir los pagos (4_pagos/efectivo no la lee; motor_matching la usa
solo para cruce de deuda opcional). NO correr downstream de 4_pagos para sembrar caja.

**Set de meses:** quedó pendiente confirmar cuáles hay con libro (¿oct/nov/dic 2025?,
¿marzo?). El usuario mencionó mayo/abril/feb/ene + junio (excesos abiertos, al final)
+ julio (reclamos, al final).

---

## Qué falta — orden sugerido

```
FASE 1 (terminar) — Sonnet, mecánico:
  [ ] diagrama_flujo_8_caja.html + diagrama_8_caja.html
  [ ] diagrama_flujo_8b_estado_cuenta.html + diagrama_8b_estado_cuenta.html
  [ ] formato_*.html por cada output (store de caja, store de cuenta corriente,
      vista/PDF de estado_cuenta)
  [ ] git mv 7b_historial_pagos/ → 8_caja/ (historial_repo.py→caja_repo.py,
      consultar.py se queda; importar_libros.py DEPRECADO, decidir borrar)
  [ ] actualizar README.md raíz (pipeline: +8_caja +8b_estado_cuenta, quitar 7b)
  [ ] borrar dato sucio 2026-05_historial.xlsx

FASE 2 (implementar) — Opus para los repos, Sonnet para lo mecánico:
  [ ] caja_repo.py (writer único, ABONO_ID determinista, TIPO abono/devol)
  [ ] importar_efectivo.py + importar_yape.py + importar_devoluciones.py
  [ ] cuenta_repo.py (evoluciona seguimiento_repo: +consumo/corte en schema,
      +MES_CARGO en aplicación, +SALDO_A_FAVOR concepto, +link ABONO_ID)
  [ ] tools/: estado_cuenta.py, explicar_reclamo.py, auditoria_conservacion.py
  [ ] migrar datos de seguimiento_pueblo.xlsx al nuevo schema de cuenta corriente

FASE 2 SEGUNDA ETAPA (aislada, tocar módulo vivo):
  [ ] 2_planilla emite cargos CONSUMO/CORTE al ledger de cuenta corriente
```

---

## Sub-decisiones diferidas (no bloquean, anotar)

- **Consumo/corte HISTÓRICO:** ¿se backfillean los cargos de consumo/corte de meses
  viejos (necesita las planillas históricas), o solo desde que 2_planilla los emite?
  Decidir en Fase 2 según lo pida un reclamo real.
- **`CARGO_ID` sintético:** hoy la clave natural del cargo es `(CONCEPTO, MES_CARGO)`.
  Si algún concepto tiene >1 cargo en el mismo mes, promover a id explícito. No antes.
- **Colisión de nombres `registrar_pago`:** en la caja = recibo; en cuenta corriente
  = aplicación. Al implementar, renombrar la de 8b a `registrar_aplicacion` para que
  el agente no las confunda.
- **Destino de `importar_libros.py`:** deprecado por el enfoque de staging. Borrar o
  dejar documentado como enfoque descartado.

---

## Pendientes sueltos (heredados, sin relación directa)

- mesa_4 A-1/D-6/S-8 (reclasificados en el input, siguen como reclamo activo).
- El ciclo julio en vivo (5_cobranza/5b desactualizados, 6_corte, 4b_reclamos) —
  ver `docs/RETOMAR_ciclo_julio_2026-07-09.md`. NO se tocó esta sesión.

---

## Modelo por tarea (recordatorio)

| Tarea | Modelo |
|---|---|
| Terminar Fase 1: diagramas, formato HTML, git mv, README raíz | Sonnet |
| Implementar caja_repo / cuenta_repo / tools (lógica no trivial) | Opus |
| Correr el backfill, validar outputs | Sonnet |
| Cualquier duda de diseño que reabra Fase 1 | Opus |
