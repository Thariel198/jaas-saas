# libro_mayor — Sistema de registro (bounded context)

**No es un módulo del pipeline.** Es la capa de registro permanente sobre la que
todo el sistema asienta y de la que todos leen — el equivalente al **libro mayor**
contable. Los módulos `0_padron … 7_cierre` son el *proceso* mensual (batch);
`libro_mayor/` es el *substrato*.

Por eso no lleva número: el número comunica orden dentro del pipeline, y esto no
es un paso del pipeline. `7_cierre` es el último módulo del proceso; después de
cerrar el mes, el proceso **asienta** el mes en este libro.

---

## Por qué un bounded context propio (y no `shared/` ni "módulo 8")

La frontera de esta carpeta es 1:1 con las fronteras del endgame:

```
libro_mayor/  (un bounded context)
   = 1 esquema Postgres     → tablas abonos, cargos, aplicaciones (todas con JASS_ID)
   = 1 servicio Docker      → docker build libro_mayor/  anda solo
   = 1 set de tools de agente → registrar_abono, registrar_cargo, aplicar, estado_cuenta…
```

- **No va en `shared/`**: `shared/` es un cajón de utilidades (templates, seguimiento_repo).
  Meter el system-of-record en un cajón de utilidades obliga a desenredarlo el día
  que se dockerice. El libro mayor es un dominio propio, no un util suelto.
- **No es "módulo 8"**: tener `.py` no lo hace un paso del pipeline —
  `shared/seguimiento_repo.py` también tiene `.py` y es infra, no módulo. El
  criterio es *qué* hace: un módulo transforma el flujo de **este mes**; el libro
  mayor es el registro **permanente** que acumula todos los meses.

---

## Un solo contexto, dos agregados — no dos módulos

`caja/` (abonos) y `estado_cuenta/` (cargos + aplicaciones) viven en el **mismo**
bounded context, no como dos módulos separados. Razón dura, no estética:

> Decisión ⑥+consistencia fuerte: el **motor de aplicación** es la única pieza que
> ve caja+deuda juntas e imputa abonos a cargos **atómicamente**. Con consistencia
> fuerte para la plata (ya decidido), separarlos en dos contextos forzaría
> transacciones distribuidas. Un solo contexto = una frontera transaccional =
> un esquema = un contenedor.

```
libro_mayor/
├── dominio/          reglas + entidades puras (MovimientoCaja, Cargo, Aplicacion,
│                     Ajuste), céntimos int, sin I/O — ver docs/retomar/RETOMAR_dominio_saldo_unico
├── caja/             caja_repo.py — writer único (abonos), se escribe desde cero.
├── estado_cuenta/    cuenta_repo.py + motor_aplicacion.py (cargos, aplicaciones)
├── tools/            estado_cuenta · explicar_reclamo · auditoria (agent-callable)
├── stores/           persistencia — xlsx HOY = adapter. Se reemplaza por Postgres;
│                     el resto del contexto NO cambia (patrón repo = puerto/adapter).
├── caja/README.md            CONTRATO del ledger (8 decisiones) — byte-idéntico ↓
└── estado_cuenta/README.md   CONTRATO del ledger (8 decisiones) — byte-idéntico ↑
```

---

## Estado

- **Fase 1 cerrada** (contrato de 8 decisiones en `caja/README.md` y
  `estado_cuenta/README.md`). Sin código todavía.
- **Diagramas y formatos HTML — hechos** (12/07/2026):
  `caja/docs/{diagrama_flujo_caja, diagrama_caja, formato_evento_caja}.html`,
  `estado_cuenta/docs/{diagrama_flujo_estado_cuenta, diagrama_estado_cuenta,
  formato_cargo, formato_aplicacion, formato_vista_estado_cuenta}.html`.
  El motor de aplicación aparece como caja propia destacada en el flujo.
- **Alimentación**: después de `7_cierre`, no en caliente (mes inmutable antes de
  asentar — invariante append-only).
- **Multi-tenant**: `JASS_ID` en todo evento; núcleo tenant-agnóstico, lo
  específico de cada JASS entra por config.
- **`7b_historial_pagos` eliminado**: era arquitectura distinta (pre-contrato).
  Se descartó en vez de migrar — `caja_repo.py` se escribe desde cero según el
  contrato. Recuperable en git history si alguna vez hace falta mirar cómo resolvió
  algún primitivo.
- **Decisión ⑨ — 4b_reclamos → ledger** (12/07/2026): al autorizar un reclamo de
  dinero, `4b_reclamos` invoca `identificar_abono` / `reasignar_abono` (caja) o
  `registrar_ajuste` (estado_cuenta); el motor deriva la aplicación que baja el
  saldo. 4b nunca escribe el saldo. Contrato en ambos README; READMEs de
  `4b_reclamos` y `2_planilla` actualizados (columna `BLANCO`/`DEVOLUCION` se retira).
- **Corrección ⑩ + decisiones ⑪ ⑫ (13/07/2026)** — ver detalle completo abajo.
  Dos errores de dominio corregidos (`comunitario` no es balde ·
  `deuda_directiva` no cruza a estado_cuenta) + `SUB_CONCEPTO`/cascada P1-P6
  completa + capacidad **Extracto de cuenta** diseñada y cerrada.

## Sesión 13/07/2026 — corrección de dominio + Extracto + arquitectura de render

**Contexto:** al generar la lista de corte de julio, dos bugs de datos reales
(ancla de corte mal calculada, canal yape no protegía de corte) llevaron a auditar
el contrato del ledger contra el código real. Se encontraron y corrigieron 2
errores de dominio en la decisión ⑩, y se cerró el diseño completo de una
capacidad nueva pedida por el usuario (extracto de cuenta). Todo documentado en
`caja/README.md` + `estado_cuenta/README.md` (contrato byte-idéntico, verificado).

**① Corrección — `comunitario` NO es un balde.** Colisión de nombres: "tanque
comunitario" (adjetivo, `5b_validacion/main.py:560`) se confundió con
`CONCEPTO=comunitario` de `motor_matching` — un mecanismo de **segregación** (1
depósito de un cobrador cubre N predios, se desgloza por lote,
`PADRE_SEGREGADO`→N×`HIJO_SEGREGADO`). Efecto: `buscar_abono` debe mirar también
los `HIJO_SEGREGADO`, no solo depósitos de primer nivel.

**② Corrección — `deuda_directiva` NO cruza a `estado_cuenta`.** Es un caso
específico (2 miembros de la directiva anterior repagando un faltante de caja) —
mismo tratamiento que `tanque`: balde caja-only, ya reconciliado en
`5b_validacion` como "otros conceptos". Nunca genera CARGO ni entra a la cascada.

**③ Decisión ⑪ — taxonomía completa + `SUB_CONCEPTO` en el CARGO.** Árbol
verificado contra `shared/seguimiento_pueblo.xlsx` (3 conceptos: MULTA/ACUERDOS/
CONVENIO) + `5_cobranza/main.py::_descomponer_saldo()` (cascada P1→P5 en el código
viejo; el destino agrega P6 OTROS). El CARGO ahora referencia `(CONCEPTO, SUB_CONCEPTO, MES_CARGO)`:

```
P1 AGUA·MANTENIMIENTO·arrastre → P2 CORTE_RECONEXION
   → P3 CONVENIO(medidor→instalación→reactivación) → P4 ACUERDOS(techado→campo)
   → P5 MULTA(reunión→faena) → P6 OTROS (slot residual, sin emisor hoy)
DEUDA_DIRECTIVA fuera de la cascada (balde, no cargo)
```
> Tramo pueblo corregido 2026-07-23 (dominio CA1): era MULTA→ACUERDOS→CONVENIO.
Sub-orden por regla de negocio (no técnica) — ver tabla completa en el contrato.
Data histórica sembrada sin desglose queda con `SUB_CONCEPTO` genérico hasta
re-sembrar (Fase 2, sembrado — no ahora).

**④ Decisión ⑫ — Extracto de cuenta, diseño CERRADO.** Responde el pedido
directo del usuario: *"quiero mi reporte de N meses con todos mis pagos, en PDF."*
5 decisiones (por predio · tanque separado informativo · rango=todo por default ·
on-demand usuario · template nuevo) + arquitectura (`extracto_predio()` tool
read-only cross-agregado) + layout de 3 secciones (deuda por concepto/sub ·
pagos y cómo se aplicaron · aportes voluntarios). Detalle completo en el
contrato de ambos README, sección "Extracto de cuenta — vista cross-agregado (⑫)".

**⑤ Arquitectura de render — decisión tomada, no numerada como decisión de
contrato** (es infraestructura, no dominio): ningún módulo de negocio imprime.
Cada bounded context arma sus filas (`extracto_predio()` en estado_cuenta,
`data_boletas()` en 3_boletas); un servicio stateless `render(plantilla, filas)
→ PDF` (hoy `shared/utils_render.py`, candidato a contenedor propio en Docker)
hace la conversión. `3_boletas` sigue en su lugar — NO se renombra a
`3_impresor` (mezclaría 2 mundos de datos). Detalle + diagrama de cajas en el
contrato de ambos README.

## Sesión 13/07/2026 (2ª) — decisión: migrar a ledger único antes de más tools

Al empezar a diseñar `riesgo_corte` (backlog de la sesión anterior) surgió que
`SALDO`/`MES_ANTERIOR` hoy tienen 3 dueños (`2_planilla`, `5_cobranza`,
`seguimiento_pueblo`) — la misma raíz de los bugs B4/B5/B7. Decisión: **Opción B**
— cerrar `libro_mayor/dominio/` (reglas puras extraídas de `6_corte`/`5_cobranza`)
y migrar agua/mant/corte al ledger ANTES de seguir con el catálogo de capacidades.
Detalle completo, inventario de ledgers existentes y roadmap B1-B4:
`docs/retomar/RETOMAR_dominio_saldo_unico_2026-07-13.md`.

## Sesión 24/07/2026 — decisión ⑬: `reactivación` como 3er SUB_CONCEPTO de CONVENIO

Caso real (M-12/Ramon Requez, S/266 "deuda 2019-oct.2025") destapó un balde no
modelado: predio dormido años, deuda que nace de una vez al reactivar, pagada en
cuotas negociadas — misma forma que medidor/instalación, por eso entra a CONVENIO
en vez de ser un balde nuevo. Cascada CONVENIO ahora medidor→instalación→reactivación.
Guardarraíl: verificar que el predio no tenía `MES_ANTERIOR` corriendo mes a mes en
ese período antes de sembrar (si no, es doble conteo con AGUA/P1). Fuente nueva:
hoja `REACTIVACION` en `obligaciones/inputs/SEGUMIENTO INSTALACIONES...xlsx`. Detalle
completo del contrato en `caja/README.md` y `estado_cuenta/README.md` (⑬, byte-idéntico).

## Pendientes Fase 2 (acumulados)

- **`3_boletas`**: alimentarse de `estado_cuenta` (saldo derivado) en vez de la
  columna de descuento de la planilla; mostrar un bloque "pagos reconocidos este mes"
  (mismo diseño que `estado_cuenta/docs/formato_vista_estado_cuenta.html`). La
  plantilla `PLANTILLA_boletas.docx` no tiene hoy línea de descuento/reconocimiento.
- **Retirar `BLANCO`/`DEVOLUCION`** de la planilla y de `5_cobranza` una vez que el
  ledger emita las aplicaciones equivalentes.
- **`formato_extracto.html`** — crear (contrato visual del extracto, 3 secciones,
  Sonnet — el diseño ya está cerrado).
- ~~`formato_cargo`/`formato_aplicacion.html` con `SUB_CONCEPTO` y cascada completa~~
  **HECHO 2026-07-16:** ambos actualizados a `SUB_CONCEPTO` + cascada P1-P6 + nombres
  canónicos (AGUA/MANTENIMIENTO/CORTE_RECONEXION) + `CARGO_ID` determinista. Falta aún
  `diagrama_flujo_estado_cuenta`/`diagrama_estado_cuenta.html` (visual, no bloquea código).
- **Re-sembrar MULTA/ACUERDOS/CONVENIO con `SUB_CONCEPTO`** cuando se backfillee
  el histórico — hoy siguen sumados (faena+reunión, techado+campo, medidor+instalación).
- **Capacidades nuevas — backlog aditivo (no bloquean el core del ledger), 13/07/2026:**
  `morosidad_total` · `lista_morosos` · `historial_predio` · `trazar_abono` (livianas) ·
  `explicar_saldo` · `recaudado_por_concepto` · `estado a-una-fecha` (medias).
  `riesgo_corte` (destino de 6_corte) y `arqueo_caja` + `conciliar_caja` (destino de 5b)
  **ya tienen su spec en los README de `6_corte`/`5b_validacion`** — no son punteros abiertos.
  El resto son tools **aditivas**: se agregan sin tocar el contrato (no generan deuda de
  integración). El umbral de corte ("2 meses" ≠ `MES_ANTERIOR>=8`) ya se resolvió: es
  amount-based, umbral = deuda mínima por config (ver `6_corte/README` + `dominio/politica_corte`).
