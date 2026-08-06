# RETOMAR — libro_mayor: bounded context + baldes + conexión 4b · 2026-07-12 15:57

Sesión Opus (diseño). Amplía y **reemplaza** como handoff activo a
`docs/RETOMAR_ledger_contrato_final_2026-07-11.md` (ese queda como historia — esta
sesión lo superó en 3 frentes: reubicación a `libro_mayor/`, eliminación de `7b`, y
las decisiones ⑨ y ⑩).

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **El ledger ya no es "módulo 8".** Es el bounded context `libro_mayor/` (sin
   número, substrato tipo `shared/`), con dos agregados: `caja/` y `estado_cuenta/`.
   `7_cierre` es el último módulo del pipeline.
2. **`7b_historial_pagos` fue ELIMINADO** (era arquitectura pre-contrato). `caja_repo.py`
   se escribe desde cero. Recuperable en git history.
3. **El contrato tiene ahora 10 decisiones (①–⑩)**, byte-idéntico en
   `libro_mayor/caja/README.md` y `libro_mayor/estado_cuenta/README.md` (verificado con `diff`).
4. **Nada de código todavía.** Sigue Fase 1. Faltan HTML (van en **Sonnet**) y luego Fase 2.
5. **Trabajo operativo que disparó todo esto:** 54 reclamos abiertos en `4b_reclamos`
   que necesitan cargar 10 meses históricos para resolverse. Ver sección "Trabajo operativo".

---

## Qué pasó esta sesión (4 bloques)

### 1 · Reubicación arquitectónica — `8_caja`/`8b` → `libro_mayor/`
Decisión senior: el ledger es un **bounded context**, no un módulo del pipeline. La
frontera de la carpeta = 1 esquema Postgres = 1 servicio Docker = 1 set de tools de
agente. Meterlo en `shared/` (cajón de utils) o numerarlo como "paso 8" eran errores
de categoría. Un solo contexto con dos agregados adentro porque el **motor de aplicación**
ve caja+deuda juntas y exige consistencia fuerte (una frontera transaccional).

```
libro_mayor/                  ← bounded context (sin número, substrato)
├── README.md
├── caja/            (agregado: abonos + egresos)   caja_repo.py
├── estado_cuenta/   (agregado: cargos + aplicaciones + MOTOR)   cuenta_repo.py + motor_aplicacion.py
├── dominio/  tools/  stores/   (esqueleto .gitkeep)
```

`7_cierre` = último módulo del pipeline; `libro_mayor/` se alimenta DESPUÉS de cierre.

### 2 · `7b_historial_pagos` eliminado
Tenía código (historial_repo.py 303 líneas, etc.) de arquitectura distinta (pre-contrato).
Decisión del usuario: descartar en vez de migrar — reescribir limpio cuesta menos que
reconciliar. Recuperable en git history.

### 3 · Decisión ⑨ — conexión `4b_reclamos` → ledger
Cómo se corrige la deuda cuando se resuelve un reclamo (ej. "ya pagué mayo" = un blanco):

```
4b_reclamos AUTORIZA → emite el HECHO que faltaba (NO escribe la aplicación)
        │
   invoca una tool del ledger según el tipo:
     · blanco reclamado    → identificar_abono(abono_id, mz, lt, reclamo_id)   [caja]
     · pago mal atribuido  → reasignar_abono(abono_id, mz, lt, reclamo_id)     [caja]
     · cargo incorrecto    → registrar_ajuste(mz, lt, concepto, monto, reclamo_id)  [estado_cuenta]
        │
   el MOTOR re-corre → deriva la APLICACIÓN → el saldo baja (derivado)
        │
   7_cierre consolida → boleta del próximo mes muestra el saldo corregido
```

Principio: **4b nunca escribe el saldo.** Un blanco es un abono con `DESTINO=PENDIENTE`;
identificarlo NO edita el abono (append-only) — agrega un evento que le asigna predio.
La columna `BLANCO`/`DEVOLUCION` de la planilla **se retira** (era descuento manual que
se pisa al regenerar y nunca cuadra en 5b). La boleta pasa a leer estado_cuenta.

### 4 · Decisión ⑩ — baldes + GASTO (auditoría de tesorería)
El evento de caja se generalizó de "ABONO" a **MOVIMIENTO DE CAJA** con tres campos
nuevos: `DIRECCION` (INGRESO/EGRESO), `BALDE`, `DESTINO` (PREDIO/CONCEPTO/PENDIENTE).
Motivo: la caja debe responder "cuánto entró y cuánto salió por mes" y reproducir la
validación de `5b_validacion`, que ya desglosa en baldes hoy.

```
                A PREDIO (MZ-LT)          POR CONCEPTO (balde, sin predio)
INGRESO (+)     agua                       tanque · deuda_directiva
  (TE PAGÓ)     blancos → PENDIENTE
EGRESO  (−)     devolucion · retorno       honorario · gasto · comunitario  (GASTO institucional)
```

Cruzan a estado_cuenta (tocan deuda): `agua`, `deuda_directiva`, `devolucion`/`retorno`,
y `blancos` tras identificarse (⑨). **Solo caja** (no tocan deuda): `tanque`/`comunitario`
(**aporte voluntario confirmado, no obligatorio**) y `honorario`/`gasto` (egreso institucional).

Clave: **`BALDE` (caja) ≠ `CONCEPTO` (deuda)**. Son vocabularios distintos que el motor conecta.
ids: `ABONO_ID` (ingreso) · `DEVOLUCION_ID`/`GASTO_ID` (egreso).

---

## El contrato — 10 decisiones (resumen; texto completo en los README byte-idénticos)

| # | Decisión |
|---|---|
| ① | `ABONO_ID` determinista sin mz/lt; 1 depósito = 1 abono |
| ② | Aplicación referencia el cargo vía `(CONCEPTO, MES_CARGO)` |
| ③ | `SALDO_A_FAVOR` = concepto explícito, no residual |
| ④ | `DEVOLUCION` baja `SALDO_A_FAVOR` FIFO |
| ⑤ | Conceptos en el contrato hoy; 2_planilla escribe CONSUMO/CORTE en Fase 2 |
| ⑥ | MOTOR DE APLICACIÓN pieza propia; los feeders solo emiten cargos |
| ⑦ | `JASS_ID` en todo evento; núcleo tenant-agnóstico (meta 25k JASS) |
| ⑧ | Toda deuda = CARGO (varias fuentes); histórico re-derivado, no migrado |
| ⑨ | **4b_reclamos emite la resolución como HECHO; el motor deriva la aplicación** |
| ⑩ | **Movimiento de caja lleva DIRECCION+BALDE+DESTINO (ingreso+egreso, incl. GASTO); BALDE≠CONCEPTO** |

---

## Archivos tocados esta sesión (todos ya escritos)

```
CREADOS:
  libro_mayor/README.md                                   (doc del bounded context)
  libro_mayor/caja/docs/diagrama_flujo_caja.html
  libro_mayor/caja/docs/diagrama_caja.html
  libro_mayor/caja/docs/formato_evento_caja.html          (era formato_abono.html)
  libro_mayor/estado_cuenta/docs/diagrama_flujo_estado_cuenta.html
  libro_mayor/estado_cuenta/docs/diagrama_estado_cuenta.html
  libro_mayor/estado_cuenta/docs/formato_cargo.html
  libro_mayor/estado_cuenta/docs/formato_aplicacion.html
  libro_mayor/estado_cuenta/docs/formato_vista_estado_cuenta.html

MOVIDOS/RENOMBRADOS:
  8_caja/README.md          → libro_mayor/caja/README.md          (+ contrato ⑨ ⑩)
  8b_estado_cuenta/README.md → libro_mayor/estado_cuenta/README.md (+ contrato ⑨ ⑩)

ELIMINADOS:
  7b_historial_pagos/  (git rm — recuperable)
  shared/reporte_acumulado_procesado/2026-05_historial.xlsx  (dato sucio)

EDITADOS (compatibilidad):
  README.md (raíz)          — pipeline termina en 7_cierre; libro_mayor como substrato
  4b_reclamos/README.md     — rol dual: datos→DATA_boletas (hoy) · dinero→ledger (Fase 2, ⑨)
  2_planilla/README.md      — BLANCO/DEVOLUCION marcadas para retiro (⑨)
  4_pagos/efectivo/README.md · motor_matching README — rutas actualizadas a libro_mayor/
```

Verificado: contrato byte-idéntico entre los dos README; las 3 tools (`identificar_abono`,
`reasignar_abono`, `registrar_ajuste`) nombradas igual en los 5 README; sin residuos de
`TIPO=ABONO/DEVOLUCION` viejo.

---

## Pendientes

### A · HTML — van en SONNET (no en Opus)
> ⚠ Error de esta sesión: se escribió `formato_evento_caja.html` y se editaron diagramas
> estando en Opus. El HTML se transcribe en Sonnet (el diseño ya está cerrado). Corregir el proceder.

- [ ] Diagramas de flujo `caja`/`estado_cuenta`: mostrar los **baldes** y la **entrada de 4b**
      (hoy solo se alinearon las referencias que contradecían el contrato).
- [ ] `diagrama_caja.html` / `diagrama_estado_cuenta.html`: pasar al vocabulario `BALDE`/`DIRECCION`.

### B · Trabajo operativo — 54 reclamos (lo que disparó el rediseño)
Ver sección "Trabajo operativo" abajo. Bloqueado por Fase 2 (los repos/motor no existen).

### C · Fase 2 — implementar (OPUS para lógica no trivial)
- [ ] `caja/caja_repo.py` (writer único, `registrar_movimiento` con DIRECCION/BALDE/DESTINO,
      id determinista, `identificar_abono`/`reasignar_abono`)
- [ ] `caja/importar_efectivo.py` · `importar_yape.py` · `importar_egresos.py`
- [ ] `estado_cuenta/cuenta_repo.py` (`registrar_cargo`/`registrar_aplicacion`/`registrar_ajuste`)
- [ ] `estado_cuenta/motor_aplicacion.py` (PIEZA NUEVA — `aplicar(cargos, abonos) → aplicaciones`, FIFO)
- [ ] `estado_cuenta/tools/` (estado_cuenta · explicar_reclamo · auditoria_conservacion)
- [ ] `5_cobranza`: dejar de aplicar, pasar a emitir solo cargos (cambia comportamiento — Opus)
- [ ] `3_boletas`: alimentarse de estado_cuenta (saldo derivado) + bloque "pagos reconocidos";
      `PLANTILLA_boletas.docx` no tiene línea de descuento/reconocimiento hoy
- [ ] Retirar columnas `BLANCO`/`DEVOLUCION` de planilla + 5_cobranza
- [ ] `5b_validacion` se simplifica: validar leyendo baldes de caja (1 query vs 8 lecturas)

---

## Trabajo operativo — 54 reclamos (contexto que dio el usuario)

El usuario tiene **54 reclamos abiertos** en `4b_reclamos` (tipo "ya te pagué medidor",
"pagué tal mes"). Para resolverlos hay que **verificar los pagos históricos**: si de verdad
pagó, buscar a qué deuda fue ese pago.

**Meses a cargar (re-correr desde crudo — confirmado por el usuario):**
```
ago · sep · oct · nov · dic 2025   +   ene · feb · mar · abr · may 2026   = 10 meses
Junio  → ya en una copia aparte del pipeline (tiene excesos abiertos)
Julio  → ciclo en vivo, en curso
```

**Cómo se resuelven (decidido: cargar directo a libro_mayor):** sembrar cada mes como
HECHOS (abonos + cargos) en el modelo nuevo; el motor re-deriva las aplicaciones y eso
resuelve los reclamos. Un blanco reclamado se corrige con `identificar_abono` (⑨).

**Bloqueo real:** esto necesita `caja_repo.py`, `cuenta_repo.py` y `motor_aplicacion.py`,
que aún NO existen → Fase 2 primero. El backfill (repo-copia, PARTE 0/1/2) está descrito
en `docs/RETOMAR_ledger_contrato_final_2026-07-11.md` y sigue vigente con un ajuste:
al sembrar, NO traer aplicaciones viejas de seguimiento — se re-derivan.

Trampas ya identificadas: (A) replay independiente por mes (no arrastrar estado acumulado);
(B) usuarios_id/maestro de la copia serán los de hoy → algún pago cae "blanco", se llena a mano.

---

## Sub-decisiones cerradas / notas

- **Tanque = aporte voluntario, NO obligatorio** → balde solo-caja, no genera cargo. CERRADO.
- **BLANCO/DEVOLUCION de la planilla**: se retiran en Fase 2 (⑨).
- **`deuda_directiva`**: es cargo OTROS (cruza a estado_cuenta). Ya en el contrato.
- **Boletas (`3_boletas`)**: hay que corregirlas para leer estado_cuenta — DIFERIDO a Fase 2.
- **`devoluciones_acumulados.xlsx`** (motor_matching) está mal nombrado: contiene todos los
  PAGASTE (egresos), no solo devoluciones. Renombrar es alcance de motor_matching, no bloquea.

---

## Modelo por tarea (recordatorio)

| Tarea | Modelo |
|---|---|
| HTML (diagramas de flujo a baldes, alinear diagramas) | **Sonnet** |
| Implementar caja_repo / cuenta_repo / motor_aplicacion / tools | **Opus** |
| Reescribir 5_cobranza para dejar de aplicar | **Opus** |
| Correr backfill, sembrar meses, validar outputs | Sonnet (Haiku para lo mecánico) |
| Resolver reclamos con causa raíz no obvia / decisión de negocio | Opus |
| /cierre, actualizar memoria/docs | Sonnet/Haiku |

**Al retomar:** decidir si primero se terminan los HTML (Sonnet) o se arranca Fase 2 (Opus).
El trabajo operativo (54 reclamos) depende de Fase 2, así que el camino crítico es Opus.
