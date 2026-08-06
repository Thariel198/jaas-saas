# obligaciones — Emisor de cargos no-agua (event-driven)

**No es un módulo del pipeline** (no lleva número) y **no es parte de `libro_mayor`**.
Es el **emisor de cargos** de las obligaciones que no nacen de la lectura del medidor —
multa, acuerdos, convenio— hacia el ledger. Su trigger es un **evento** (una asamblea,
una faena, un convenio firmado), no el ciclo mensual; por eso no tiene número, igual que
el ledger.

> **Estado:** Fase 1 — diseño **detallado cerrado** (2026-07-15): fuentes crudas de la
> secretaria mapeadas con columnas, split `SUB_CONCEPTO` resuelto, tabla de reasignaciones
> COFOPRI cerrada. **Sin código todavía** — se codifica después de `dominio/` + repos +
> motor (roadmap B). MULTA se siembra **bruta** (cada ausencia = cargo); la reconciliación
> del gap de asistencia la hace el **motor** al aplicar los abonos de caja (ver "Reconciliación
> de MULTA" abajo), no un paso previo manual.

---

## Por qué existe

En el ledger, cada deuda es un **CARGO** (un HECHO: "nació una obligación"). Algo tiene
que **emitir** ese hecho. Hay dos emisores, por trigger distinto:

```
  2_planilla     → agua · mantenimiento · corte_reconexion   (driven por LECTURA, mensual)
  obligaciones   → multa · acuerdos · convenio                (driven por EVENTO, ad-hoc)
```

`5_cobranza` **nunca** emitió estos cargos — el código lo confirma: solo llamaba
`registrar_pago` / `registrar_ajuste` (reconciliaba pagos), nunca `registrar_cargo`. La
deuda de multa/acuerdos/convenio la **siembra** hoy `shared/sembrar_seguimiento_pueblo.py`
(un script marcado "desechable"). `obligaciones/` es la versión **permanente y de primera
clase** de ese emisor.

## Por qué afuera del ledger

El ledger es **writer-only-vía-repo**: `estado_cuenta` expone `registrar_cargo(...)` y
nada entra sin pasar por ahí. Los emisores son **clientes**, no parte del ledger:

```
   obligaciones  ──registrar_cargo──►  libro_mayor/estado_cuenta (cuenta_repo, writer único)
```

El ledger **no debe saber qué es una "faena" ni un "convenio firmado"** — solo recibe un
CARGO. Ese conocimiento de dominio vive acá, en el cliente. (Mismo patrón que `2_planilla`
y `4_pagos`: emiten hechos al ledger sin ser parte de él.)

---

## Qué hace (dos modos, un solo código)

```
obligaciones/
   registrar_multa(...)      nace una multa    → registrar_cargo(MULTA, sub, mes) → estado_cuenta
   registrar_acuerdo(...)    asamblea acuerda   → registrar_cargo(ACUERDOS, sub, mes)
   registrar_convenio(...)   se firma convenio  → registrar_cargo(CONVENIO, sub, mes)
   backfill(...)             MODO MIGRACIÓN: corre en lote sobre lo histórico
                             = re-siembra los cargos que hoy están en seguimiento_pueblo
```

- **Steady-state:** cuando ocurre el evento, emite el cargo. Es una **tool** que un
  humano/agente invoca (`registrar_multa`, …).
- **Backfill:** la **misma** lógica corrida en lote sobre la data histórica — así el
  motor re-deriva las aplicaciones (Hueco 3 del contrato). No es otra pieza; es la
  primera corrida del mismo emisor. Parte del cutover del ledger (roadmap B4).

## Relación con la migración

```
seguimiento_pueblo  =  store VIEJO → lo REEMPLAZA estado_cuenta → DESAPARECE
obligaciones        =  el que POBLA estado_cuenta con los cargos no-agua
                       · durante la migración: backfill re-siembra el histórico
                       · para siempre: emite cada obligación nueva
```

`obligaciones` **no es "parte de" la migración** ni desechable — es permanente, y la
migración es una de sus corridas.

---

## Alcance — qué emite y qué NO

| Concepto | ¿Lo emite `obligaciones`? | Nota |
|---|---|---|
| `multa` (reunión / faena) | **Sí** | evento: asamblea / faena |
| `acuerdos` (techado / campo) | **Sí** | evento: acuerdo de asamblea |
| `convenio` (medidor / instalación) | **Sí** | evento: firma de convenio |
| `agua` · `mantenimiento` · `corte_reconexion` | No | los emite `2_planilla` (lectura) |
| `deuda_directiva` | **No** | balde caja-only (⑩) — nunca es CARGO de predio |
| `tanque` | **No** | aporte voluntario, caja-only |

> Hoy el `convenio` de instalación también lo emitía `arrastre_consolidado` de
> `5_cobranza`. Bajo este diseño, esa emisión se unifica acá.

---

## Diseño detallado (Fase 1 — CERRADO 2026-07-15)

### Fuentes crudas de la secretaria → CARGO

El trigger real es el **registro del evento de la secretaria** (asistencia a asamblea/
faena, convenio/instalación firmados), NO un artefacto derivado como `DATA_boletas`. Cada
fuente vive en `obligaciones/inputs/` (cruda, intocable — nunca se regenera). El `MONTO`
del CARGO es la **deuda original**; los pagos de esas mismas hojas (`Pago1..5`, columnas de
mes) entran como **ABONOS a `caja`**, no acá — el saldo lo deriva el motor (contrato: CARGO
≠ pago).

| CONCEPTO | SUB_CONCEPTO | Archivo · hoja | Predio | MONTO (deuda original) | MES_CARGO |
|---|---|---|---|---|---|
| `MULTA` | `REUNION` | `FAENAS REUNIONES JASS PUEBLO.xlsx` · Hoja1 · col evento `REU dd/mm` | `MZ`·`LOTE` | **tarifa fija 20** (config); celda vacía = faltó = 1 cargo | del header del evento (ej. `2026-04`) |
| `MULTA` | `FAENA` | `FAENAS REUNIONES JAS.xlsx` · Hoja1 · col evento (col5) | `MZ`·`LOTE` | **tarifa fija 30** (config); vacía = faltó. **col6 = typo, ignorar** | 1ra faena `2026-04` · 2da `2026-05` (aún no cargada) |
| `ACUERDOS` | `TECHADO` | `DEUDORES Y PAGOS DEL TECHADO Y CAMPO.xlsx` · hoja **`Corregido`** · col `TECHADO` | `MZ`·`LT` | **monto directo** de la celda | génesis |
| `ACUERDOS` | `CAMPO` | idem · hoja `Corregido` · col `CAMPO` | `MZ`·`LT` | **monto directo** de la celda | génesis |
| `CONVENIO` | `MEDIDOR` | `mayo-planilla 2026-03-11 A 2026-04-10.xlsx` · hoja `Cobro medidores` · col `Deuda` | `MZ`·`LT` | **monto directo** (50 ó 100); auto-reconcilia `Deuda − ΣPago = Saldo` | `2026-04` |
| `CONVENIO` | `INSTALACION` | `SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx` · hoja `NUEVAS INSTALACIONES` · col `TOTAL` | `MZ`·`LT` | **monto directo** (`TOTAL`) | génesis (fecha de instalación) |

- **Split `SUB_CONCEPTO` resuelto:** MULTA (2 hojas: reunion/faena) y CONVENIO (2 hojas:
  medidor/instalacion) vienen partidos; ACUERDOS también (hoja `Corregido` trae `TECHADO`
  y `CAMPO` en columnas separadas — la `deuda_total=75` de otros archivos = techado+campo).
  Ningún concepto queda con sub genérico. Ignorar siempre la `Hoja1` cruda desordenada.
- **Normalizar al leer:** `P`/`p`, `MZ` minúscula, `LT` alfanumérico (`6A`, `8-A`); flag de
  celdas anómalas (hubo un `17` suelto en reunión). Reusar `_norm_mz`/`_norm_lt`.

### CARGO que emite (schema del contrato `estado_cuenta` · Entidad 2)

```
JASS_ID · CARGO_ID(determinista) · MZ · LT · CONCEPTO · SUB_CONCEPTO · MES_CARGO · MONTO · SOURCE=obligaciones
```
`CARGO_ID = sha256[:8](JASS_ID, MZ, LT, CONCEPTO, SUB_CONCEPTO, MES_CARGO)` → **idempotente**:
re-correr backfill no duplica. `MES_CARGO` es parte de la llave, así carry (mes real del
evento) y nuevas (mes del evento futuro) **no se solapan** — anti-doble-conteo automático.
`obligaciones` emite TODO el convenio, incluida instalación (unifica lo que antes tocaba
`arrastre_consolidado`); el hack `PREDIOS_INSTALACION_EXCLUIDOS` desaparece — el cargo de
instalación ahora existe y el pago se aplica normal.

### Reasignación de predio (COFOPRI) — remap obligatorio al leer

Las fuentes crudas usan predios **viejos** de la JASS (pre-COFOPRI). Antes de emitir el
cargo, cada `(MZ, LT)` cruda pasa por la tabla canónica de reasignaciones
(`0_padron/reasignaciones_candidata.xlsx`, 32 filas: C-45→C-43, B-29→B-20, cascadas A1/H1,
etc.). **Regla de alcance:** el remap se aplica SOLO al leer estas fuentes para sembrar; NO
reescribe `padron_reconciliado` ni la planilla. Conflictos no autorizados (ej. I-2B) se
dejan tal cual — el remap deja continuar el flujo sin resolver la identidad. Aplicar como
**lookup RAW→SYS de golpe, nunca secuencial** (hay swaps y cadenas → double-shift).

### Idempotencia y qué pasa si falta un archivo/columna

- Idempotente por `CARGO_ID` determinista (arriba). `audit_ref` → fila cruda de origen.
- Validación de inputs al inicio (patrón del sistema): si falta un archivo o una columna
  requerida → `FileNotFoundError`/`ValueError` descriptivo, no siembra parcial.

### Reconciliación de MULTA — la hace el motor, no un paso previo (2026-07-15)

Los archivos de asistencia dan deuda **bruta** (reunión 234 ausentes×20 = 4,680 · faena
262×30 = 7,860 ≈ S/12,540), pero el residual real es ínfimo (`Deuda faena` 9 filas ·
`Techado` 1 en la mayo-planilla). **No se reconcilia a mano.** Se siembra MULTA bruta (cada
ausencia = un cargo) y se importan **todos** los abonos a `caja` (los pagos de multa entran
por mesa/yape — confirmado); el **motor** aplica la cascada y el residual cae solo → se
**valida** contra `Deuda faena` (9). No se puede parquear MULTA: la cascada acopla los
conceptos (un pago "para el medidor" pudo irse a agua/multa por prioridad P1/P3), así que
falta MULTA → todo lo de abajo (convenio) se atribuye mal en el extracto.

---

## Lo que NO hace

- **No aplica pagos** — eso es el motor de aplicación del ledger.
- **No calcula saldos** — el saldo es una query derivada de `estado_cuenta`.
- **No toca la caja** — no conoce abonos; solo emite cargos.

## Contexto

Este módulo cierra el diseño de la disolución de `5_cobranza` post-ledger: todo lo que
`5_cobranza` hacía se reparte entre el motor (aplicar), queries (saldo/arrastre), tools
de lectura (reportes) y **este emisor** (los cargos no-agua). Ver
`docs/RETOMAR_dominio_saldo_unico_2026-07-13.md` y el contrato en
`libro_mayor/estado_cuenta/README.md`.
