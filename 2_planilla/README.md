# 2_planilla

Genera la planilla mensual de cobro consolidando las lecturas cerradas del operario con los saldos comprometidos en `shared/seguimiento_pueblo.xlsx`.

**Transición D-005:** agosto de 2026 abre AGUA/CORTE con el consolidado cerrado de julio. Desde
septiembre, `MES_ANTERIOR` y `CORTE_RECONEXION` salen del `estado_cuenta` comprometido por `7_cierre`;
si el mes anterior no está comprometido, la generación aborta en vez de asumir deuda cero.

## Cuándo correr

Después de que `1_lecturas` cierre el ciclo (sin bloqueantes pendientes).

```
1_lecturas → 2_planilla → 4_pagos → 5_cobranza
```

---

## Inputs — esquema exacto de cada archivo

### 1. `../1_lecturas/outputs/lecturas_planilla_YYYY-MM.xlsx`

Es la fuente directa de lecturas: `2_planilla` no conserva ni consulta una copia en
`inputs/`. El módulo detecta el mes leyendo la columna `MES_ANO` de la primera fila de
datos.

Columnas requeridas:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `MZ` | texto | Manzana (ej: "A", "B1") |
| `LT` | texto | Lote normalizado (ej: "1", "11A") |
| `NOMBRE` | texto | Nombre del usuario |
| `MES_ANO` | texto | Formato `YYYY-MM` (ej: "2026-06") |
| `MARC_ANT` | número | Marcación anterior |
| `MARC_ACT` | número | Marcación actual |
| `M3` | número | Consumo declarado por el operario |

### 2. `shared/seguimiento_pueblo.xlsx` — saldo comprometido del ciclo anterior

`YYYY-MM` es el ciclo anterior al de las lecturas. Desde septiembre de 2026, y solo
cuando `7_cierre` lo haya comprometido en `estado_ciclo.json`, se leen los saldos al
cierre de ese mes mediante `seguimiento_repo`:

| Columna planilla | Concepto de saldo |
|---|---|
| `MES_ANTERIOR` | `AGUA` + `MANTENIMIENTO` |
| `CORTE_RECONEXION` | `CORTE_RECONEXION` |
| `CONVENIO` | `CONVENIO` |
| `MULTA` | `MULTA` |
| `ACUERDOS_ASAMBLEA` | `ACUERDOS` |

El libro es un ledger append-only: sus cambios se registran como eventos mediante
`seguimiento_repo`, nunca editando un saldo manualmente. Para agosto de 2026 se conserva
la apertura excepcional desde `arrastre_consolidado` conforme a D-005.

---

## Lógica de join

La clave de unión entre todos los archivos es **(MZ, LT)** normalizada:
- `MZ` → uppercase, sin espacios
- `LT` → si es número entero flotante (ej: `1.0`) convertir a `"1"` · si tiene letras conservar uppercase (ej: `"11A"`)

El archivo base es `lecturas_planilla` de `1_lecturas/outputs` — todos los usuarios que aparecen ahí
aparecen en la planilla de salida. Los arrastres se unen sobre esa base.
Un usuario en el arrastre que no esté en lecturas se **ignora** (log de advertencia).

---

## Reglas de cálculo

```
── Consumo del mes ──────────────────────────────────────────────────
MES_ACTUAL           = max(M3, 5)          ← tarifa S/1/m³, mínimo S/5
MANTENIMIENTO        = 3                    ← fijo para todos, sin excepción

── Arrastres ────────────────────────────────────────────────────────
MES_ANTERIOR         = saldo AGUA + MANTENIMIENTO al cierre del ciclo anterior
CORTE_RECONEXION     = saldo CORTE_RECONEXION al mismo cierre

── Seguimiento (ledger comprometido) ─────────────────────────────────
CONVENIO             = saldo CONVENIO al cierre del ciclo anterior
MULTA                = saldo MULTA al cierre del ciclo anterior
ACUERDOS_ASAMBLEA    = saldo ACUERDOS al cierre del ciclo anterior

── Descuentos (los llena 5_cobranza — valores negativos) ────────────
BLANCO               = 0 al generar · 5_cobranza aplica valor negativo
DEVOLUCION           = 0 al generar · 5_cobranza aplica valor negativo

── Total ─────────────────────────────────────────────────────────────
TOTAL_A_PAGAR        = MES_ACTUAL + MANTENIMIENTO
                     + MES_ANTERIOR + CORTE_RECONEXION
                     + CONVENIO + MULTA + ACUERDOS_ASAMBLEA
                     + BLANCO + DEVOLUCION       ← negativos reducen el total
```

> **Nota sobre BLANCO y DEVOLUCION:** al generar la planilla valen 0. Cuando 5\_cobranza
> aplica un blanco o registra una devolución, escribe el monto como valor negativo
> en esa celda y Excel recalcula TOTAL\_A\_PAGAR automáticamente (fórmula Excel, no valor fijo).
>
> **⚠ Se retiran en Fase 2 (decisión ⑨ del ledger):** estas dos columnas son un
> descuento manual en un archivo regenerable — se pisan al regenerar y nunca cuadran
> en `5b_validacion`. Cuando exista `libro_mayor/estado_cuenta`, el descuento por
> blanco reclamado o por exceso deja de ser una celda de la planilla y pasa a ser una
> **aplicación auditable** en el ledger, linkeada al `ABONO_ID` del pago y al
> `reclamo_id` que la autorizó. La boleta mostrará el saldo ya corregido leyendo el
> ledger, no esta columna. Ver `libro_mayor/estado_cuenta/README.md`.

---

## Output — `outputs/planilla_YYYY-MM.xlsx`

Una sola hoja llamada `Planilla`. Columnas en este orden exacto:

| # | Columna | Origen | Descripción |
|---|---------|--------|-------------|
| 1 | `MZ` | lecturas | Manzana |
| 2 | `LT` | lecturas | Lote |
| 3 | `NOMBRE` | lecturas | Nombre del usuario |
| 4 | `MES_ANO` | lecturas | Mes del ciclo (YYYY-MM) |
| 5 | `MARC_ANT` | lecturas | Marcación anterior |
| 6 | `MARC_ACT` | lecturas | Marcación actual |
| 7 | `M3` | lecturas | Consumo declarado |
| 8 | `MES_ACTUAL` | calculado | max(M3, 5) |
| 9 | `MANTENIMIENTO` | fijo | 3 |
| 10 | `MES_ANTERIOR` | saldo `AGUA` + `MANTENIMIENTO` del ledger | 0 si no hay |
| 11 | `CORTE_RECONEXION` | saldo `CORTE_RECONEXION` del ledger | 0 si no hay |
| 12 | `CONVENIO` | saldo `CONVENIO` del ledger | 0 si no hay |
| 13 | `MULTA` | saldo `MULTA` del ledger | 0 si no hay |
| 14 | `ACUERDOS_ASAMBLEA` | saldo `ACUERDOS` del ledger | 0 si no hay |
| 15 | `BLANCO` | **0 al generar** | 5_cobranza escribe valor negativo cuando aplica descuento por blanco |
| 16 | `DEVOLUCION` | **0 al generar** | 5_cobranza escribe valor negativo cuando devuelve exceso |
| 17 | `TOTAL_A_PAGAR` | fórmula Excel | suma cols 8–16 (BLANCO y DEVOLUCION negativos reducen el total) |
| 18 | `MONTO_YAPE` | **vacío** | lo llena 5_cobranza |
| 19 | `MONTO_EFECTIVO` | **vacío** | lo llena 5_cobranza |
| 20 | `ESTADO` | **vacío** | lo llena 5_cobranza |
| 21 | `FECHA_PAGO` | **vacío** | lo llena 5_cobranza |

---

## Alimentación del ledger `libro_mayor/estado_cuenta`

> **Estado transitorio desde 2026-08:** la planilla alimenta el snapshot de `5_cobranza`;
> `7_cierre` escribe los cargos oficialmente junto con las aplicaciones del mes.

`2_planilla` es una **fuente de CARGOS** del ledger de cuenta corriente. Al generar la
planilla del mes emite a `libro_mayor/estado_cuenta` un cargo por cada obligación de
**agua**, **mantenimiento** y **corte** (nombres canónicos del contrato — el feeder traduce
sus columnas viejas al emitir, ver `dominio/taxonomia`):

```
  planilla_YYYY-MM  ──►  cuenta_repo.registrar_cargo(
                            jass_id, mz, lt,
                            concepto = "AGUA" | "MANTENIMIENTO" | "CORTE_RECONEXION",
                            sub_concepto = "",        # estos conceptos no tienen sub
                            mes_cargo = MES_ANO,
                            monto,
                            source = "2_planilla")
```

Reglas de compatibilidad con el contrato del ledger (ver `libro_mayor/estado_cuenta/README.md`):
- **Solo emite HECHOS (cargos). NO aplica pagos** — el reparto abono→cargo lo hace el
  **motor de aplicación** de `libro_mayor/estado_cuenta`, que es el único que ve caja y deuda juntas.
- Cada cargo lleva `JASS_ID` y `MES_CARGO = MES_ANO` (el mes en que nació la deuda).
- **Emite un cargo por concepto** (AGUA y MANTENIMIENTO son 2 conceptos distintos de P1,
  no se fusionan): `AGUA` = `MES_ACTUAL` (col 8) · `MANTENIMIENTO` = col 9 · `CORTE_RECONEXION`
  = col 11 (cuando aplica). Ninguno lleva `SUB_CONCEPTO`.
- Idempotente por `CARGO_ID` determinista = `sha256[:8](JASS_ID, MZ, LT, CONCEPTO,
  SUB_CONCEPTO, MES_CARGO)`: regenerar la planilla no duplica el cargo.

---

## Idempotencia

Correr el módulo dos veces con los mismos inputs produce exactamente el mismo output.
Si `outputs/planilla_YYYY-MM.xlsx` ya existe, se sobreescribe sin preguntar.

---

## Lo que NO hace

- No valida lecturas — eso es `1_lecturas`
- No registra pagos — eso es `4_pagos`
- No genera el arrastre del mes siguiente — eso es `5b_validacion`
- No aplica cortes de servicio — eso es `6_corte`
- No actualiza saldos de convenios ni multas — el tesorero actualiza esos archivos manualmente

---

## Errores comunes

| Error | Causa | Solución |
|-------|-------|----------|
| `FileNotFoundError: lecturas_planilla` | `1_lecturas` no cerró el ciclo | Verificar que no haya bloqueantes en `1_lecturas/outputs/` |
| `TOTAL_A_PAGAR` negativo | Arrastre con valor negativo | Revisar archivo de arrastre — no se aceptan montos negativos |
| Filas duplicadas en output | Lecturas con (MZ, LT) duplicado | No debería ocurrir — `1_lecturas` ya detecta DUPLICADO |
