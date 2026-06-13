# 2_planilla

Genera la planilla mensual de cobro consolidando lecturas del operario con deudas arrastradas.

## Cuándo correr

Después de que `1_lecturas` cierre el ciclo (sin bloqueantes pendientes).

```
1_lecturas → 2_planilla → 4_pagos → 5_cobranza
```

---

## Inputs — esquema exacto de cada archivo

### 1. `inputs/lecturas/lecturas_planilla_YYYY-MM.xlsx`

Viene de `1_lecturas/outputs/`. El módulo detecta el mes leyendo la columna `MES_ANO`
de la primera fila de datos.

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

### 2. `inputs/deuda_anterior/arrastre_deuda_YYYY-MM.xlsx`

Generado por `5b_validacion` (hoy se crea manualmente).
El nombre del archivo debe coincidir con el YYYY-MM de las lecturas.

Columnas requeridas:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `MZ` | texto | Manzana |
| `LT` | texto | Lote |
| `monto` | número | Monto de deuda que arrastra del mes anterior |

Si el archivo **no existe** → `MES_ANTERIOR = 0` para todos los usuarios. El log registra advertencia.
Si un usuario **no aparece** en el archivo → `MES_ANTERIOR = 0` para ese usuario.

### 3. `inputs/corte/arrastre_corte_YYYY-MM.xlsx`

Generado por `6_corte` (hoy se crea manualmente).
El nombre del archivo debe coincidir con el YYYY-MM de las lecturas.

Columnas requeridas:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `MZ` | texto | Manzana |
| `LT` | texto | Lote |
| `monto` | número | Cargo por corte y reconexión |

Si el archivo **no existe** → `CORTE_RECONEXION = 0` para todos. El log registra advertencia.
Si un usuario **no aparece** → `CORTE_RECONEXION = 0` para ese usuario.

### 4. `inputs/convenios/convenios.xlsx`

Mantenido manualmente por el tesorero. Un solo archivo que se actualiza cada mes.

Columnas requeridas:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `MZ` | texto | Manzana |
| `LT` | texto | Lote |
| `cuota_mes` | número | Cuota que le corresponde pagar este mes |

Si el archivo **no existe** → `CONVENIO = 0` para todos. El log registra advertencia.
Si un usuario **no aparece** → `CONVENIO = 0` para ese usuario.

### 5. `inputs/multas/multas.xlsx`

Mantenido manualmente por el tesorero. Un solo archivo que se actualiza cada mes.

Columnas requeridas:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `MZ` | texto | Manzana |
| `LT` | texto | Lote |
| `monto_mes` | número | Monto de multa que le corresponde pagar este mes |

Si el archivo **no existe** → `MULTA = 0` para todos. El log registra advertencia.
Si un usuario **no aparece** → `MULTA = 0` para ese usuario.

### 6. `inputs/acuerdos_asamblea/acuerdos_asamblea.xlsx`

Mantenido manualmente por el tesorero. Recoge los aportes acordados en asamblea para obras de infraestructura (techado de local, tanque de agua, etc.). Un solo archivo que se actualiza cada mes conforme a los acuerdos vigentes.

Columnas requeridas:

| Columna | Tipo | Descripción |
|---------|------|-------------|
| `MZ` | texto | Manzana |
| `LT` | texto | Lote |
| `monto_mes` | número | Aporte que le corresponde pagar este mes según el acuerdo |

Si el archivo **no existe** → `ACUERDOS_ASAMBLEA = 0` para todos. El log registra advertencia.
Si un usuario **no aparece** → `ACUERDOS_ASAMBLEA = 0` para ese usuario.

---

## Lógica de join

La clave de unión entre todos los archivos es **(MZ, LT)** normalizada:
- `MZ` → uppercase, sin espacios
- `LT` → si es número entero flotante (ej: `1.0`) convertir a `"1"` · si tiene letras conservar uppercase (ej: `"11A"`)

El archivo base es `lecturas_planilla` — todos los usuarios que aparecen ahí
aparecen en la planilla de salida. Los arrastres se unen sobre esa base.
Un usuario en el arrastre que no esté en lecturas se **ignora** (log de advertencia).

---

## Reglas de cálculo

```
── Consumo del mes ──────────────────────────────────────────────────
MES_ACTUAL           = max(M3, 5)          ← tarifa S/1/m³, mínimo S/5
MANTENIMIENTO        = 3                    ← fijo para todos, sin excepción

── Arrastres ────────────────────────────────────────────────────────
MES_ANTERIOR         = monto de arrastre_deuda        (0 si no hay)
CORTE_RECONEXION     = monto de arrastre_corte         (0 si no hay)

── Seguimiento (archivos de tesorero) ───────────────────────────────
CONVENIO             = cuota_mes de convenios.xlsx     (0 si no hay)
MULTA                = monto_mes de multas.xlsx         (0 si no hay)
ACUERDOS_ASAMBLEA    = monto_mes de acuerdos_asamblea.xlsx (0 si no hay)

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
| 10 | `MES_ANTERIOR` | arrastre_deuda | 0 si no hay |
| 11 | `CORTE_RECONEXION` | arrastre_corte | 0 si no hay |
| 12 | `CONVENIO` | convenios | 0 si no hay |
| 13 | `MULTA` | multas | 0 si no hay |
| 14 | `ACUERDOS_ASAMBLEA` | acuerdos_asamblea | 0 si no hay — aporte acordado en asamblea |
| 15 | `BLANCO` | **0 al generar** | 5_cobranza escribe valor negativo cuando aplica descuento por blanco |
| 16 | `DEVOLUCION` | **0 al generar** | 5_cobranza escribe valor negativo cuando devuelve exceso |
| 17 | `TOTAL_A_PAGAR` | fórmula Excel | suma cols 8–16 (BLANCO y DEVOLUCION negativos reducen el total) |
| 18 | `MONTO_YAPE` | **vacío** | lo llena 5_cobranza |
| 19 | `MONTO_EFECTIVO` | **vacío** | lo llena 5_cobranza |
| 20 | `ESTADO` | **vacío** | lo llena 5_cobranza |
| 21 | `FECHA_PAGO` | **vacío** | lo llena 5_cobranza |

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
