# RETOMAR — Precursores del ledger + reclamos julio · Sesión 2026-07-24

## ⚠️ PRIMER PASO al retomar

**Correr `5_cobranza --force` (julio).** Todo lo de esta sesión (reversión F1-4,
génesis tardía de 6 lotes techado/campo + A1-13, ajuste de consumo A1-13, blanco
E-8, abono rezagado F1-4) está verificado en aislado (`_cargar_planilla` directo)
pero **`planilla_cobrado.xlsx` sigue con los números viejos** — confirmado
recién: F1-4 sigue con `CORTE=60` (no revertido) y `SALDO=200`; A1-12/A1-13 con
su exceso viejo. Nada de esto se perdió — está todo en los precursores — pero
hace falta la corrida para que se refleje.

Después de esa corrida, revisar si conviene correr también:
- `4_pagos` de nuevo (el `sin_confirmar.xlsx`/`pendientes.xlsx` PAGASTE de Janet
  ya estaban bien llenados a las 14:12, nunca se re-confirmó con una corrida).
- `4b_reclamos --mes 2026-07` de nuevo si se resuelve algo más.

---

## 1. Lo hecho hoy (resumen técnico)

### Bug de código resuelto
`5_cobranza/main.py::_actualizar_blancos` — escribía en una columna 19 de
`blancos_acumulados.xlsx` que nunca existió (nunca se había ejercitado porque
nunca antes se había identificado un blanco ahí). Arreglado, verificado.

### Precursores nuevos creados (además de `reidentificacion`, `abonos_rezagados`,
`devoluciones_aplicadas`, `blancos_efectivo`/`blancos_acumulados`,
`aportes_tanque_manuales` que ya existían):

| Archivo | Evento ledger | Para qué | Casos hoy |
|---|---|---|---|
| `shared/deuda_correcciones.xlsx` | `reasignar_abono` (mitad espejo) | el lote de origen de una reidentificación SÍ tenía deuda real (a diferencia del caso genérico donde solo hay exceso) | B-15 (deferred agosto) |
| `shared/ajustes_cargo.xlsx` | `registrar_ajuste` | un CARGO nació MAL y se anula (sin plata de más) — extendido hoy a soportar `AGUA`/`MANTENIMIENTO`, no solo `CORTE_RECONEXION` | F1-4 (corte jun+jul, 60), A1-13 (consumo 38→5, tarifa mínima) |
| `shared/genesis_tardia.xlsx` | `registrar_cargo` tardío | un CARGO legítimo cuya fuente en `obligaciones/` se corrigió DESPUÉS de que el ciclo ya estaba congelado | 6 lotes techado/campo: A1-12, A1-13, C1-13, C1-15, H1-13, H1-36 (S/75 c/u) |

Además: `shared/abonos_rezagados.xlsx` ganó columna `RESPALDO`
(`documentado` vs `solo_declaracion`) para distinguir abonos con rastro real
(mesa+yape) de los que solo tienen la palabra de alguien (F1-4, S/101,
retenido por Yanet — no por un cobrador de ruta).

Documentación del ledger sincronizada byte-idéntica en `libro_mayor/caja/README.md`
y `libro_mayor/estado_cuenta/README.md`: decisiones ⑬ (reactivación, sub-concepto
de CONVENIO), ⑭ (`ajustes_cargo`/F1-4), ⑮ (`genesis_tardia`/techado-campo), ⑯
(`ajustes_cargo` extendido a AGUA/A1-13). También `libro_mayor/README.md`,
`libro_mayor/dominio/README.md`, `backfill_ledger/docs/cuaderno_backfill.html`,
`docs/cuaderno/libro_mayor.html`.

### Fuentes reales corregidas (no solo overlays — quedan bien para el backfill)
- `obligaciones/inputs/DEUDORES Y PAGOS DEL TECHADO Y CAMPO.xlsx` (hoja
  `Corregido`) — 6 filas agregadas (los mismos lotes de arriba), SUBTOTAL corregido.
- `obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx` —
  hoja `REACTIVACION` nueva (Ramon Requez Mendoza / M-12, S/266).
- `0_padron/overrides_padron.xlsx` — M-12 ya tenía el override de nombre
  (Iglesia Evangélica → Ramon Requez), confirmado, no tocado hoy.

### Casos de reclamos/exceso resueltos hoy (con RESOLUCION + ESTADO en
`4b_reclamos/outputs/reclamos_2026-07.xlsx` y/o `arrastre_devolucion_2026-07.xlsx`)

| Caso | Motivo |
|---|---|
| Q-16 | Sin deuda (SALDO=0), cerrado sin explicación antes — completado |
| M-15 + G-1 | Multa mal atribuida entre ambos (mismo cobrador/día) — reidentificación |
| G-18 | Exceso de junio aplicado a convenio/medidor |
| F-9 | Abono rezagado cubrió todo — CANCELADO |
| A1-12, A1-13, C1-13, C1-15, H1-13, H1-36 | Techado y campo (génesis tardía) |
| A1-13 (extra) | + consumo corregido 38→5 (tarifa mínima, orden directa secretaria) |
| F1-4 | Abono rezagado (101, `solo_declaracion`) + penalidad de corte revertida (60) |
| E-8 | Blanco Yape S/58 identificado + penalidad de corte revertida (20) |
| F1-11 | Convenio/instalación, ya registrado en SEGUIMIENTO INSTALACIONES (no exceso de agua) |

### Revisados pero dejados `EN_REVISION` (el precursor no resuelve lo que pide el reclamo específicamente)
T-7, G-12, D-16, S-5 — tienen abono/reidentificación aplicado pero el concepto
que reclamaban sigue con SALDO>0 en ese componente.

---

## 2. Pendiente — sin resolver, sin evidencia o sin decisión

- **26 excesos sin explicar** en `arrastre_devolucion_2026-07.xlsx` (de 37
  totales, 11 ya resueltos) — nadie los reclamó, no hay precursor ni patrón.
- **Z-14** — sigue `PENDIENTE`, nunca se confirmó revertir su penalidad de
  corte (se preguntó, no se confirmó).
- **G1-14** — candidato a 7º lote de techado/campo (reclamo dice "ya pagué
  techado y campo, no debo nada") — no verificado todavía.
- **D-5** — reclamo "Reclamo" sin ningún detalle, ya routeado a VERIFICANDO en
  `pendientes_convenio_multas.xlsx`, pero sigue sin saber qué es. Falta
  preguntarle a Maximo Encarnacion (igual que Q-16 antes de resolverse).
- **B-19, F1-4 (la vieja duda de origen)** — ver memoria previa, F1-4 ya se
  resolvió hoy vía abono rezagado; B-19 seguía resuelto de antes.
- **`sin_confirmar.xlsx`** (yape, 74 filas) y **PAGASTE pendientes** (5,
  Correcciones/) — el archivo de Janet (5 PAGASTE) ya estaba bien llenado a
  las 14:12 pero nunca se re-confirmó con una corrida de `4_pagos`.
- **`4b_reclamos`**: `resolucion_reclamos_2026-07.xlsx` no existe — hay 22
  reclamos RESUELTO/INFORMADO sin CAMPO/VALOR_ANTERIOR/VALOR_APLICADO
  registrado (warning explícito en la corrida de hoy). Correr `resolucion.py`
  si existe, o diseñar ese paso si no.

---

## 3. Archivos nuevos/tocados hoy (para saber qué mirar mañana)

**Código:**
- `5_cobranza/main.py` — +392/-19 líneas (ver `git diff --stat`).

**Precursores (`shared/`, todos "writer único humano"):**
- `deuda_correcciones.xlsx`, `ajustes_cargo.xlsx`, `genesis_tardia.xlsx` (nuevos)
- `reidentificacion.xlsx`, `abonos_rezagados.xlsx` (filas nuevas + columna RESPALDO)
- `blancos_acumulados.xlsx` (E-8 identificado)

**Fuentes reales:**
- `obligaciones/inputs/DEUDORES Y PAGOS DEL TECHADO Y CAMPO.xlsx`
- `obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx`

**Auditorías de módulo:**
- `6_corte/outputs/audit_penalidad.xlsx` (E-8 y F1-4 revertidos)

**Docs del ledger:**
- `libro_mayor/README.md`, `libro_mayor/dominio/README.md`,
  `libro_mayor/caja/README.md`, `libro_mayor/estado_cuenta/README.md`,
  `backfill_ledger/docs/cuaderno_backfill.html`, `docs/cuaderno/libro_mayor.html`

**Outputs regenerados (gitignored, no importan para commit):**
`4_pagos/**/outputs/*`, `5_cobranza/outputs/*` (⚠ desactualizado, ver §PRIMER PASO),
`4b_reclamos/outputs/reclamos_2026-07.xlsx`, `4b_reclamos/trazabilidad/trazabilidad_reclamos.xlsx`

**Repo separado `jass_system - junio`** (congelado, no confundir con este repo):
se tocó `5_cobranza/outputs/arrastre_devolucion_2026-06.xlsx` (M-12 resuelto) —
por pedido explícito, es la copia que conserva el estado de junio.

---

## 4. Decisión de metodología (candidata, no aplicada a CLAUDE.md todavía)

**Un precursor por cada "forma" distinta de corrección**, no reusar uno
existente solo porque el mecanismo (sumar/restar un campo) es igual. El
criterio no es la mecánica, es la **historia que el backfill va a contar**:
`reasignar_abono` ≠ `registrar_ajuste` ≠ `registrar_cargo` tardío, aunque los
3 se implementen como "sumale/restale X a este campo". Esta sesión encontró y
corrigió 2 veces un archivo mal elegido (F1-4 en `devoluciones_aplicadas` →
movido a `ajustes_cargo`; casi se repite con techado/campo). Ver propuesta en
§5 de esta sesión para agregarlo a `docs/metodologia_desarrollo.md`.
