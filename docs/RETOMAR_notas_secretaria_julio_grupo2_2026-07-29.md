# RETOMAR — Notas de la secretaria (GRUPO 2, notas_2026-07.xlsx) · Sesión 2026-07-28/29

## Punto de entrada

Sesión larga triando `4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx` (GRUPO 2, 8 fotos
de WhatsApp de la secretaria, 132 filas). El pedido central: para las notas donde ella afirma
"ya pagó" / "está al día" (no las que piden verificar), cancelar la deuda exacta que
`seguimiento_pueblo.xlsx` todavía mostraba — sin inventar montos, sacándolos del ledger real.

En el camino aparecieron 3 descubrimientos técnicos importantes (ver sección 3) y 2 errores de
cálculo que se encontraron y corrigieron en la misma sesión (ver sección 4). Todo quedó verificado
contra `planilla_2026-08.xlsx` antes de cerrar.

**Detalle técnico completo (qué se escribió, en qué archivo, con qué mecanismo) ya está en
`LEER_ANTES.md`, sección "correcciones notas_2026-07 (GRUPO 2)"** — este documento es el resumen
narrativo de la sesión, no duplica esa parte.

---

## 1. Resuelto esta sesión (verificado, guardado)

### 1.1 — 17 predios con deuda cancelada (MULTA/ACUERDOS → 0)

K-8 (solo MULTA, ACUERDOS=30 no se tocó), L-16, N-5, O-16, Q-5, Q-10, Q-11, R-4, S-1, S-8, S-9,
S-12, V-6, B1-12, D1-3, F1-5, G1-14.

Mecanismo final: `registrar_pago` **directo** en `shared/seguimiento_pueblo.xlsx`
(`SOURCE=manual`, `AUDIT_REF=notas_2026-07|MZ-LT-CONCEPTO`) — no el overlay `abonos_rezagados.xlsx`
que se probó primero (ver sección 3).

### 1.2 — 2 cargos nuevos sembrados

- **P-9**: ACUERDOS 75 (techado+campo). `CARGO` directo en el ledger + fila en
  `shared/genesis_tardia.xlsx`. Ya visible en `planilla_2026-08.xlsx`.
- **M-12**: MANTENIMIENTO +80. Sin ledger propio → **parche manual** directo en
  `2_planilla/outputs/planilla_2026-08.xlsx` y `shared/planilla_mes/planilla_2026-08.xlsx`
  (celda MANTENIMIENTO fila M-12: 3→83). ⚠️ **se pierde si se regenera `2_planilla` de nuevo** —
  hay que repetir el parche (no sumar 80 de nuevo si ya dice 83).

### 1.3 — 4 "Verificando" sacados de boleta

`3_boletas/inputs/pendientes_convenio_multas.xlsx`: K-8, G1-3, G1-14, B1-12 pasados de
`ESTADO=VERIFICANDO` a `RESUELTO` — ya mostraban el monto real (0 en los 4 casos), solo se
sacó el flag de impresión. No tocó el ledger (son cosas separadas, ver sección 3.3).

### 1.4 — 2 marcados para confirmar, NO ejecutados

`notas_2026-07.xlsx`, `TIPO_RECLAMO=CONFIRMACION`:
- **P-6**: pide "poner en convenio 350, instalación de Flor Valdivia" — es un convenio NUEVO,
  no una corrección, no se ejecuta sin que la secretaria confirme.
- **C1-3**: montos con "?" y ella misma escribió "Verificar".

---

## 2. Dónde quedó cada archivo

| Archivo | Estado |
|---|---|
| `shared/seguimiento_pueblo.xlsx` (repo actual) | Corregido y verificado — 17 predios en 0, P-9 en 75 |
| `shared/abonos_rezagados.xlsx` | Con las filas de agua vieja de O-16/Q-11 + Q-5 corregido (69, no 89) — documentación para el backfill futuro, no es lo que arregló el ledger |
| `shared/genesis_tardia.xlsx` | Filas M-12 (MANTENIMIENTO) y P-9 (ACUERDOS) agregadas |
| `5_cobranza/outputs/arrastre_consolidado_2026-07.xlsx` | Corregido a mano (Q-5 CONVENIO revertido, O-16/Q-11 agua a 0) |
| `2_planilla/outputs/planilla_2026-08.xlsx` + `shared/planilla_mes/planilla_2026-08.xlsx` | Regenerado 2 veces + parche manual de M-12. **Es el que hay que verificar antes de imprimir boletas** |
| `4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx` | 17+2+2 filas con ESTADO/TIPO_RECLAMO actualizados y documentados |
| `3_boletas/inputs/pendientes_convenio_multas.xlsx` | 4 filas pasadas a RESUELTO |
| `C:\Users\wilde\PycharmProjects\jass_system - Julio` | Copia usada para poder re-correr `5_cobranza --force` sobre julio sin tocar el repo activo. Tiene su propia versión (parcial, superada) de estos mismos archivos — **no es la fuente de verdad**, el repo actual ya tiene todo portado y corregido. Se puede dejar como está o borrar, no se usa más para esto. |
| `LEER_ANTES.md` | Sección completa con el detalle técnico, los 2 bugs, y la advertencia de M-12 |

---

## 3. Descubrimientos técnicos de la sesión (para no repetir el error)

### 3.1 — `abonos_rezagados`/`ajustes_cargo` NO tocan el ledger de MULTA/ACUERDOS sin plata real

Estos overlays reducen el campo de deuda (`u["multa"]`, `u["acuerdos_asamblea"]`) en el cálculo
de `5_cobranza`, pero `_reconciliar_pagos_pueblo()` solo escribe en `seguimiento_pueblo.xlsx`
si hay una diferencia contra el `total_pagado` REAL. Si nunca hubo plata real (la mayoría de estas
notas), el overlay reduce el reporte del mes pero el ledger queda igual — y como `2_planilla` de
agosto lee MULTA/ACUERDOS **exclusivamente** de `seguimiento_pueblo`, la corrección nunca llega.
Nunca se había usado este overlay para MULTA/ACUERDOS antes (solo para AGUA/CORTE, que no pasan
por este ledger) — por eso no se había topado con el hueco.

**Solución que sí funciona:** `registrar_pago`/`registrar_ajuste`/`registrar_cargo` directo,
llamando `shared/seguimiento_repo.py` a mano.

### 3.2 — El ledger es descartable, no hace falta que el backfill lo lea

Se discutió si escribir directo rompía la trazabilidad para el backfill (`backfill_ledger/`)
de agosto. Conclusión: `seguimiento_pueblo.xlsx` es una herramienta operativa de HOY (alimenta
boletas y reportes) — cuando exista `libro_mayor`, este archivo se descarta y se resiembra desde
los precursores crudos (`abonos_rezagados.xlsx`, `notas_2026-07.xlsx`, mesas de `4_pagos`, etc.),
no desde cómo quedó el ledger. Por eso escribir directo acá es seguro.

### 3.3 — `pendientes_convenio_multas.xlsx` (VERIFICANDO) nunca toca el ledger

Es un interruptor de impresión de `3_boletas`, nada más. El monto que imprime siempre sale del
ledger real, esté marcado VERIFICANDO o no. Sacar el flag no cancela ni crea deuda.

### 3.4 — Reabrir un ciclo cerrado: usar la copia paralela, no mover archivos en el repo activo

`5_cobranza` no tiene selector de mes — toma el último `planilla_*.xlsx` de
`shared/planilla_mes/`. En vez de mover `planilla_2026-08.xlsx` a un costado en el repo activo
(riesgoso), se usó `jass_system - Julio` — una copia ya congelada antes de que existiera agosto
— para correr `5_cobranza --force` apuntando a julio sin ningún riesgo sobre el repo real.

---

## 4. Los 2 errores encontrados y corregidos en el camino (mirar antes de confiar ciegamente en overlays de cascada)

1. **Q-5**: el "MES_ANTERIOR=20" que se iba a cancelar como "agua vieja" en realidad era su
   consumo REAL de julio sin pagar — la nota decía textual "ponle su consumo de este mes" (no
   tocar). Como `mes_anterior` real ya era 0, el monto de más se coló a CONVENIO (medidor) vía
   la cascada genérica de `abonos_rezagados`. Revertido en `arrastre_consolidado_2026-07.xlsx`
   — nunca tocó el ledger real. **Lección: antes de meter un monto a una cascada genérica,
   verificar de dónde sale cada componente — "MES_ANTERIOR" en la planilla del mes siguiente no
   siempre es deuda vieja, puede ser consumo corriente sin pagar.**
2. **O-16 y Q-11**: sus filas de agua se agregaron a `abonos_rezagados.xlsx` DESPUÉS de correr
   `5_cobranza --force` (por instrucción de no volver a correr `--force`), así que nunca se
   procesaron por la cascada. Se aplicaron a mano, directo en `arrastre_consolidado_2026-07.xlsx`.
   **Lección: un overlay que se agrega después de la corrida que lo consume no se aplica solo —
   hay que verificar el resultado real, no asumir por el nombre del archivo.**

Ambos se descubrieron haciendo una verificación fila por fila de `planilla_2026-08.xlsx` contra
lo esperado — no asumir que "el script corrió sin error" significa "el resultado es correcto".

---

## 5. Pendiente real (qué falta al retomar)

1. **Confirmar con la secretaria**: P-6 (convenio nuevo 350, Flor Valdivia) y C1-3 (montos "?").
   No ejecutar hasta que ella confirme — quedaron en `TIPO_RECLAMO=CONFIRMACION`.
2. **Q-3 y Q-1** (`pendientes_convenio_multas.xlsx`, ACUERDOS_ASAMBLEA=VERIFICANDO): se
   identificaron como candidatos a sacarles el "Verificando" (mostrarían 20 y 50 reales, sin
   cancelar nada), pero la sesión terminó sin la confirmación final del usuario para ejecutarlo.
3. **O-22, K-9, Q-8** (`pendientes_convenio_multas.xlsx`, CONVENIO=VERIFICANDO): excluidos a
   propósito toda la sesión por instrucción explícita de "no tocar medidor". Sin decisión tomada,
   siguen en VERIFICANDO.
4. **Imprimir/generar boletas de agosto (`3_boletas`) y confirmar visualmente** que salen con
   los valores ya corregidos (los 17 en 0, P-9=75, M-12=83) — esta sesión verificó
   `planilla_2026-08.xlsx` (el insumo), no corrió `3_boletas` todavía.
5. **Si se vuelve a correr `2_planilla` para agosto por cualquier motivo**: repetir el parche
   manual de M-12 (MANTENIMIENTO=83) — no tiene preservación automática.
6. **`jass_system - Julio`** (la copia usada): decidir si se borra o se deja — ya no hace falta
   para esto, tiene una versión superada/parcial de los mismos archivos.

## Cómo cerrar este RETOMAR

Cuando se confirmen los puntos 1-4 de la sección 5 (o se decida explícitamente diferirlos),
borrar este archivo y la sección correspondiente de `LEER_ANTES.md`.
