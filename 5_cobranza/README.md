# 5_cobranza — Consolidación de la cobranza del mes

Cruza la planilla del mes con los pagos identificados (Yape + efectivo), calcula el
**saldo** de cada predio y produce el estado de cobro del ciclo + los arrastres para
el mes siguiente. Es el módulo que responde "¿quién pagó, cuánto, y cuánto sigue debiendo?".

> Documentación de detalle (diseño, columnas, contratos visuales) en
> `docs/README_cobranza.md` + los `docs/*.html`. Este README es la vista de módulo.

---

## ⚠ DISEÑO POST-LEDGER — este módulo SE DISUELVE (Fase 1 cerrada · 2026-07-14)

Bajo `libro_mayor`, `5_cobranza` **desaparece**. Nunca fue dueño de nada que el ledger no
haga mejor; sus responsabilidades se reparten:

| Responsabilidad de 5_cobranza | Nuevo dueño |
|---|---|
| aplicar pagos (cascada P1-P6) | **motor de aplicación** (`libro_mayor`) |
| calcular SALDO por predio | **query derivada** (`dominio/saldo.py`) |
| `arrastre_deuda` / `arrastre_devolucion` (archivos) | query de cargos abiertos / `SALDO_A_FAVOR` — **el arrastre deja de ser archivo** |
| reconciliar pagos pueblo → `seguimiento_pueblo` | el motor deriva aplicaciones desde los abonos de caja |
| reportes (`planilla_cobrado`, `resumen`) | tools de lectura del ledger |
| emitir cargos multa/acuerdos/convenio | **nunca lo hizo** → eso es `obligaciones/` (génesis) |

**Verificado contra código:** `5_cobranza` solo llama `registrar_pago`/`registrar_ajuste`,
nunca `registrar_cargo`. Por eso el contrato ⑧ (que lo nombraba emisor de cargos) se
corrigió a `obligaciones`.

**Estado:** diseño cerrado, ledger **sin implementar**. El código descrito abajo (pre-ledger)
**sigue corriendo** hasta que el ledger esté codificado. Ver
`docs/retomar/RETOMAR_dominio_saldo_unico_2026-07-13.md` §11.

---

## Implementación actual (pre-ledger, transitoria)

---

## Qué hace

Lee la planilla del mes y los pagos en vivo desde los outputs de `4_pagos`, aplica los
blancos automáticos de `shared/`, descompone cada saldo por prioridad de concepto
(cascada de prioridad P1→P5, tramo pueblo ya en CA1: CONVENIO→ACUERDOS→MULTA), reconcilia los pagos de pueblo al ledger
`seguimiento_pueblo`, y emite 5 archivos.

Las corridas normales y `--force` son proyecciones regenerables: generan un
`snapshot_ledger_YYYY-MM.json`, pero no escriben `seguimiento_pueblo`. El único commit
mensual lo ejecuta `7_cierre` después de la proyección final y 5b_validacion.

Desde agosto de 2026 el snapshot contiene la cuenta completa: apertura de AGUA/CORTE desde julio,
cargos nuevos de AGUA/MANTENIMIENTO/CORTE_RECONEXION y aplicaciones para los seis conceptos activos.
Un abono rezagado conserva `MES_CICLO`: si pertenece a un ciclo cerrado, no paga consumo nuevo de agosto.

## Cuándo se corre

Después de que `4_pagos` está completo y sin pendientes:
- `4_pagos/yape/motor_matching/outputs/pagos_yape_tepago.xlsx` sin pendientes
- `4_pagos/efectivo/outputs/pagos_efectivo.xlsx` sin discrepancias

**Idempotente:** si los pagos no cambiaron respecto a la trazabilidad existente, sale
sin tocar nada. Para forzar recálculo cuando cambió la planilla pero no los pagos:
`python main.py --force`.

---

## Qué lee y qué genera

```
LEE (en vivo, no copias):
  shared/planilla_mes/planilla_YYYY-MM.xlsx        (master — 2_planilla la publica ahí)
  4_pagos/yape/motor_matching/outputs/pagos_yape_tepago.xlsx
  4_pagos/efectivo/outputs/pagos_efectivo.xlsx
  shared/blancos_acumulados.xlsx                    (blancos automáticos)
  6_corte/outputs/audit_penalidad.xlsx              (overlay de penalidad, re-derivado en vivo)

GENERA (outputs/):
  planilla_cobrado_YYYY-MM.xlsx   hojas planilla_cobrado + arrastre_consolidado + arrastre_devolucion
  trazabilidad_cobranza.xlsx      un registro por pago cargado (acumulada)
  resumen_recaudacion.xlsx        totales del mes
  arrastre_deuda_YYYY-MM.xlsx      SALDO>0  → lo pre-carga 2_planilla del próximo mes
  run.log
```

> **`lista_corte` NO se genera acá.** `SALDO` sale como columna explícita en
> `planilla_cobrado_YYYY-MM.xlsx`; la lista de corte la produce `6_corte/generar_lista.py`
> leyendo `SALDO` + `MES_ANTERIOR` desde acá. (El `docs/README_cobranza.md` viejo aún
> la lista como output de este módulo — quedó desactualizado.)

---

## Reglas de negocio

- **Cascada de prioridad** (`_descomponer_saldo`): un pago que no cubre todo el
  saldo se aplica en orden `P1 AGUA·MANTENIMIENTO·arrastre → P2 CORTE_RECONEXION →
  P3 CONVENIO → P4 ACUERDOS → P5 MULTA`. El código viejo llega hasta P5; el motor del
  ledger que la replica **agrega P6 OTROS** (cascada destino P1-P6, contrato ⑪).
  El tramo pueblo se invirtió el **2026-08-13** (decisión CA1 del 23/07): el dinero
  cubre primero lo que SOLO el dinero salda — la multa se recupera con faena o se
  exonera. Ver `docs/decisiones/reimputacion_cascada_ca1.md`.
  ⚠ El orden es **posicional** contra `_AC_P` (columnas de arrastre_consolidado),
  `_CONCEPTOS_PUEBLO` (índice → concepto del ledger) y
  `_CAMPOS_WATERFALL_REIDENTIFICACION`. Reordenar uno solo desplaza montos de columna.
- **Writer único de `seguimiento_pueblo`:** solo `shared/seguimiento_repo.py` escribe
  ese ledger (patrón repo). 5_cobranza reconcilia por **delta**: `SET_DEBE` (recién
  calculado) − `SET_TIENE` (Σ ya anotado) = lo que falta registrar. No pisa totales
  (append-only).
- **Pagos en vivo, no copias:** lee los outputs reales de `4_pagos` para no perder
  pagos agregados después (ej. un 2º Yape del mismo predio). Retroescribe
  `CICLO_COBRANZA` en esos archivos (le corresponde por contrato).
- **Overlay de penalidad (Modelo A):** `6_corte`/`6b` ya **no** escriben
  `CORTE_RECONEXION` en `shared/planilla_mes`; la penalidad vive solo en sus audits y
  5_cobranza la re-deriva en vivo sumando `PENALIDAD_APLICADA`. `shared` queda como
  base pura de `2_planilla`.
- **Excluidos de convenio de instalación** (`PREDIOS_INSTALACION_EXCLUIDOS`): su deuda
  ya está completa en `arrastre_consolidado`; no se les reconcilia pago de CONVENIO en
  `seguimiento_pueblo` para no crear un PAGO sin CARGO (saldo negativo falso). Este hack
  desaparece en Fase 2 (ver abajo).
- **Guard de abonos rezagados:** `inputs/abonos_rezagados_manifest_2026-08.json` es el
  manifest aprobado de 18 filas. Antes de cargar un ciclo, `5_cobranza` compara la fila
  completa por `MZ + LT + MONTO + MES_CICLO + MES_ANO_APLICA`; una fila extra, faltante,
  duplicada o bloqueada detiene la corrida antes de reconciliar el ledger. El `BALDE`
  no decide la cascada del abono; el manifest identifica el hecho y la cascada combina
  pago normal + abono.

---

## Acoplamiento

```
2_planilla ─(shared/planilla_mes)─►  5_cobranza  ─(planilla_cobrado: SALDO)─►  6_corte
4_pagos    ─(pagos en vivo)───────►               ─(arrastre_deuda)────────►  2_planilla (mes+1)
                                                   ─(mismo workbook)────────►  7_cierre
                                                   ─(reconcilia)────────────►  shared/seguimiento_pueblo
```

---

## Pendiente Fase 2 — el ledger le quita la aplicación

Bajo `libro_mayor` (contrato ⑥⑧), 5_cobranza **deja de aplicar pagos**: solo **emite
cargos** de los conceptos que le corresponden (multa reunión/faena · acuerdos
techado/campo · convenio medidor/instalación) hacia `estado_cuenta`. El **motor de
aplicación** —única pieza que ve caja + deuda— deriva las aplicaciones. Consecuencias:

- El `arrastre_deuda_*.xlsx` (memoria-por-archivo, raíz de los bugs B4/B5/B7) se
  reemplaza por una **query derivada** del ledger (`SALDO = Σcargos − Σaplicaciones ± Σajustes`).
- El hack `PREDIOS_INSTALACION_EXCLUIDOS` desaparece: el cargo de instalación existe en
  el ledger y el pago se aplica normal.
- Las columnas `BLANCO`/`DEVOLUCION` de la planilla se retiran (pasan a ser aplicaciones
  auditables del ledger).
- La cascada `_descomponer_saldo` migra a `libro_mayor/dominio/cascada.py` (lógica pura).

Ver `docs/retomar/RETOMAR_dominio_saldo_unico_2026-07-13.md` (roadmap B2).

---

## Lo que NO hace

- **No identifica pagos** — eso es `4_pagos` (motor de matching Yape + mesas efectivo).
- **No genera la lista de corte** — eso es `6_corte`, leyendo el `SALDO` de acá.
- **No emite boletas** — eso es `3_boletas`.

## Errores comunes

- Correr sin que `4_pagos` esté cerrado → pagos incompletos, saldos mal.
- Esperar que regenere outputs cuando la planilla cambió pero los pagos no: usar `--force`.
- Buscar `lista_corte.xlsx` en `outputs/` — no está acá, está en `6_corte/`.
