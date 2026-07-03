# 7_cierre — Transición de período (freeze · foto · limpieza)

Cierra un ciclo mensual: **congela** el mes, **cosecha** su registro inmutable a una
carpeta trackeada, y **limpia** los temporales para que el próximo mes arranque limpio.

> **Estado:** rediseñado 2026-07-03 · pendiente de implementación.
> Reemplaza el diseño original (generar arrastres) — ese trabajo ya lo absorbió
> `5_cobranza` (`arrastre_consolidado`) + `seguimiento_pueblo`. Ver "Por qué cambió".

---

## Qué hace — en una frase

`7_cierre` **no genera** arrastres ni saldos. Solo **transiciona el período**:
sella `estado_ciclo → CERRADO`, materializa la foto inmutable del mes en
`7_cierre/archivo/YYYY-MM/`, y resetea los slots mutables para el mes siguiente.

```
python 7_cierre/consolidar_cierre.py --mes 2026-06

  PASO 1 · GATE      ¿estado_ciclo[mes].arrastre.validado == true?  NO → aborta
  PASO 2 · COSECHAR  copia canónicos + fuentes de pago → archivo/YYYY-MM/ (trackeado)
  PASO 3 · FREEZE    estado_ciclo[mes].estado = CERRADO
  PASO 4 · LIMPIAR   reset templates · borra basura · deja lo month-stamped

  (persistir = paso SEPARADO — el script imprime el comando git, no auto-commitea)
```

---

## Por qué cambió (el diseño viejo quedó obsoleto)

El README original hacía que `7_cierre` generara `arrastre_corte_final` y
`arrastre_deuda_final`. Ese trabajo migró:

| Lo que decía el README viejo | Quién lo hace hoy |
|---|---|
| `arrastre_corte_final` · `arrastre_deuda_final` | `arrastre_consolidado` (5_cobranza, OUTPUT 6) |
| copiar arrastres a `2_planilla/inputs/` | ELIMINADO — 2_planilla lee en vivo (Opción A) |
| saldo de multa/acuerdos/convenio | `seguimiento_pueblo` (ledger event-sourced) |

Todo el scope original está absorbido. Lo que **nadie** hacía todavía —y es la razón
de existir de este módulo— es la **transición de mes**: freeze + foto + limpieza.

---

## Cuándo correr

Al **cierre del período**, después de que `5b_validacion` selló `validado:true` en
`estado_ciclo.json` para el mes. Ese sello es el gate (PASO 1): sin él, `7_cierre` aborta.

```
… → 5_cobranza → 5b_validacion (valida) → 7_cierre (cierra) → (próximo mes)
```

**Nueva dependencia dura:** `7_cierre` del mes N es prerequisito de `2_planilla` del mes
N+1 — julio lee el arrastre del congelado (`archivo/2026-06/`), que solo existe si junio
se cerró. Es el candado deseado: no se construye julio sobre un junio sin cerrar.

Es idempotente: re-correrlo re-cosecha la foto y re-sella (no duplica).

---

## Estructura de carpetas

```
7_cierre/
├── README.md
├── config.py
├── consolidar_cierre.py        ← script único
├── docs/
│   ├── diagrama_flujo_7_cierre.html          ← vista de 5 segundos
│   └── diagrama_consolidador_cierre.html      ← detalle de reglas
├── archivo/                    ← TRACKEADO en git · la foto inmutable por período
│   └── 2026-06/
│       ├── arrastre_consolidado_2026-06.xlsx
│       ├── planilla_cobrado.xlsx
│       ├── arrastre_devolucion_2026-06.xlsx
│       ├── mesa_1..7.xlsx
│       └── correcciones_lote.xlsx
└── outputs/
    └── run.log
```

`archivo/` NO cae bajo la regla `.gitignore: outputs/` → se versiona. Es la excepción
deliberada: los outputs canónicos de un mes cerrado NO son regenerables (trabajo manual +
historia de ledger + números validados) → se congelan como registro de auditoría.

---

## Los 3 baldes — qué se congela, qué se limpia, qué no se toca

### BALDE 1 · Permanente por naturaleza (ledger / acumulador) — 7_cierre NO lo toca
Vive siempre; el próximo mes le hace *append*. Se consulta vivo, filtrando por período.

```
shared/seguimiento_pueblo.xlsx (+ vista)   shared/blancos_acumulados.xlsx
shared/registro_cortes.xlsx                shared/usuarios_id.xlsx
shared/data_boletas_audit.xlsx             shared/…/estado_ciclo.json
1_lecturas/inputs/registro_operario_acumulado.xlsx
trazabilidad_cobranza.xlsx · trazabilidad_reclamos.xlsx  (append-only)
```

### BALDE 2 · Arrastre del mes — se COSECHA a `archivo/` (foto congelada)
Canónicos derivados + fuentes de pago del período. Julio lee el arrastre de acá.

```
CANÓNICOS         arrastre_consolidado_2026-06 · planilla_cobrado · arrastre_devolucion
FUENTES DE PAGO   mesa_1..7.xlsx · correcciones_lote.xlsx · reportes yape del mes
```

### BALDE 3 · Temporal del mes — se LIMPIA
```
RESET a template   mesa_1..7.xlsx · correcciones_lote.xlsx   (ya archivados → slots vacíos)
BORRA basura       trazabilidad_*_pre_dedup_*.xlsx · discrepancias · validacion_errores
DEJA quieto        lecturas_2026-06 · arrastre_deuda_2026-06 (residuo)  — month-stamped,
                   julio usa *_2026-07, no colisionan
```

---

## Decisiones de diseño (2026-07-03)

| # | Decisión | Por qué |
|---|---|---|
| A1 | Limpiar = **reset a template** (no borrar) | Julio necesita los `mesa_*` vacíos con fila-ejemplo, no ausentes |
| A2 | `archivo/` guarda **también las fuentes de pago mutables** (mesa, correcciones) | `git-history` versiona código, no datos; un slot mutable destruye el dato del período al reusarse. `archivo/YYYY-MM/` = ancestro file-era de la partición por período en Postgres |
| A3 | `arrastre_deuda` = residuo (retirar su generación aparte) · `lista_multas` = NO residuo (working de 6b) | `arrastre_deuda` sin consumidor (2_planilla migró al consolidado); `lista_multas` se regenera cada ciclo dentro de 6b |
| Q1 | **NO auto-commit.** El script sella + imprime el comando `git` | Cerrar ≠ persistir. El backend (git hoy, DB mañana) es pluggable — un agente hace `cerrar()` → `persistir()` sin tocar 7_cierre |
| Q2 | **NO snapshot de ledgers.** Se consultan vivos con filtro por período | El ledger es append-only → cualquier estado pasado es reconstruible (`get_saldo(mes)`). Fotografiarlo = segunda copia = riesgo de desync (anti-patrón B7). El aporte de junio ya quedó en `arrastre_consolidado` |
| — | Julio lee el arrastre del **congelado** (`archivo/`), no del `outputs/` vivo | `outputs/` = churn de dev; `archivo/` = contrato inmutable del período |

---

## Inputs

| Archivo | Origen | Rol |
|---|---|---|
| `estado_ciclo.json` | `shared/reporte_acumulado_procesado/` | Gate (validado) + destino del freeze |
| `arrastre_consolidado_YYYY-MM.xlsx` | `5_cobranza/outputs/` | Canónico a cosechar |
| `planilla_cobrado.xlsx` | `5_cobranza/outputs/` | Canónico a cosechar |
| `arrastre_devolucion_YYYY-MM.xlsx` | `5_cobranza/outputs/` | Canónico a cosechar (tiene REVISION manual) |
| `mesa_1..7.xlsx` | `4_pagos/efectivo/inputs/` | Fuente de pago a cosechar + resetear |
| `correcciones_lote.xlsx` | `5_cobranza/inputs/` | Fuente de corrección a cosechar + resetear |

---

## Outputs

| Archivo | Descripción |
|---|---|
| `archivo/YYYY-MM/*.xlsx` | La foto inmutable del período (trackeada en git) |
| `estado_ciclo.json` (actualizado) | `estado: CERRADO` para el mes |
| `outputs/run.log` | Log del cierre |
| (consola) | El comando `git add … && git commit …` listo para copiar/pegar |

---

## Flujo paso a paso

```
python 7_cierre/consolidar_cierre.py --mes 2026-06
```

1. **GATE** — leer `estado_ciclo.json`; si `[mes].arrastre.validado != true` → abortar
2. **COSECHAR** — crear `archivo/2026-06/`, copiar canónicos + fuentes de pago del BALDE 2
3. **FREEZE** — `estado_ciclo["2026-06"].estado = "CERRADO"` + `cerrado_en` timestamp
4. **LIMPIAR** — reset `mesa_*`/`correcciones_lote` a template · borrar basura del BALDE 3
5. **run.log** + imprimir el comando git de persistencia (paso manual/agente separado)

---

## Lo que NO hace este módulo

- No genera arrastres ni saldos (lo hace `5_cobranza`)
- No regenera un mes cerrado (un mes cerrado se fotografía, no se re-deriva)
- No hace `git commit` (lo emite; persistir es un paso separado, pluggable)
- No fotografía los ledgers (se consultan vivos con filtro por período)
- No toca el BALDE 1 (acumuladores/ledgers viven; el próximo mes les hace append)

---

## Errores comunes

| Error | Causa | Solución |
|---|---|---|
| `Ciclo no validado — aborta` | `5b_validacion` no selló `validado:true` | Correr `5b_validacion` primero |
| `2_planilla no encuentra el consolidado del mes anterior` | `7_cierre` del mes previo no corrió | Cerrar el mes anterior antes de generar la planilla |
| `mesa_N.xlsx tiene datos de dos meses` | Se corrió `4_pagos` de julio sin cerrar junio | El reset del PASO 4 evita esto — correr el cierre entre ciclos |
