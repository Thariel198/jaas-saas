# jass_system

Sistema de gestión mensual de cobros de agua para la JASS. Automatiza el ciclo completo: desde la lectura de medidores hasta el cierre del mes con arrastres para el siguiente ciclo.

---

## Pipeline de módulos

```
0_padron → 1_lecturas → 2_planilla → 3_boletas
→ [4_pagos + 4b_reclamos]
→ [5_cobranza + 5b_validacion]
→ [6_corte + 6b_corte_multas]
→ 7_cierre                       ← último módulo del pipeline (proceso mensual)

libro_mayor/  (substrato, sin número — NO es un paso del pipeline)
   ├── caja/            registro del dinero (abonos)
   └── estado_cuenta/   registro de la deuda (cargos) + motor de aplicación
```

`7_cierre` cierra el mes; recién ahí el proceso **asienta** el mes en `libro_mayor/`,
el registro permanente. `libro_mayor/` no lleva número porque no transforma el flujo
del mes — es el libro mayor sobre el que todos los meses asientan y del que todos leen.

| Módulo | Qué hace | Estado |
|---|---|---|
| `0_padron` | Limpia, reconcilia y mantiene el padrón de socios (MZ+LT), con COFOPRI como fuente de verdad. Sub-módulos: `01_limpieza`, `02_matching` + override auditado. | Operativo |
| `1_lecturas` | Sincroniza el registro del operario con el padrón y procesa las lecturas mensuales del medidor. Produce `lecturas_planilla_YYYY-MM.xlsx`. | Operativo |
| `2_planilla` | Genera la planilla mensual de cobro consolidando lecturas con deudas arrastradas del mes anterior. | Operativo |
| `3_boletas` | Genera las boletas de cobro por socio. | En desarrollo |
| `4_pagos` | Registra pagos recibidos — efectivo y Yape — y los aplica a la planilla. | Operativo |
| `4b_reclamos` | Gestiona el ciclo de vida de reclamos detectados en cobros de efectivo: detección, clasificación, corrección y cierre con auditoría. | Operativo |
| `5_cobranza` | Consolida la cobranza del mes: planilla cobrado, correcciones, arrastres y devoluciones. | Operativo |
| `5b_validacion` | Valida que el dinero cobrado cuadre con lo registrado en el sistema. | Operativo |
| `6_corte` | Genera la lista de usuarios en mora elegibles para corte, aplica penalidad de S/20, gestiona ventana de gracia de 2 días. Ciclo: `BORRADOR → PUBLICADA → COMPROMETIDA`. **Destino post-ledger:** corte unificado **multi-motivo** (agua + multa/acuerdos) — absorbe `6b_corte_multas` como trigger `multa`. | Operativo · reshaped |
| `6b_corte_multas` | Espejo de `6_corte` para deuda de multas y acuerdos de asamblea. Penalidad de S/40. **Destino post-ledger:** SE DISUELVE en `6_corte` (trigger `multa` del motor de corte único); el spec destino vive en `6_corte/README.md`. | Operativo (pre-ledger) · se disuelve |
| `7_cierre` | Consolida decisiones del mes (reclamos + cortes) y produce arrastres finales para que `2_planilla` del próximo mes los pre-cargue. Último módulo del pipeline. | Diseñado — pendiente de implementación |

Los módulos `b` son dependientes del principal: `4b` valida pagos, `5b` cuadra el dinero, `6b` gestiona multas paralelas al corte. **En el destino post-ledger, `6b` se disuelve** en `6_corte` (pasa a ser el trigger `multa` del motor de corte único); `5b` también se disuelve (→ `arqueo_caja` + `conciliar_caja`). El código pre-ledger sigue corriéndolos como módulos separados hasta Fase 2.

### `libro_mayor/` — sistema de registro (bounded context, sin número)

No es un módulo del pipeline sino el **substrato** permanente que `7_cierre` alimenta y del que todos leen. Un solo bounded context (→ 1 esquema Postgres, 1 servicio Docker, 1 set de tools de agente) con dos agregados adentro:

| Sub-contexto | Qué hace | Estado |
|---|---|---|
| `libro_mayor/caja` | Ledger append-only del **dinero** (HECHO): abonos y devoluciones por canal, `ABONO_ID` determinista, multi-tenant (`JASS_ID`). Reemplaza al descartado `7b_historial_pagos`. Responde "¿pagué? ¿cuándo? ¿canal? ¿cuánto?". | Fase 1 cerrada — pendiente de implementación |
| `libro_mayor/estado_cuenta` | Ledger append-only de la **deuda** (cargos) + **motor de aplicación** que imputa abonos a cargos por prioridad. Rediseña `seguimiento_pueblo`. Responde "tu pago fue a multa, aún debes medidor". | Fase 1 cerrada — pendiente de implementación |

`caja` y `estado_cuenta` son **un solo contexto con dos agregados**, no dos módulos: el motor de aplicación imputa abonos a cargos atómicamente (consistencia fuerte), lo que exige una sola frontera transaccional. El **contrato de interfaz del ledger** (entidades ABONO/CARGO/APLICACIÓN, `JASS_ID`, motor) vive byte-idéntico en `libro_mayor/caja/README.md` y `libro_mayor/estado_cuenta/README.md`. Ver `libro_mayor/README.md`.

**Reglas puras — `libro_mayor/dominio/`** (Fase 1, spec cerrado 2026-07-14): taxonomía de conceptos, cascada de prioridad P1-P6, política de corte, saldo derivado e identidad determinista, extraídas del código real como funciones **puras** y **tenant-agnósticas** (montos en `int` de céntimos, sin `TOL`, cero I/O). Las importan los dos agregados; el motor de aplicación las invoca. Detalle de las 6 firmas en `docs/RETOMAR_dominio_saldo_unico_2026-07-13.md`.

**Extracto de cuenta + arquitectura de render** (decisión ⑫, 2026-07-13): el historial multi-mes de un predio ("esto pasó en tu cuenta") es una tool de solo lectura de `estado_cuenta` (`extracto_predio`), distinta de la boleta ("esto debes ahora", 1 mes). **Ningún módulo de negocio imprime**: cada dueño de datos arma sus filas (`3_boletas` la boleta, `estado_cuenta` el extracto) y un servicio **stateless** `render(plantilla, filas) → PDF` las convierte. `3_boletas` **no se renombra** a `3_impresor` — se queda como dueño de datos de la boleta.

---

## Cómo correr el ciclo mensual

Cada módulo tiene su propio `README.md` con el flujo paso a paso y los comandos exactos.
El orden del pipeline es el de la tabla de arriba.

```bash
# Ejemplo: correr un módulo
cd 1_lecturas
python main.py
# El resultado queda en outputs/ junto con run.log
```

---

## Estructura general de cada módulo

```
MODULO/
├── inputs/          ← archivos de entrada del ciclo
├── outputs/         ← archivos generados + run.log
├── docs/            ← diagramas HTML y contratos de formato
├── tests/           ← tests sintéticos (donde aplica)
├── backup/          ← respaldos automáticos pre-corrida
├── config.py        ← paths y constantes del módulo
├── main.py          ← punto de entrada principal
└── README.md        ← fuente de verdad del módulo
```

---

## Recursos compartidos

```
shared/          ← funciones puras reutilizables entre módulos
docs/            ← metodología, arquitectura del sistema, skill tracker
recursos/        ← archivos de referencia (tarifas, padrón base, etc.)
```

---

## Instalación

```bash
pip install -r requirements.txt
```

Requiere Python 3.10+, pandas, openpyxl.
