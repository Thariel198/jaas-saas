# Pendientes y plan de sesiones — jass_system

**Actualizado:** 2026-06-28 | Inventario completo del estado del ciclo 2026-06 y pipeline hacia 2026-07.

> **Nota:** memoria del proyecto vive en `C:\Users\wilde\.claude\projects\C--Users-wilde\memory\project_jass_system.md` — leer al iniciar cada sesión.

---

## Estado actual del pipeline (2026-06-26)

```
0_padron   → operativo (scripts sueltos) · sin estructura de módulo estándar
1_lecturas → operativo · ciclo 2026-06 CERRADO · template 2026-07 generado con 15 bloqueantes
2_planilla → planilla_2026-06.xlsx existe en outputs/ · run.log incompleto (sin SUCCESS ni ERROR post 15:02)
3_boletas  → funciona manualmente · sin convenciones del sistema
4_pagos    → efectivo OK · Yape OK · rename DEVOLUCION→RETORNO hecho en código · archivo histórico sin renombrar
4b_reclamos → ciclo 2026-06 completo
5_cobranza → ciclo corrido · planilla_cobrado.xlsx OK · trazabilidad OK
5b_validacion → OPERATIVO · efectivo OK · yape por-MZ OK · crudo gap Yape -S/706 pendiente · DE6 pendiente (diseño devolucion/retorno)
6_corte    → completo · 1 salvado · 4 cortes físicos · arrastre_corte_2026-06.xlsx
6b_corte_multas → lista_multas.xlsx generada (91 elegibles) · aplicar_penalidad y seguimiento NO corrieron
7_cierre   → diseñado · README + 2 HTML · consolidar_cierre.py NO EXISTE
```

---

## 1. Bugs de código

| # | Descripción | Módulo | Causa raíz | Complejidad |
|---|---|---|---|---|
| ~~B2~~ ✅ (2026-06-29) | **RESUELTO.** Duplicados en `trazabilidad_cobranza.xlsx` (997 filas, 839 dup). Causa raíz: `_fecha_hora_str` hacía `str(Timestamp)`+reparse `dayfirst=True` → invertía día/mes (5 jun→6 may). Trazabilidad guardaba fecha corrupta pero la identidad de idempotencia usaba la fecha correcta → re-agregaba ~277 pagos por corrida. Fix 1: `_fecha_hora_str` formatea Timestamp/datetime directo. Fix 2: excluir huérfanos (BLANCO sin lote) de `ids_actuales`. Limpieza one-time: dedup 997→448 + fechas corregidas desde fuente (backup `trazabilidad_cobranza_pre_dedup_*.xlsx`). Verificado: idempotencia exacta (0 nuevos), 5b_validacion OK | `5_cobranza` | — | — |
| B4 | **NUEVO** `2_planilla/main.py` escribe `TOTAL_A_PAGAR = None` (fórmula Excel, no valor Python). Si la planilla se re-genera y no se abre en Excel antes de correr `5_cobranza`, `TOTAL_A_PAGAR=NaN` → 5_cobranza ve deuda 0 → 309 EXCESO falso. Evidencia: run 2026-06-26 17:19 produjo CANCELADO=101 EXCESO=309 (vs Jun 24: CANCELADO=291 EXCESO=17). Fix: (A) abrir planilla en Excel y guardar antes de correr 5_cobranza, (B) calcular TOTAL_A_PAGAR en Python en 2_planilla. Decisión pendiente. **5_cobranza/outputs/ de Jun 26 17:19 son inválidos — no usar.** | `2_planilla`, `5_cobranza` | Diseño — fórmula vs Python | Media |

---

## 2. Datos / registros manuales pendientes

| # | Descripción | Módulo | Complejidad |
|---|---|---|---|
| D1 | `2_planilla/inputs/` tiene todos los subdirectorios vacíos (corte, deuda_anterior, multas, convenios, acuerdos_asamblea). Para ciclo 2026-07 estos arrastres deben copiarse desde sus módulos fuente antes de correr `main.py` | `2_planilla` | Mecánica |
| D2 | Dos pagos en blanco (O-6: S/107 y R-7: S/24) quedan sin dueño asignado. Sus dueños deben reclamar en ciclo 2026-07 — ya documentado en `deciciones/efectivo_2026-06_conflictos_lotes.md` | `4_pagos/efectivo` | Mecánica — proceso ya definido |
| D3 | `1_lecturas` produjo `orden_verificacion_2026-07.pdf` con 15 bloqueantes de campo. El operario debe resolverlos antes del siguiente ciclo de correcciones | `1_lecturas` | Mecánica — operativa de campo |

---

## 3. Módulos sin terminar o por codificar

| # | Descripción | Módulo | Causa raíz | Complejidad |
|---|---|---|---|---|
| M1 | `7_cierre/consolidar_cierre.py` no existe. README completo, diseño cerrado (2 HTML en docs/). Implementación pendiente | `7_cierre` | No se ha iniciado la Fase 2 | Media |
| M2 | `6b_corte_multas`: faltan correr `aplicar_penalidad_multas.py` (Día 0) y `seguimiento_multas.py` (Día 2) para el ciclo 2026-06 | `6b_corte_multas` | Scripts implementados pero no ejecutados | Mecánica |
| M3 | `3_boletas`: funciona pero no sigue las convenciones del sistema (usa `Outputs/` con mayúscula, sin `config.py`, sin `README.md`, sin `docs/`) | `3_boletas` | Módulo heredado pre-convenciones · estado "en desarrollo" | Alta (refactor) |
| M4 | `0_padron`: sin `README.md` principal, sin `main.py`, sin estructura de módulo. Solo scripts sueltos con sub-módulos. El output `padron_reconciliado.xlsx` funciona pero el pipeline no está documentado ni orquestado | `0_padron` | Módulo creado antes de la metodología actual | Alta |

---

## 4. HTML / contratos pendientes

| # | Descripción | Módulo | Complejidad |
|---|---|---|---|
| H1 | `0_padron/docs/`: solo tiene `flujo_comparacion.html`. Falta `arquitectura_0_padron.html`, `diagrama_0_padron.html`, contratos de formato de outputs (`padron_reconciliado`, `correcciones`, `reporte_conflictos`) | `0_padron` | Mecánica |
| H2 | `3_boletas/docs/`: carpeta no existe. Falta todo el stack: `diagrama_flujo_3_boletas.html`, `arquitectura_3_boletas.html`, `formato_CONSOLIDADO.html` | `3_boletas` | Mecánica (bloqueado por M3) |
| H3 | `4_pagos/docs/`: solo tiene `manual_orquestacion.md`. Falta `diagrama_flujo_4_pagos.html`, `arquitectura_4_pagos.html`, contratos de formato de `pagos_efectivo.xlsx` y outputs Yape | `4_pagos` | Mecánica |
| H4 | `2_planilla/docs/formato_planilla.html`: existe pero puede estar desactualizado respecto a las columnas actuales (MULTA, ACUERDOS_ASAMBLEA, PENALIDAD_MULTA). Verificar | `2_planilla` | Mecánica |
| H5 | **NUEVO** `5b_validacion/docs/`: solo `manual_uso.md` y `validacion_diseno.html`. Falta `diagrama_flujo_5b_validacion.html`, `arquitectura_5b_validacion.html` | `5b_validacion` | Mecánica |

---

## 5. Decisiones de negocio / diseño sin definir

| # | Descripción | Módulo | Complejidad |
|---|---|---|---|
| DE2 | `6b_corte_multas` Día 0+2: ¿ejecutar `aplicar_penalidad_multas.py` y `seguimiento_multas.py` para ciclo 2026-06 o diferir? La ventana de gracia ya venció. La penalidad no se aplicó a la planilla del mes. `compromisos.xlsx` tiene más allá de los 23 del tag — editar a mano o re-correr antes de aplicar penalidad | `6b_corte_multas` | Decisión de negocio — Alta (impacto en planilla vigente) |
| DE3 | `7_cierre` arrastre de reclamos EN_REVISION: el README define la regla especial de junio 2026 (diferir deuda, no cortar). ¿Esta regla aplica también a julio 2026 o se normaliza? | `7_cierre` | Decisión de negocio — Media |
| DE4 | `3_boletas`: ¿refactor completo para seguir convenciones del sistema (M3) o mantener como herramienta manual independiente? Afecta si `DATA_boletas_repo` integra con `3_boletas` | `3_boletas` | Diseño — Alta |
| DE5 | **Arrastre consolidado (Opción B — APROBADA):** `5_cobranza` debe generar UN solo arrastre con todos los componentes (agua + corte + multa) en orden de prioridad, reemplazando los 3 arrastres separados (5/6/6b). La descomposición por prioridad vive en `5_cobranza`, no en `6b`. Diseño pendiente de especificar antes de implementar | `5_cobranza`, `6_corte`, `6b_corte_multas` | Diseño — Alta |

---

## 6. Tests que faltan correr o están fallando

| # | Descripción | Módulo | Complejidad |
|---|---|---|---|
| T1 | `5b_validacion/tests/test_validacion.py`: puede intentar correr post-B3 (B1 ya resuelto). Verificar que los totales cuadran con datos reales | `5b_validacion` | Mecánica (post B3) |
| T2 | `1_lecturas/tests/test_anomalias_integracion.py`: verificar que integración con `sin_servicio` sigue pasando después del último commit `feat(sin_servicio)` | `1_lecturas` | Mecánica |
| T3 | `6_corte/tests/` y `6b_corte_multas/tests/`: verificar que los tests existen y pasan antes del ciclo 2026-07 | `6_corte`, `6b_corte_multas` | Mecánica |
| T4 | `4b_reclamos/tests/`: verificar estado post-commit del ciclo 2026-06 | `4b_reclamos` | Mecánica |

---

## 7. Scripts (main.py) que faltan correr o necesitan verificación

| # | Descripción | Módulo | Complejidad |
|---|---|---|---|
| S2 | `6b_corte_multas/aplicar_penalidad_multas.py` + `seguimiento_multas.py`: nunca corrieron para 2026-06. Bloqueado por decisión DE2 | `6b_corte_multas` | Mecánica (post DE2) |
| S3 | `2_planilla/main.py`: fix `header=1` aplicado. PERO planilla regenerada tiene `TOTAL_A_PAGAR=NaN` (ver B4). Planilla ciclo 2026-06 restaurada del backup — operativa. Para ciclo 2026-07 debe resolverse B4 antes de correr. | `2_planilla` | Bloqueado por B4 |

---

## Plan por modelo y sesión

### Sesión A — Sonnet ✅ COMPLETADA (2026-06-28)

B3 resuelto, S1 corrida (efectivo OK, yape por-MZ OK), S4 re-corrida OK (CANCELADO=291), D4 false positive confirmado. Pendiente menor: S3 (2_planilla run.log).

1. ~~[B3]~~ ✅
2. ~~[S1]~~ ✅ — crudo gap Yape -S/706 no es bug, es TE PAGÓ sin identificar
3. **[S3]** (pendiente menor) — verificar que `2_planilla/main.py` corre bien para 2026-07
4. ~~[D4]~~ ✅ false positive

---

### Sesión B — Opus (decisiones de negocio + causa raíz B2)

**Orden:** Después de Sesión A. Desbloquea 6b_corte_multas, 7_cierre, arrastre consolidado.

1. ~~**[B2]**~~ ✅ (2026-06-29) RESUELTO — causa raíz: `_fecha_hora_str` corrompía fechas (dayfirst reparse) rompiendo idempotencia. Ver tabla Bugs.
2. **[DE2]** Decidir qué hacer con `6b_corte_multas` ciclo 2026-06: ¿aplicar penalidad ahora o cerrar el ciclo sin ella? (verificar `compromisos.xlsx` primero)
3. **[DE3]** Definir si la regla de reclamos EN_REVISION de junio 2026 se normaliza en julio o continúa
4. **[DE4]** Decidir si `3_boletas` se refactoriza o permanece como herramienta manual independiente
5. **[DE5]** Especificar el diseño del arrastre consolidado (Opción B): schema exacto, orden de prioridad, qué reemplaza
6. ~~**[DE6]** Diseño~~ ✅ (2026-06-28) — ~~**[DE6-CODE]**~~ ✅ (2026-06-29) split en `motor_matching` + loader `_cargar_devueltos_yape` + columna DEVUELTO en `planilla_cobrado` y `trazabilidad`

---

### Sesión C — Haiku (ejecución mecánica post-decisiones)

**Orden:** Después de Sesión B.

1. **[T1]** Correr `5b_validacion/tests/test_validacion.py`
2. **[T2]** Correr `1_lecturas/tests/test_anomalias_integracion.py`
3. **[T3]** Verificar tests `6_corte/tests/` y `6b_corte_multas/tests/`
4. **[T4]** Verificar tests `4b_reclamos/tests/`
5. **[M2]** Si DE2 = sí: correr `aplicar_penalidad_multas.py` y `seguimiento_multas.py` para 2026-06
6. **[D1]** Preparar inputs de arrastres para ciclo 2026-07: copiar outputs de 6_corte, 5_cobranza, 6b_corte_multas → `2_planilla/inputs/`
7. **[H1]** Crear HTMLs `0_padron/docs/` (proponer en consola → aprobación → crear)
8. **[H3]** Crear HTMLs `4_pagos/docs/` (proponer en consola → aprobación → crear)
9. **[H4]** Auditar y actualizar `2_planilla/docs/formato_planilla.html` respecto a columnas actuales
10. **[H5]** Crear HTMLs faltantes `5b_validacion/docs/`

---

### Sesión DE7-CODE — Sonnet (codificar CONCEPTO en el pipeline)

**Orden:** Puede hacerse antes de Sesión B. Diseño HTML cerrado (2026-06-29). Formatos actualizados: `formato_pagos_efectivo.html`, `trazabilidad_cobranza.html`, `pagos_yape_tepago_diseno.html`.

1. **`4_pagos/efectivo/main.py`**: leer columna CONCEPTO de `pagos_efectivo.xlsx` y propagarla al output (hoy se ignora). Vacío = agua.
2. **`4_pagos/yape/motor_matching/exportar_motor.py`**: agregar `tanque` al dropdown de validación de datos de CONCEPTO en hoja Ambiguos de `pendientes.xlsx`. Vocabulario: `comunitario | tanque | honorario | gasto`.
3. **`5_cobranza/main.py`**: leer CONCEPTO desde `pagos_efectivo` (efectivo) y desde `pagos_yape_tepago` (yape). Excluir del cálculo de agua todo pago con CONCEPTO no vacío. Propagar CONCEPTO a `trazabilidad_cobranza.xlsx`.
4. Verificar regresión: correr `5_cobranza/main.py` + `5b_validacion/main.py` y confirmar que CANCELADO=291 EXCESO≈0 (los 17 EXCESO originales pasan a CONCEPTO=tanque/etc.).

---

### Sesión D — Sonnet (implementación 7_cierre)

**Orden:** Después de Sesión B (DE3 resuelto). El spec de 7_cierre está cerrado en `README.md`.

1. **[M1]** Implementar `7_cierre/config.py`
2. **[M1]** Implementar `7_cierre/consolidar_cierre.py` (flujo de 9 pasos documentado en README)
3. Crear `7_cierre/inputs/` y `7_cierre/outputs/` con archivos de prueba
4. Correr contra datos reales 2026-06

---

### Sesión E — Opus + Sonnet (0_padron y 3_boletas — alta complejidad)

**Orden:** Puede ser paralela a C/D. Son módulos independientes del pipeline principal.

1. **[M4]** Diseñar estructura de módulo `0_padron` (README principal, main.py, convenciones)
2. **[DE4]** Si se decidió refactorizar: **[M3]** Refactor `3_boletas` para seguir convenciones

---

### Pendientes sin clasificación clara

- **Directorio `Pendiente/`** (raíz del proyecto): tiene `6b_corte_multas_README.md`. Evaluar si es draft o resto.
- **`comparar_padrones.py` y `tmp_explorar_mesas.py`** en raíz: probablemente desechables, confirmar antes de borrar.
- **`docs/trazabilidad_2026_06.xlsx`**: hay un xlsx dentro de `docs/`. Parece estar en el lugar equivocado — debería estar en un módulo o en `shared/`.
- **`H2` `3_boletas/docs/`**: bloqueado por DE4 (decisión de refactor) — no crear HTMLs si el módulo va a cambiar de estructura.
- **Skill tracker** (baja urgencia): documentar "por qué" en código, manejo de errores descriptivo, thin layer en `0_padron`, enriquecimiento `3_boletas` (sub-módulos 3.1/3.2).

---

## Blockers actuales por módulo

```
5_cobranza     → B4 (TOTAL_A_PAGAR=NaN si se re-genera planilla — fix pendiente para 2026-07)
5b_validacion  → crudo gap Yape -S/706 (TE PAGÓ sin identificar en banco) · DE6-CODE (implementación código pendiente)
6b_corte_multas (Día 0+2) → DE2 (decisión de negocio)
7_cierre       → M1 (código no existe) + DE3 (regla EN_REVISION)
ciclo 2026-07  → D1 (arrastres vacíos) + D3 (15 bloqueantes de campo) + B4 (fix antes de correr 2_planilla)
3_boletas docs → DE4 (decisión de refactor)
```

---

## Sesión 2026-06-29 (Opus) — decisiones y hallazgos nuevos

**B2 RESUELTO** (ver tabla Bugs): `_fecha_hora_str` corrompía fechas (dayfirst) rompiendo idempotencia → trazabilidad re-agregaba ~277 pagos/corrida. Fix + dedup 997→448 + huérfanos fuera de `ids_actuales`. Verificado.

**DE7 (NUEVO) — Generalizar CONCEPTO como vocabulario de ruteo de conceptos.** *(Aprobado · el usuario lo implementa.)*
Agregar columna CONCEPTO a los documentos del pipeline (los 7 + pagos_efectivo + trazabilidad), no solo Yape. CONCEPTO ya tiene 3 usos (comunitario, gasto/honorario, **tanque**) → generalizar ahora cumple Regla del Tres y el criterio agentic SaaS (etiqueta explícita ruteable por agente, vs BLANCO sobrecargado o columna-por-concepto rígida). 5_cobranza ya ignora pagos con CONCEPTO seteado → un pago marcado sale del cálculo de agua. Pendiente de DISEÑO antes de código (trigger C): qué documentos lo llevan, vocabulario válido, quién escribe/consume, actualizar `formato_*.html` + `diagrama_flujo`. | `todos` | Diseño — Alta |

**DE8 (NUEVO) — Regla de ventas para el contador.** ventas = consumo + mantenimiento (corte EXCLUIDO: concepto agua pero no es venta por ley peruana). Base = dinero **real** cobrado con parciales: `min(PAGO, MES_ACTUAL + MANTENIMIENTO + MES_ANTERIOR)` sobre los que pagaron. Jun-2026 = **S/ 7,585.50**. Inmune a exceso/duplicados (el cap a lo facturado los neutraliza). El reporte solo es oficial tras 5b_validacion OK. Falta: generar `ventas_contador_YYYY-MM.xlsx` (¿en 5b o 5_cobranza?). | `5b_validacion` / `5_cobranza` | Implementación |

**Hallazgo — el "exceso" NO son duplicados.** Los 17 EXCESO son pagos por **conceptos no modelados** en la planilla: tanque (A-3 S/200, E-1 S/100), deuda histórica (M-12 S/266, "2019 hasta octubre 2025"), reclamo (C1-17 S/170), + redondeos y P-6 (S/300 por averiguar). Mesa/cobrador/comentario en `4_pagos/efectivo/outputs/pagos_efectivo.xlsx`; origen/mensaje yape en `pagos_yape_tepago.xlsx`. **Aporte tanque** → futura base de datos, alimentada vía CONCEPTO=tanque (DE7).

---

## SIGUIENTE_ACCION
modelo: Sonnet
sesion: Sesión DE7-CODE — codificar CONCEPTO en el pipeline
razon: Diseño HTML de DE7 cerrado (2026-06-29): formato_pagos_efectivo.html + trazabilidad_cobranza.html + pagos_yape_tepago_diseno.html actualizados con vocabulario tanque/honorario/gasto/comunitario. Siguiente paso: codificar las 4 tareas de DE7-CODE (ver sesión arriba). Después de DE7-CODE, seguir con Sesión B (Opus): DE2, DE3, DE4, DE5.
