# Pendientes y plan de sesiones — jass_system

**Actualizado:** 2026-07-02 | Inventario completo del estado del ciclo 2026-06 y pipeline hacia 2026-07.

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
| ~~B5~~ ✅ (2026-07-01) | **RESUELTO.** `correcciones_lote.xlsx` perdía correcciones manuales aplicadas en corridas previas (C-88→C-8B, P1-4→D1-4 desaparecieron, solo quedó B-21→B-14). Causa raíz: el archivo tipeado (`discrepancias_cobranza.xlsx` MZ_CORRECTO/LT_CORRECTO) se regenera desde cero cada corrida → borra lo tipeado; y `correcciones_lote.xlsx` es un archivo mutable único que un `git checkout`/revert puede devolver a su estado commiteado, perdiendo correcciones agregadas solo en working-tree. La única prueba durable de una corrección ya aplicada quedó en `trazabilidad_cobranza.xlsx` (columnas MZ_ORIGEN/LT_ORIGEN). Fix: `_recuperar_correcciones_trazabilidad()` reconstruye `correcciones_lote.xlsx` desde la trazabilidad en cada corrida (self-healing); `_leer_correcciones_tipeadas()` preserva MZ_CORRECTO/LT_CORRECTO ya tipeados al regenerar `discrepancias_cobranza.xlsx`. Verificado: recuperó las 2 correcciones perdidas, CANCELADO 291→292, EXCESO 17→16, idempotente en re-corrida. | `5_cobranza` | Archivo mutable único sin reconstrucción desde fuente derivada | — |
| B4 | **NUEVO** `2_planilla/main.py` escribe `TOTAL_A_PAGAR = None` (fórmula Excel, no valor Python). Si la planilla se re-genera y no se abre en Excel antes de correr `5_cobranza`, `TOTAL_A_PAGAR=NaN` → 5_cobranza ve deuda 0 → 309 EXCESO falso. Evidencia: run 2026-06-26 17:19 produjo CANCELADO=101 EXCESO=309 (vs Jun 24: CANCELADO=291 EXCESO=17). Fix: (A) abrir planilla en Excel y guardar antes de correr 5_cobranza, (B) calcular TOTAL_A_PAGAR en Python en 2_planilla. Decisión pendiente. **5_cobranza/outputs/ de Jun 26 17:19 son inválidos — no usar.** | `2_planilla`, `5_cobranza` | Diseño — fórmula vs Python | Media |
| B6 | **NUEVO (2026-07-01)** `4_pagos/efectivo/main.py:341` (`leer_hoja`) descarta en silencio cualquier pago con `MZ` o `LT` vacío (`if not mz or not lt: continue`) — incluye blancos reales (`MZ=BLANCO`, `LT=''`). A diferencia de yape (que rutea blancos a `blancos_mes.xlsx` → `shared/blancos_acumulados.xlsx`), efectivo no tiene ruteo de blancos: la plata desaparece del pipeline sin dejar rastro. Evidencia ciclo 2026-06: 3 filas descartadas por este bug, S/151 (mesa_1: Wilder Trujillo S/69 + S/37; mesa_6: Wagner Trujillo S/45, nota "Hernestina Valladares?"). Fix: detectar blanco real (cobrador + monto>0 + fecha, sin mz/lt) vs fila de subtotal (sin cobrador/fecha, se descarta igual) y emitir los blancos reales hacia el store canal-agnóstico `shared/blancos_acumulados.xlsx`, igual que yape. | `4_pagos/efectivo` | Mismo eje que DE10 (concepto/estado canal-agnóstico, no por canal) | Media |
| B7 | **PARCIAL (2026-07-01, 5ª sesión): Fase 1 (síntoma) HECHA · Fase 2 (arquitectura) = deuda.** Escaneo reveló **TRIPLE-writer** (2_planilla + 6_corte + **6b_corte_multas**), no dual. Fase 1: 7 CORTE_RECONEXION negativos (−20) corregidos a 0 → arrastre_devolucion 16→13, 5b OK. Fase 2 (Modelo A / overlay de AMBOS audits por **delta** col 5, no absoluto col 6; ajustar guard DESYNC de 6b) = sesión dedicada. Ver sección "Sesión 2026-07-01 (5ª)". Texto original abajo: **fix aprobado (Modelo A), NO codificado.** Violación del principio *writer único* sobre `shared/planilla_mes/planilla_YYYY-MM.xlsx`: **dos módulos lo escriben** — `2_planilla` (`publicar_a_shared`, `shutil.copy2` = sobreescritura total ciega) y `6_corte/aplicar_penalidad.py` (+20/−20 a `CORTE_RECONEXION`). Al re-correr `2_planilla` **después** de 6_corte, su copy2 pisa el +20; `audit_penalidad.xlsx` sigue marcando APLICADO → la siguiente reconciliación de 6_corte resta 20 sobre una base que ya volvió a 0 → `CORTE_RECONEXION = −20`. Efecto: **C-7 y C-8B aparecen como EXCESO S/20 falso** en `arrastre_devolucion` (2 de los 16 de D4 NO son sobrepagos reales — son este bug). Evidencia: backups en `6_corte/backup/planilla_mes/` muestran la secuencia 0→20→0→−20 (16/06→17/06). **Causa raíz:** el archivo publicado cumple dos roles incompatibles (foto inmutable de 2_planilla + libro mutable de 6_corte) = *lost update*. **Fix aprobado — Modelo A (overlay / writer único):** 6_corte deja de escribir en shared; la penalidad vive SOLO en `6_corte/outputs/audit_penalidad.xlsx`; 5_cobranza la lee en vivo y hace overlay (`CORTE_efectivo = base + neto(audit)`), materializando en su propio `planilla_cobrado.xlsx`. Doc didáctico con diagramas de cajas: `docs/aprendizaje/writer_unico_desincronizacion_20260701.html`. **Bug SIGUE VIVO hoy** en `shared/planilla_mes/planilla_2026-06.xlsx`. | `2_planilla`, `5_cobranza`, `6_corte` | Dual-writer / lost update | Media-Alta |

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
| H5 | **NUEVO** `5b_validacion/docs/`: solo `manual_uso.md` y `validacion_diseno.html`. Falta `diagrama_flujo_5b_validacion.html`, `arquitectura_5b_validacion.html`. Actualizar también para reflejar la columna REFERENCIA (col 12, grupo "¿Por qué?") agregada 2026-07-01 a `validacion_diferencias.xlsx` (hojas `yape_por_mz` y `efectivo_por_mz`) | `5b_validacion` | Mecánica |

---

## 5. Decisiones de negocio / diseño sin definir

| # | Descripción | Módulo | Complejidad |
|---|---|---|---|
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

### ~~Sesión DE7-CODE~~ ✅ HECHA (2026-06-29 · Sonnet) — codificar CONCEPTO en el pipeline

**Orden:** Puede hacerse antes de Sesión B. Diseño HTML cerrado (2026-06-29). Formatos actualizados: `formato_pagos_efectivo.html`, `trazabilidad_cobranza.html`, `pagos_yape_tepago_diseno.html`.

1. ~~**`4_pagos/efectivo/main.py`**: leer columna CONCEPTO de `pagos_efectivo.xlsx` y propagarla al output (hoy se ignora). Vacío = agua.~~ ✅ HECHO — `leer_hoja()` lee CONCEPTO, `exportar_pagos_efectivo()` escribe columna con badges de color. `pagos_efectivo.xlsx` ahora tiene 10 cols.
2. ~~**`4_pagos/yape/motor_matching/exportar_motor.py`**~~ ✅ NO-OP confirmado — CONCEPTO es texto libre en `pendientes.xlsx`, no hay dropdown.
3. ~~**`5_cobranza/main.py`**: leer CONCEPTO desde `pagos_efectivo` (efectivo) y desde `pagos_yape_tepago` (yape). Excluir del cálculo de agua todo pago con CONCEPTO no vacío. Propagar CONCEPTO a `trazabilidad_cobranza.xlsx`.~~ ✅ HECHO — loaders leen `concepto`, `_calcular()` filtra `ys_agua`/`es_agua`, `_exportar_trazabilidad_cobranza()` escribe CONCEPTO col 7 (RETORNO→8, DEVUELTO→9, grupos y separadores desplazados).
4. Verificar regresión: correr `5_cobranza/main.py` + `5b_validacion/main.py` y confirmar que CANCELADO=291 EXCESO≈0 (los 17 EXCESO originales pasan a CONCEPTO=tanque/etc.). **PENDIENTE** — el plombing está, pero los cobradores aún no han dividido pagos en mesa_N.xlsx con CONCEPTO. Re-correr el pipeline después de que llenen CONCEPTO en mesa files para verificar.

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
5b_validacion  → crudo gap Yape -S/706 (TE PAGÓ sin identificar en banco)
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

## Sesión 2026-07-01 (Opus, 4ª del día) — B7 descubierto + diseñado · REFERENCIA/COMENTARIO en arrastre_devolucion

**Origen:** al empezar D4 (revisar los 16 EXCESO), C-7 exceso = exactamente S/20 = penalidad de corte. El rastreo de `6_corte/backup/planilla_mes/*.xlsx` reveló la secuencia `CORTE_RECONEXION` 0→20→0→−20 → **bug B7 (dual-writer)**. Ver tabla Bugs.

**B7 — writer único violado en `shared/planilla_mes`.** `2_planilla` (copy2 total) y `6_corte/aplicar_penalidad` (+20 CORTE_RECON) escriben el mismo archivo. Re-correr 2_planilla pisa el +20; 6_corte después resta 20 sobre base 0 → −20. C-7 y C-8B son EXCESO falso (salen de D4). **Fix aprobado: Modelo A (overlay).** 6_corte deja de escribir en shared; penalidad solo en `audit_penalidad.xlsx`; 5_cobranza lee en vivo + overlay + materializa en `planilla_cobrado`. Bug NO codificado — ver SIGUIENTE_ACCION (prioridad absoluta). Doc: `docs/aprendizaje/writer_unico_desincronizacion_20260701.html`.

**Principio de arquitectura consolidado (candidato para metodología):** *writer único* = un archivo, un dueño que lo ESCRIBE. Los demás LEEN en vivo su `outputs/`, nunca escriben en él ni copian a su `inputs/`. Escribir gobierna exclusividad; leer no. Un archivo no puede ser "foto inmutable" y "libro mutable" a la vez — si dos módulos aportan datos, cada aporte va a un archivo propio y el consumidor los combina con overlay al leer.

**REFERENCIA + COMENTARIO agregadas a `arrastre_devolucion_YYYY-MM.xlsx`** (grupo "¿Cómo ubicarlo?", cols 9-10). Efectivo: `mesa / cobrador / fecha` + nota del cobrador. Yape: `origen / S/monto / fecha-hora-seg` + mensaje del banco. Si un lote tiene >1 pago-agua, se concatenan con ` · ` (mismo orden en ambas columnas). Fuente: los outputs procesados que 5_cobranza ya lee (`pagos_efectivo.xlsx` / `pagos_yape_tepago.xlsx`), sin dependencias nuevas. Helper nuevo `_fecha_hora_seg_str` (con segundos, solo para esta columna). HTML de diseño `arrastre_devolucion_diseno.html` actualizado. Verificado con datos reales (16 excesos) + test suite + idempotencia. Hallazgo colateral: E-1 tiene 2 pagos (efectivo S/8 + yape S/100 "tanque") — el yape cuenta como agua porque el "tanque" está en MENSAJE, no en CONCEPTO → para sacarlo del exceso hay que tagear CONCEPTO=tanque en la fuente.

---

## Sesión 2026-07-01 (Opus, 5ª del día) — B7 Fase 1 (desbloqueo) HECHA · hallazgo triple-writer

**Escaneo de consumidores de `CORTE_RECONEXION` (paso 2 de B7) reveló que son TRES writers, no dos:**
- `2_planilla/main.py` (copy2 total), `6_corte/aplicar_penalidad.py` (+40/−40), y **`6b_corte_multas/aplicar_penalidad_multas.py` (+40/−40, misma columna)**. Este último nunca corrió para 2026-06 (M2/DE2) pero el code-path es un tercer writer latente.
- Consumidores lectores (todos overlay-limpios): `5_cobranza:399,933` (aquí iría el overlay), `6_corte/seguimiento` (deriva de planilla_cobrado), `6b/seguimiento_multas:207` (lee planilla_cobrado, no shared — pero su guard DESYNC `:284` compara contra `CORTE_RECON_DESPUES` **absoluto**, que se pudre bajo overlay → hay que cambiarlo a base+delta). `7_cierre` no existe (M1), no lee la penalidad.

**Dos hallazgos que invalidan parte del diseño Modelo A cerrado el 2026-07-01 (4ª sesión):**
1. **Triple-writer, no dual.** Fix parcial (solo 6_corte→overlay, 6b sigue escribiendo) reintroduce el lost-update en 6b. Es "ambos o ninguno".
2. **El audit guarda ABSOLUTO, no delta.** `CORTE_RECON_DESPUES` (col 6) = base_al_momento+penalidad → se vuelve obsoleto si la base cambia. El overlay debe sumar el **delta** (`PENALIDAD_APLICADA` col 5), no el absoluto. Además `6_corte/outputs/audit_penalidad.xlsx` (net activo = F1-13, F1-4, S-1, S-5, V-14, todos +40) **no coincide con ningún valor vivo de shared/planilla_mes** → el copy2 de 2_planilla ya borró las penalidades del audit (el lost-update en crudo).

**FASE 1 — desbloqueo mínimo (HECHO, verificado):** el radio del bug en el dato vivo eran **7 predios con CORTE_RECONEXION = −20** (cargo de corte negativo = imposible, patrón `0→+20→0→−20`): C-7, C-8B, C-43, J-1, V-2, Z-12, Z-14. Corregidos a 0 en `shared/planilla_mes/planilla_2026-06.xlsx` (backup en `shared/planilla_mes/backups/`). Re-corrida `5_cobranza --force`: **arrastre_devolucion 16→13 excesos** (salieron C-7, C-8B, V-2 — eran exceso falso por el −20). Regresión `5b_validacion` = VALIDACION OK. **B7 sigue vivo arquitectónicamente** (el triple-writer sigue en el código) — solo se limpió el síntoma en el dato de 2026-06.

**FASE 2 — Modelo A completo (DEUDA, sesión dedicada):** 6_corte Y 6b dejan de escribir shared; 5_cobranza `_cargar_penalidades()` overlaya AMBOS audits sumando **deltas** (col 5), no absolutos; ajustar guard DESYNC de `6b/seguimiento_multas` a base+delta. Requiere decidir cómo reconstruir la base viva (hoy shared/audit están desincronizados). Diagramas de flujo (`5_cobranza`, `6_corte`) pendientes de actualizar antes de codificar (regla 4). Doc base: `docs/aprendizaje/writer_unico_desincronizacion_20260701.html`.

---

## Sesión 2026-07-01 (Opus, 6ª del día) — D4 resuelto (negocio) · B7 Fase 1 (diseño) CERRADA

**D4 (decisión de negocio) — RESUELTO.** Reconciliados los 13 excesos (pago − deuda_agua). Decisión del usuario: **no se devuelve nada ahora — todos quedan pendientes en `arrastre_devolucion` hasta que el dueño reclame.** P-6 (S/300) confirmado legítimo (convenio, no sobrepago de agua; falta ubicar dónde se asienta el exceso). El resto (convenio/tanque/reclamo/exoneración/deuda histórica + menores) los rutea el usuario a mano a pagos/convenios/multas después. No se saca ninguna fila del arrastre por ahora.

**Feature nueva (de D4) — columna `REVISION` en `arrastre_devolucion_YYYY-MM.xlsx`** (texto libre, grupo "¿Revisado?"). El usuario anota origen/legitimidad de cada exceso (legítimo vs error del sistema). NO cambia lógica: nada entra/sale del arrastre automáticamente. Requiere las **3 capas de preservación de trabajo manual** (el arrastre se regenera → sin preservación la columna se borra, = bug B5). Análogo cercano ya en `5_cobranza/main.py` (`_recuperar_correcciones_trazabilidad`/`_leer_correcciones_tipeadas`). **Diseño cerrado · falta codificar (Sonnet).**

**Regla nueva de metodología (escrita este turno).** CLAUDE.md Regla 9 + `docs/metodologia_desarrollo.md` 3.6d ampliado + changelog 3.0: agregar una columna humana a un output que un `main.py` regenera obliga a cablear la preservación (3 capas) desde el diseño, no como parche. B5 documentado como caso.

**B7 Fase 1 (diseño Modelo A) — CERRADA. 3 decisiones confirmadas por el usuario:**
1. **Overlay por delta** — 5_cobranza suma `Σ col 5 (PENALIDAD_APLICADA)` de AMBOS audits, NO el absoluto (col 6, queda informativa).
2. **Ambos writers salen** — 6_corte Y 6b dejan de escribir shared (fix parcial reintroduce lost-update en 6b).
3. **Materialización en `planilla_cobrado.xlsx`** (5_cobranza) — `seguimiento` de 6/6b ya lee de ahí, downstream cubierto.

**Insight que de-riesga el refactor (lo que trababa la 5ª sesión):** no hay que "reconstruir la base viva desincronizada". Al dejar 2_planilla como único writer de shared, el `copy2` pasa a ser proyección pura de la base (idempotente); 5_cobranza re-deriva la penalidad del audit en cada lectura → **auto-sana**. Ninguna celda de shared guarda jamás la penalidad → nunca hay revert-sobre-base-reseteada → cero negativos. Es **más simple** que hoy (borra `_modificar_planilla`+`_backup_planilla` de los 2 scripts de penalidad; agrega 1 overlay en 5_cobranza).

**Urgencia confirmada:** junio congelado (sin más corridas) → la curita de Fase 1 aguanta. Ciclo julio re-corre 2_planilla→6_corte → el −20 renace. Urgente ANTES de julio.

**Cambios concretos de Fase 2 (código):**
```
6_corte/aplicar_penalidad.py     → quitar escritura a shared (queda: reconciliar + append audit)
6b/aplicar_penalidad_multas.py   → idem (espejo)
5_cobranza/main.py               → _cargar_penalidades(): net delta por (mz,lt) de los 2 audits;
                                    aplicar a u["corte_reconexion"] tras leer base (l.399, usado en total l.933)
6b/seguimiento_multas.py:284     → guard DESYNC a base+delta (no se ejerce en junio, se corrige de paso)
```

---

## Sesión 2026-07-01 (Opus, 7ª del día) — B7 Fase 2 (Modelo A) CODIFICADA + verificada

**B7 Fase 2 — HECHA.** El triple-writer se eliminó del código:
- `6_corte/aplicar_penalidad.py` y `6b/aplicar_penalidad_multas.py`: se quitaron las escrituras a `CORTE_RECONEXION` en shared (`ws.cell(...).value = nuevo` ×2 + `wb.save`). Ahora solo reconcilian y escriben su audit. Docstrings + prints actualizados.
- `5_cobranza/main.py`: nuevo `_cargar_penalidades(mes_ano)` que suma el net delta (col 5 `PENALIDAD_APLICADA`, con signo) de AMBOS audits; overlay aplicado en `_cargar_planilla` sobre `u["corte_reconexion"]` tras leer la base. Constantes `AUDIT_CORTE_PATH`/`AUDIT_MULTA_PATH`.
- `6b/seguimiento_multas.py`: guard DESYNC (compara absoluto col 6, se pudre bajo overlay) marcado con TODO — **deferido** (necesita separar base/delta; 6b no corre en junio → sale temprano por `if not p.exists()`).

**Verificación (datos reales 2026-06):** overlay aplica 5 predios ×+40 (F1-13, F1-4, S-1, S-5, V-14 → CORTE=40, SALDO+40, la reconexión que la curita perdía); arrastre_devolucion=13 (sin cambio), arrastre_deuda=269 (mismo count), CANCELADO=293 EXCESO=13; 5b_validacion=OK; idempotente (2 corridas). Sin doble-conteo (shared=0 para los 5, verificado antes de codificar). El FAIL `trazabilidad huérfano esperado=6 obtenido=5` en `test_cobranza.py` es **pre-existente** (falla igual contra HEAD, fixture viejo de B2/B5) — no introducido por este cambio.

**Invariante nueva:** shared/planilla_mes = base pura de 2_planilla (nadie más escribe CORTE_RECONEXION). `copy2` de 2_planilla y re-run de 6_corte ya NO pueden reintroducir el −20. Penalidad se re-deriva del audit en cada lectura de 5_cobranza.

**Pendiente (deuda de verificación, NO bloquea junio):**
- `6_corte/seguimiento.py` (consumidor de planilla_cobrado): re-correr cuando el ciclo de corte se reactive; no se re-corrió ahora para no mutar estado cerrado de junio.
- Guard DESYNC de 6b: rediseñar cuando se active 6b.
- Fixture pre-existente `test_cobranza.py` trazabilidad huérfano (esperado=6 obtenido=5): revisar aparte.
- `aplicar_penalidad_multas.py` / `seguimiento_multas.py` de 6b aparecen como untracked en git (no matchea pathspec en stash) — verificar que estén bajo control de versiones.

---

## Sesión 2026-07-02 (Sonnet + Haiku) — columna REVISION + docs B7 sincronizadas

**Columna REVISION en `arrastre_devolucion_YYYY-MM.xlsx`** — HECHA. Col 12, grupo "¿Revisado?" (sep col 11), paleta verde `GH_AV_REVIS` (mismo código que MZ_CORRECTO/LT_CORRECTO, "acá tipeás vos"). 3 capas de preservación (Regla 9), mismo patrón que `1_lecturas/sin_servicio/actualizar_lista.py:_backup_lista` + `_leer_correcciones_tipeadas`:
1. `_backup_arrastre_devolucion(ruta)` — copia el archivo a `5_cobranza/outputs/backups/arrastre_devolucion_{mes_ano}_{timestamp}.xlsx` antes de sobreescribir.
2. `_leer_revision_previa(ruta)` — lee REVISION existente keyed por `(mz, lt)`.
3. Reaplicación en el loop de escritura de `_exportar_arrastre_devolucion`.

**Verificado con datos reales (2026-06):** corrida `--force` → 13 excesos (sin cambio), CANCELADO=293 EXCESO=13. Prueba de preservación: escribí REVISION manual en C1/15 ("legitimo - tanque"), re-corrida → valor preservado exacto; backup nuevo por corrida confirmado. Limpiado el valor de prueba antes de cerrar. `test_cobranza.py`: mismo FAIL pre-existente (trazabilidad huérfano esperado=6 obtenido=5, ajeno a este cambio). `5b_validacion` → VALIDACION OK.

**Docs sincronizadas (Haiku):**
1. `5_cobranza/docs/arrastre_devolucion_diseno.html` — REVISION col 12, 3 capas preservación en instrucciones, ejemplos vacío/lleno, leyenda. Fix de alineación de columnas (faltaba 1 separador entre TRAZ y REFERENCIA en las filas de datos).
2. `6_corte/docs/diagrama_flujo_6_corte.html` — paso 2 (aplicar_penalidad): "registra S/20 en audit_penalidad.xlsx (NO escribe shared)".
3. `5_cobranza/docs/diagrama_5_cobranza.html` — paso "Aplica overlay penalidades", output planilla_cobrado actualizado, nota writer único.

---

## Sesión 2026-07-02 (Opus) — DE2 resuelto · diseño génesis de obligaciones + siembra julio

**DE2 RESUELTO — DIFERIR.** NO aplicar penalidad de multas para junio. La lista (`lista_multas.xlsx`, 91 elegibles) ya generada es lo único que se necesita este ciclo. Penalidades → **próximo mes con las 2 listas juntas (junio + julio)**. M2/S2 quedan diferidos a julio.

**DE3 se reencuadró:** los 16 reclamos `EN_REVISION` de `resolucion_reclamos_2026-06.xlsx` resultaron ser en su mayoría **reasignaciones de lote**, no reclamos puntuales. La resolución de esos reclamos ahora fluye por el diseño de génesis + `DATA_boletas`, no por la regla EN_REVISION de 7_cierre (DE3 original sigue formalmente abierta pero de-priorizada).

**Diseño de fondo cerrado (Fase 1, ~90%) — génesis de obligaciones para ciclo julio:**
- **Problema:** `cangrejo_jun2026.py` es de un solo uso (siembra `planilla_2026-06` desde `DATA_boletas`). Desde julio `main.py` arranca de `1_lecturas` (solo agua) → **sin fuente para multa/acuerdos/convenio**.
- **Hallazgo:** `2_planilla/main.py:99-107` YA lee convenios/multas/acuerdos vía `_load_optional` (si el archivo no existe → 0). El código existe; falta la fuente.
- **Modelo cerrado:** génesis (obligación original, input humano) vs saldo (derivado por 5_cobranza en `arrastre_consolidado`, writer-único). Nunca mantener el saldo a mano. CONCEPTO (DE7) = llave de join. `saldo = génesis − Σpagos[CONCEPTO]`.
- **Junta decidió: mostrar deuda COMPLETA** (no por cuotas) → convenio deja de ser caso especial, modelo uniforme (un solo `MONTO`/saldo por obligación). Parcial descuenta por prioridad **Agua (consumo→mant→corte) → Pueblo (multa→acuerdos→convenio)** — ya coincide con las columnas P1-P5 de `arrastre_consolidado`.
- **Corte es conductual:** pagó CERO 2 meses seguidos ⟹ corte; parcial ⟹ inmune. Mostrar deuda grande no manda a corte.
- **Siembra desde `DATA_boletas`** (post-reclamo, post-reasignación): fiabilidad multa ~94% (35 diffs/575), cuota directa ~97% (15/575); el resto = área común (~9 lotes→0) + reasignación (pares ±) + reclamo real. `DATA_boletas.Multa` y `Cuota directa` vienen AGREGADAS (pierden split reunion/faena, techado/campo) — alcanza para facturar julio; el split es sync de mañana.
- **Reasignación de lote YA capturada** en el diff padron_secundario vs DATA_boletas (casos "dueño distinto en la misma clave"). Guardar como `reasignaciones_lotes_2026-06.xlsx` + commit (audit trazable en git, pedido del usuario).

**Fuentes de datos confirmadas** (en `C:\Users\wilde\Downloads\Base de datos\` y en el repo):
- `padron_secundario.xlsx` = DB maestra con hojas `reunion`/`faena`/`techado y campo`/`medidor`/`instalacion` (cada una con SALDO + LLAVE) — split por concepto, pero PRE-reasignación.
- `DATA_boletas.xlsx` = verdad ACTUAL (post-reclamo) pero agregada. → génesis se siembra de acá.
- medidor vigente = `mayo-planilla · "Cobro medidores"`; inscripción = `SEGUIMIENTO · "NUEVAS INSTALACIONES"` (saldo = TOTAL − Σmeses).

**Tanque queda fuera del génesis de obligaciones** (voluntario, `4_pagos/outputs/aportes_tanque.xlsx`, alimentado por CONCEPTO=tanque). Regla futura: no-aportantes que quieran beneficiarse pagan inscripción alta (~S/2000) = usuario nuevo.

**Preguntas abiertas para mañana (cierran el spec antes de codificar):**
1. ¿Opción **A** (main.py lee saldos multa/acuerdos/convenio desde `arrastre_consolidado_{mes_anterior}`, writer-único, en vivo) — recomendada — o **B** (archivos estáticos en `2_planilla/inputs`, a mano)?
2. ¿Snapshot génesis + `reasignaciones_lotes` se congelan primero (con commit), o se va directo al rewire de main.py?

**Nada de código tocado** (freno de modelo: Opus no codifica). Implementación = Sonnet.

---

## Sesión 2026-07-02 (Opus, 3ª) — arrastre_consolidado CODIFICADO (DE5) + 2_planilla → Opción A

**HECHO Y VERIFICADO (Opus codificó bajo instrucción explícita del usuario, freno saltado a pedido):**

*Confirmaciones del usuario:* Opción A (writer único, no estáticos). Sin problema de bootstrap: junio ya cargó toda la deuda en `planilla_2026-06` (génesis vía cangrejo) y `5_cobranza` ya computa el saldo COMPLETO incluyendo multa/acuerdos/convenio (`main.py` `total` = mes_act+mant+mes_ant+corte+convenio+multa+acuerdos+blanco+devolucion). Por eso `arrastre_consolidado_2026-06` es generable desde el `resultado` de junio.

**Mitad 1 — `5_cobranza/main.py` (genera el archivo):**
- Nuevo `_exportar_arrastre_consolidado(resultado, mes_ano)` — OUTPUT 6. Descompone el saldo por prioridad P1 DEUDA_AGUA → P2 CORTE → P3 MULTA → P4 ACUERDOS → P5 CONVENIO (waterfall del pago sobre componentes), solo filas `TOTAL_ARRASTRE>0`. Schema = `formato_arrastre_consolidado.html`. Salida `outputs/arrastre_consolidado_YYYY-MM.xlsx`.
- `_descomponer_saldo(r)` — helper del waterfall. `_leer_estado_ciclo` / `_ciclo_validado` / `_marcar_generado`. Gate: no emite si `estado_ciclo.json[mes].arrastre.validado` ≠ true. `_marcar_generado` setea `generado=true` (lo que 5b lee para sellar — antes nadie lo seteaba, era manual).
- Call en `main()` tras `_exportar_arrastre_devolucion`. Constantes de paleta `_AC_P` (P1-P5).
- **Verificado:** 269 filas, `suma(P1..P5)==TOTAL_ARRASTRE` en todas, `TOTAL=S/21,876.50==arrastre_deuda` exacto. Componentes: DEUDA_AGUA 3230.50 · CORTE 500 · MULTA 6170 · ACUERDOS 8775 · CONVENIO 3201. 5 predios overlay-penalidad muestran CORTE=40. `5b_validacion` = VALIDACION OK.

**Mitad 2 — `2_planilla` (lee el archivo, Opción A):**
- `config.py`: `COBRANZA_OUTPUTS_DIR`, `consolidado_path(mes)` (lee en vivo de `5_cobranza/outputs`, no copia a inputs), `ESTADO_CICLO_PATH`, `COLS_CONSOLIDADO`. Paths estáticos viejos (CONVENIOS/MULTAS/ACUERDOS_PATH, deuda_path, corte_path) → marcados LEGACY, se conservan solo para tests, ya no se leen.
- `main.py`: `_mes_anterior()`, `_ciclo_validado()`, `_load_consolidado()` (header=1; gate: aborta si el ciclo anterior existe pero no está validado; None si no existe → arrastres=0 genesis). `build_planilla`: 5 `_load_optional` → 1 read del consolidado. Mapeo `DEUDA_AGUA→MES_ANTERIOR · CORTE→CORTE · MULTA→MULTA · ACUERDOS→ACUERDOS · CONVENIO→CONVENIO`. Param `warn` en `_join_optional` (evita 5× el mismo warning).
- `test.py`: fixtures deuda/corte/acuerdos → un `arrastre_consolidado_2026-05` (startrow=1 para header=1) + `estado_ciclo.json` validado.
- **Verificado:** test integración PASÓ (A/12 MES_ANTERIOR=15 · B1/11A CORTE=25 · ACUERDOS=10 · MULTA/CONVENIO=0). Wiring real: `_load_consolidado('2026-07')` lee consolidado_2026-06 (269 filas, validado:true, F1/4 CORTE=40).

**NO tocado:** `shared/planilla_2026-06.xlsx` intacto (2_planilla solo corrió en tmp) → 4_pagos/5_cobranza sin regresión. **2_planilla real julio no corre aún** (falta `lecturas_planilla_2026-07`, blocker D3).

**DEUDA DE DOCS (trigger C / regla 4, saltada a pedido):** actualizar `diagrama_flujo_5_cobranza.html` (OUTPUT 6 nuevo), `diagrama_flujo_2_planilla`/`arquitectura` (input único que reemplaza 5 loads), y README (nuevo output de 5_cobranza). Nada commiteado aún.

---

## GÉNESIS de obligaciones — NECESITA DISEÑO OPUS antes de código (no Sonnet-ready)

**Problema de diseño sin resolver (bloquea codificar génesis):** el consolidado ya carga multa/acuerdos/convenio como SALDO de junio. `DATA_boletas` = "verdad ACTUAL post-reclamo". Si julio siembra desde DATA_boletas Y lo lee del consolidado → **doble conteo**. Hay que decidir cuál manda para esos 3 conceptos.

**Preguntas que Opus debe cerrar (spec de génesis):**
1. **Carry vs re-baseline:** para multa/acuerdos/convenio en julio, ¿manda el arrastre (consolidado_2026-06) o DATA_boletas re-siembra? Si re-siembra, ¿reemplaza el arrastre o solo aporta obligaciones NUEVAS (las que no existían en junio)?
2. **¿`DATA_boletas.Multa`/`Cuota directa` es génesis (original) o saldo (restante)?** Determina si se le restan pagos o no.
3. **Persistencia:** génesis = store append-only (shared, writer único) o snapshot por mes (`2_planilla/inputs`).
4. **¿2_planilla lee 1 fuente o 2?** Si génesis + consolidado son dos inputs, regla de merge por concepto (¿suma? ¿override? ¿solo NUEVAS del génesis, saldo del arrastre?). Hoy 2_planilla lee SOLO el consolidado → si se agrega génesis, es un `_join` extra + regla anti-doble-conteo.
5. **Mapeo hojas DATA_boletas → conceptos** (multa/acuerdos/convenio) sabiendo que vienen AGREGADAS (pierden split reunion/faena, techado/campo). Fuentes ya identificadas: `padron_secundario.xlsx` (hojas reunion/faena/techado y campo/medidor/instalacion, PRE-reasignación) y `DATA_boletas.xlsx` (post-reclamo, agregada). En `C:\Users\wilde\Downloads\Base de datos\`.
6. **`reasignaciones_lotes_2026-06.xlsx`:** cómo derivarlo del diff `padron_secundario` vs `DATA_boletas` (casos "dueño distinto en la misma clave") + commit (audit en git, pedido del usuario).

**Sonnet (SOLO cuando 1-6 estén decididas):** codificar la lectura de génesis en 2_planilla (según la regla de merge que cierre Opus), generar el snapshot génesis desde DATA_boletas, generar `reasignaciones_lotes_2026-06.xlsx` + commit.

---

## SIGUIENTE_ACCION
modelo: **Opus** (diseño — cerrar las 6 preguntas del spec de génesis; ninguna es codificable sin decisión de negocio) → luego Sonnet (código)
sesion: cerrar el spec de génesis de obligaciones para julio. Resolver las 6 preguntas de la sección "GÉNESIS de obligaciones" — la crítica es la #1 (carry vs re-baseline: evitar doble-conteo de multa/acuerdos/convenio entre el consolidado y DATA_boletas). Producir HTML de diseño (trigger C). Recién después: Sonnet codifica la lectura de génesis en 2_planilla + snapshot + reasignaciones.
razon: el arrastre_consolidado (DE5) quedó CODIFICADO y verificado esta sesión — la mitad "saldo/arrastre" del ciclo julio está resuelta. Falta la mitad "génesis" (obligaciones nuevas + verdad post-reclamo de DATA_boletas), que NO es code-ready: tiene un conflicto de doble-conteo con el arrastre que solo Opus/negocio puede zanjar. Codificar antes de decidir #1 arriesga facturar multa/acuerdos/convenio dos veces en julio. 2_planilla real julio además espera `lecturas_planilla_2026-07` (D3, campo).
