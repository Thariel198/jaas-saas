# Decisión de diseño — 4b_reclamos

Fecha: 2026-06-14
Estado: Aprobado en conversación · Fase 2.0.6

---

**Problema:**
Los reclamos detectados en pagos de efectivo (el cobrador escribe "reclamo" en COMENTARIO de `mesa_N.xlsx`) requieren un flujo completo: detectar, clasificar el tipo, mostrar el dato actual a corregir desde `DATA_boletas.xlsx`, registrar la corrección aplicada, y dejar auditoría permanente de cada cambio. Hoy `reclamos.py` vive en `4_pagos/efectivo/` y solo cubre las dos primeras etapas; la corrección sobre `DATA_boletas` no existe.

**Criterios:**
- Auditabilidad — cada reclamo cerrado queda con su corrección registrada en una sola fila (CAMPO, VALOR_ANTERIOR, VALOR_APLICADO, RESOLUCION, FECHA_CIERRE).
- Una historia por reclamo — la trazabilidad responde *"¿qué pasó con este reclamo?"* sin requerir JOIN entre archivos.
- Sin perder trabajo manual — TIPO_RECLAMO, VALOR_A_CORREGIR y RESOLUCION sobreviven cada re-corrida del módulo.
- Cruza límites de módulo correctamente — el módulo lee `3_boletas/inputs/DATA_boletas.xlsx` sin convertirse en sub-módulo de efectivo ni de boletas.
- Idempotente — re-correr el módulo sin cambios en los inputs no modifica los outputs.

**Enfoque elegido:**
Módulo independiente `4b_reclamos`, paralelo en numeración a `5b_validacion` (sub-proceso de su módulo padre, mismo patrón).
- Dos scripts:
  - `main.py` — detecta reclamos desde `pagos_efectivo.xlsx`, preserva trabajo manual, arrastra pendientes, y cierra los marcados RESUELTO/RECHAZADO a trazabilidad (mergeando los datos de corrección de `resolucion_reclamos.xlsx` cuando existen).
  - `resolucion.py` — cruza `reclamos_YYYY-MM.xlsx` (con TIPO_RECLAMO ya clasificado) contra `DATA_boletas.xlsx`, genera `resolucion_reclamos_YYYY-MM.xlsx` con CAMPO y VALOR_ACTUAL auto-poblados.
- Una sola trazabilidad unificada (`trazabilidad_reclamos.xlsx`) con columnas de corrección que quedan vacías si ESTADO_FINAL=RECHAZADO. Sigue el patrón "registro de auditoría" de la metodología 2.6.
- **Un solo archivo para el cierre:** ESTADO + FECHA_RESOLUCION + RESOLUCION viven como master state en `resolucion_reclamos.xlsx`. `main.py` los sincroniza hacia `reclamos.xlsx` automáticamente — el supervisor no edita el mismo dato en dos archivos. La regla: si en resolucion hay un valor no-vacío, sobrescribe lo de reclamos; si está vacío, conserva lo de reclamos (no degrada RESUELTO a PENDIENTE).
- `DATA_boletas.xlsx` se lee directo desde `3_boletas/inputs/` — sin copia a `inputs/`, porque el supervisor necesita ver el VALOR_ACTUAL vigente al momento de corregir, no un snapshot.
- `reclamos.py` se migra desde `4_pagos/efectivo/` a `4b_reclamos/main.py` junto con sus 3 HTMLs de diseño (`diagrama_reclamos.html`, `formato_reclamos.html`, `formato_trazabilidad_reclamos.html`), el output `reclamos_2026-06.xlsx`, y la trazabilidad existente.

**Alternativas descartadas:**
- *Dejar reclamos en `4_pagos/efectivo/`* — el script ya cruza límites (lee `3_boletas/DATA_boletas.xlsx`), no es responsabilidad de efectivo. El cross-check de mesas y el flujo de reclamos son ciclos de vida distintos.
- *Dos trazabilidades separadas (`trazabilidad_reclamos` + `trazabilidad_correcciones`)* — violaría "un archivo = una historia" de la metodología 2.6. Requeriría JOIN para responder *"¿qué pasó con este reclamo?"*.
- *Un solo script con fases* — ambigüedad sobre en qué fase está el ciclo (¿detección nueva o cierre de resoluciones?), más difícil de mantener y testear.
- *Copiar `DATA_boletas.xlsx` a `inputs/`* — el snapshot puede quedar desactualizado mid-mes si `3_boletas` regenera; queremos el valor vigente al momento del cruce.

**Señal de alerta:**
La trazabilidad crece mes a mes pero la vista `reclamos_YYYY-MM.xlsx` no decrece → nada se está cerrando. Diagnóstico en orden: (1) revisar columna ESTADO en `reclamos_YYYY-MM.xlsx` — ¿el supervisor está marcando RESUELTO/RECHAZADO?; (2) revisar `run.log` de `resolucion.py` por warnings de MZ-LT no encontrado en `DATA_boletas` — si falla el cruce, el supervisor no puede cerrar correcciones; (3) revisar que el merge de corrección en trazabilidad esté usando la clave `(MESA, MZ, LT, FECHA_COBRO)` y no esté perdiendo filas.
