# Registro de decisiones — 2_planilla

Sigue la Fase 2.0.6 de la metodología: documentar enfoque elegido + alternativas descartadas con razón.

---

## Decisión 1 — Enfoque único: un script, todos los inputs de una vez

**Problema:** `2_planilla` necesita juntar 5 fuentes distintas (lecturas, arrastre de deuda,
arrastre de corte, convenios, multas) para generar la planilla mensual. Los módulos que
generarán esas fuentes automáticamente (5b_validacion, 6_corte) aún no existen.

**Criterios:**
- Corre aunque los módulos fuente no existan aún (inputs creados manualmente por ahora)
- La planilla generada es la entrada directa de 4_pagos — sin transformaciones intermedias
- Separación Yape/Efectivo en la estructura desde el primer día
- Total por usuario auditable fila por fila sin abrir otros archivos
- Idempotente: mismos inputs → mismo output exacto

**Enfoque elegido:** un solo script que lee todos los inputs, hace join por (MZ, LT),
calcula en una pasada y genera la planilla. Si un archivo de input no existe → ese campo
queda en 0 y el log avisa. En el futuro los módulos correspondientes generarán esos archivos
automáticamente sin cambiar el código de 2_planilla.

**Alternativas descartadas:**
- *Validación estricta (falla si falta un input)* — útil cuando todos los módulos existan;
  hoy bloquea el trabajo porque 04b y 6_corte aún no están construidos. Descartado.
- *Pipeline por pasos (un script por fuente)* — permite correr parcialmente pero rompe
  la idempotencia y complica la orquestación. Viola el criterio de entrada directa a 4_pagos. Descartado.

**Señal de alerta:** si los campos que hoy llegan en 0 (arrastre, corte) siguen en 0
cuando los módulos fuente ya existan, el join por (MZ, LT) no está funcionando — revisar
normalización de claves.

---

## Decisión 2 — Dos columnas separadas: MONTO_YAPE y MONTO_EFECTIVO

**Problema:** la planilla anterior tenía un solo campo de pago con código de medio ('Y'=Yape,
'E'=Efectivo). Si un usuario pagaba mitad Yape mitad efectivo y se marcaba 'Y', todo el
pago quedaba registrado como Yape — generando excesos en Yape y déficits en efectivo que
había que corregir manualmente mes a mes.

**Criterios:**
- Que 4_pagos pueda asignar exactamente a cada medio sin generar excesos
- Auditable: el total pagado = MONTO_YAPE + MONTO_EFECTIVO
- Compatible con el motor de matching de 4_pagos/yape

**Enfoque elegido:** dos columnas separadas en la planilla: `MONTO_YAPE` y `MONTO_EFECTIVO`.
El total cobrado = suma de ambas. 4_pagos escribe en cada columna según el medio que matcheó.

**Alternativas descartadas:**
- *Una columna de monto + columna de medio* — no permite pagos mixtos sin perder información.
  Era el esquema anterior. Descartado.
- *Una columna por medio de pago (Yape, Efectivo, Transferencia...)* — escala mal si se agregan
  medios nuevos. Por ahora solo existen Yape y Efectivo. Descartado.

**Señal de alerta:** si aparece un tercer medio de pago (transferencia, POS), agregar columna
`Monto_Transferencia` siguiendo el mismo patrón.

---

## Historial de cambios

| Fecha | Cambio |
|---|---|
| 2026-06-06 | Versión inicial — decisiones del enfoque de integración y separación Yape/Efectivo |
