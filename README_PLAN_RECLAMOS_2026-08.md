# Plan para Cerrar Reclamos — Agosto/Septiembre 2026

**Fecha de control:** 2026-08-15  
**Objetivo:** publicar una lista de corte correcta en agosto y entregar boletas de septiembre corregidas, trazables y defendibles ante reclamos.

```text
PRECURSORES Y PAGOS
        |
        v
4_pagos cerrado sin pendientes
        |
        v
5_cobranza --force
        |
        v
5b_validacion = OK
        |
        v
6_corte: BORRADOR -> PUBLICADA
        |
        v
CIERRE DE AGOSTO
        |
        v
REIMPUTACION DEL LEDGER
        |
        v
EXONERACIONES CON RESPALDO
        |
        v
PLANILLA Y BOLETAS DE SEPTIEMBRE
```

## 1. Resultado que debe quedar

- [ ] Lista de corte de agosto publicada con los pagos y abonos rezagados considerados.
- [ ] Ningún usuario pagado permanece en la lista por un dato atrasado o mal atribuido.
- [ ] Ningún usuario que cumple el criterio vigente de corte queda fuera sin motivo auditado.
- [ ] Agosto cerrado y congelado antes de modificar históricamente la imputación del ledger.
- [ ] Reimputación validada sin crear ni destruir dinero ni deuda.
- [ ] Multas exoneradas solo con respaldo de asistencia o decisión documentada.
- [ ] Boletas de septiembre calculadas desde el estado corregido, no desde parches manuales.
- [ ] Cada reclamo puede responderse con: pago, fecha, concepto, aplicación, saldo y motivo de ajuste.

## 2. Estado actual

### Hecho y verificado

- [x] La cascada de pagos fue diseñada para el orden `CONVENIO -> ACUERDOS -> MULTA`.
- [x] El reporte de reimputación fue corregido para mover solo plata real y limitar cada movimiento por concepto.
- [x] La simulación tiene validaciones de conservación y no escribe el ledger.
- [x] El código de `5_cobranza` ya usa el orden CA1 para agosto.
- [x] Existe `LEER_ANTES.md` con el orden obligatorio y las protecciones del ciclo.
- [x] Se restauró el ledger al checkpoint limpio de abonos antes de la nueva corrida.
- [x] `5_cobranza/main.py` compila.
- [x] La fuente de abonos rezagados está centralizada en `shared/abonos_rezagados.xlsx`.
- [x] Los abonos normales y rezagados fueron separados conceptualmente (`COBRANZA` vs `ABONO_REZAGADO`).
- [x] El overlay de penalidades evita que `6_corte` vuelva a escribir la planilla compartida.
- [x] Las boletas de agosto ya fueron generadas y validadas; no deben regenerarse por este plan.

### Pendiente de ejecución o decisión

- [ ] Incorporar los 18 abonos seguros confirmados: 6 de julio y los 12 confirmados ayer para la corrida de agosto; confirmar sus eventos en el ledger.
- [ ] Pendiente posterior: integrar los 13 eventos ya resueltos manualmente para que la lista de corte reconozca esos pagos; no trabajarlos en esta corrida.
- [ ] Mantener sin escribir los cuatro casos bloqueados: `S-5`, `D-16`, `D1-6`, `Q-11`.
- [ ] Resolver la colisión del aporte de tanque `K-2/C1-2`.
- [ ] Ejecutar el fix de matching de `K-3/K-4` y verificar sus pagos separados.
- [ ] Eliminar o corregir la fila duplicada de `E-14B` antes de otra corrida forzada.
- [ ] Resolver pendientes con capacidad de cambiar saldos: `S-16`, `X-11`, `F-4/F1-4` y excesos sin explicación.
- [ ] Ejecutar `4_pagos` completo con los datos finales de agosto.
- [ ] Ejecutar `5_cobranza --force` desde el checkpoint correcto.
- [ ] Ejecutar `5b_validacion` y obtener cero alertas nuevas.
- [ ] Generar y revisar la lista de corte de agosto.
- [ ] Publicar la lista y congelar su versión publicada.
- [ ] Cerrar agosto con `7_cierre/consolidar_cierre.py`: primero dry-run, luego `--confirmar` y consentimiento `SI`.
- [ ] Regenerar la simulación de reimputación usando saldos post-cierre.
- [ ] Escribir y validar los asientos de reimputación.
- [ ] Diseñar/ejecutar el lote de exoneraciones con evidencia independiente.
- [ ] Generar planilla y boletas de septiembre.

### Los 12 abonos confirmados ayer

Estos registros ya se consideran seguros para la corrida de agosto. No son los cuatro casos bloqueados (`S-5`, `D-16`, `D1-6`, `Q-11`) ni reemplazan los seis abonos históricos de julio sin colisión.

| MZ | LT | MONTO | CICLO_ORIGEN | APLICA | RETENIDO_POR |
|---|---:|---:|---|---|---|
| D | 18 | S/32 | 2026-08 | 2026-08 | Yerald Romero |
| K | 12 | S/122 | 2026-08 | 2026-08 | Yerald Romero |
| X | 33 | S/89 | 2026-08 | 2026-08 | Yerald Romero |
| C | 3A | S/10 | 2026-08 | 2026-08 | Maximo Encarnacion |
| F1 | 8 | S/13 | 2026-08 | 2026-08 | Wagner Trujillo |
| P | 3 | S/33 | 2026-07 | 2026-08 | Yerald Romero |
| W | 5 | S/15 | 2026-07 | 2026-08 | Wagner Trujillo |
| F1 | 8 | S/14 | 2026-07 | 2026-08 | Wagner Trujillo |
| H1 | 15 | S/36 | 2026-07 | 2026-08 | Wagner Trujillo |
| E1 | 9 | S/3 | 2026-07 | 2026-08 | Wagner Trujillo |
| S | 2 | S/60 | 2026-07 | 2026-08 | Wagner Trujillo |
| P | 12 | S/30 | 2026-07 | 2026-08 | Wagner Trujillo |

**Total confirmado:** 12 filas, `S/457`. Cada fila debe conservar su propio `SOURCE=abonos_rezagados` y `AUDIT_REF`; no fusionar los dos abonos de `F1-8`.

## 3. Fase A — Preparar la lista de corte

### A0. Congelar el punto de partida

- [ ] Confirmar que el ciclo activo es `2026-08` en `shared/ciclo_activo.json`.
- [ ] Crear un backup identificable de los archivos que serán modificados.
- [ ] Registrar en este documento la fecha/hora de inicio de la corrida.
- [ ] No ejecutar reimputación histórica ni exoneraciones durante esta fase.

### A1. Abonos rezagados

- [ ] Mantener separados los 6 abonos de julio sin colisión y los 12 abonos de agosto confirmados ayer.
- [ ] Confirmar que cada uno tiene respaldo de mesa, Yape o declaración válida.
- [ ] Confirmar que el pago normal y el abono tienen `SOURCE` y `AUDIT_REF` distintos.
- [ ] Verificar que ningún caso de `Pendiente/abonos_rezagados_pendientes_2026-07.md` se escriba prematuramente; los 12 de agosto confirmados sí entran en la corrida actual.
- [ ] Recalcular al cierre el saldo disponible de `S-5`, `D-16`, `D1-6` y `Q-11`.

### A2. Otros precursores que pueden sacar usuarios del corte

Antes de generar la lista, revisar cualquier cambio que afecte el saldo de agua o la antigüedad:

- [ ] Pagos Yape nuevos o previamente no identificados.
- [ ] Pagos de efectivo y discrepancias de mesa.
- [ ] Abonos rezagados con respaldo.
- [ ] Correcciones de lote o reidentificaciones confirmadas.
- [ ] Reclamos ya resueltos que modifiquen el saldo.
- [ ] Correcciones de padrón/lecturas que cambien el universo o `MES_ANTERIOR`.
- [ ] Aportes de tanque y conceptos no-agua correctamente excluidos del saldo de agua.
- [ ] Duplicados de predio o planilla eliminados antes de calcular la deuda.

**Regla:** una observación verbal no cambia la lista; debe existir un precursor, fuente o decisión auditada.

### A3. Ejecutar la cadena de agosto

```text
datos finales de pagos
        |
        v
4_pagos
        |
        +--> Yape sin identificar = 0 o casos explícitamente aprobados
        +--> efectivo con discrepancias = 0
        |
        v
5_cobranza --force
        |
        v
5b_validacion
        |
        v
6_corte/generar_lista.py
```

- [ ] Ejecutar `4_pagos` y guardar el log.
- [ ] Confirmar que no quedan pendientes de Yape ni discrepancias de efectivo.
- [ ] Ejecutar `py 5_cobranza/main.py --force`.
- [ ] Confirmar que la corrida lee los archivos de `2026-08`, no restos de junio/julio.
- [ ] Confirmar que los pagos normales quedan como `COBRANZA`.
- [ ] Confirmar que los abonos quedan como `ABONO_REZAGADO`.
- [ ] Confirmar `SALDO` no negativo y trazabilidad sin duplicados nuevos.
- [ ] Ejecutar `5b_validacion`.
- [ ] Detenerse si la validación no es `OK`.

### A4. Revisar y publicar

- [ ] Generar `6_corte/outputs/lista_corte.xlsx`.
- [ ] Comparar la lista contra la versión anterior y explicar cada alta y baja.
- [ ] Revisar especialmente usuarios con pagos recientes, abonos y reclamos.
- [ ] Confirmar que el corte usa únicamente deuda de agua y el criterio vigente.
- [ ] Confirmar que la reimputación de multa/acuerdos/convenio no altera esta lista.
- [ ] Guardar una copia de la lista revisada como versión de publicación.
- [ ] Cambiar el estado a `PUBLICADA`.
- [ ] Registrar fecha, responsable y hash/nombre del archivo publicado.
- [ ] No regenerar silenciosamente después de publicar; cualquier cambio requiere nueva revisión.

## 4. Fase B — Cerrar agosto

- [ ] Terminar la ventana operativa de la lista de corte.
- [ ] Procesar los pagos de la ventana de gracia según el procedimiento de `6_corte`.
- [ ] Generar salvados, cortes físicos y arrastre de corte.
- [ ] Ejecutar la validación final del ciclo.
- [ ] Confirmar que todos los outputs canónicos de agosto están identificados.
- [ ] Crear la foto inmutable de agosto.
- [ ] Marcar agosto como `CERRADO`.
- [ ] Resetear únicamente los slots temporales del siguiente ciclo.
- [ ] Cambiar la lista a `COMPROMETIDA` solo después de que no se acepten cambios ordinarios.

**Riesgo actual:** `7_cierre/consolidar_cierre.py` todavía no existe. Hasta implementarlo, el cierre debe hacerse con una checklist manual equivalente y backups verificables; no se debe afirmar que agosto está cerrado solo porque terminó la corrida.

## 5. Fase C — Reimputar el ledger

La reimputación no es una devolución ni una exoneración. Es un ajuste de concepto:

```text
plata ya registrada
        |
        v
MULTA --> ACUERDOS --> CONVENIO   (orden viejo)
        |
        v
CONVENIO --> ACUERDOS --> MULTA   (orden nuevo)
```

- [ ] Re-correr el reporte después del cierre de agosto.
- [ ] No reutilizar ciegamente las cifras de simulaciones anteriores.
- [ ] Congelar el contrato post-cierre.
- [ ] Escribir todos los asientos con `MES=2026-08`.
- [ ] Escribir un precursor por movimiento en `shared/reasignaciones_aplicacion.xlsx`.
- [ ] Escribir el par de asientos: sube deuda del concepto que cede y baja deuda del concepto que recibe.
- [ ] Usar `CLASE=REASIGNACION`, no `COBRANZA` ni `ABONO_REZAGADO`.
- [ ] No mover pagos pre-génesis.
- [ ] No mover créditos de `EXONERACION` ni `CORRECCION_SISTEMA`.
- [ ] No mover deuda de `INSTALACION` ni `REACTIVACION` al convenio de medidor.
- [ ] Verificar ambos lados del movimiento: concepto que cede y concepto que recibe.
- [ ] Ejecutar la herramienta de reversión por predio antes de cerrar el lote.

### Validaciones obligatorias

- [ ] Plata real total antes = después.
- [ ] Deuda total antes = después.
- [ ] Deuda por predio antes = después.
- [ ] Deuda por concepto antes/después coincide con el contrato.
- [ ] Ningún saldo negativo nuevo.
- [ ] Ningún movimiento supera la plata real del concepto que cede.
- [ ] Cero movimientos sin mes de origen.
- [ ] `AUDIT_REF` determinista e idempotente.
- [ ] Cero asientos de reimputación duplicados.

## 6. Fase D — Exonerar multas

- [ ] Separar candidatos por evidencia: asistencia, decisión institucional o reclamo.
- [ ] No exonerar porque un predio apareció en un reporte de auditoría.
- [ ] Mantener las exoneraciones históricas que ya tienen `MOTIVO` válido.
- [ ] Exonerar solo `MULTA`, salvo decisión explícita documentada para `ACUERDOS`.
- [ ] Registrar motivo, fuente, fecha, responsable y período.
- [ ] Aplicar la exoneración después de la reimputación.
- [ ] Validar que la exoneración no se interprete como plata cobrada.
- [ ] Regenerar el estado de cuenta y verificar que el saldo total baje exactamente por el crédito autorizado.

## 7. Fase E — Septiembre y boletas sin reclamos

```text
ledger post-reimputación + exoneraciones
        |
        v
2_planilla septiembre
        |
        v
3_boletas/enriquecimiento
        |
        v
validación de datos
        |
        v
3_boletas/main.py
        |
        v
validación PDF + entrega
```

- [ ] Generar la planilla de septiembre desde el estado corregido.
- [ ] Confirmar que la planilla no cobra dos veces el convenio.
- [ ] Confirmar que la multa exonerada aparece en cero.
- [ ] Confirmar que acuerdos y convenio quedan en el concepto correcto.
- [ ] Validar cada predio contra el ledger, no solo totales agregados.
- [ ] Ejecutar el enriquecimiento de `3_boletas`.
- [ ] Validar nombres, MZ/LT, saldo, conceptos y número de recibo.
- [ ] Generar las boletas.
- [ ] Revisar el PDF consolidado y una muestra de casos críticos.
- [ ] Entregar las boletas y conservar el lote final como evidencia.

## 8. Muestra mínima de control

Antes de entregar septiembre, revisar manualmente al menos:

- [ ] Un predio limpio de reimputación.
- [ ] Un predio `AMPLIADO`.
- [ ] Un predio `RECORTADO`.
- [ ] Un predio `EXCLUIDO` por exoneración.
- [ ] Un predio con abono rezagado.
- [ ] Un predio con pago Yape.
- [ ] Un predio con pago efectivo.
- [ ] Un predio con convenio de medidor.
- [ ] Un predio con acuerdos de techado/campo.
- [ ] Un predio con multa exonerada.
- [ ] Un predio sin servicio o sin lectura.
- [ ] Un predio con reclamo corregido.

## 9. Registro de corridas

| Fecha/hora | Ciclo | Comando/script | Resultado | Responsable | Observaciones |
|---|---|---|---|---|---|
| 2026-08-15 | 2026-08 | — | PENDIENTE | — | Plan creado; no se ha publicado la lista |
| 2026-08-15 08:18 | 2026-08 | `py 4_pagos/main.py` | BLOQUEADO | — | Efectivo OK; Yape V1/V2/V4 fallan: banco S/3721 vs salida S/3671, 1 pago sin pareja y 2 filas extra. No ejecutar `5_cobranza` hasta resolver. |
| 2026-08-15 08:31–08:49 | 2026-08 | `py 5_cobranza/main.py --force` | COMPLETADO | — | 559 usuarios; 273 con arrastre de deuda; 9 excesos + 1 no identificado; 61 elegibles para corte; 28 pagos de pueblo y 0 ajustes. |
| 2026-08-15 08:49 | 2026-08 | `py 5b_validacion/main.py` | ALERTA | — | TE PAGÓ +S/700 por tanque incluido; efectivo -S/22 en `S-LT16` (`mesa_1_Wilder Trujillo`). No avanzar a `6_corte`. |
| 2026-08-15 09:17 | 2026-08 | `py 5_cobranza/main.py --force` | BLOQUEADO POR GUARD | — | No escribió datos. Detectó `Q-5 S/45`, ciclo `2026-07`, aplica `2026-08`, fuera del manifest; pendiente de decisión específica. |

## 10. Criterio de cierre del plan

```text
lista agosto PUBLICADA
        AND
agosto CERRADO
        AND
reimputacion VALIDADA
        AND
exoneraciones AUDITADAS
        AND
boletas septiembre VALIDAS
        =
objetivo cumplido
```

Mientras una condición esté pendiente, el problema no está cerrado aunque las boletas se hayan generado.
