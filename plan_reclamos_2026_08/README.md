# Centro de control — reclamos agosto/septiembre 2026

Esta carpeta contiene el plan operativo del mes y los diagramas visuales del sistema.
No es un modulo productivo y no debe contener datos vivos ni reemplazar los README de
cada modulo.

**Importante:** los diagramas distinguen el flujo que corre hoy del diseño futuro.
Actualmente el registro vivo de deuda de pueblo es `shared/seguimiento_pueblo.xlsx`,
escrito mediante `shared/seguimiento_repo.py`. `libro_mayor/` todavía no participa del
flujo ejecutable; sus READMEs describen una migración futura.

```text
plan_reclamos_2026_08/
    README.md
    diagrama_flujo_sistema.html
    diagrama_flujo_plan_reclamos.html
```

## Objetivo

```text
pagos y precursores correctos
        -> lista de corte de agosto publicada
        -> agosto cerrado
        -> reimputacion validada
        -> exoneraciones auditadas
        -> boletas de septiembre correctas
```

## Estado al 2026-08-15

- [ ] Incorporar los 18 abonos seguros confirmados: 6 de julio y 12 confirmados ayer para la corrida de agosto.
- [ ] Dejar pendiente la integración de 13 pagos/manuales ya presentes en el ledger que todavía no llegan a la lista de corte.
- [ ] Resolver colisiones y cambios que puedan sacar predios de corte.
- [ ] Ejecutar `4_pagos` final de agosto.
- [ ] Ejecutar `5_cobranza --force`.
- [ ] Ejecutar `5b_validacion` con resultado OK.
- [ ] Generar, revisar y publicar `6_corte/outputs/lista_corte.xlsx`.
- [ ] Cerrar y congelar agosto.
- [ ] Recalcular la reimputacion con saldos post-cierre.
- [ ] Escribir y validar la reimputacion del ledger.
- [ ] Exonerar multas con evidencia independiente.
- [ ] Generar y validar las boletas de septiembre.

## Archivos de referencia

- Plan detallado: `../README_PLAN_RECLAMOS_2026-08.md`
- Restricciones activas: `../LEER_ANTES.md`
- Decisiones de reimputacion: `../docs/decisiones/reimputacion_cascada_ca1.md`
- Ultimo estado de abonos: `../docs/retomar/RETOMAR_ABONOS_REZAGADOS_2026-08-14.md`

## Diagramas

- [Flujo completo del sistema](diagrama_flujo_sistema.html)
- [Flujo operativo de este problema](diagrama_flujo_plan_reclamos.html)

## Reglas de lectura

- Verde: resultado validado o salida operativa.
- Azul: entrada, fuente o archivo de referencia.
- Morado: proceso o motor de reglas.
- Ambar: decision, gate o revision humana.
- Rojo: bloqueo o riesgo.
- El ledger es append-only: corregir significa agregar eventos inversos, no borrar.
- En el flujo actual, ese ledger es `shared/seguimiento_pueblo.xlsx`; no `libro_mayor/`.
- `libro_mayor/` aparece solo como arquitectura futura y no como dependencia de ejecución.
- La lista de corte se calcula antes de la reimputacion historica.
- Las exoneraciones se ejecutan despues de la reimputacion.
- Las boletas de agosto no se regeneran dentro de este plan.

## La zona confusa: 4_pagos -> precursores -> 5_cobranza

```text
4_pagos
  | pagos de agua identificados
  | pagos con CONCEPTO != vacio salen del agua
  v
5_cobranza carga planilla + pagos + overlays
  | calcula todo en memoria
  | no escribe cada overlay como pago nuevo
  v
planilla_cobrado + arrastres + reconciliacion
  |                         |
  |                         v
  |                 shared/seguimiento_pueblo.xlsx
  v
arrastre_devolucion
```

### Entradas que lee `5_cobranza`

| Grupo | Archivos reales | Efecto |
|---|---|---|
| Planilla | `shared/planilla_mes/planilla_YYYY-MM.xlsx` | Cargos base de agua, corte, convenio, acuerdos y multa |
| Pagos Yape | `4_pagos/yape/motor_matching/outputs/pagos_yape_tepago_YYYY-MM.xlsx` | Pago de agua si `CONCEPTO` está vacío; con concepto sale del agua |
| Pagos efectivo | `4_pagos/efectivo/outputs/pagos_efectivo_YYYY-MM.xlsx` | Pago de agua si `CONCEPTO` está vacío; con concepto sale del agua |
| Retornos/devoluciones | `pagos_yape_retorno`, `pagos_efectivo_devolucion`, `pagos_yape_devolucion` | Reducen el Yape neto; no son un nuevo pago |
| Penalidades | `6_corte/outputs/audit_penalidad.xlsx` y audit de 6b | Overlay de `CORTE_RECONEXION` por delta |
| Correcciones de exceso | `shared/devoluciones_aplicadas.xlsx` | Baja un concepto una sola vez en el ciclo indicado |
| Correcciones de cargo | `shared/ajustes_cargo.xlsx` | Anula un cargo que no correspondía; no representa plata |
| Genesis tardia | `shared/genesis_tardia.xlsx` | Agrega un cargo legítimo que llegó tarde |
| Reidentificacion | `shared/reidentificacion.xlsx` | Acredita un pago al predio correcto |
| Deuda del origen | `shared/deuda_correcciones.xlsx` | Devuelve deuda al predio de origen de una reidentificación |
| Abonos rezagados | `shared/abonos_rezagados.xlsx` | Agrega plata real fuera de la ventana; separa ciclo cerrado/vigente |
| Blancos efectivo | `shared/blancos_efectivo.xlsx` | Acredita efectivo que existía pero no tenía MZ/LT |
| Reasignacion de aplicacion | `shared/reasignaciones_aplicacion.xlsx` | Cambia el concepto que recibe el pago; no cambia el cargo |
| Tanque | `shared/aportes_tanque_manuales.xlsx` | Resta aporte voluntario antes de la cascada de deuda |

### Qué significa cada salida

- `planilla_cobrado.xlsx`: foto calculada del ciclo, con pagos, `SALDO` y estado.
- `arrastre_deuda_YYYY-MM.xlsx`: saldo positivo pendiente para el siguiente ciclo.
- `arrastre_consolidado_YYYY-MM.xlsx`: saldo positivo descompuesto por P1 agua, P2 corte, P3 convenio, P4 acuerdos y P5 multa; lo consume `2_planilla` cuando el ciclo está validado.
- `arrastre_devolucion_YYYY-MM.xlsx`: dos universos distintos: predios con `SALDO < 0` y pagos cuyo MZ/LT no existe en la planilla. El segundo es un huérfano, no un saldo negativo de un predio.
- `shared/seguimiento_pueblo.xlsx`: eventos de `MULTA`, `ACUERDOS` y `CONVENIO`; no guarda agua ni corte.

### Qué ocurre en `seguimiento_pueblo`

`5_cobranza` no reemplaza totales. Recalcula el mes entero y compara:

```text
SET_DEBE = pago normal y abono que la corrida acaba de calcular
SET_TIENE = PAGO + AJUSTE que ya escribió SOURCE=5_cobranza
DELTA = SET_DEBE - SET_TIENE
```

- `DELTA > 0`: escribe un `PAGO`.
- `DELTA < 0`: escribe un `AJUSTE` `CORRECCION_SISTEMA`.
- El mismo evento no duplica si conserva `SOURCE + AUDIT_REF + MZ + LT + CONCEPTO`.
- Los pagos normales se escriben como `SOURCE=5_cobranza`, `CLASE=COBRANZA`.
- Los abonos se escriben como `SOURCE=abonos_rezagados`, `CLASE=ABONO_REZAGADO`.
- Los registros manuales existentes no se vuelven automáticamente pagos de `5_cobranza`.
- `DECLARACION`, `DECLARACION_SECRETARIA`, `EXONERACION`, `CORRECCION_SISTEMA` y `REASIGNACION` explican hechos distintos; solo `COBRANZA` y `ABONO_REZAGADO` suman caja.

### Auditoría puntual del archivo actual

Snapshot leído el 2026-08-15, antes de una nueva corrida:

| Campo | Conteo |
|---|---:|
| Eventos totales | 1.569 |
| `CARGO` | 755 |
| `PAGO` | 601 |
| `AJUSTE` | 213 |
| `SOURCE=5_cobranza` | 625 |
| `SOURCE=manual` | 75 |
| `SOURCE=correccion_genesis_formula` | 109 |
| `SOURCE=genesis_tardia` | 42 |
| `AJUSTE` sin `MOTIVO` | 209 |

Los 209 ajustes sin motivo son una cola de auditoría: pueden mover saldo, pero no explican por sí solos por qué se movió. No se deben borrar ni reinterpretar sin revisar su `SOURCE`, `AUDIT_REF`, predio, concepto y mes.

Los 12 abonos confirmados ayer están listados individualmente en `../README_PLAN_RECLAMOS_2026-08.md`, con total `S/457`. Son parte de la corrida de agosto; no deben volver a clasificarse como casos históricos pendientes.
