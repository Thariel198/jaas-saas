# Estado de decisiones

Este archivo es el registro corto que usa el agente para distinguir reglas
CERRADAS de trabajo ABIERTA. Una decision solo pasa a CERRADA con confirmacion
explicita del usuario.

## CERRADAS

| ID | Regla | Evidencia de cierre |
|---|---|---|
| D-001 | La referencia de pago confirma si el pago existe. Sin referencia no se reparte en el historial. | Confirmacion explicita del usuario: "Ok" en sesion 2026-08-17 |
| D-002 | Un `SOURCE` tecnico del ledger no es evidencia de pago. | Confirmacion explicita del usuario: "Ok" en sesion 2026-08-17 |
| D-003 | Los reportes usan planillas historicas hasta 2026-05 y el ledger desde 2026-06; esta frontera no se reinterpreta ni autoriza cambios de codigo. | Confirmacion explicita del usuario en sesion 2026-08-17 |
| D-004 | El codigo del reporte validado de siete lotes queda protegido por `guard_codigo.json` en estado `CERRADO`; el agente no puede modificarlo ni desbloquearlo. Solo un cambio manual a `ABIERTO` permite editarlo. | Confirmacion explicita del usuario en sesion 2026-08-17 |

| D-005 | `estado_cuenta` empieza en agosto de 2026 sin backfill: agua y corte se abren con el arrastre cerrado de julio; los cargos nuevos salen de lecturas/planilla y `6_corte`; `7_cierre` es el único commit oficial. | Confirmación explícita del usuario, 2026-08-21 |

## ABIERTAS

| ID | Tema | Siguiente accion |
|---|---|---|
| — | No hay decisiones nuevas registradas. | — |

## BLOQUEADAS

| ID | Motivo | Que falta |
|---|---|---|
| — | No hay decisiones bloqueadas registradas. | — |
