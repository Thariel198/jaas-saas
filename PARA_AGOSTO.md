# PARA AGOSTO — condonar multas/corte a estos usuarios

Pendiente de la reconciliación junio→julio (sesión 2026-07-18/19, ver
`jass_system - junio\docs\RETOMAR_junio_cierre_reconciliacion_2026-07-18.md`
y `jass_system - junio\4_pagos\efectivo\outputs\reconciliacion_junio_a_julio.xlsx`).
Los pagos ya se cargaron en julio (`4_pagos/efectivo/inputs/mesa_5.xlsx`), pero
estos 4 ajustes de MULTA/CORTE **todavía no se ejecutaron** — quedan para agosto.

**No son bugs ni correcciones de sistema — son condonaciones.** El sistema
generó la MULTA/CORTE correctamente según sus reglas (falla de faena, 2 meses
sin pago detectado); la directiva decidió, caso por caso, perdonarla por una
razón puntual (no es una regla general que se automatiza). Quedan como
decisión humana registrada acá, no como fix de código.

## Condonaciones pendientes

| Lote | Usuario | Condonar | Motivo de la condonación | Decidido por |
|---|---|---|---|---|
| S-5 | Valerio Porfilio Javier Santiago | CORTE (40) + MULTA (30) | El corte se disparó porque el sistema vio junio sin pago (el yape nunca lo remitió el cobrador) como si hubiera fallado 2 meses seguidos — pero sí pagó. `215-40-30=145=71(jun)+74(jul)`, calza exacto: no fue falta del usuario. | Usuario, sesión 2026-07-18 |
| D-16 | Esteban Guerrero Chingel | MULTA (50) | La directiva falla en registrar asistencia a faena; el usuario sí asiste. `169-50=119=85(jun)+34(jul)`, calza exacto. Convenio/acuerdos NO se condonan, son deuda real. | Usuario, sesión 2026-07-18 |
| C1-17 | Macarlopu | MULTA (30) | Dice haber asistido a faena, la directiva no lo registró (mismo criterio que D-16). | Usuario, sesión 2026-07-18 — ✅ EJECUTADO 28/07/2026 (ver `LEER_ANTES.md`, fila C1-17 eliminada de `DATA_boletas.xlsx` y `padron_reconciliado.xlsx`) |
| G-12 | Hernestina Valladares | MULTA (30) | Donó troncos para la pachamanca del pueblo — se decidió liberarla de la falta de faena a cambio de esa donación. | Usuario, sesión 2026-07-18 |

Ninguna de estas 4 es transferible a otro caso parecido sin que la directiva
lo decida de nuevo — cada una es puntual (no crear una regla automática de
"perdonar si dice que asistió").

## Cómo ejecutarlo

No hay mecanismo de código para esto todavía (no confundir con el override de
`5_cobranza`, que no alcanza para este caso — ver memoria "Override C1-9
Roberto"). Se resuelve a mano, editando directamente MULTA/CORTE de estos 4
lotes en la planilla de agosto antes de correr `5_cobranza`.
