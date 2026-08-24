# RETOMAR - Cierre de agosto y lista de corte

**Fecha de corte:** 2026-08-17 21:05

## Estado actual

```text
Abono G-3 aplicado
        |
        v
Cobranza 2026-08 recalculada (ciclo 29)
        |
        v
5b ejecutado con diferencias conocidas y aceptadas
        |
        v
estado_ciclo.json: validado=true
        |
        v
lista_corte.xlsx: 38 filas, 5 SI, 33 NO
        |
        v
PENDIENTE: revisar y aplicar penalidad a los 5 SI
```

## Decisiones cerradas vigentes

- `D-001`: sin referencia de pago no se reparte en el historial.
- `D-002`: un `SOURCE` tecnico no es evidencia de pago.
- `D-003`: planillas historicas hasta 2026-05; ledger desde 2026-06.
- `D-004`: los generadores protegidos del reporte de siete lotes siguen bloqueados.
- No reinterpretar ni modificar estas decisiones.

## Abono G-3

Se agrego a `shared/abonos_rezagados.xlsx` y al manifest de agosto:

| Campo | Valor |
|---|---|
| MZ-LT | `G-3` |
| Monto | `S/23` |
| Mes ciclo | `2026-08` |
| Mes aplica | `2026-08` |
| Canal | `efectivo` |
| Cobrador / retenido por | `Wilder Trujillo` |
| Motivo | `se sacara el monto de pagos blanco efectivo` |
| Estado manifest | `CONFIRMADO` |

Validacion concreta del motor:

```text
G-3 total a pagar  S/23
abono rezagado     S/23
saldo              S/0
estado             CANCELADO
```

El guard del manifest paso con `20 activos`; fuente y manifest coinciden.

Backups creados antes de la aplicacion:

- `shared/abonos_rezagados_pre_G-3_20260817.xlsx`
- `5_cobranza/inputs/abonos_rezagados_manifest_2026-08_pre_G-3_20260817.json`
- `shared/seguimiento_pueblo_pre_G-3_20260817.xlsx`

## Cobranza

Comando ejecutado:

```powershell
py -u -X utf8 5_cobranza/main.py --force
```

Resultado:

- Ciclo de cobranza `29`.
- `560` usuarios.
- Estados: `CANCELADO=288`, `EXCESO=15`, `PARCIAL=72`, `PENDIENTE=185`.
- Quedo una discrepancia de efectivo para `S-16` por `S/22`.
- `arrastre_consolidado_2026-08.xlsx` no se genero en esa corrida porque el JSON todavia no estaba validado.

## Validacion 5b aceptada

Comando ejecutado despues de cobranza:

```powershell
py -u -X utf8 5b_validacion/main.py
```

Resultado:

```text
Yape procesado vs planilla: OK
Nivel 1a TE PAGO:          +S/700 ALERTA
Efectivo:                  -S/22 ALERTA
```

Las diferencias fueron investigadas y aceptadas expresamente por el usuario como
esperables:

```text
Tanque valido                         S/1,220
Reasignaciones duplicadas de tanque   +S/750
Nelson Mon* pendiente                 -S/50
                                      -------
Alerta Nivel 1a actual                +S/700

Efectivo S-16, Wilder Trujillo        -S/22
Abono G-3                              S/23  (no forma el -S/22)
```

Las cuatro filas duplicadas en `4_pagos/outputs/aportes_tanque.xlsx` suman
`S/750`: reasignaciones de `C1-2`, `V-14`, `I-13` y `H1-36`.

El pendiente esperado sigue en
`4_pagos/yape/motor_matching/correcciones/pendientes.xlsx`:
`Nelson Mon*`, `S/50`, sin mensaje ni maestro.

Por instruccion del usuario se actualizo:

`shared/reporte_acumulado_procesado/estado_ciclo.json`

```text
2026-08.arrastre.validado = true
validado_en = 2026-08-17T21:04:07
gap_conocido = +S/700 TE PAGO y -S/22 efectivo, con causas documentadas
```

## Lista de corte generada

Comando ejecutado:

```powershell
py -u -X utf8 6_corte/generar_lista.py
```

Salida:

`6_corte/outputs/lista_corte.xlsx`

Resumen:

```text
38 usuarios en lista
EJECUTAR_CORTE=SI:  5
EJECUTAR_CORTE=NO: 33
NO por reclamo:    23
NO por pago parcial: 10
```

Los cinco candidatos `SI` son:

| MZ-LT | Nombre | Saldo | Mes anterior | Penalidad | Total |
|---|---|---:|---:|---:|---:|
| `D-9` | RUFINA CABELLO TICLIO | 30 | 15 | 20 | 50 |
| `F-6` | ALBERTO TARAZONA CALDAS | 91 | 9 | 20 | 111 |
| `H-16` | GREGORIO TOLENTINO SANCHEZ | 66 | 11 | 20 | 86 |
| `P-6` | FLOR VALDIVIA MILLA | 16 | 8 | 20 | 36 |
| `X-12` | CARLOS CASHPA CORAQUILLO | 18 | 9 | 20 | 38 |

Durante la generacion, `6_corte` leyo `pagos_efectivo.xlsx` con `359` filas y
el manifest con `20` filas confirmadas / `18` usuarios. No cambiar esas fuentes
sin seguir primero el lector del modulo.

## Siguiente accion

```text
Revisar los 5 EJECUTAR_CORTE=SI
        |
        v
Esperar instruccion explicita del usuario
        |
        v
py -u -X utf8 6_corte/aplicar_penalidad.py
```

No ejecutar `aplicar_penalidad.py`, no modificar la lista y no corregir las
fuentes de las diferencias aceptadas sin una nueva instruccion del usuario.

## Archivos relevantes

- `RETOMAR_CIERRE_AGOSTO_LISTA_CORTE_2026-08-17.md`
- `shared/reporte_acumulado_procesado/estado_ciclo.json`
- `shared/abonos_rezagados.xlsx`
- `5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json`
- `5_cobranza/outputs/planilla_cobrado.xlsx`
- `5b_validacion/outputs/validacion_diferencias.xlsx`
- `6_corte/outputs/lista_corte.xlsx`
- `4_pagos/outputs/aportes_tanque.xlsx`
- `4_pagos/efectivo/outputs/pagos_efectivo_2026-08.xlsx`
- `4_pagos/yape/motor_matching/correcciones/pendientes.xlsx`

## Worktree

El repositorio ya estaba ampliamente modificado y contiene archivos no
rastreados de sesiones anteriores. No revertir ni limpiar cambios ajenos; revisar
`git status --short` antes de cualquier commit.
