# Manual de ejecucion del ledger real

## 7 lotes con abonos rezagados · ciclo 2026-08

Este manual transforma la corrida aislada validada en un procedimiento para
aplicar los cambios al ledger real. La mini-corrida es la referencia de calculo;
el ledger real solo se modifica despues de aprobar la comparacion completa.

```text
fuentes + correcciones
        |
        v
mini pipeline aislado
        |
        v
comparacion exacta de 7 lotes
        |
        v
backup inmediato
        |
        v
reconciliacion append-only del ledger real
        |
        v
validacion + reporte final
```

## 1. Alcance

Los lotes de la interseccion `lista_corte ∩ abonos_rezagados` son:

```text
I-9, L-5, P-12, P-3, Q-5, S-2, W-5
```

El mini resultado validado deja estos saldos para agosto:

| Lote | Abono incluido | Saldo consumo | Saldo mant. | Saldo anterior | Saldo corte | Saldo convenio | Saldo acuerdos | Saldo multa | Total |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| I-9 | S/144 | S/0 | S/0 | S/0 | S/0 | S/0 | S/0 | S/0 | **S/0** |
| L-5 | S/126 | S/0 | S/0 | S/0 | S/0 | S/0 | S/0 | S/50 | **S/50** |
| P-12 | S/30 | S/0 | S/3 | S/0 | S/0 | S/20 | S/0 | S/0 | **S/23** |
| P-3 | S/33 | S/22 | S/3 | S/0 | S/0 | S/0 | S/0 | S/0 | **S/25** |
| Q-5 | S/114 | S/13 | S/3 | S/0 | S/0 | S/0 | S/0 | S/0 | **S/16** |
| S-2 | S/60 | S/24 | S/3 | S/0 | S/0 | S/0 | S/20 | S/0 | **S/47** |
| W-5 | S/15 | S/12 | S/3 | S/0 | S/0 | S/0 | S/37 | S/0 | **S/40** |
| **Total** | **S/522** |  |  |  |  |  |  |  | **S/201** |

El mini-pipeline actualizado debe coincidir con esta tabla. El PDF anterior de
S/334 corresponde al escenario previo, antes de corregir I-9.

## 2. Abonos que deben conservarse

La fuente es `shared/abonos_rezagados.xlsx`. Para estos 7 lotes hay ocho filas:

| Lote | Monto | `MES_CICLO` | `MES_ANO_APLICA` | Canal / retenido por |
|---|---:|---|---|---|
| I-9 | S/86 | 2026-06 | 2026-08 | Yape / Wagner Trujillo |
| I-9 | S/58 | 2026-07 | 2026-08 | Efectivo / Secretaria |
| L-5 | S/126 | 2026-08 | 2026-08 | Decision de usuario |
| P-12 | S/30 | 2026-07 | 2026-08 | Yape / Wagner Trujillo |
| P-3 | S/33 | 2026-07 | 2026-08 | Yape / Yerald Romero |
| Q-5 | S/114 | 2026-07 | 2026-08 | Efectivo / Secretaria |
| S-2 | S/60 | 2026-07 | 2026-08 | Yape / Wagner Trujillo |
| W-5 | S/15 | 2026-07 | 2026-08 | Yape / Wagner Trujillo |

No fusionar filas que tengan distinto origen o `MES_CICLO`. Q-5 ya es una
excepcion aprobada: sus dos partes anteriores fueron unificadas en una fila de
S/114.

## 3. Correcciones ya realizadas

Estas correcciones forman parte de la base que debe usarse antes de reconciliar:

| Lote | Correccion | Backup |
|---|---|---|
| L-5 | Retiro de pagos prematuros invalidos: convenio S/60 y acuerdos S/50 | `shared/backups_ledger/seguimiento_pueblo_pre_remove_L5_invalid_20260816.xlsx` |
| S-2 | Retiro del pago prematuro de acuerdos S/3 | `shared/backups_ledger/seguimiento_pueblo_pre_remove_S2_invalid_20260816.xlsx` |
| Q-5 | Retiro de pagos manuales y ajustes de julio; abono unificado S/114 | `shared/backups_ledger/seguimiento_pueblo_pre_remove_Q5_manual_y_ajustes_20260817.xlsx` |

No borrar eventos historicos durante la nueva corrida. Las correcciones ya
hechas quedan auditadas por sus backups; los abonos nuevos deben entrar como
eventos append-only con `SOURCE=abonos_rezagados` y
`CLASE=ABONO_REZAGADO`.

## 4. Reglas del calculo

```text
MES_CICLO       -> deuda que cubre el abono
MES_ANO_APLICA  -> corrida en que se registra
```

La cascada usada por el mini pipeline es:

```text
abono cerrado:   mes anterior -> corte -> convenio -> acuerdos -> multa
abono vigente:   consumo -> mantenimiento -> mes anterior -> corte -> convenio -> multa -> acuerdos
```

El abono no es un pago normal, no debe duplicar Yape o efectivo y no debe
convertirse en exoneracion. La referencia de pago confirma que el pago existe;
un `SOURCE` tecnico no reemplaza evidencia de pago.

## 5. Correccion I-9 en el mini-pipeline

La fuente real debe actualizarse antes de producción. El mini-pipeline crea una
copia aislada y aplica:

```text
Wagner:    S/86, MES_CICLO 2026-06, MES_ANO_APLICA 2026-08
Secretaria:S/58, MES_CICLO 2026-07, MES_ANO_APLICA 2026-08
Total:     S/144
```

Resultado mini validado:

```text
Deuda antes:  S/152
Pago normal:   S/8
Abono:       S/144
Pago total:  S/152
Saldo final:   S/0
Exceso:        S/0
```

I-9 ya no es bloqueante para el mini-pipeline. Antes de produccion se debe
actualizar la fuente, el manifest y ejecutar la reconciliacion por el writer unico.

## 6. Preparacion sin escribir el ledger

1. Verificar que `shared/ciclo_activo.json` siga en `2026-08`.
2. Verificar que no haya cambios concurrentes en las fuentes.
3. Crear backup de `shared/seguimiento_pueblo.xlsx`.
4. Crear backup de `shared/abonos_rezagados.xlsx`.
5. Crear backup de `5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json`.
6. No editar originales durante esta fase.
7. Ejecutar la mini-corrida:

```powershell
py 5_cobranza/tests/generar_mini_corrida_abonos.py
```

La salida esperada es:

```text
C:\Users\wilde\AppData\Local\Temp\opencode\mini_corrida_lista_corte_7_20260816\outputs\mini_resultado_cascada.xlsx
C:\Users\wilde\AppData\Local\Temp\opencode\mini_corrida_lista_corte_7_20260816\outputs\mini_ledger_predicho.xlsx
```

## 7. Validacion de la mini-corrida

Revisar cada lote contra la tabla de la seccion 1:

1. El abono fuente coincide con la seccion 2.
2. El pago normal no aparece duplicado como abono.
3. El saldo por concepto coincide, no solo el total.
4. Q-5 termina en consumo S/13 y mantenimiento S/3.
5. No hay saldos negativos inesperados.
6. Cada fila conserva origen, retenido por, evidencia y respaldo.
7. El archivo `mini_ledger_predicho.xlsx` contiene solo los eventos previstos.

Comandos de verificacion:

```powershell
py 5_cobranza/tests/test_abonos_manifest.py
py 5_cobranza/tests/test_abonos_rezagados_mini.py
py -m py_compile 5_cobranza/main.py
py -m py_compile 4b_reclamos/reporte_historico.py 4b_reclamos/reporte_deuda_ledger.py
```

No continuar si falla una prueba o si cambia la tabla objetivo.

## 8. Preparacion de archivos reales

Solo despues de aprobar la mini-corrida:

1. Actualizar `shared/abonos_rezagados.xlsx` unicamente con los cambios aprobados.
2. Mantener las ocho filas y sus referencias, salvo una decision explicita de fusion.
3. Actualizar el manifest para que refleje exactamente monto, lote, ciclo y aplicacion.
4. Ejecutar nuevamente `test_abonos_manifest.py`.
5. Crear un segundo backup inmediato del ledger antes de escribir.
6. Comparar el segundo backup contra el primero y registrar la fecha.

No usar `shared/parches_manuales_pendientes_julio.xlsx` como sustituto de la
fuente de abonos ni insertar pagos manualmente en `seguimiento_pueblo.xlsx`.

## 9. Reconciliacion del ledger real

La escritura debe pasar por el flujo normal y por su writer unico:

```text
5_cobranza/main.py
        ↓
_reconciliar_pagos_pueblo()
        ↓
shared/seguimiento_repo.py
        ↓
seguimiento_pueblo.xlsx
```

Ejecutar una sola vez, solo despues de las fases anteriores:

```powershell
py 5_cobranza/main.py --force
```

Durante la corrida comprobar:

- no se ejecuta reimputacion historica;
- no se borran eventos anteriores;
- los nuevos eventos son append-only;
- `SOURCE=abonos_rezagados`;
- `CLASE=ABONO_REZAGADO`;
- cada origen mantiene su `AUDIT_REF`;
- no se vuelve a aplicar un pago normal ya existente.

## 10. Validacion posterior

1. Ejecutar `5b_validacion`.
2. Leer los saldos vivos de los siete lotes desde el ledger real.
3. Compararlos por concepto contra la tabla de la seccion 1.
4. Regenerar `lista_corte.xlsx`.
5. Regenerar el reporte historico de los siete lotes.
6. Confirmar que Q-5 muestre saldo S/16.
7. Confirmar que no existan saldos negativos, pagos duplicados o pagos fantasma.
8. Confirmar que el total de los siete saldos sea S/201.
9. Guardar comandos, backups, salida de validacion y PDF final en el registro de cierre.

## 11. Rollback

Si la validacion falla:

```text
detener flujo
    ↓
conservar logs y outputs
    ↓
no ejecutar otra vez --force
    ↓
comparar mini_ledger_predicho vs eventos reales
    ↓
usar el backup correspondiente, con autorizacion
```

Nunca corregir un resultado inesperado borrando eventos a mano. Primero
identificar si el problema esta en la fuente, el manifest, la planilla, la
cascada o la escritura.

## 12. Evidencia de cierre

El cierre solo es valido cuando existe este conjunto:

```text
backup fuente de abonos
backup manifest
backup ledger pre-escritura
mini_resultado_cascada.xlsx
mini_ledger_predicho.xlsx
resultado de tests
resultado de 5b_validacion
ledger real post-escritura
lista de corte regenerada
PDF final de los 7 lotes
fuente y manifest de I-9 actualizados a S/86 + S/58
```

Este manual no autoriza por si solo la escritura del ledger. La aprobacion de la
tabla mini y el porte de la cascada al pipeline real son las compuertas previas
a produccion; I-9 ya no es un bloqueo separado.

## 13. Pendiente antes de la corrida real

La mini-corrida ahora separa las fuentes de pago antes de repartir por concepto
para los siete lotes:

```text
pago efectivo del ciclo actual
    -> consumo y mantenimiento del ciclo actual

abono rezagado de ciclo cerrado
    -> deuda arrastrada del ciclo anterior
```

Este cambio esta validado en el aislamiento y debe verificarse en la
reconciliacion real de los siete lotes. No se modifica el pipeline para copiar al
mini-pipeline; I-9 ya no es un bloqueo separado. Antes de ejecutar
`5_cobranza/main.py --force` se debe:

1. Verificar que `_reconciliar_pagos_pueblo()` conserve la separacion general por fuente, sin crear un caso especial para I-9.
2. Ejecutar el mini-pipeline de los 7 lotes nuevamente.
3. Comparar la distribucion por concepto del mini contra la proyeccion del pipeline real.
4. Verificar al menos I-9 y otro lote sin abono de ciclo actual.
5. Actualizar las pruebas del pipeline real y obtener validacion sin diferencias.
6. Solo despues ejecutar la reconciliacion real con backup inmediato.
