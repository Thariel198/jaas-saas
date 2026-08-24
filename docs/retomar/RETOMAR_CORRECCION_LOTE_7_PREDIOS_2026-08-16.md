# RETOMAR - Correccion de lote de 7 predios

**Fecha de corte:** 2026-08-16

**Proposito:** continuar la proxima sesion con la correccion conjunta de los 7
predios que aparecen simultaneamente en la lista de corte y en
`shared/abonos_rezagados.xlsx`, sin repetir la investigacion ni escribir el
ledger real antes de aprobar el mini resultado.

## Estado en una mirada

```text
7 lotes C ∩ A
        |
        v
Reporte foto real ledger-only ya generado
        |
        v
Clasificacion documentada en Problema 03
        |
        v
I-9 probado en aislamiento: saldo S/0
        |
        v
PENDIENTE: probar los otros 6
        |
        v
PENDIENTE: actualizar fuente + manifest
        |
        v
PENDIENTE: backup -> ledger real -> validacion
```

## Regla de seguridad

Hasta el cierre de la siguiente sesion:

- No escribir en `shared/seguimiento_pueblo.xlsx`.
- No escribir en `shared/abonos_rezagados.xlsx`.
- No editar el manifest real para simular una aprobacion.
- No ejecutar `py 5_cobranza/main.py --force`.
- No ejecutar reimputacion historica.
- No usar `shared/abonos_rezagados.xlsx` para la foto real de reclamos.
- El mini pipeline debe trabajar con copias temporales.

## Documentos de referencia

| Documento | Uso |
|---|---|
| `docs/aprendizaje/solucion de proble/03_reporte_correccion_pagos.md` | Problema 03, clasificacion de los 7 lotes y solucion del caso I-9 |
| `docs/aprendizaje/solucion de proble/02_visualizacion_por_conjuntos.md` | Prueba de que los 7 lotes son `lista de corte ∩ abonos rezagados` |
| `docs/aprendizaje/solucion de proble/01_tiempo_20_minutos.md` | Contexto del aislamiento para evitar repetir corridas lentas |
| `RETOMAR_MINI_PIPELINE_ABONOS_REZAGADOS_2026-08-15.md` | Estado general del mini pipeline y advertencias previas |
| `README_PLAN_RECLAMOS_2026-08.md` | Orden mayor del cierre de agosto |
| `LEER_ANTES.md` | Restricciones activas y eventos excepcionales del flujo |

## Los 7 lotes

```text
C = lista de corte actual
A = abonos rezagados
C ∩ A = I-9, L-5, P-12, P-3, Q-5, S-2, W-5
```

| Lote | Abono fuente actual | Filas | Clasificacion | Problema que se corrige | Estado para correr real |
|---|---:|---:|---|---|---|
| `I-9` | S/136, objetivo S/133 | 2 | Confirmado + ajuste | Secretaria declara al dia; S/86 + S/50 no cuadra con saldo pendiente S/133 | Mini aprobado; fuente real aun no cambiada |
| `L-5` | S/126 | 1 | Confirmado | Revisar aplicacion entre convenio, multa y acuerdos | Pendiente mini |
| `P-12` | S/30 | 1 | Confirmado | Corregir convenio y evitar repetir el bug de signo | Pendiente mini |
| `P-3` | S/33 | 1 | Confirmado | Separar abono rezagado de deuda vigente | Pendiente mini |
| `Q-5` | S/114 | 2 | Revisar | Dos filas; una parte requiere decision antes de aplicar | Pendiente separar y aprobar |
| `S-2` | S/60 | 1 | Confirmado | Revisar aplicacion a campo/acuerdos | Pendiente mini |
| `W-5` | S/15 | 1 | Confirmado | Secretaria indica convenio cancelado; revisar campo y convenio | Pendiente mini |

## Caso I-9 aprobado en mini

### Hechos de la foto real

```text
Deuda total ultimo mes       S/141
Pago normal ya registrado      S/8
Saldo antes del nuevo abono  S/133
```

### Dos filas que se deben conservar

| Fuente | Monto final | `MES_CICLO` | Accion |
|---|---:|---|---|
| Wagner | S/86 | `2026-06` | Conservar |
| Secretaria | S/47 | `2026-07` | Ajustar desde S/50 a S/47 |
| **Total** | **S/133** | — | No fusionar las filas |

### Cargos que no se deben borrar

| Concepto | Mes | Monto |
|---|---|---:|
| `MULTA` | `2026-06` | S/50 |
| `ACUERDOS` | `2026-06` | S/75 |

### Resultado mini aprobado

```text
DEUDA_TOTAL_MINI = S/141
PAGO_NORMAL      = S/8
ABONO_WAGNER     = S/86
ABONO_SECRETARIA = S/47
ABONOS_TOTAL     = S/133
PAGO_TOTAL       = S/141
SALDO_FINAL      = S/0
ESTADO           = CANCELADO
```

La prueba uso el motor real `_calcular()` con una copia logica del escenario,
incluyendo el pago previo S/8, los cargos MULTA S/50 y ACUERDOS S/75 y el abono
total S/133. No modifico archivos reales.

### Diferencia importante encontrada

La planilla viva de agosto para `I-9` trae `ACUERDOS=58`, mientras la foto
ledger-only conserva `ACUERDOS=75`. Por eso el mini aprobado debe usar la deuda
ledger que se esta corrigiendo, no aceptar ciegamente el total parcial de la
planilla actual. Si se prueba sin ajustar el escenario, S/133 produce exceso;
eso no es el caso de negocio aprobado.

## Cambios que todavia NO se hicieron

```text
shared/abonos_rezagados.xlsx
    I-9 Secretaria: S/50 -> S/47              NO HECHO
    MES_ANO_APLICA de ambas filas -> 2026-08  NO HECHO

5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json
    agregar/ajustar las dos filas de I-9     NO HECHO

shared/seguimiento_pueblo.xlsx
    nuevos eventos ABONO_REZAGADO             NO HECHO
```

La fuente real inspeccionada tiene actualmente para `I-9`:

| Monto | Origen | `MES_CICLO` | `MES_ANO_APLICA` |
|---:|---|---|---|
| S/86 | Wagner | `2026-06` | `2026-07` |
| S/50 | Secretaria | `2026-07` | `2026-07` |

El manifest actual inspeccionado contiene `I-9 Wagner S/86` para julio, pero no
la fila de Secretaria. No ejecutar el guard hasta actualizar ambas cosas de forma
intencional y respaldada.

## Orden exacto de la proxima sesion

### Fase 1 - Preparar copias

1. Confirmar que el ciclo activo sigue siendo `2026-08` en `shared/ciclo_activo.json`.
2. Leer nuevamente `LEER_ANTES.md` y este documento.
3. Crear backup inmediato de `shared/seguimiento_pueblo.xlsx`.
4. Crear backup inmediato de `shared/abonos_rezagados.xlsx`.
5. Respaldar `5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json`.
6. No modificar todavia los originales; crear una carpeta temporal de mini corrida.

### Fase 2 - Preparar la fuente temporal

1. Copiar a la carpeta temporal todas las filas de los 7 lotes.
2. En la copia de `I-9`, mantener Wagner S/86.
3. En la copia de `I-9`, cambiar Secretaria de S/50 a S/47.
4. En la copia, poner `MES_ANO_APLICA=2026-08` para las dos filas.
5. Mantener `MES_CICLO=2026-06` para Wagner y `MES_CICLO=2026-07` para Secretaria.
6. Mantener `SOURCE` y evidencia distintos para poder auditar ambos hechos.
7. No fusionar las filas en una sola fila de S/133.

### Fase 3 - Ejecutar mini pipeline

1. Usar `5_cobranza/tests/generar_mini_corrida_abonos.py` o un runner equivalente
   apuntando exclusivamente a la copia temporal.
2. Revisar `outputs/mini_resultado_cascada.xlsx`.
3. Revisar `outputs/mini_ledger_predicho.xlsx`.
4. Comparar por cada lote: abono fuente, pago normal, concepto de aplicacion y saldo.
5. Para `I-9`, exigir exactamente S/0 y `CANCELADO`.
6. Para `Q-5`, detenerse si la segunda fila no tiene decision de negocio.
7. No pasar a producción por el solo hecho de que el script termine sin error.

### Fase 4 - Aprobar los 7

1. Registrar el resultado mini de cada lote en una tabla de aprobación.
2. Confirmar que ningún saldo esperado quede negativo por error de imputacion.
3. Confirmar que ningún pago normal aparezca duplicado como abono.
4. Confirmar que cada fila conserve su `AUDIT_REF` y origen.
5. Confirmar que `Q-5` tenga decisión específica antes de incluirlo.
6. Si cualquiera falla, corregir la copia temporal y repetir solo el mini.

### Fase 5 - Actualizar los archivos reales

1. Solo después de aprobar los 7, actualizar `shared/abonos_rezagados.xlsx`.
2. Cambiar I-9 Secretaria S/50 a S/47, sin eliminar la fila.
3. Cambiar la aplicación de ambas filas I-9 a `2026-08`.
4. Actualizar el manifest con las dos filas I-9 y estado `CONFIRMADO`.
5. Revisar que la suma de las dos filas sea S/133.
6. Ejecutar el test del manifest antes de cualquier reconciliación.

### Fase 6 - Correr el ledger real

1. Crear un segundo backup inmediatamente antes de escribir.
2. Ejecutar la reconciliación mediante `5_cobranza/main.py`; no insertar filas a mano.
3. Confirmar que el writer unico sea `shared/seguimiento_repo.py`.
4. Confirmar eventos append-only, sin borrar ni editar eventos anteriores.
5. Confirmar `SOURCE=abonos_rezagados` y `CLASE=ABONO_REZAGADO`.
6. Confirmar `AUDIT_REF` diferente para Wagner y Secretaria.
7. Ejecutar `5_cobranza --force` solo si la secuencia documentada lo exige y una sola vez.

### Fase 7 - Validar y cerrar

1. Ejecutar `5b_validacion`.
2. Revisar el saldo de los 7 lotes en `shared/seguimiento_pueblo.xlsx`.
3. Regenerar la lista de corte de agosto.
4. Regenerar el reporte de foto real de los 7 predios.
5. Confirmar que el PDF no contenga `ABONO REZ.` cuando se use modo foto real ledger-only.
6. Confirmar que `I-9` quede en saldo cero por la correccion aprobada.
7. Revisar que no haya pagos fantasma, duplicados ni saldos negativos inesperados.
8. Registrar en este documento la fecha, comandos, backups y resultados finales.

## Comandos de verificacion previstos

```powershell
py 5_cobranza/tests/test_abonos_manifest.py
py 5_cobranza/tests/test_abonos_rezagados_mini.py
py -m py_compile 5_cobranza/main.py
py -m py_compile 4b_reclamos/reporte_historico.py 4b_reclamos/reporte_deuda_ledger.py
```

No correr todavia el comando de producción hasta completar las fases 1 a 5.

## Salidas esperadas

```text
mini_resultado_cascada.xlsx
mini_ledger_predicho.xlsx
backup de fuente de abonos
backup de manifest
backup de seguimiento_pueblo
ledger real con eventos append-only
validacion sin alertas nuevas
reporte foto real de 7 predios
```

## Archivos tocados en esta sesión

- `4b_reclamos/reporte_historico.py`: agrega `SALDO` solo en el ultimo mes de cada predio; `DEUDA` y `PAGO` permanecen mensuales.
- `4b_reclamos/reporte_deuda_ledger.py`: genera la foto real ledger-only de los 7 predios.
- `docs/aprendizaje/solucion de proble/03_reporte_correccion_pagos.md`: clasifica los 7 lotes y documenta la solucion I-9.
- `RETOMAR_CORRECCION_LOTE_7_PREDIOS_2026-08-16.md`: este punto de continuidad.

## Evidencia ya verificada

```text
PDF de foto real:
  7 predios
  SALDO: 1 por predio, solo en el ultimo mes
  ABONO REZ.: ausente
  SIMULACION: ausente
  I-9: S/141 - S/8 = S/133

Mini I-9:
  Wagner S/86 + Secretaria S/47 = S/133
  deuda S/141
  pago total S/141
  saldo S/0
  estado CANCELADO
```

## Soluciones mini aprobadas para preparar el ledger real

```text
I-9
  abono: S/133 (Wagner S/86 + Secretaria S/47)
  resultado: saldo S/0 · CANCELADO
  no escribir todavía: la fuente real aún conserva Secretaria S/50

L-5
  abono: S/126
  cascada vigente: agua → convenio → acuerdos → multa
  resultado esperado: queda MULTA S/50
  no escribir todavía: los pagos inválidos S/110 ya fueron retirados del ledger

S-2
  abono: S/60 · Yape de Wagner Trujillo, retenido por Wagner
  resultado: saldo S/47 · solo ACUERDOS
  no escribir todavía: el PAGO inválido S/3 ya fue retirado del ledger
```

Estas tres soluciones están probadas en mini pipeline y quedan preparadas para la
corrida real append-only, después de aprobar los cuatro lotes restantes y resolver
Q-5. No ejecutar `5_cobranza --force` todavía.

## Punto exacto para continuar

```text
NO empezar por el ledger real.
Empezar por construir la copia temporal de los 7 lotes,
resolver Q-5,
probar los 6 lotes restantes,
y comparar cada resultado contra la foto real.
```
