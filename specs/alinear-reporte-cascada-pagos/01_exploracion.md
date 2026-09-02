# Exploración: Alinear reporte con cascada de pagos

## Resumen ejecutivo (lectura de 1 minuto)

```text
5_cobranza reparte correctamente
MES_ANTERIOR -> MES_ACTUAL -> MANTENIMIENTO -> resto
                         |
                         v
ledger agrupa anterior + actual como AGUA
                         |
                         v
vista y reporte pierden la antiguedad visible
```

- I-9 confirma que no falta dinero ni deuda: el arrastre S/8, el consumo S/5 y el
  mantenimiento S/3 estan asentados y pagados.
- El error es de presentacion: agosto muestra `CONSUMO=13` y `MES_ANTERIOR=0` en vez de
  `CONSUMO=5` y `MES_ANTERIOR=8`.
- No se cambia la cascada de `5_cobranza`, no se agregan conceptos y no se escriben
  ajustes al ledger.
- La salida visible debe separar `MES_ANTERIOR`, `MES_ACTUAL` y `MANTENIMIENTO`;
  mantenimiento anterior se suma dentro de `MES_ANTERIOR`.
- Falta elegir el formato de `vista_seguimiento_pueblo.xlsx` y la regla segura para
  ajustes cuya antiguedad no esta identificada en el ledger transitorio.

## Detalle completo

## Sistema actual

### Flujo ejecutable observado

```text
shared/abonos_rezagados.xlsx / Abonos_Raw
  MONTO + MES_CICLO + MES_ANO_APLICA
                |
                +------------------------------+
                |                              |
                v                              v
comun._overlays_de_plata()          comun._abonos_rezagados_predio()
referencia visible                  monto que entra al reparto
MES = MES_CICLO                     filtro = MES_ANO_APLICA
                |                              |
                |                              v
                |                    comun._datos_ciclo()
                |                    deuda + cascada del reporte
                |                              |
                +---------------+--------------+
                                v
                    comun._filas_recientes()
                    planillas + eventos ledger
                                |
                                v
                    comun.tabla_predio()
                                |
                +---------------+----------------+
                v                                v
reporte_historico.py PDF                  buscar_pago.py
```

### Fuentes con trabajos distintos

- `Abonos_Raw` representa el dinero manual y conserva el ciclo original y el ciclo de
  regularizacion.
- `shared/planilla_mes/planilla_YYYY-MM.xlsx` aporta la deuda corregida de cada ciclo.
- `planilla_cobrado_YYYY-MM.xlsx` conserva el snapshot usado al repartir el pago en la
  corrida de ese ciclo.
- `shared/seguimiento_pueblo.xlsx` conserva cargos, pagos y ajustes por concepto. Un
  evento `ABONO_REZAGADO` es una aplicacion del dinero, no el movimiento completo.

### Lectura y reparto en el reporte

- `comun.py:455-480::_abonos_rezagados_predio()` selecciona filas por
  `MES_ANO_APLICA == mes_ano`. Luego clasifica como cerrado o vigente comparando
  `MES_CICLO` con ese mismo mes de aplicacion.
- `comun.py:552-577::_datos_ciclo()` aplica un abono cerrado sin consumo vigente, pero
  aplica el pago del ciclo en orden
  `consumo -> mantenimiento -> mes anterior -> corte -> convenio -> acuerdos -> multa`.
- `comun.py:615-727::_filas_recientes()` combina dos representaciones: pagos ledger de
  `MULTA/ACUERDOS/CONVENIO` y el reparto reconstruido desde planillas. Los conceptos de
  pueblo solo se reemplazan por la reconstruccion en ciertos modos o en el ciclo
  proyectado; agua, mantenimiento, mes anterior y corte se toman siempre de `_datos_ciclo`.
- `reporte_historico.py:71-93` ya usa `MES_APLICA` para el estado de las referencias, pero
  el estado de la fila mensual depende de `MES` y del booleano `APLICACION_PENDIENTE`.
- `reporte_historico.py:425-433` considera tanto el mes original como el mes de aplicacion
  al decidir que filas conservan pagos en la generacion por lote.

### Cascada de referencia

```text
5_cobranza/main.py:651-654
MES_ANTERIOR -> MES_ACTUAL -> MANTENIMIENTO
             -> CORTE -> CONVENIO -> ACUERDOS -> MULTA
```

Para abonos cuyo ciclo ya cerro, `5_cobranza/main.py:639-641` usa
`MES_ANTERIOR -> CORTE -> CONVENIO -> ACUERDOS -> MULTA` y no cancela consumo nuevo del
mes de regularizacion. La diferencia es intencional en el motor operativo; el problema
del reporte es que ubica primero el dinero en el mes de regularizacion y desde ahi decide
que cascada reconstruir.

### Estado ejecutable tras la fusion del reporte

`reporte_historico.py` ya no usa `comun.tabla_predio()` para la salida principal. Desde
la fusion vigente construye agosto en `tabla_predio_ledger()` y conserva
`comun.tabla_predio()` como consumidor de `buscar_pago.py`. Esto deja dos proyecciones que
deben alinearse, pero no autoriza a cambiar el ledger para hacerlas coincidir.

```text
2_planilla: MES_ANTERIOR / MES_ACTUAL / MANTENIMIENTO
                         |
                         v
5_cobranza._componentes_cuenta()
  AGUA_ANT + MANT_ANT -> AGUA_ACT -> MANT_ACT -> resto de cascada
                         |
                         v
_aplicar_componentes(): colapsa AGUA_ANT + AGUA_ACT como CONCEPTO=AGUA
                         |
                         v
seguimiento_pueblo.xlsx: saldo correcto, sin antiguedad visible del pago AGUA
                         |
             +-----------+-----------+
             v                       v
vista_seguimiento_pueblo.xlsx   reporte_historico.py
AGUA agregado                  todo AGUA se muestra como CONSUMO
```

## Hallazgos

| ID | Hallazgo verificado | Consecuencia observable |
|---|---|---|
| H-01 | `5_cobranza` prioriza mes anterior; `_datos_ciclo()` prioriza consumo vigente. | El mismo pago se explica con conceptos distintos segun el consumidor. |
| H-02 | La referencia usa `MES_CICLO`, pero el reparto monetario agrupa por `MES_ANO_APLICA`. | El PDF puede nombrar un pago en un mes y mostrarlo aplicado en otro. |
| H-03 | `_resumen_y_historial()` suma todo `PAGO` ledger sin distinguir `CLASE`; `_datos_ciclo()` vuelve a incorporar el abono fuente. | La aplicacion contable y el ingreso pueden contarse como dos pagos si no se concilian. |
| H-04 | Los eventos ledger guardan porciones por concepto y no `ABONO_ID`. | El adapter actual solo puede conciliar agregados; no puede enlazar individualmente dos abonos del mismo predio y mes. |
| H-05 | El estado de la referencia usa `MES_APLICA`, pero el estado de la fila mensual no conserva ese periodo. | Una fila de un mes cerrado puede decir `ASENTADO` aunque su aplicacion sea posterior y siga pendiente. |
| H-06 | La generacion por lote filtra pagos usando referencias; la generacion individual no ejecuta ese mismo bloque. | Dos entrypoints del mismo reporte pueden divergir para el mismo predio. |
| H-07 | `buscar_pago.py` consume `tabla_predio()` y usa `MES_ANT` para explicar reclamos. | Corregir el reporte tambien puede cambiar veredictos de busqueda; es un consumidor obligatorio de regresion. |
| H-08 | `_abonos_rezagados_predio()` lee todas las filas de `Abonos_Raw` sin consultar clasificacion ni manifest. | El cambio no debe ampliar ni redefinir que pagos manuales estan autorizados; solo conservar el contrato actual. |
| H-09 | La cascada acumula solo lo aplicado a deuda y descarta el remanente de sus columnas, mientras `total_pagado` conserva el monto fuente. | `PAGO_TOTAL` puede ser menor al abono; hay que distinguir exceso real de monto desplazado al ciclo equivocado. |
| H-10 | Las pruebas cubren abono cerrado y snapshots corregidos, pero no el orden del pago vigente ni los dos periodos. | La divergencia actual no esta protegida por una regresion sintetica directa. |
| H-11 | `5_cobranza/main.py::_ORDEN_CICLO` ya aplica `AGUA_ANT, MANT_ANT, AGUA_ACT, MANT_ACT` antes del resto de la cascada. | No hace falta cambiar la distribucion operativa de pagos. |
| H-12 | `_aplicar_componentes()` agrega `AGUA_ANT` y `AGUA_ACT` en un solo `PAGO` de concepto `AGUA`; el ledger transitorio no guarda `CARGO_ID` ni `MES_CARGO` en la aplicacion. | La vista no puede leer una columna fisica que ya contenga el split; debe proyectarlo por FIFO desde eventos existentes. |
| H-13 | `tabla_predio_ledger()` mapea todo `CONCEPTO=AGUA` a `CONSUMO`. | Una apertura anterior y un cargo actual aparecen sumados como consumo del mes. |
| H-14 | `generar_vista()` crea una hoja por concepto y muestra `AGUA` agregado. | El saldo es correcto, pero la vista no explica cuanto corresponde a mes anterior y cuanto al actual. |
| H-15 | `2_planilla._load_saldos_cuenta()` suma los saldos anteriores de `AGUA` y `MANTENIMIENTO` en `MES_ANTERIOR`. | `MANT_ANT` es detalle interno; visualmente pertenece a `MES_ANTERIOR`, no a una cuarta categoria. |

### Evidencia I-9

```text
Julio arrastre_consolidado       DEUDA_AGUA = 8
Agosto planilla_cobrado          MES_ANTERIOR = 8

Ledger agosto
  CARGO AGUA 8  source=saldo_inicial   audit_ref=apertura|2026-08|AGUA|I|9
  CARGO AGUA 5  source=2_planilla
  CARGO MANT 3  source=2_planilla
  PAGO  AGUA 8  source=abonos_rezagados
  PAGO  AGUA 5  source=5_cobranza
  PAGO  MANT 3  source=5_cobranza

Reporte actual agosto
  CONSUMO deuda/pago = 13/13
  MANT.   deuda/pago =  3/3
  MES ANT deuda/pago =  0/0    <- presentacion incorrecta

Proyeccion esperada
  MES ANT deuda/pago = 8/8
  CONSUMO deuda/pago = 5/5
  MANT.   deuda/pago = 3/3
```

No falta deuda ni pago en el ledger de I-9. El defecto es exclusivamente de proyeccion.

### Evidencia F1-4

```text
Fuente manual: S/101
MES_CICLO: 2026-06
MES_ANO_APLICA: 2026-08

Deuda reconstruida de junio, con corte exonerado:
mes anterior 29 + consumo 29 + mantenimiento 3 + multa 50 = 111

Reparto esperado:
29 mes anterior + 29 consumo + 3 mantenimiento + 40 multa = 101
saldo multa = 10
```

- La ejecucion de lectura del reporte actual produjo para F1-4:
  `2026-06 PAGO_TOTAL=0`, `2026-07 PAGO_TOTAL=0` y
  `2026-08 PAGO_TOTAL=79` (`29 mes anterior + 50 multa`).
- La referencia ya sale con `MES=2026-06`, `MES_APLICA=2026-08` y `MONTO=101`.
- El ledger conserva `2026-06 CARGO MULTA=50` y
  `2026-08 PAGO MULTA=50, SOURCE=abonos_rezagados, CLASE=ABONO_REZAGADO`.
- Una ejecucion aislada de `_datos_ciclo()` sobre junio con el abono S/101 y la cascada
  del ciclo produjo exactamente `29 + 29 + 3 + 40`.

### Alcance medido del ledger transitorio

- Se observaron 19 eventos con `SOURCE=abonos_rezagados` o
  `CLASE=ABONO_REZAGADO`, agrupados en 11 claves `MZ + LT + MES_APLICA`.
- Los 19 eventos tienen al menos una fila fuente coincidente en `Abonos_Raw`.
- Para esas claves, la fuente suma S/2,231 y el ledger S/668. La diferencia no prueba
  perdida: el ledger observado solo contiene las porciones aplicadas a conceptos de
  pueblo y omite agua, mantenimiento, corte, exceso y filas aun no comprometidas.

### Consumidores encontrados

- `reporte_historico.py`: cinco llamadas a `tabla_predio()` para salida individual,
  lotes por deuda y ejecucion de consola.
- `buscar_pago.py`: usa `tabla_predio()` para explicar reclamos de mes anterior.
- `auditar_pago_vs_ledger.py`: compara el reporte contra dinero y ledger, normalmente
  con abonos manuales excluidos.
- Mini-pipelines de `5_cobranza/tests/`: reutilizan la tabla para validacion sintetica.

## Restricciones

### Decisiones de negocio confirmadas

- `MES_CICLO` es el mes real del pago y el mes donde se muestra en el reporte.
- `MES_ANO_APLICA` es el mes de regularizacion contable.
- Si ambos coinciden, el pago se muestra y aplica en ese mismo ciclo.
- La cascada del ciclo empieza por `MES_ANTERIOR`, sigue con `MES_ACTUAL` y
  `MANTENIMIENTO`, y termina con `CORTE -> CONVENIO -> ACUERDOS -> MULTA`.
- La regla es general. F1-4 es evidencia de regresion, no una excepcion por predio.

### Contratos cerrados y compatibilidad

- D-001: sin referencia de pago no se reparte dinero en el historial.
- D-002: un `SOURCE` tecnico no demuestra que el pago existe.
- D-003: el reporte usa archivos historicos hasta mayo de 2026 y el flujo reciente desde
  junio; esta frontera no cambia.
- D-005: el estado de cuenta completo empieza en agosto sin backfill. La reconstruccion
  visual anterior sigue leyendo las fuentes historicas existentes.
- El ledger es append-only: este cambio no borra, mueve ni reescribe eventos.
- `5_cobranza` ya contiene la cascada confirmada y queda fuera de cambios funcionales.
- `shared/abonos_rezagados.xlsx` sigue teniendo writer humano. El reporte solo lee.
- No se agregan condiciones por lote, monto, fecha o resultado esperado.

### Operacion y verificacion

- Las pruebas deben usar DataFrames y temporales; no pueden escribir Excel o PDF sobre
  outputs reales.
- Debe verificarse un caso afectado, uno no afectado y el consumidor `buscar_pago.py`.
- La salida debe conservar la deuda del ledger aun cuando deje de mostrar su evento de
  aplicacion como una segunda entrada de caja.
- Una diferencia entre fuente, ledger y PDF se informa; no se fuerza ningun saldo para
  hacer coincidir F1-4.
- La presentacion confirmada por el usuario tiene tres categorias visibles:
  `MES_ANTERIOR`, `MES_ACTUAL` y `MANTENIMIENTO`.
- `MES_ANTERIOR` incluye el saldo anterior de `AGUA` y el saldo anterior de
  `MANTENIMIENTO`; `MANT_ANT` no se presenta como categoria separada.
- La prioridad visible y operativa es `MES_ANTERIOR -> MES_ACTUAL -> MANTENIMIENTO`
  antes del resto de la cascada.
- El cambio no agrega conceptos al ledger ni escribe ajustes. La taxonomia T2 se conserva:
  arrastre es deuda antigua de AGUA/MANTENIMIENTO, no un concepto nuevo.

## Incertidumbres

No quedan preguntas de negocio abiertas para esta fase. Las incertidumbres tecnicas que
deben compararse en `03_opciones.md` son:

- conciliar temporalmente por agregado `MZ + LT + MES_ANO_APLICA`, aprovechando que hoy
  todos los eventos observados tienen fuente, o introducir una identidad puente antes de
  corregir el reporte;
- filtrar los eventos `ABONO_REZAGADO` solo en la vista de dinero y conservarlos completos
  en el calculo de saldo, o construir una proyeccion separada de ingreso y aplicacion;
- centralizar la cascada del reporte sin importar `5_cobranza/main.py`, para evitar acoplar
  una herramienta de lectura al orquestador operativo, o extraer una primitiva compartida;
- representar multiples aplicaciones futuras de un mismo abono cuando el adapter actual
  solo guarda un `MES_ANO_APLICA`;
- definir la regresion para dos abonos del mismo predio y mes, caso donde no existe
  identidad individual en el ledger transitorio.
- elegir si `vista_seguimiento_pueblo.xlsx` muestra las tres categorias en una hoja
  compuesta o en hojas separadas; ambas alternativas deben conservar las hojas y
  consumidores no relacionados.
- definir como exponer `AJUSTE` de AGUA/MANTENIMIENTO cuando el ledger transitorio no
  identifica el cargo concreto al que apunta; la vista no debe inventar antiguedad.
