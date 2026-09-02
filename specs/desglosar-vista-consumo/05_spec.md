# Especificación: Desglosar consumo en vista seguimiento

## Alcance

```text
eventos efectivos AGUA/MANTENIMIENTO
                  |
                  v
       proyección temporal FIFO
                  |
                  v
MES_ANTERIOR + MES_ACTUAL + MANTENIMIENTO -> Excel/PDF
```

El cambio transforma, en modo de solo lectura, los eventos efectivos de
`seguimiento_pueblo.xlsx` en una vista temporal de deuda, pagos y saldo. La entrada es el
ledger luego de aplicar sus anulaciones lógicas; las salidas son
`vista_seguimiento_pueblo.xlsx` y su PDF derivado. La vista está destinada a secretaría,
tesorería y directiva, y continúa siendo regenerable: nunca es fuente contable.

La proyección separa los eventos de `AGUA` y `MANTENIMIENTO` en tres bloques visibles por
predio y ciclo:

- `MES_ANTERIOR`: agua y mantenimiento pendientes de ciclos anteriores.
- `MES_ACTUAL`: cargo de agua correspondiente al ciclo consultado.
- `MANTENIMIENTO`: cargo de mantenimiento correspondiente al ciclo consultado.

Los demás conceptos conservan su contenido y semántica. El ledger, su taxonomía y la
cascada de cobranza no cambian.

## Requisitos

- `REQ-001`: La proyección debe clasificar cada cargo efectivo de `AGUA` o
  `MANTENIMIENTO` como anterior o actual usando evidencia temporal y causal del ledger,
  sin modificar el evento original.
- `REQ-002`: Para cada concepto, predio y ciclo, la proyección debe imputar `PAGO` y
  `DECLARADO` primero al saldo anterior y después al cargo actual, sin exceder la deuda
  disponible de cada bloque.
- `REQ-003`: La salida debe mostrar `MANTENIMIENTO` anterior dentro de
  `MES_ANTERIOR`; la hoja `MANTENIMIENTO` debe contener únicamente el cargo, pago,
  declarado, ajuste y saldo del ciclo actual.
- `REQ-004`: Un saldo actual impago al finalizar un ciclo debe aparecer como anterior en
  el ciclo siguiente. La proyección no debe reconstruir periodos anteriores a agosto de
  2026.
- `REQ-005`: Las hojas temporales deben conservar las columnas visibles `DEUDA`, `PAGO`,
  `DECLARADO`, `AJUSTE` y `SALDO`, además de las claves y columnas identificativas que ya
  expone la vista vigente.
- `REQ-006`: El workbook debe presentar, en este orden, `MES_ANTERIOR`, `MES_ACTUAL`,
  `MANTENIMIENTO`, `CORTE_RECONEXION`, `CONVENIO`, `ACUERDOS`, `MULTA`, `OTROS` y las
  hojas auxiliares vigentes que no representan conceptos. No debe contener la hoja
  `AGUA`.
- `REQ-007`: Los conceptos distintos de `AGUA` y `MANTENIMIENTO`, incluida
  `CONVENIO_HISTORIAL`, deben conservar sus filas, importes y comportamiento previo.
- `REQ-008`: El Excel y el PDF deben representar los mismos bloques, orden e importes; el
  exportador no debe depender de la existencia de la hoja `AGUA`.
- `REQ-009`: Si un `AJUSTE` de agua o mantenimiento no contiene evidencia suficiente para
  identificar el bloque afectado, la generación debe fallar explícita y atómicamente. No
  debe asignarlo por una heurística silenciosa ni reemplazar parcialmente la salida
  existente.
- `REQ-010`: Repetir la generación sobre la misma entrada efectiva debe producir el mismo
  contenido funcional y no debe agregar, modificar ni eliminar eventos del ledger.
- `REQ-011`: La regla de proyección debe ser independiente del formato Excel y de valores
  particulares de una JASS, para permitir que otro adaptador entregue la misma vista.

## Invariantes

- `INV-001`: Por predio y ciclo, la suma de `DEUDA`, `PAGO`, `DECLARADO`, `AJUSTE` y
  `SALDO` de los tres bloques proyectados debe ser igual a la suma de esos campos para
  `AGUA` y `MANTENIMIENTO` en la vista agregada equivalente.
- `INV-002`: Para cada fila proyectada se mantiene la identidad contable
  `SALDO = DEUDA - PAGO - DECLARADO + AJUSTE`, de acuerdo con la semántica vigente de la
  vista.
- `INV-003`: Ningún bloque puede recibir más pago o declarado que la deuda disponible que
  le corresponde, ni producir saldo negativo por efecto de la proyección.
- `INV-004`: La proyección no escribe en `seguimiento_pueblo.xlsx`,
  `anulaciones_ledger.json` ni ningún input contable.
- `INV-005`: Las anulaciones lógicas vigentes se respetan antes de clasificar o sumar
  eventos; un evento anulado no reaparece en la vista.
- `INV-006`: `AGUA` y `MANTENIMIENTO` continúan siendo los conceptos persistidos. Los
  nombres temporales son categorías de presentación, no conceptos nuevos del ledger.
- `INV-007`: La generación es completa o nula: una ambigüedad o una ruptura de totales no
  deja un Excel o PDF parcialmente actualizado.
- `INV-008`: El orden FIFO se conserva dentro de `AGUA` y dentro de `MANTENIMIENTO`; la
  vista no reinterpreta el orden entre conceptos decidido por cobranza.

## Criterios de aceptación

- `AC-001` (`REQ-001`, `REQ-002`, `REQ-003`, `REQ-005`): Con los eventos reales de I-9,
  la vista muestra `MES_ANTERIOR` con deuda S/8, pago S/8 y saldo S/0;
  `MES_ACTUAL` con deuda S/5, pago S/5 y saldo S/0; y `MANTENIMIENTO` con deuda S/3,
  pago S/3 y saldo S/0.
- `AC-002` (`REQ-002`, `REQ-003`): Dado un pago parcial menor que la deuda anterior, todo
  el pago queda en `MES_ANTERIOR` y `MES_ACTUAL` conserva íntegro su saldo.
- `AC-003` (`REQ-002`, `REQ-003`): Dado un pago que supera la deuda anterior, la vista
  salda primero `MES_ANTERIOR` y asigna únicamente el remanente al bloque actual del mismo
  concepto.
- `AC-004` (`REQ-004`): En un caso sintético agosto-septiembre, el saldo impago de agua y
  mantenimiento actuales de agosto aparece sumado en `MES_ANTERIOR` de septiembre; los
  cargos nuevos aparecen en `MES_ACTUAL` y `MANTENIMIENTO`.
- `AC-005` (`REQ-006`): El workbook generado no contiene `AGUA` y presenta las ocho hojas
  de conceptos en el orden requerido, sin alterar la posición relativa de las hojas
  auxiliares vigentes.
- `AC-006` (`REQ-001`, `REQ-009`, `INV-007`): Ante un ajuste efectivo sin evidencia del
  bloque afectado, la generación informa el predio, ciclo y referencia ambigua, termina
  con error y conserva intactas las salidas preexistentes.
- `AC-007` (`REQ-007`): Una cuenta no afectada de `CONVENIO` y la hoja
  `CONVENIO_HISTORIAL` conservan exactamente sus importes y filas tras regenerar la vista.
- `AC-008` (`REQ-008`): El PDF generado contiene los mismos bloques e importes del Excel y
  se exporta correctamente sin una hoja `AGUA`.
- `AC-009` (`REQ-010`, `INV-001`, `INV-004`): Dos generaciones consecutivas sobre una
  copia temporal del mismo ledger producen iguales valores por hoja; el hash y el conteo
  de eventos de la entrada permanecen iguales.
- `AC-010` (`REQ-010`, `INV-001`): Para los 509 predios actuales con eventos de agua o
  mantenimiento, la comparación por predio y ciclo entre la vista agregada y los tres
  bloques proyectados arroja diferencia S/0.00 en todos los campos monetarios.
- `AC-011` (`REQ-005`, `REQ-011`): Una prueba de la regla de proyección con estructuras en
  memoria, sin leer ni escribir Excel, obtiene los mismos bloques y totales esperados que
  el adaptador de vista para el mismo caso.

## Fuera de alcance

- Cambiar la cascada o cualquier cálculo de `5_cobranza`.
- Agregar `MES_ANTERIOR` o `MES_ACTUAL` a la taxonomía persistida.
- Escribir, migrar, reclasificar o reconstruir eventos del ledger.
- Implementar `CARGO_ID`, `MES_CARGO` o aplicaciones individuales pago-cargo.
- Inferir la antigüedad de un ajuste ambiguo.
- Reconstruir deuda anterior al inicio aprobado de agosto de 2026.
- Cambiar `reporte_historico.py`, boletas, planillas o reportes de cobranza.
- Cambiar importes, filas o reglas de conceptos distintos de `AGUA` y `MANTENIMIENTO`.
