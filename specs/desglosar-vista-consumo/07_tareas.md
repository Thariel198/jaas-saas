# Tareas: Desglosar consumo en vista seguimiento

## Definition of Done

```text
documentación aprobada -> proyector puro -> adapters -> pruebas -> evidencia
```

El cambio satisface `REQ-001` a `REQ-011` y `AC-001` a `AC-011`; conserva el ledger
real byte a byte; genera Excel y PDF temporales con las hojas aprobadas; valida un caso
afectado y uno no afectado; y deja README, contrato HTML, diagrama, código y cuaderno SDD
sin divergencias. Ninguna tarea se marca completa sin ejecutar su validación indicada.

## Tareas

- [x] `TASK-001` Actualizar primero el contrato documental de la vista.
  Dependencias: diseño aprobado. Archivos: `shared/README.md`,
  `shared/docs/formato_vista_seguimiento_pueblo.html` y
  `shared/docs/diagrama_flujo_seguimiento_pueblo.html`. Cambio: documentar las hojas
  `MES_ANTERIOR`, `MES_ACTUAL`, `MANTENIMIENTO`, su orden, FIFO, rollover, error ante
  ajustes ambiguos y condición regenerable. Validación: revisar que inventario, columnas,
  diagramas y consumidores coincidan con `05_spec.md` y `06_diseno.md`. Cubre:
  `REQ-003`, `REQ-005` a `REQ-009`, `REQ-011`, `AC-005`, `AC-008`.

- [x] `TASK-002` Implementar y probar aisladamente el proyector temporal puro.
  Dependencia: `TASK-001`. Archivo: `shared/seguimiento_repo.py`. Cambio: agregar el
  estado `AGUA_ANT/MANT_ANT/AGUA_ACT/MANT_ACT`, clasificación de apertura, rollover,
  reparto FIFO de pago/declarado, calendario global y
  `ProyeccionTemporalAmbiguaError`; validar totales y saldos antes de cualquier I/O.
  Validación: casos en memoria para I-9, pago parcial, pago cruzado, rollover, orden
  estable, pago excedente y ajuste ambiguo. Cubre: `REQ-001` a `REQ-004`, `REQ-009` a
  `REQ-011`, `AC-001` a `AC-004`, `AC-006`, `AC-011`.

- [x] `TASK-003` Integrar la proyección en las vistas oficial y provisional.
  Dependencia: `TASK-002`. Archivos: `shared/seguimiento_repo.py` y
  `shared/tests/test_vista_provisional.py`. Cambio: renderizar filas mensuales
  preagregadas, retirar la hoja visible `AGUA`, ordenar las ocho hojas de conceptos y
  conservar `Ajustes`, `CONVENIO_HISTORIAL` y `PROVISIONAL`. El PDF seguirá consumiendo
  hojas genéricamente. Validación: generar Excel/PDF en temporales y comparar estructura e
  importes; comprobar por bytes que la vista provisional no muta su ledger de entrada.
  Cubre: `REQ-005` a `REQ-010`, `AC-005`, `AC-008`, `AC-009`.

- [x] `TASK-004` Completar pruebas del caso afectado y propiedades globales.
  Dependencias: `TASK-002`, `TASK-003`. Archivo:
  `shared/tests/test_seguimiento_repo.py` y, solo si mejora el aislamiento sin duplicar
  fixtures, una prueba nueva bajo `shared/tests/`. Validación: ejecutar mediante
  `py -m test_safety.run`; verificar invariantes por predio/ciclo/campo en temporales,
  dos generaciones equivalentes y preservación de la salida previa ante error. Cubre:
  `INV-001` a `INV-008`, `AC-001` a `AC-006`, `AC-009` a `AC-011`.

- [x] `TASK-005` Ejecutar regresiones de consumidores no afectados.
  Dependencia: `TASK-004`. Archivos de prueba: `shared/tests/test_vista_provisional.py`,
  `7_cierre/test.py` y la prueba aplicable de `3_boletas` para
  `CONVENIO_HISTORIAL`. Validación: runners seguros con todas las rutas de escritura en
  temporales; comparar `CONVENIO`, `CONVENIO_HISTORIAL` y `Ajustes`; confirmar que
  `7_cierre` conserva las llamadas a Excel/PDF sin comprometer el ciclo real. Cubre:
  `REQ-007` a `REQ-010`, `AC-007` a `AC-009`.

- [x] `TASK-006` Registrar evidencia y reconciliar artefactos.
  Dependencia: `TASK-005`. Archivos: este registro de implementación y cuaderno SDD;
  `08_verificacion.md` se completa en el gate de verificación posterior a la aprobación
  humana de implementación. Validación: registrar comandos, resultados, caso
  afectado/no afectado, hashes o conteos de integridad, limitaciones y veredicto;
  ejecutar `git diff --check`, comparar `git status --short` antes/después y regenerar
  `index.html`. Cubre todos los requisitos y criterios; no autoriza `converge` sin
  aprobación humana posterior.

## Evidencia de implementación

```text
proyector puro -> Excel/PDF temporal -> consumidores -> ledger intacto
```

| Evidencia | Resultado |
|---|---|
| `py -m test_safety.run pytest shared/tests/test_proyeccion_consumo.py -q` | 8 passed |
| `py -m test_safety.run pytest shared/tests/test_vista_provisional.py -q` | 1 passed |
| `py -m test_safety.run script shared/tests/test_seguimiento_repo.py` | Todos los checks pasaron |
| `py -m test_safety.run pytest 3_boletas/tests/test_recibos_medidor_pagado.py -q` | 1 passed; aviso externo de deprecación de PyPDF2 |
| `py -m test_safety.run script 7_cierre/test.py` | PASÓ; todas las rutas del fixture fueron temporales |
| Proyección efectiva de AGUA/MANTENIMIENTO | 1,843 eventos, 509 predios, 509 filas por hoja |
| I-9 en Excel temporal real | anterior 8/8/0; actual 5/5/0; mantenimiento 3/3/0 |
| Excel + PDF reales dirigidos a `%TEMP%` | 10 hojas correctas; PDF generado |
| Integridad de `shared/seguimiento_pueblo.xlsx` | hash antes/después idéntico: `8e551b49d7e7...` |
| `git diff --check` en archivos trackeados afectados | sin errores; solo avisos de conversión LF/CRLF |

La primera validación temporal detectó que `exportar_vista_pdf()` dejaba abierto su
`pd.ExcelFile`; se agregó cierre explícito y una regresión que borra el Excel después de
exportar. La repetición completa terminó sin error y limpió el directorio temporal.
