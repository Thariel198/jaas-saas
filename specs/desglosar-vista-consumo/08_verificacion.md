# Verificación: Desglosar consumo en vista seguimiento

## Matriz de trazabilidad

```text
requisito -> implementación -> prueba ejecutada -> evidencia observable
```

| ID | Implementación | Evidencia ejecutada | Resultado |
|---|---|---|---|
| `REQ-001` | `seguimiento_repo._proyectar_consumo_temporal()` clasifica apertura y cargo actual. | `test_i9_separa_anterior_actual_y_mantenimiento` | PASS |
| `REQ-002` | Reparto `ANT -> ACT` por concepto y `CLASES_SUMAN_CAJA`. | `test_pago_aplica_fifo_dentro_de_agua`, `test_pago_y_declarado_se_conservan_separados_con_orden_estable` | PASS |
| `REQ-003` | Tres DataFrames proyectados y writer resumido. | I-9 sintético y real temporal | PASS |
| `REQ-004` | Rollover antes de procesar el ciclo siguiente. | `test_rollover_pasa_saldo_actual_al_mes_anterior` | PASS |
| `REQ-005` | `_COLUMNAS_PROYECCION` y `_escribir_hoja_vista(..., resumida=True)`. | Test de Excel/PDF y test standalone del repo | PASS |
| `REQ-006` | `CONCEPTOS_VISTA_ORDEN`; `AGUA` no se crea. | Test standalone y workbook real temporal | PASS |
| `REQ-007` | Los demás conceptos conservan el writer existente. | `CONVENIO`, `Ajustes`, `CONVENIO_HISTORIAL` y recibos | PASS |
| `REQ-008` | PDF recorre genéricamente las hojas nuevas. | `test_excel_y_pdf_exponen_las_mismas_hojas_temporales` | PASS |
| `REQ-009` | `ProyeccionTemporalAmbiguaError` antes de crear el workbook. | `test_ajuste_ambiguo_no_reemplaza_vista` | PASS |
| `REQ-010` | Proyección pura, orden estable y `_save_atomic`. | Dos proyecciones iguales + hash real antes/después | PASS |
| `REQ-011` | Proyector recibe/devuelve DataFrames, sin rutas ni workbook. | Ocho pruebas unitarias en memoria | PASS |
| `AC-001` | Caso I-9. | Real: anterior `8/8/0`, actual `5/5/0`, mantenimiento `3/3/0`. | PASS |
| `AC-002` | Pago parcial anterior. | Parámetro `pago=6`: anterior saldo 2, actual saldo 5. | PASS |
| `AC-003` | Pago cruza bloques. | Parámetro `pago=10`: anterior pago 8, actual pago 2. | PASS |
| `AC-004` | Rollover agosto-septiembre. | Septiembre: anterior deuda/saldo 6, actual deuda/saldo 5. | PASS |
| `AC-005` | Orden de ocho hojas visibles. | Workbook temporal con diez hojas totales y sin `AGUA`. | PASS |
| `AC-006` | Falla atómica por ajuste ambiguo. | Salida previa conserva exactamente sus bytes. | PASS |
| `AC-007` | Consumidor no afectado. | `test_recibos_medidor_pagado` y test standalone del repo. | PASS |
| `AC-008` | Exportación PDF. | PDF temporal contiene las tres secciones e importe de control. | PASS |
| `AC-009` | Idempotencia e integridad. | Frames iguales; ledger real hash `8e551b49d7e7...` intacto. | PASS |
| `AC-010` | Validación global previa al writer. | 1,843 eventos, 509 predios, 509 filas por hoja, sin divergencias. | PASS |
| `AC-011` | Regla pura en memoria. | `test_proyeccion_consumo.py`, sin I/O salvo tests explícitos de adapters. | PASS |

## Evidencia

Entorno: Windows, Python 3.13, ejecución desde la raíz del repositorio el 25/08/2026.

| Comando | Resultado |
|---|---|
| `py -m test_safety.run pytest shared/tests/test_proyeccion_consumo.py -q` | `8 passed` |
| `py -m test_safety.run pytest shared/tests/test_vista_provisional.py -q` | `1 passed` |
| `py -m test_safety.run script shared/tests/test_seguimiento_repo.py` | `TODOS LOS CHECKS PASARON` |
| `py -m test_safety.run pytest 3_boletas/tests/test_recibos_medidor_pagado.py -q` | `1 passed`, una advertencia de deprecación externa de PyPDF2 |
| `py -m test_safety.run script 7_cierre/test.py` | `PASÓ`, commit simulado únicamente en fixture temporal |
| Proyección real de solo lectura mediante `py -X utf8 -c ...` | 1,843 eventos, 509 predios; I-9 correcto |
| Generación real con Excel/PDF dirigidos a `TemporaryDirectory` | diez hojas, PDF creado, ledger intacto |
| `git diff --check -- shared/...` | sin errores; avisos de conversión LF/CRLF del worktree |

La primera corrida temporal detectó que `exportar_vista_pdf()` no cerraba
`pd.ExcelFile`, por lo que Windows impedía limpiar el Excel temporal. Se añadió
`xl.close()` y una regresión que elimina el `.xlsx` después de exportar; la repetición
terminó sin error.

Limitaciones verificadas:

- No se reemplazaron `shared/vista_seguimiento_pueblo.xlsx` ni su PDF operativo.
- El ledger efectivo actual no contiene ajustes de `AGUA`/`MANTENIMIENTO`; si aparece
  uno sin cargo objetivo, la vista se bloqueará según el spec.
- La advertencia de PyPDF2 pertenece a una dependencia de `3_boletas` y no afecta este
  cambio.

## Regresiones

- Caso afectado: I-9 real y sintético muestran el desglose requerido y conservan S/16.
- Bordes afectados: pago parcial, pago que cruza bloques, PAGO/DECLARADO con igual
  timestamp, rollover, pago excedente y ajuste ambiguo.
- Caso no afectado: `CONVENIO`, `Ajustes` y `CONVENIO_HISTORIAL` mantienen estructura e
  importes en el test del repo.
- Consumidor `3_boletas`: `_cargar_pagados()` sigue leyendo `CONVENIO_HISTORIAL` sin
  depender de `AGUA`.
- Consumidor `7_cierre`: el test de integración conserva las llamadas a vista/PDF y pasa
  con rutas redirigidas; no se ejecutó el cierre real.
- Vista provisional: usa el mismo generador, antepone `PROVISIONAL` y no muta el ledger.

## Veredicto

`PASS`

Problema, opción A, spec, diseño, tareas, implementación y evidencia coinciden. No se
detectan desvíos funcionales ni escrituras sobre datos reales. Queda pendiente únicamente
la aprobación humana del gate de verificación antes de converger el cambio.
