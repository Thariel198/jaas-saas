# Diseño: Desglosar consumo en vista seguimiento

## Arquitectura

```text
seguimiento_pueblo.xlsx + anulaciones_ledger.json
                       |
                       v
          _leer_eventos() efectivos
                       |
          +------------+-------------+
          |                          |
          v                          v
 AGUA/MANTENIMIENTO           demás conceptos
          |                          |
          v                          |
 proyector FIFO puro                 |
          |                          |
          v                          v
3 resúmenes mensuales ------> writer de vista
 MES_ANTERIOR                         |
 MES_ACTUAL                           v
 MANTENIMIENTO          vista_seguimiento_pueblo.xlsx
                                      |
                                      v
                         exportar_vista_pdf() genérico
```

La solución se incorpora a `shared/seguimiento_repo.py` como una proyección pura en
memoria llamada únicamente por `generar_vista()`. No se crea un concepto, archivo
persistente ni writer adicional. `_leer_eventos()` continúa resolviendo las anulaciones
antes de que el proyector reciba datos.

El proyector entrega filas mensuales preagregadas. No fabrica eventos con apariencia de
ledger: separa explícitamente la representación derivada de la persistencia append-only.
Los conceptos no afectados siguen el camino actual y el exportador PDF continúa leyendo
las hojas de forma genérica.

### Algoritmo temporal

Por cada predio se mantienen cuatro saldos internos:
`AGUA_ANT`, `MANT_ANT`, `AGUA_ACT` y `MANT_ACT`.

1. Se seleccionan eventos efectivos de `AGUA` y `MANTENIMIENTO` y se ordenan por predio,
   `MES`, `TIMESTAMP` y posición original para desempatar de forma estable.
2. En el primer ciclo, un cargo con `SOURCE=saldo_inicial` abre el bloque anterior. Los
   demás cargos del ciclo abren el bloque actual de su concepto.
3. Al comenzar cada ciclo posterior, el saldo que quedó en `AGUA_ACT` pasa a `AGUA_ANT` y
   el que quedó en `MANT_ACT` pasa a `MANT_ANT`; los bloques actuales vuelven a cero antes
   de recibir cargos nuevos.
4. Cada evento `PAGO` se aplica dentro de su concepto: anterior primero, actual después.
   La misma distribución conserva por separado `PAGO` si su `CLASE` pertenece a
   `CLASES_SUMAN_CAJA`, o `DECLARADO` en caso contrario.
5. La salida mensual suma `AGUA_ANT + MANT_ANT` en `MES_ANTERIOR`, deja `AGUA_ACT` en
   `MES_ACTUAL` y `MANT_ACT` en `MANTENIMIENTO`.
6. Se materializan todos los ciclos globales desde la primera aparición del predio hasta
   el último ciclo disponible, aunque en uno no tenga movimiento, para que el rollover no
   desaparezca de la vista.
7. Antes de crear el workbook se validan conservación de totales, identidad de saldo y
   ausencia de saldos negativos. Cualquier fallo aborta antes de `_save_atomic()`.

El esquema vigente no identifica el cargo objetivo de un `AJUSTE`. Por eso, en esta
versión, cualquier ajuste efectivo de `AGUA` o `MANTENIMIENTO` produce
`ProyeccionTemporalAmbiguaError` con predio, mes y `AUDIT_REF`. Es preferible bloquear la
vista a inventar antigüedad. Las columnas `AJUSTE` permanecen en las tres hojas y quedan
en cero mientras no exista un contrato explícito de cargo objetivo.

## Componentes

| Archivo | Cambio | Responsabilidad |
|---|---|---|
| `shared/seguimiento_repo.py` | Modificar | Proyectar AGUA/MANTENIMIENTO, validar invariantes, ordenar hojas y renderizar los tres resúmenes. |
| `shared/tests/test_seguimiento_repo.py` | Modificar | Actualizar contrato de hojas y cubrir FIFO, parciales, exceso, rollover, ajuste ambiguo, atomicidad e idempotencia. |
| `shared/tests/test_vista_provisional.py` | Modificar | Verificar que la vista provisional usa las hojas nuevas sin mutar el ledger real. |
| `shared/README.md` | Modificar antes del código | Documentar la API visual, las hojas derivadas y el bloqueo ante ajustes ambiguos. |
| `shared/docs/formato_vista_seguimiento_pueblo.html` | Modificar antes del código | Reemplazar el contrato visual agregado por el desglose aprobado y corregir el inventario de hojas. |
| `shared/docs/diagrama_flujo_seguimiento_pueblo.html` | Modificar antes del código | Mostrar la proyección entre lectura efectiva y adapters Excel/PDF. |
| `7_cierre/consolidar_cierre.py` | Sin cambio | Continúa llamando `generar_vista()` y luego `exportar_vista_pdf()` después del commit. |
| `3_boletas/recibos_medidor_pagado.py` | Sin cambio | Continúa leyendo únicamente `CONVENIO_HISTORIAL`. |

### Contrato interno

La función pura recibe un `DataFrame` de eventos efectivos y devuelve un diccionario con
las claves `MES_ANTERIOR`, `MES_ACTUAL` y `MANTENIMIENTO`. Cada valor contiene:

```text
MZ | LT | MES | DEUDA | PAGO | DECLARADO | AJUSTE | SALDO
```

No recibe rutas, workbooks ni configuración de una JASS. Los importes siguen la unidad
del ledger actual durante esta transición; no se introduce una conversión parcial a
céntimos dentro de un adapter que todavía opera con Excel.

`generar_vista()` crea primero las ocho hojas de conceptos en el orden aprobado, luego
`Ajustes` y, si existe su fuente, `CONVENIO_HISTORIAL`. La vista provisional conserva
`PROVISIONAL` en la posición cero y reutiliza exactamente el mismo generador.

## Datos y migración

- No hay migración ni cambio de esquema en `seguimiento_pueblo.xlsx`.
- `CONCEPTO` conserva `AGUA` y `MANTENIMIENTO`; los tres nombres nuevos existen solo en
  adapters de lectura.
- La hoja `AGUA` deja de ser parte del contrato de la vista. No se encontró un consumidor
  operativo de esa hoja; la prueba estructural y la vista provisional sí deben cambiar.
- `Ajustes`, `CONVENIO_HISTORIAL` y todos los conceptos no afectados conservan su forma.
- El Excel se sigue reemplazando mediante `_save_atomic()`. La validación temporal ocurre
  antes de construir o reemplazar la salida.
- El despliegue consiste en actualizar documentación, código y pruebas, y luego regenerar
  Excel/PDF desde una copia temporal para revisión. No se ejecuta `7_cierre` ni se escribe
  el ledger real como parte de la implementación.
- Rollback: retirar la llamada al proyector, restaurar el orden visible anterior y
  regenerar las vistas. No hay datos que deshacer.
- La ruta futura a PostgreSQL conserva la frontera: la consulta de eventos reemplaza a
  `_leer_eventos()`, mientras el proyector tenant-agnóstico mantiene su contrato.

## Riesgos

| Riesgo | Impacto | Mitigación y evidencia |
|---|---|---|
| Un ajuste no indica antigüedad. | Saldo asignado al bloque incorrecto. | Fallo explícito antes de escribir; prueba conserva salida previa. |
| Orden inestable de eventos con igual timestamp. | Reparto distinto entre PAGO y DECLARADO. | Posición original como desempate estable y prueba de repetición. |
| Pago supera la deuda conocida del concepto. | Saldo negativo o pérdida de dinero visible. | Rechazo con predio, mes y referencia; no recortar silenciosamente. |
| El rollover omite un ciclo sin movimientos. | Deuda pendiente desaparece visualmente. | Materializar calendario global y prueba agosto-septiembre. |
| Totales proyectados divergen del agregado. | Vista deja de representar el ledger. | Reconciliación por predio/ciclo/campo antes del writer. |
| Consumidor manual espera la hoja `AGUA`. | Cambio operativo no detectado por búsqueda de código. | Contrato HTML y README explícitos; revisión humana del cuaderno y workbook temporal. |
| Vista provisional conserva expectativas antiguas. | Simulación y cierre muestran estructuras distintas. | Un solo `generar_vista()` y regresión específica provisional. |
| PDF no refleja el Excel. | Dos lecturas operativas incompatibles. | Comparar hojas/importes Excel-PDF tras exportación temporal. |

No se toma ninguna decisión irreversible. La única decisión futura pendiente es el
contrato de identidad para ajustes de agua/mantenimiento; queda fuera de este cambio y no
se simula con `SOURCE`, texto de `MOTIVO` ni patrones ad hoc.

## Estrategia de pruebas

```text
proyector puro -> workbook temporal -> PDF temporal -> consumidores no afectados
```

- Unitarias (`AC-001` a `AC-004`, `AC-006`, `AC-011`): I-9, pago parcial, pago que cruza
  de anterior a actual, rollover agosto-septiembre, orden estable y ajuste ambiguo.
- Integración Excel (`AC-005`, `AC-009`, `AC-010`): generar en temporales, comprobar hojas,
  columnas, orden, invariantes por los 509 predios y dos corridas funcionalmente iguales.
- Integración provisional (`REQ-010`): confirmar por bytes que el ledger temporal de
  entrada no cambia y que `PROVISIONAL` precede a las hojas nuevas.
- Integración PDF (`AC-008`): exportar desde el workbook temporal y verificar que existen
  las hojas/secciones nuevas y sus importes de control.
- Regresión no afectada (`AC-007`): comparar `CONVENIO`, `CONVENIO_HISTORIAL`, `Ajustes` y
  un recibo de medidor pagado antes/después.
- Seguridad operativa: ejecutar pruebas con `test_safety.run`, comparar `git status
  --short` antes y después, y no usar `shared/seguimiento_pueblo.xlsx` como fixture
  escribible.
- Validación del consumidor: probar `7_cierre` con repositorios y rutas redirigidos a
  temporales; no ejecutar el commit real del ciclo.
