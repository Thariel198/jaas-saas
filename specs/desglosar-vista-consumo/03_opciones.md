# Opciones: Desglosar consumo en vista seguimiento

## Resumen ejecutivo (lectura de 1 minuto)

```text
OPCION-A: ledger intacto -> proyector FIFO -> 3 hojas       RECOMENDADA
OPCION-B: migrar ledger -> cargos/aplicaciones identificados -> 3 hojas
```

La opcion A resuelve la vista actual con bajo riesgo y es completamente reversible. La B
da identidad perfecta por cargo, pero convierte un cambio visual en la migracion completa
del ledger destino.

## Detalle completo

## OPCION-A

### Proyeccion FIFO regenerable desde eventos existentes

Al generar la vista, un proyector de solo lectura conserva saldos internos
`AGUA_ANT/MANT_ANT/AGUA_ACT/MANT_ACT`, procesa cargos y pagos cronologicamente y produce
filas mensuales preagregadas para `MES_ANTERIOR`, `MES_ACTUAL` y `MANTENIMIENTO`.
`MANT_ANT` se suma a la primera hoja. El ledger y su taxonomia no cambian.

- Valor y cobertura: resuelve I-9, rollover futuro y las 509 cuentas actuales.
- Complejidad y costo: medio-bajo; un proyector puro y adaptación del writer visual.
- Riesgos e integridad: bajo; totales se validan contra AGUA+MANTENIMIENTO. Un ajuste sin
  cargo objetivo bloquea atomicamente la nueva vista en vez de inventar antiguedad.
- Migración y compatibilidad: ninguna migracion; cambia solo el conjunto de hojas.
- Reversibilidad: alta; retirar el proyector restaura la hoja AGUA.
- Tiempo hasta evidencia: corto; I-9 y un caso sintetico agosto-septiembre.

## OPCION-B

### Migrar el ledger a cargos y aplicaciones identificados

Agregar `CARGO_ID`, `MES_CARGO` y aplicaciones pago-cargo. Cada pago apuntaria al cargo
exacto y la vista solo agruparia aplicaciones por antiguedad. Requiere migrar todos los
eventos transitorios y cambiar writers/lectores antes de cambiar la vista.

- Valor y cobertura: identidad exacta, incluidos ajustes futuros y multiples meses.
- Complejidad y costo: muy alto; implementa una parte central del ledger destino.
- Riesgos e integridad: alto; exige backfill, conciliacion y cutover de datos reales.
- Migración y compatibilidad: cambia `seguimiento_repo`, `2_planilla`, `5_cobranza`,
  `7_cierre`, reportes y pruebas.
- Reversibilidad: baja despues del cutover.
- Tiempo hasta evidencia: largo; no entrega rapido la mejora visual solicitada.

## Comparación

| Criterio | OPCION-A | OPCION-B |
|---|---|---|
| Ledger real | Intacto | Migrado |
| Exactitud con datos actuales | Completa, salvo ajuste sin identidad | Completa |
| Ajustes AGUA/MANT hoy | 0; no bloquean | Soportados por cargo |
| Superficie afectada | Vista, PDF y pruebas | Pipeline completo |
| Riesgo | Bajo | Alto |
| Reversible | Si | Dificil |
| Evidencia inicial | I-9 en una iteracion | Tras migracion |
| Alineacion con pedido | Directa | Excede alcance |

**Recomendacion: OPCION-A.** Mantiene el ledger como fuente, evita una segunda verdad y
resuelve la necesidad visible sin adelantar parcialmente la migracion de arquitectura.
