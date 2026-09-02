# Intake: Desglosar consumo en vista seguimiento

## Resumen ejecutivo (lectura de 1 minuto)

```text
vista actual: [AGUA agregado] [MANTENIMIENTO]
                         |
                         v
vista esperada: [MES_ANTERIOR] [MES_ACTUAL] [MANTENIMIENTO]
```

- El cambio afecta `shared/vista_seguimiento_pueblo.xlsx`, no el ledger.
- La hoja `AGUA` desaparece de la vista.
- `MES_ANTERIOR` incluye agua y mantenimiento anteriores.
- Las hojas quedan ordenadas según la cascada visible.
- I-9 debe mostrar anterior S/8, actual S/5 y mantenimiento S/3.

## Detalle completo

## Pedido original

Wilde solicita cambiar `shared/vista_seguimiento_pueblo.xlsx` porque la hoja `AGUA`
mezcla deuda del mes anterior y del mes actual. La vista no permite observar cuanto se
debia antes, cuanto cargo el ciclo actual ni que bloque cubrieron primero los pagos.

Resultado solicitado:

```text
MES_ANTERIOR -> MES_ACTUAL -> MANTENIMIENTO -> CORTE_RECONEXION
             -> CONVENIO -> ACUERDOS -> MULTA -> OTROS
```

- `AGUA` desaparece como hoja visible.
- Se agregan `MES_ANTERIOR` y `MES_ACTUAL` al lado de `MANTENIMIENTO`.
- Mantenimiento anterior no tiene hoja propia: se suma dentro de `MES_ANTERIOR`.
- El ledger conserva `CONCEPTO=AGUA` y `CONCEPTO=MANTENIMIENTO`; el cambio es una vista
  derivada y no escribe ajustes ni reclasifica eventos.

## Evidencia disponible

### Hechos verificados

- `shared/seguimiento_repo.py::generar_vista()` crea una hoja por cada valor de
  `CONCEPTOS_ORDEN`; hoy genera `AGUA` y `MANTENIMIENTO` por separado.
- Cada hoja mensual expone `DEUDA`, `PAGO`, `DECLARADO`, `AJUSTE` y `SALDO`.
- `5_cobranza._componentes_cuenta()` ya separa internamente `AGUA_ANT`, `MANT_ANT`,
  `AGUA_ACT` y `MANT_ACT`.
- La cascada vigente aplica primero componentes anteriores y luego actuales.
- `_aplicar_componentes()` colapsa `AGUA_ANT` y `AGUA_ACT` al escribir
  `CONCEPTO=AGUA`; `CLASE` describe el hecho (`GENESIS`, `COBRANZA`,
  `ABONO_REZAGADO`), no la antiguedad.
- Existen cero eventos `AJUSTE` de `AGUA` o `MANTENIMIENTO` en el ledger actual.

### Caso I-9

```text
arrastre julio                         S/8
agosto CARGO AGUA source=saldo_inicial S/8
agosto CARGO AGUA source=2_planilla    S/5
agosto CARGO MANTENIMIENTO             S/3
agosto PAGO AGUA                       S/13
agosto PAGO MANTENIMIENTO               S/3
saldo final                              S/0
```

La vista actual muestra `AGUA=13`; la vista solicitada debe explicar `MES_ANTERIOR=8`,
`MES_ACTUAL=5` y `MANTENIMIENTO=3`, conservando total S/16.

## Preguntas iniciales

- Definir en exploracion la reconstruccion FIFO general para meses posteriores a agosto.
- Definir el comportamiento fail-safe si aparece un `AJUSTE` de agua/mantenimiento sin
  identificador del cargo afectado; no se puede inventar si corresponde a deuda anterior
  o actual.
- Confirmar consumidores que dependen de la hoja fisica `AGUA` antes de retirarla.
- `reporte_historico.py` queda fuera de este cambio; pertenece a otro problema/spec.
