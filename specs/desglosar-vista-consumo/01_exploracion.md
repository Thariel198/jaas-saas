# Exploración: Desglosar consumo en vista seguimiento

## Resumen ejecutivo (lectura de 1 minuto)

```text
eventos AGUA/MANTENIMIENTO
          |
          +-- cargo anterior: source=saldo_inicial
          +-- cargo actual:   source=2_planilla
          +-- pagos:          FIFO dentro de cada concepto
          |
          v
[MES_ANTERIOR] [MES_ACTUAL] [MANTENIMIENTO]
```

- El ledger actual tiene 1,843 eventos de agua/mantenimiento para 509 predios, todos de
  agosto de 2026.
- Los cargos permiten distinguir apertura anterior y cargo actual.
- Los pagos estan agregados por concepto, pero el split anterior/actual se puede derivar
  sin cambiar saldos porque ambos ordenes operativos pagan primero lo antiguo.
- No existen ajustes de agua/mantenimiento que introduzcan ambiguedad hoy.
- Ningun consumidor operativo encontrado lee directamente la hoja fisica `AGUA`.

## Detalle completo

## Sistema actual

### Flujo ejecutable

```text
2_planilla
  MES_ANTERIOR + MES_ACTUAL + MANTENIMIENTO
                         |
                         v
5_cobranza._componentes_cuenta()
  AGUA_ANT + MANT_ANT + AGUA_ACT + MANT_ACT
                         |
                         v
5_cobranza._aplicar_componentes()
  aplica FIFO, luego colapsa por concepto
                         |
                         v
shared/seguimiento_pueblo.xlsx
  CONCEPTO=AGUA o MANTENIMIENTO
                         |
                         v
seguimiento_repo.generar_vista()
  una hoja por CONCEPTOS_ORDEN
                         |
                         v
vista_seguimiento_pueblo.xlsx
  AGUA + MANTENIMIENTO agregados
```

### Writer y forma persistida

- `shared/seguimiento_repo.py` es el writer unico del ledger.
- El registro transitorio persiste `MZ`, `LT`, `CONCEPTO`, `MES`, `TIPO_EVENTO`,
  `CARGO`, `PAGO`, `AJUSTE`, `SALDO`, `SOURCE`, `AUDIT_REF`, `TIMESTAMP`, `CLASE` y
  `MOTIVO`.
- No persiste `CARGO_ID`, `MES_CARGO` ni una aplicacion pago-cargo.
- `CLASE` no representa antiguedad; representa el hecho contable, por ejemplo
  `GENESIS`, `COBRANZA` o `ABONO_REZAGADO`.

### Cascada real

```text
pago del ciclo:
AGUA_ANT -> MANT_ANT -> AGUA_ACT -> MANT_ACT -> CORTE -> CONVENIO -> ACUERDOS -> MULTA

abono de ciclo cerrado:
AGUA_ANT -> MANT_ANT -> CORTE -> CONVENIO -> ACUERDOS -> MULTA -> AGUA_ACT -> MANT_ACT
```

Los dos caminos conservan FIFO dentro de `AGUA` y dentro de `MANTENIMIENTO`: una
aplicacion agregada de `AGUA` siempre cubre primero `AGUA_ANT` y luego `AGUA_ACT`; una
aplicacion agregada de mantenimiento cubre primero `MANT_ANT` y luego `MANT_ACT`.

### Proyeccion temporal disponible

Para agosto, un `CARGO AGUA` con `SOURCE=saldo_inicial` y `AUDIT_REF=apertura|...` es
deuda anterior. Un `CARGO AGUA` con `SOURCE=2_planilla` es cargo actual. Desde el mes
siguiente, todo saldo final no cubierto del mes previo pasa al bloque anterior y los cargos
del nuevo mes entran al bloque actual. `MANT_ANT` se mantiene internamente para respetar
FIFO, pero se suma visualmente dentro de `MES_ANTERIOR`.

### Consumidores de la vista

- `7_cierre/consolidar_cierre.py` regenera la vista; no lee hojas individuales.
- `exportar_vista_pdf()` recorre las hojas de forma generica y admite nombres nuevos.
- `3_boletas/recibos_medidor_pagado.py` lee solo `CONVENIO_HISTORIAL`.
- `shared/validar_vista_boletas.py` consulta el ledger por API; no abre `AGUA`.
- `shared/tests/test_seguimiento_repo.py` si fija el conjunto de hojas y debe actualizarse.
- No se encontro codigo operativo que lea directamente `vista_seguimiento_pueblo.xlsx`
  hoja `AGUA`.

## Hallazgos

| ID | Hallazgo verificado | Consecuencia |
|---|---|---|
| H-01 | `CONCEPTOS_ORDEN` genera `AGUA` y `MANTENIMIENTO`. | La forma actual copia la taxonomia, no la necesidad visual. |
| H-02 | `_componentes_cuenta()` ya conserva antiguedad durante el calculo. | No se cambia la cascada de `5_cobranza`. |
| H-03 | `_aplicar_componentes()` agrega anterior y actual al escribir el concepto. | La vista debe re-derivar el split; no puede leerlo de `CLASE`. |
| H-04 | Ambos ordenes pagan anterior antes que actual dentro de cada concepto. | El split de pagos por FIFO es determinista con los eventos existentes. |
| H-05 | Agosto contiene 509 cargos AGUA actuales por S/6,187 y 105 aperturas por S/3,380. | La fuente distingue los dos tipos de deuda. |
| H-06 | Hay 374 pagos AGUA por S/5,107 y 346 pagos MANTENIMIENTO por S/1,036. | La proyeccion debe separar pagos efectivos y declarados sin alterar el total. |
| H-07 | Hay 0 ajustes AGUA/MANTENIMIENTO. | El caso vigente no exige inventar antiguedad para ajustes. |
| H-08 | Solo existe el mes 2026-08 en la cuenta completa. | La primera version se valida en agosto y debe probar sinteticamente el rollover a septiembre. |
| H-09 | La vista es regenerable y nunca fuente. | Se pueden reemplazar hojas sin migrar datos persistidos. |
| H-10 | PDF y consumidores no dependen de la hoja AGUA, salvo una prueba de estructura. | Impacto externo bajo y verificable. |

### Caso I-9 verificado

```text
entrada anterior:  8   -> MES_ANTERIOR deuda 8
cargo actual:      5   -> MES_ACTUAL deuda 5
mantenimiento:     3   -> MANTENIMIENTO deuda 3

pago AGUA:        13   -> anterior 8 + actual 5
pago MANT:         3   -> mantenimiento actual 3

saldo visible:     0 + 0 + 0
```

## Restricciones

- D-005: la cuenta completa empieza en agosto sin backfill.
- El ledger es append-only; generar la vista no escribe eventos.
- `AGUA` y `MANTENIMIENTO` siguen siendo conceptos validos internos.
- `MES_ANTERIOR` no se agrega como concepto: es antiguedad visible.
- La hoja `AGUA` desaparece solo del Excel/PDF de vista.
- Orden visible confirmado: `MES_ANTERIOR`, `MES_ACTUAL`, `MANTENIMIENTO`,
  `CORTE_RECONEXION`, `CONVENIO`, `ACUERDOS`, `MULTA`, `OTROS`.
- Mantenimiento anterior se suma a `MES_ANTERIOR`.
- Deben preservarse `PAGO` y `DECLARADO` como columnas distintas.
- Pruebas y regeneraciones usan temporales; no pisan el ledger real.
- El diseño destino con `CARGO_ID/MES_CARGO` no se implementa en este cambio.

## Incertidumbres

- Si aparece en el futuro un `AJUSTE` AGUA/MANTENIMIENTO sin cargo objetivo, no existe
  evidencia para asignarlo a anterior o actual. Las opciones deben comparar bloqueo
  explicito contra una categoria visible no clasificada; queda prohibido elegir un mes por
  heuristica silenciosa.
- Debe decidirse si la proyeccion se construye como filas mensuales preagregadas para
  reutilizar el writer visual o como eventos sinteticos. La salida debe ser identica.
- El orden entre conceptos para abonos cerrados difiere del pago del ciclo, pero no afecta
  el split anterior/actual dentro de `AGUA` y `MANTENIMIENTO`.
