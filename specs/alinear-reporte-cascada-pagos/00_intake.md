# Intake: Alinear reporte con cascada de pagos

## Pedido original

Wilder solicita corregir el reporte historico para que los pagos agregados manualmente a
`shared/abonos_rezagados.xlsx` se muestren y repartan con la misma cascada oficial que
usa `5_cobranza`.

```text
abono manual
    |
    +-> MES_CICLO: mes real del pago y mes visible en el reporte
    +-> MES_ANO_APLICA: mes de regularizacion contable
    |
    v
MES_ANTERIOR -> MES_ACTUAL -> MANTENIMIENTO
             -> CORTE -> CONVENIO -> ACUERDOS -> MULTA
```

El flujo OCR de abonos rezagados queda congelado y fuera de este cambio. El resultado
esperado es que el reporte no cambie el orden de la cascada, no desplace el pago al mes
de regularizacion y no cuente la aplicacion ledger como un segundo ingreso.

## Evidencia disponible

### Hechos observados

- `5_cobranza/main.py::_CAMPOS_WATERFALL_CICLO_VIGENTE` ya prioriza
  `mes_anterior -> mes_actual -> mantenimiento -> corte -> convenio -> acuerdos -> multa`.
- `4b_reclamos/herramienta/comun.py::_datos_ciclo()` reparte hoy el pago vigente como
  `consumo -> mantenimiento -> mes_anterior -> corte -> convenio -> acuerdos -> multa`.
- `comun.py::_abonos_rezagados_predio()` agrupa actualmente por `MES_ANO_APLICA`, aunque
  la referencia de pago ya puede mostrar `MES_CICLO`.
- F1-4 tiene una sola fila manual por S/101 con `MES_CICLO=2026-06` y
  `MES_ANO_APLICA=2026-08`, respaldada por `shared/abono_rezagado/5.jpeg`.
- Para F1-4, la deuda reconstruida de junio es mes anterior S/29, consumo S/29,
  mantenimiento S/3 y multa S/50. La cascada acordada aplica los S/101 como
  `29 + 29 + 3 + 40` y deja S/10 de multa.
- El reporte actual muestra la referencia en junio, pero ubica S/79 aplicados en agosto.
- El ledger conserva en agosto un `PAGO MULTA S/50` con `CLASE=ABONO_REZAGADO`; representa
  la aplicacion contable de parte del abono, no otro ingreso de caja.
- Los 19 eventos ledger vinculados a abonos rezagados observados tienen una fila fuente
  coincidente por predio y mes de aplicacion, pero el ledger transitorio no conserva
  `ABONO_ID` individual.

### Fuentes y consumidores relevantes

- Fuente manual: `shared/abonos_rezagados.xlsx`, hoja `Abonos_Raw`.
- Lector y transformacion del reporte: `4b_reclamos/herramienta/comun.py`.
- Render y estado de aplicacion: `4b_reclamos/herramienta/reporte_historico.py`.
- Cascada de referencia: `5_cobranza/main.py`.
- Pruebas enfocadas: `4b_reclamos/tests/test_reporte_provisional.py` y
  `5_cobranza/tests/test_abonos_manifest.py`.

## Preguntas iniciales

### Confirmado

- Alcance: solo alinear el reporte de pagos manuales con la cascada oficial.
- `MES_CICLO` determina donde se muestra el pago; `MES_ANO_APLICA` conserva cuando se
  regularizo.
- El orden dentro del bloque de agua comienza por `MES_ANTERIOR`.
- La solucion debe ser general; F1-4 es un caso de regresion, no una condicion especial.
- No se modifica el flujo OCR congelado.
- No se modifica ni reescribe el ledger append-only.
- No se cambia la cascada de `5_cobranza`, que ya refleja la regla confirmada.
- Wilder elige opciones y aprueba cada gate SDD.

### Por resolver en exploracion y opciones

- Como representar por separado el ingreso en `MES_CICLO` y su aplicacion ledger en
  `MES_ANO_APLICA` sin doble conteo.
- Como conciliar de forma segura eventos agregados que todavia no tienen `ABONO_ID`.
- Que pruebas sinteticas cubren multiples abonos del mismo predio y preservan pagos
  normales no afectados.
