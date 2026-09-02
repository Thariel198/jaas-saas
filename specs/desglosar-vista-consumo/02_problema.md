# Problema: Desglosar consumo en vista seguimiento

## Resumen ejecutivo (lectura de 1 minuto)

```text
saldo correcto en ledger -> hoja AGUA agregada -> operador no ve antiguedad
```

- La vista no permite distinguir deuda anterior, consumo actual y mantenimiento actual.
- Tampoco permite comprobar en que bloque termino el pago agregado de `AGUA`.
- La salida necesaria son tres hojas ordenadas por cascada, sin modificar el ledger.

## Detalle completo

## Brecha

El ledger conserva correctamente los totales de `AGUA` y `MANTENIMIENTO`, pero
`vista_seguimiento_pueblo.xlsx` replica esos conceptos como hojas agregadas. Para un
operador, `AGUA=13` no explica si corresponde a deuda anterior, consumo actual o ambos,
ni permite comprobar que un pago cubrio primero la parte antigua.

La vista necesaria debe mostrar por separado `MES_ANTERIOR`, `MES_ACTUAL` y
`MANTENIMIENTO`, en ese orden, manteniendo exactamente los totales y saldos del ledger.
El mantenimiento impago de ciclos previos pertenece visualmente a `MES_ANTERIOR`.

## Objetivos

1. Reemplazar la hoja visible `AGUA` por `MES_ANTERIOR` y `MES_ACTUAL`.
2. Mantener `MANTENIMIENTO` solo para el cargo y saldo del ciclo actual.
3. Ordenar hojas conforme a la cascada visible confirmada.
4. Derivar pagos por FIFO sin cambiar eventos, conceptos ni saldos persistidos.
5. Conservar `DEUDA`, `PAGO`, `DECLARADO`, `AJUSTE` y `SALDO` en la vista.
6. Mantener el Excel y PDF regenerables e idempotentes.

## No-objetivos

- No cambiar `5_cobranza` ni su cascada.
- No agregar conceptos al ledger.
- No escribir ajustes ni migrar `seguimiento_pueblo.xlsx`.
- No modificar `reporte_historico.py`.
- No implementar `CARGO_ID`, `MES_CARGO` o aplicaciones individuales del ledger destino.
- No reconstruir meses anteriores a agosto de 2026.

## Métricas de éxito

| Metrica | Evidencia | Umbral |
|---|---|---|
| I-9 se entiende | Vista temporal | Anterior 8/8/0, actual 5/5/0, mantenimiento 3/3/0 en deuda/pago/saldo |
| Totales invariantes | Comparacion AGUA+MANT antes/despues | Diferencia 0.00 por predio y mes |
| Ledger intacto | Hash y conteo | 0 eventos modificados |
| Orden de hojas | Workbook generado | MES_ANTERIOR, MES_ACTUAL, MANTENIMIENTO, CORTE_RECONEXION, CONVENIO, ACUERDOS, MULTA, OTROS |
| AGUA retirada | Workbook generado | No existe hoja AGUA |
| Rollover correcto | Prueba sintetica agosto-septiembre | Saldo actual impago pasa a MES_ANTERIOR |
| Consumidores | PDF, recibos y validadores | Caso afectado y no afectado pasan |

## Afectados

- Secretaria, tesorero y directiva que consultan deuda.
- `shared/seguimiento_repo.py::generar_vista()` y `exportar_vista_pdf()`.
- `shared/tests/test_seguimiento_repo.py` y pruebas nuevas de proyeccion.
- `7_cierre`, que regenera la vista sin cambiar su contrato contable.
- `CONVENIO_HISTORIAL`, recibos y validadores no cambian, pero se verifican por regresion.
