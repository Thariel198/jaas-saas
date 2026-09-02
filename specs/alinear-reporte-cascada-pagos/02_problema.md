# Problema: Alinear reporte con cascada de pagos

## Resumen ejecutivo (lectura de 1 minuto)

```text
distribucion correcta en 5_cobranza + saldo correcto en ledger
                              |
                              v
vista y reporte muestran otra distribucion o pierden antiguedad
                              |
                              v
operador no puede explicar que deuda cubrio cada pago
```

- El problema no es monetario: los saldos de I-9 cuadran.
- El problema es de lectura: S/8 de mes anterior aparecen sumados al consumo actual.
- Los abonos rezagados tambien pueden mostrarse en un periodo o cascada distintos de los
  usados por `5_cobranza`.
- El cambio debe alinear vista, reporte y consumidores sin modificar eventos del ledger.

## Detalle completo

## Brecha

`5_cobranza` distribuye pagos por antiguedad y prioridad, y el ledger conserva el saldo
resultante por concepto. Sin embargo, sus vistas de lectura no siempre conservan el mismo
significado:

- `vista_seguimiento_pueblo.xlsx` agrega toda la deuda y pago de `AGUA`, sin distinguir
  saldo anterior de cargo actual;
- `reporte_historico.py` muestra todo `AGUA` de agosto como `CONSUMO`, aunque parte sea la
  apertura del mes anterior;
- la reconstruccion usada por herramientas historicas puede aplicar abonos rezagados en el
  periodo de regularizacion y con una cascada distinta de la operativa;
- la generacion individual, por lote y `buscar_pago.py` pueden explicar de forma distinta
  un mismo pago.

El resultado observable es una cuenta cuyo total cuadra, pero cuya distribucion visual no
permite saber que parte pago el mes anterior, el consumo actual o el mantenimiento. I-9
muestra la brecha exacta: el sistema presenta `CONSUMO=13` y `MES_ANTERIOR=0` cuando la
deuda real visible es `MES_ANTERIOR=8`, `MES_ACTUAL=5`, `MANTENIMIENTO=3`.

## Objetivos

1. Mostrar la misma distribucion que aplico `5_cobranza`, respetando periodo real del pago,
   periodo de aplicacion y prioridad por antiguedad.
2. Exponer `MES_ANTERIOR`, `MES_ACTUAL` y `MANTENIMIENTO` como categorias visibles; el
   mantenimiento anterior se incluye dentro de `MES_ANTERIOR`.
3. Mantener iguales los totales de deuda, pago, ajuste y saldo del ledger.
4. Dar el mismo resultado en reporte individual, reporte por lote y consumidores de la
   tabla historica.
5. Mantener la vista regenerable: borrar y recrear el Excel no cambia ningun hecho.
6. Presentar cada artefacto SDD con resumen ejecutivo y detalle completo.

## No-objetivos

- No cambiar la cascada operativa de `5_cobranza`.
- No agregar `MES_ANTERIOR` como concepto del ledger.
- No modificar, borrar ni compensar eventos reales con ajustes.
- No hacer backfill anterior a agosto ni cambiar la frontera historica D-003/D-005.
- No redefinir que abonos rezagados estan autorizados.
- No implementar todavia el esquema destino con `CARGO_ID`, `ABONO_ID` y `MES_CARGO`.
- No corregir casos por lote, monto o fecha mediante excepciones.

## Métricas de éxito

| Metrica | Evidencia | Umbral |
|---|---|---|
| I-9 agosto se distribuye correctamente | PDF, Excel y vista regenerados desde datos reales | Anterior 8/8, actual 5/5, mantenimiento 3/3 en deuda/pago; saldo 0 |
| F1-4 conserva periodo y cascada | Prueba sintetica y reporte | La referencia conserva ciclo real y la aplicacion no se desplaza silenciosamente |
| Totales invariantes | Comparacion antes/despues por predio y mes | Diferencia 0.00 en deuda, pago, ajuste y saldo total |
| Ledger intacto | Hash/conteo del archivo antes y despues de generar vistas | 0 eventos agregados, editados o borrados |
| Entry points consistentes | Caso individual y lote para los mismos predios | Mismos importes por categoria y mes |
| Consumidor sin regresion | Pruebas de `buscar_pago.py` | Caso afectado y caso no afectado pasan |
| Vista regenerable | Generacion en temporal repetida | Dos corridas producen el mismo contenido logico |

## Afectados

- Secretaria, tesorero y directiva que leen la vista y el reporte para explicar deudas.
- Vecinos que reciben un historial de cuenta.
- `shared/seguimiento_repo.py`, generador de `vista_seguimiento_pueblo.xlsx` y PDF.
- `4b_reclamos/herramienta/reporte_historico.py` y helpers de lectura compartidos.
- `buscar_pago.py`, que consume la distribucion historica para resolver reclamos.
- Pruebas y validadores de `shared`, `4b_reclamos` y consumidores directos de la vista.
- No se afectan el writer del ledger, `2_planilla`, `5_cobranza` ni `7_cierre` en su
  comportamiento contable.
