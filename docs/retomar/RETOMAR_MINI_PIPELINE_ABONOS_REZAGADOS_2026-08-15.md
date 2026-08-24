# RETOMAR - Mini pipeline de abonos rezagados

## Estado confirmado

La mini-corrida reproduce el resultado de la corrida real para los tres casos revisados:

| Predio | Mini | Corrida real | Estado |
|---|---:|---:|---|
| F1-4 | S/0 | S/0 | coincide |
| H-21 | S/45 | S/45 | coincide |
| N-6 | -S/7 | -S/7 | coincide |

La equivalencia es de **entradas y calculo**:

- Usa `_cargar_pagos_yape()` de `5_cobranza`.
- Usa `_cargar_pagos_efectivo()` de `5_cobranza`.
- Usa `_cargar_planilla()` y los overlays reales.
- Usa `_cargar_abonos_rezagados()`.
- Usa `_calcular()` y `_descomponer_pago()` reales.
- Filtra despues de cargar por las claves presentes en `shared/abonos_rezagados.xlsx`.
- No escribe `shared/seguimiento_pueblo.xlsx` ni outputs reales.

No es aun una equivalencia de escritura: `mini_ledger_predicho.xlsx` es la proyeccion de eventos; la escritura real mediante `_reconciliar_pagos_pueblo()` sigue siendo la etapa posterior.

## Manifiesto actual

- Confirmados historicos: 21.
- Activos para agosto: 16.
- Total de abonos activos: S/820.
- F1-4: un abono de S/101 del ciclo 2026-08, confirmado visualmente en
  `shared/abono_rezagado/5.jpeg`; reemplaza el consolidado S/117 sin respaldo exacto
  (decision explicita del usuario, 2026-08-22).
- H-21: un abono de S/50.
- N-6: un abono de S/70.
- Q-5 S/45 fue retirado de la fuente bloqueante.

## Mini pipeline

Runner:

`5_cobranza/tests/generar_mini_corrida_abonos.py`

Carpeta de la ultima corrida:

`C:\Users\wilde\AppData\Local\Temp\opencode\mini_corrida_abonos_20260815`

Archivos principales:

- `inputs/abonos_rezagados.xlsx`
- `inputs/pagos_yape_filtrados.xlsx`
- `inputs/pagos_efectivo_filtrados.xlsx`
- `inputs/ledger_subset.xlsx`
- `outputs/mini_resultado_cascada.xlsx`
- `outputs/mini_ledger_predicho.xlsx`

Ultimos conteos:

- 42 filas en la fuente de abonos.
- 4 pagos Yape filtrados, S/35.
- 12 pagos efectivo filtrados, S/237.
- 175 filas del ledger reducido.

## Backup del ledger

`shared/backups_ledger/seguimiento_pueblo_pre_mini_ledger_20260815_140000.xlsx`

Tambien existe el checkpoint Git:

`checkpoint-abonos-20260815`

## Correccion aplicada

`_cargar_abonos_rezagados()` devuelve claves `(MZ, LT)`, mientras `_calcular()` las buscaba como string `MZ-LT`. Eso hacia que el manifiesto pasara pero el abono no entrara al calculo. Se corrigio la busqueda en `5_cobranza/main.py` y se agrego una asercion de F1-4 al test.

## Forma de trabajo siguiente sesion

```text
1. Modificar la politica de aplicacion retroactiva en el mini pipeline.
2. Correr el mini pipeline con todos los abonos.
3. Revisar por predio: abono, agua, corte, multa, acuerdos, convenio y saldo.
4. Corregir la proyeccion del mini ledger hasta que el resultado sea aprobado.
5. Aplicar exactamente esas modificaciones al ledger real.
6. Ejecutar 5_cobranza --force una sola vez.
7. Verificar 5b_validacion, 6_corte y los outputs finales.
```

## Advertencias pendientes

- H-21 termina en S/45 y N-6 en exceso de S/7 con la cascada actual; no se deben aceptar como resultado de negocio sin decidir la aplicacion retroactiva contra cargos de julio.
- La mini-corrida actual no escribe el ledger; solo proyecta el ledger.
- No mezclar planillas historicas con el ledger vivo actual. El mini debe usar la planilla viva del pipeline y el ledger actual como base de la nueva reimputacion.
- Antes de modificar el ledger real, crear otro backup inmediato y comparar la proyeccion contra el cambio esperado fila por fila.
