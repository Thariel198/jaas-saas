# Fix (traído desde julio): saldos negativos falsos en seguimiento_pueblo.xlsx

**Fecha del incidente:** 2026-07-13 (en `jass_system - Julio`). **Fecha del fix:** 2026-07-27.

## Contexto — por qué esto está en agosto

Al generar boletas de agosto con `3_boletas`, 27 predios salieron con Convenio/Multa/
Cuota directa en negativo en `DATA_boletas.xlsx` (un "crédito a favor" falso). El origen
real estaba en julio: `5_cobranza` leyó `pagos_yape_tepago.xlsx` a medio guardar (el
usuario estaba subiendo un reporte nuevo en simultáneo, 13/07 05:51), interpretó que el
pago había bajado y escribió un `AJUSTE` negativo; 25 minutos después, con el archivo ya
completo, volvió a registrar el mismo pago. Como `seguimiento_pueblo.xlsx` es
append-only, ambos eventos quedaron pegados para siempre — 36 pares (predio, concepto)
con saldo negativo falso.

Detalle completo del incidente y el fix en
`jass_system - Julio/docs/RETOMAR_fix_race_condition_yape_seguimiento_pueblo_2026-07-27.md`
(carpeta separada, no conectada a esta).

## Qué se hizo

1. Se corrigió el ledger en `jass_system - Julio` (36 `AJUSTE` de compensación exacta vía
   `seguimiento_repo.registrar_ajuste`, saldo a 0 en cada uno).
2. Se re-corrió `5_cobranza --force` ahí para regenerar `arrastre_consolidado_2026-07.xlsx`.
3. Se copiaron a esta carpeta (agosto) los 2 archivos ya corregidos:
   - `shared/seguimiento_pueblo.xlsx`
   - `5_cobranza/outputs/arrastre_consolidado_2026-07.xlsx`

   (Verificado antes: la copia de agosto era byte-idéntica a la de julio — sin cobranza
   corrida todavía acá — así que no se perdió nada al sobreescribir.)

## Pendiente en esta carpeta (agosto)

Como `2_planilla`/`3_boletas` de agosto ya habían corrido con los datos viejos (antes de
este fix), hay que **volver a correr `2_planilla/main.py` y regenerar `DATA_boletas.xlsx`
+ boletas** para que los 27 predios queden con el monto real (ya no negativo).
