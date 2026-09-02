# Prompt — referencias_pago(): leer columna real de Efectivo/Yape (oct-2025 → may-2026)

## CONTEXTO

`4b_reclamos/reporte_referencias_pago.py`, función `referencias_pago()`, rama
histórica (líneas 137-169, un archivo por mes vía `_ARCHIVOS_HISTORICOS`).

Hoy, para un pago en efectivo, el código hace:

```python
out.append({"MES": mes, "MEDIO": "EFECTIVO", "FECHA_HORA": "--", "MONTO": monto_mes})
```

`monto_mes` es el `TOTAL` que calculó `_fila_historica()` (suma de TODOS los
conceptos de deuda del mes: Consumo+Mant+Mes_anterior+Corte+Convenio+Multa+
Acuerdos) — no una lectura de la columna de efectivo del predio. El docstring
del módulo (líneas 6-10) promete leer "la hoja Efectivo dedicada" desde
feb-2026; eso no existe en el código — `_cargar_hojas_historicas()` (línea
103-118) solo abre "Cobranza" y "Reporte".

## REALIDAD VERIFICADA (no supuesta — confirmada contra los 8 xlsx reales)

Fuente: `obligaciones/inputs/planillas anteriores/`, hoja "Cobranza", los 8
archivos de `_ARCHIVOS_HISTORICOS` (2025-10 → 2026-05).

- `Estado == 'c'` → pagó ese mes. `NaN` → no pagó. Igual en los 8 archivos.
- `Medio == 'y'/'Y'` → yape. `Medio` vacío/`NaN` → efectivo. Igual en los 8.
- Columna **"Efectivo"** y columna **"Yape"**: existen literalmente en LOS 8
  archivos (no solo desde feb-2026 como decía el docstring).
- 0 filas reales con Efectivo>0 y Yape>0 a la vez (solo la fila de TOTAL al
  pie del sheet, que no tiene MZ/LT y ya se excluye igual que en
  `_fila_historica()`).
- `Total == Efectivo + Yape` para TODAS las filas pagadas con MZ/LT en los 8
  archivos: **0 diferencias en 2673 filas revisadas.**
- Columnas **"Monto Efectivo"** y **"Yape rep"** solo existen desde feb-2026
  (4 de 8 archivos) y NO son un duplicado confiable de "Efectivo"/"Yape":
  - 2026-03 P-9: `Monto Efectivo=NaN` pero el pago real fue Yape S/435.
  - 2026-04 K-8/K-9/O-12/W-6: `Monto Efectivo=8` fijo en las 4 filas, pero el
    pago real (columna "Efectivo") fue S/117, S/22, S/116, S/16 respectivamente.
  - Total 5 filas de 4 archivos donde "Monto Efectivo"/"Yape rep" no cuadra
    contra `Total` (que sí es confiable, ver punto anterior).
- `_col(df, "Efectivo")` y `_col(df, "Yape")` resuelven por match EXACTO a
  las columnas "Efectivo"/"Yape" literales, nunca a "Monto Efectivo"/"Yape
  rep" — confirmado ejecutando `_col()` contra el archivo de febrero (donde
  coexisten las 4 columnas). Sin colisión de nombres.
- No existe columna de fecha/hora para efectivo en ningún archivo histórico
  (a diferencia de yape, que cruza con la hoja "Reporte").

## DECISIONES YA CERRADAS (no reabrir sin evidencia nueva)

1. Usar las columnas **"Efectivo"** y **"Yape"** — están en los 8 archivos y
   siempre cuadran contra `Total`. NO usar "Monto Efectivo"/"Yape rep": están
   incompletas (feb-may únicamente) y ya se demostró que fallan en 5 filas
   reales.
2. Si no existe un dato real (fecha/hora de efectivo), no se pone. Nada de
   placeholder `"--"` que aparente ser un dato — se omite el campo o queda
   vacío explícito.
3. Sin ambigüedad de nombre de columna: `_col(df, "Efectivo")` /
   `_col(df, "Yape")` ya resuelven exacto, confirmado por ejecución real, no
   por lectura de código.

## CAMBIO A IMPLEMENTAR

En `referencias_pago()`, líneas 156-169, rama histórica: reemplazar el uso
de `monto_mes` (el `TOTAL` agregado de conceptos) por el valor real de la
columna `"Efectivo"` o `"Yape"` del predio en ese archivo/mes, usando
`_col(df, "Efectivo")` / `_col(df, "Yape")` ya existentes en
`reporte_historico.py` (importado como `rh`).

Como `Total == Efectivo + Yape` siempre (0 diferencias verificadas), este
cambio **no altera ningún monto ya mostrado** — es leer la columna que
efectivamente corresponde al medio de pago en vez de reusar el total de
deuda del mes, que hoy coincide por construcción (nunca hay pago mixto).

No tocar la rama de yape (líneas 157-167) ni la lógica de fecha/hora de
yape (`hojas.get("Reporte")`) — esa parte no está en discusión acá.

## VERIFICACIÓN DESPUÉS DEL CAMBIO

Repetir la comparación `Total vs Efectivo+Yape` sobre los 8 archivos y
confirmar que el PDF de referencias para T-12, L-4 y F-9
(`4b_reclamos/reporte_referencias_pago.py::generar_pdf`) sigue mostrando los
mismos montos en efectivo que mostraba antes del cambio (deben coincidir
número por número, solo cambia la fuente de lectura).
