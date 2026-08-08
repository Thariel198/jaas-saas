# Decisión de diseño — verificación de lotes (sub-módulo de 4_pagos/efectivo)

Fecha: 2026-08-08
Estado: Aprobado en conversación · Fase 2.0.6

---

**Problema:**
El cobrador escribe `MZ`, `LT` y `MONTO` en papel; alguien los transcribe a `mesa_N.xlsx`.
Nadie cuestiona nunca que ese `MZ-LT` sea el del vecino que pagó. Un error de una letra o
un dígito acredita el pago a otro predio: el que pagó sigue debiendo (y reclama), el que no
pagó aparece al día. El error viaja sin obstáculo hasta el ledger.

Los tres ejes de verificación que ya existen en el módulo no pueden atraparlo:
`main.py` cruza cobrador contra cobrador (inútil si la mesa tiene una sola hoja),
`arqueo.py` cuadra la plata (cuadra perfecto: el monto está bien, el lote no),
`reclamos.py` detecta lo que el cobrador marcó (no marca nada, para él no pasó nada).

Evidencia real, ciclo 2026-08: dos errores encontrados a mano, invisibles para los tres ejes.

**Criterios:**
- **Un falso OK es peor que un no-sé.** Si el monto no alcanza para probar el lote, el reporte
  lo dice. Nada de booleano.
- Auditable fila por fila por una persona no técnica — la secretaria mira una fila y entiende
  por qué está marcada, sin leer código ni un score numérico.
- El trabajo manual nunca se pierde, aunque el humano corrija el `MZ-LT` — que es justamente
  lo que va a corregir.

**Enfoque elegido:**
Sub-módulo de `4_pagos/efectivo` (el output es control, no lo consume ningún módulo downstream).
Script propio `verificar_lotes.py`, corre **antes** de `main.py` — verificar antes de consolidar,
no después de que el error ya viajó. Avisa, no bloquea (igual que `discrepancias.xlsx` hoy).
**Solo reporta: nunca escribe en `mesa_N.xlsx`**, que el README declara "manual — sagrado".

Cuatro capas, en este orden:

| # | Capa | Pregunta | Salida |
|---|---|---|---|
| ① | Cuadre con su boleta | ¿el monto es alguna combinación de los cargos de ESE lote? | CUADRA / NO CUADRA |
| ② | Fuerza de evidencia | ¿cuántos lotes del pueblo deben ese mismo importe? | ALTA / MEDIA / BAJA |
| ③ | Vecindad de confusión | ¿qué lotes *confundibles* deben ese monto? | lista de candidatos |
| ④ | Filtro de realidad | ¿el candidato está impago? | descarta pistas falsas |

- **Fuente de la deuda: `3_boletas/inputs/DATA_boletas.xlsx`, columna `Importe a pagar`** y las
  7 columnas de cargo. **No la planilla**: su `TOTAL_A_PAGAR` es fórmula Excel → pandas lee `NaN`.
  Además la boleta es la foto de lo que se le pidió al vecino; la planilla sigue mutando.
- **Capa ③ — la tabla de confusión.** Grupos de manzana que se confunden a mano o por
  autocompletar: `U/V/W` · `G/O/Q/C` · `I/L/T/J` · `P/D/O/B` · `A/D` · `B/P` · `E/F/G` ·
  `M/N/W` · `S/Z` · `X/Y/K` · `R/P/B`, más `X ↔ X1` (A/A1 … H/H1). Dígitos: `4↔9 · 1↔7 · 3↔8 ·
  5↔6 · 0↔6 · 0↔8 · 6↔8 · 2↔7 · 1↔4 · 3↔5 · 5↔8 · 7↔9`, transposición (`12↔21`), dígito de
  más/de menos (`A-4↔A-14`, `B-2↔B-12` — el autocompletar), sufijo (`3↔3A`, `8↔8B`).
- **Error simple gana sobre error doble.** Un candidato con la manzana mal *o* el lote mal se
  reporta primero; con los dos mal, en segundo nivel. Medido: en los 4 casos conocidos el
  correcto siempre fue un error simple, y el nivel doble solo aportó ruido.
- **El candidato solo se propone cuando la lista queda en 1.** Con 17 candidatos el reporte dice
  "17 posibles", no elige — proponer con esa evidencia es inventar con cara de certeza.
- **Output `outputs/verificacion_lotes_YYYY-MM.xlsx`** — vista operacional temporal, mismo patrón
  que `discrepancias.xlsx`: el supervisor llena la columna de resolución y el archivo desaparece
  cuando no queda nada sin resolver.
- **Guard de pipeline en dos capas**, antes de verificar nada: (1) `shared/ciclo_activo.json`
  existe y declara `mes_ano` (lo escribe `1_lecturas`) — si no, parar: "no corriste 1_lecturas";
  (2) las `FECHA` de pago de las mesas caen dentro de la ventana `FECHA DE EMISIÓN` →
  `FECHA DE VENCIMIENTO` de `DATA_boletas` — si no, parar: "`DATA_boletas` es vieja, no corriste
  3_boletas". Se descartó un `estado_pipeline.json` nuevo: un JSON de estado puede mentir
  (alguien lo edita, un módulo corre a medias); comparar fechas contra el dato real no puede.
- **Preservación en 3 capas**, copiando el patrón de `motor_matching`: backup + leer decisiones
  humanas (solo filas donde el humano escribió al menos un campo fuerte, como
  `_leer_pendientes_preservados`) + la decisión manual gana sobre el automático en la re-corrida
  (`main.py:1832`, ciclo 2+). **Clave de preservación `(MESA, COBRADOR, FECHA, MZ, LT, MONTO,
  FILA_EXCEL)`** — los primeros 6 campos son los mismos que el `ABONO_ID` de `libro_mayor/caja`
  (menos `jass`, single-tenant hoy); medidos contra los 165 pagos reales de este ciclo dan
  **0 colisiones**. `FILA_EXCEL` (posición física de la fila en la hoja) se agrega como
  **respaldo puro** para el caso que esos 6 campos no cubren: el mismo lote pagando el mismo
  importe dos veces el mismo día con el mismo cobrador — raro, no observado este ciclo, pero
  posible. En el Excel es **columna oculta** (`hidden=True`): la preservación nunca falla, pero
  el supervisor no la ve ni la usa — busca por `MZ` como cualquier columna visible.
  Descartada `(MESA, COBRADOR, FECHA, MONTO)` **sin** `MZ, LT` (primer intento, copiado sin
  verificar del `ABONO_ID`): esa sí colisiona en 62% de las filas — es la ausencia de `MZ, LT`
  la que causaba el choque, no la de un contador de fila.
  Nota sobre la corrección: al re-correr después de que el supervisor corrige `MZ-LT` en
  `mesa_N.xlsx` a mano, la clave visible cambia (porque `MZ, LT` cambiaron) y la fila deja de
  emparejar con su versión anterior — eso es correcto, no una pérdida: la fila recalcula limpia
  contra su nueva boleta (pasa a `CUADRA` si la corrección era la buena, o muestra un `NO CUADRA`
  nuevo y genuino si no lo era). Mientras el supervisor no haya tocado `mesa_N.xlsx` (por
  ejemplo, marcó `RESOLUCION` pero todavía no aplicó el cambio a mano), la clave no cambia y la
  preservación funciona sin más.

**Decisiones de negocio (cerradas en conversación):**
1. Se compara contra `MONTO` (efectivo + yape), no contra `MONTO_EFECTIVO`. La deuda es una sola;
   verificar media deuda no tiene sentido.
2. `MONTO = 0` se omite — son los registros de reclamo (el cobrador anota la visita sin cobro).
   `CATEGORIA=reclamo` ya los captura por otra vía.
2b. `CONCEPTO` no vacío (tanque · honorario · gasto · comunitario) también se omite —
   misma familia que el punto 2: plata que `5_cobranza` ya excluye de la deuda de agua, así
   que nunca va a cuadrar contra `DATA_boletas` y generaría ruido `NO CUADRA` permanente.
   Verificado en este ciclo: `CONCEPTO` vacío en las 163 filas — el campo existe pero no se
   usó, por eso no apareció solo en la medición. Es el mismo campo que causó el bug de A-4
   (2026-08-06) cuando quedó sin marcar.
3. Un pago = un lote = una fila. No se modela el pago multi-lote: el vecino que paga por dos
   predios genera dos filas.
4. Cuando dos combinaciones distintas de cargos dan el mismo monto (8 de 157 filas medidas, todas
   de la forma "mes anterior" vs "consumo+mantenimiento"), **la ambigüedad no cambia el veredicto**
   — el lote queda confirmado por cualquiera de las dos. Solo afecta la etiqueta informativa de
   qué conceptos se pagaron; ahí se desempata con el orden de la cascada.
5. La verificación por monto **no reemplaza al nombre**. Con montos comunes el monto no prueba
   nada, y ahí solo el nombre resuelve. El reporte cuantifica cuántas filas no puede verificar:
   ese número es la medida del problema, no un defecto del método.

**Evidencia medida (163 filas reales, ciclo 2026-08 — Fase 0):**

```
el pueblo            557 boletas · solo 116 importes distintos
                     S/8 → 101 lotes (18%) · S/16 → 24 · S/18 → 17 · S/9 → 17

capa ① sola          157 verificables → 151 cuadran · 6 no cuadran
                     prueba contrafactual: 2 de 2 errores reales detectados
                       Magda  M-19 (boleta 18) pagó 9 → no cuadra · M-14 (boleta 9) → cuadra
                       Pedro  G-13 (boleta 37) pagó 19 → no cuadra · O-13 (boleta 19) → cuadra

capa ② evidencia     ALTA 25 · MEDIA 75 · BAJA 51 (31%) · NO CUADRA 6
                     (una definición previa, contando cualquier subconjunto como rival,
                      daba 79% BAJA — inservible. Se mide contra el importe total.)

capa ③ confusión     Magda   17 candidatos →  1 (el correcto)
                     Pedro   10 candidatos →  1 (el correcto)
                     W-2/U-2  212 (parcial) → 10 → 3 aplicando ④
                     Q-5/O-5  213 (parcial) →  5 → 2 restringiendo a error simple
```

**Corrida real del módulo (2026-08-08, las 7 mesas — Fase 4.1):**

```
326 pagos · 557 boletas
   ALTA 37 · MEDIA 182 · BAJA 76 · NO CUADRA 19 · OMITIDO 7 · SIN BOLETA 5

   BAJA = 76/295 evaluables = 26%   (umbral de alerta: 50%)
   NO CUADRA = 19 filas, de las cuales 4 con candidato único propuesto:
      H1-15 → H1-16 [lote 15→16] · G-14 → E-14 [manzana G→E]
      O-17  → O-27  [lote 17→27] · A-1  → A-2  [lote 1→2]

tests   21/21 en tests/test_verificar_lotes.py — incluye los 2 CONTRAFACTUALES:
        con M-19/S9 y G-13/S19 (los lotes mal escritos) el módulo marca NO CUADRA
        y propone M-14 y O-13. Es la prueba de que habría atrapado los errores reales.
regresión  test_unitarios 27/27 · test_integracion 16/16 (main.py intacto)
```

**Corregido durante la implementación:** la primera versión proponía candidatos de
**error doble** (manzana Y lote mal) cuando no había ninguno simple — `G-17→Q-12`,
`Q-14→C-44`, ambos sin relación real. Eso es exactamente el falso-OK que el criterio ①
prohíbe. Ahora solo se propone nivel simple; los dobles se cuentan y se reportan como
`"N solo con doble error"` sin elegir.

**Alternativas descartadas:**
- *Verificador puro (solo capa ①)* — viola el criterio principal: le da verde a 128 filas con la
  misma confianza aparente que a las 25 sólidas. Un S/8 que cuadra no prueba nada y el reporte no
  lo diría.
- *Matcher inverso como salida (estilo `motor_matching` con yape)* — proponer "el lote correcto es
  X" a partir de un monto que comparten 17 vecinos es exactamente el falso-OK que el criterio ①
  prohíbe, y una lista de 17 candidatos no es auditable. Sobrevive como **componente interno**: su
  índice es lo que alimenta las capas ② y ③.
- *Score numérico de plausibilidad* — un 0.73 no le dice nada a la secretaria. Viola el criterio de
  auditabilidad. Se reemplaza por nivel categórico + motivo en texto.
- *Cruzar contra la planilla en vez de `DATA_boletas`* — `TOTAL_A_PAGAR` es fórmula (pandas lee
  `NaN`) y la planilla muta después de emitida la boleta.
- *Que el script corrija `mesa_N.xlsx`* — el README declara ese archivo "manual — sagrado", y
  corregir automáticamente un `MZ-LT` con evidencia débil es el error que el módulo existe para
  evitar.

**Señal de alerta:**
Si más del 50% de las filas cae en evidencia BAJA, el monto dejó de discriminar en esta JASS
(hoy: 31%). La respuesta entonces **no es afinar el algoritmo** — es que falta la columna `NOMBRE`
en la hoja de papel del cobrador. Segunda señal: si el supervisor deja de llenar la columna de
resolución, el reporte está marcando demasiado y perdió credibilidad; revisar el umbral de
evidencia antes que agregar reglas. Tercera: si la capa ③ empieza a devolver listas largas en vez
de 1-3 candidatos, la tabla de confusión creció de más y hay que podarla — cada grupo nuevo tiene
que justificarse con un error real observado, no con una intuición.

---

## Escala (lente de tenant)

Los grupos de confusión de **letras** son universales — cómo se confunde una U con una W no depende
de la JASS. Pero `X ↔ X1` depende de **cómo esta JASS nombra sus manzanas** (`A`…`Z`, `A1`…`H1`):
eso va al manifiesto del tenant, no al código. Lo mismo la lista de conceptos de la boleta, hoy 7
columnas de `DATA_boletas`: son data del tenant, no columnas hardcodeadas del motor.
