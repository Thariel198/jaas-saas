# RETOMAR — verificación de lotes en efectivo · 2026-08-08

Bitácora con el razonamiento completo: `docs/diario/2026-08-08_verificacion_lotes_efectivo.html`.

---

## ⚡ PRIMER PASO al retomar

**Verificar Maximo Encarnacion (mesa_3) y Wagner Trujillo (mesa_4) contra sus fotos
nuevas** — mismo método que se usó hoy con Wilder (mesa_1) y Yerald (mesa_2):

```
4_pagos/recursos/Maximo_Encarnacion/fotos_2/
4_pagos/recursos/Wagner_Trujillo/fotos_2/
```

`verificar_lotes.py` ya corrió sobre las 7 mesas y encontró **4 candidatos con lote único
propuesto**, sin verificar todavía — dos de ellos caen justo en las mesas que faltan:

```
mesa_2  (Yerald)   H1-15 → H1-16  [lote 15→16]    Marleny Samaritano Heredia
mesa_3  (Maximo)   G-14  → E-14   [manzana G→E]   Feliciano Sigueñas Ugarte
mesa_3  (Maximo)   O-17  → O-27   [lote 17→27]    Dante Ruben Medrano Mallma
mesa_4  (Wagner)   A-1   → A-2    [lote 1→2]      Rosa Aurora Roca Vidal
```

`H1-15` es la más urgente: Marleny Samaritano debe exactamente los S/58 que se cobraron, y
H1-16 aparece en la misma mesa con solo S/8 — no se agarró en la verificación manual de hoy
porque el foco estuvo en los 2 casos que preguntó el usuario (Magda, Pedro), no en correr el
módulo todavía.

Después de verificar mesa_3 y mesa_4, correr `verificar_lotes.py` de nuevo — sus otras 15
filas `NO CUADRA` sin candidato (`G-17`, `C-16`, `Z-7`, `H-9`, `B-20`, `I-11`, `C-17`, `O-15`,
`H1-2`, `Q-14`, `I-6`, `P-13`, `C1-3`, `O-12`, `H1-35`) también quedan pendientes de revisión
manual — el módulo avisa, no bloquea, así que `main.py` puede correr igual mientras se
resuelven.

---

## 1. Lo que se cerró

### 1a. Verificación manual — Wilder Trujillo (mesa_1) y Yerald Romero (mesa_2)

Fotos nuevas de la secretaria (las de ayer no se veían bien). Cruzando contra
`usuarios_id.xlsx` y la caligrafía:

```
mesa_1  fila 10   W-2 → U-2         (la Ц de "U" se leía como W)
mesa_1  fila 62   BLANCO 16 nuevo   (01/08: entregó 595, las filas explican 579)
mesa_2  fila 76   M-19 → M-14       ("Magda Maria Blas" — el lote 19 se leyó 14)
mesa_2  fila 107  G-13 → O-13       ("Pedro Mendoza Marquina" — la manzana G se leyó O)
```

**Patrón que motivó el resto de la sesión:** Yerald escribe el nombre del vecino en su hoja,
Wilder no. Los 2 errores de Yerald se cazaron cruzando el nombre — con solo `MZ+LT` (como
tiene Wilder) ninguno se hubiera detectado.

### 1b. `usuarios_id.xlsx` — 2 correcciones

```
U0500 nuevo   RAMON REQUEZ MENDOZA, M-12
   existía en 0_padron y en la planilla — el nombre viejo del lote (Iglesia Evangelica
   Bautista) no coincidía con lo que el usuario recordaba (Iglesia Adventista del Séptimo
   Día), que en realidad es un predio distinto: C-3A, sigue viva en usuarios_id.

U0457         C1-17 → C1-9 (Roberto Macarlopu Flores)
   el override de julio renombraba el predio pero nunca tocó usuarios_id — quedaba
   desalineado desde entonces.
```

### 1c. `entregas_hoja.xlsx` — 3 filas nuevas

```
Wilder  01/08   efectivo 595   (el "recibí" tachado en el papel — la nota original decía
                                 595 y 579, tachado; se tomó 595 con comentario aclarando)
Yerald  01/08   efectivo 818, yape 8
Yerald  02/08   efectivo 1643, yape 251
   ambas de Yerald son ASUMIDAS = lo que mesa_2.xlsx registra que cobró. Sus fotos NO
   tienen una declaración "Recibí X" legible como sí tiene Wilder. Comentario de Excel
   (no MOTIVO — esa columna dispara delta en el import) documentando el supuesto.
```

### 1d. `verificar_lotes.py` — sub-módulo nuevo, completo

Cruza cada pago de efectivo contra `DATA_boletas.xlsx` en 4 capas — cuadre, fuerza de
evidencia, vecindad de confusión de tipeo/OCR, filtro de realidad — antes de que `main.py`
consolide. Corre primero en el flujo del módulo.

```
diseño completo, metodología Fase 0→3:
   docs/decisiones/verificacion_lotes_efectivo.md    (CRAD, 3 enfoques, evidencia medida)
   4_pagos/efectivo/docs/diagrama_flujo_verificacion_lotes.html
   4_pagos/efectivo/docs/diagrama_verificacion_lotes.html
   4_pagos/efectivo/docs/formato_verificacion_lotes.html

código:
   4_pagos/efectivo/verificar_lotes.py
   4_pagos/efectivo/tests/test_verificar_lotes.py   (21/21, incluye 2 CONTRAFACTUALES)

corrida real (7 mesas, ciclo 2026-08):
   326 pagos · 557 boletas
   ALTA 37 · MEDIA 182 · BAJA 76 (26%) · NO CUADRA 19 · OMITIDO 7 · SIN BOLETA 5

regresión: test_unitarios 27/27 · test_integracion 16/16 (main.py sin tocar)
commit: 8fd678c
```

**Prueba contrafactual (la que justifica el diseño):** con los 2 lotes reales mal escritos
(M-19 en vez de M-14, G-13 en vez de O-13) el módulo los marca `NO CUADRA` y propone el lote
correcto en los dos. Es la evidencia de que lo habría atrapado antes de que la secretaria
tuviera que reclamar.

---

## 2. Decisiones de diseño que vale recordar

```
fuente de la deuda      DATA_boletas.xlsx, NO la planilla (TOTAL_A_PAGAR es fórmula Excel,
                         pandas la lee vacía)

evidencia se mide       contra el IMPORTE TOTAL de la boleta, no contra cualquier
                         subconjunto (esa definición daba 79% BAJA — inservible)

candidato de capa 3     SOLO nivel de error SIMPLE (manzana O lote mal, no los dos).
                         El doble se cuenta pero no se propone — medido: agregaba ruido
                         sin acertar nunca (G-17→Q-12, sin relación real)

clave de preservación   (MESA, COBRADOR, FECHA, MZ, LT, MONTO) visible + FILA_EXCEL
                         oculta como respaldo — la primera versión (sin MZ/LT, copiada
                         del ABONO_ID) colisionaba en 62% de las filas reales

CONCEPTO≠vacío           se omite del cuadre (mismo campo que causó el bug de A-4 el 06/08)
```

---

## 3. Pendiente, por prioridad

```
① verificar mesa_3 (Maximo) y mesa_4 (Wagner) contra sus fotos — PRIMER PASO arriba
② revisar los 4 candidatos únicos que propuso el módulo (H1-15, G-14, O-17, A-1)
③ las 15 filas NO CUADRA sin candidato — revisión manual, una por una
④ correr main.py recién después de resolver lo anterior (verificar_lotes NO bloquea,
   pero corre antes por diseño)
```

No relacionado con esta sesión, sigue vivo del `LEER_ANTES.md` y RETOMAR anteriores:
`5_cobranza/main.py`, `shared/usuarios_id.xlsx` (huecos previos), y las mesas 3/4 ya
traían cambios sin commitear de antes de hoy — no se tocaron ni se commitearon en esta
sesión, siguen sueltos.
