# RETOMAR — Reconciliación junio→julio + precursores del ledger · Sesión 2026-07-19

Handoff detallado para retomar mañana. Leer de arriba a abajo antes de tocar nada.
Continúa el trabajo de `jass_system - junio\docs\RETOMAR_junio_cierre_reconciliacion_2026-07-18.md`
(sesión de ayer) — hoy se ejecutó casi todo ese checklist y aparecieron 2 mecanismos
nuevos en `5_cobranza` que no existían ayer.

---

## ⚡ Contexto — dos repos, no confundir (recordatorio de ayer, sigue aplicando)

```
C:\...\jass_system            → repo principal. Ciclo julio en curso.
C:\...\jass_system - junio    → repo aparte (git propio), foto de junio YA CERRADO.
```

---

## 1. Checklist de ayer — EJECUTADO casi completo

De los 8 casos de `reconciliacion_junio_a_julio.xlsx` (repo junio,
`4_pagos\efectivo\outputs\`), se cargaron en **`jass_system\4_pagos\efectivo\inputs\mesa_5.xlsx`**
(repo principal, filas 4-12):

| Caso | Fila mesa_5 | Estado |
|---|---|---|
| T-12 (Samuel Samaritano Romero) | 4 | Cargado, S/155 |
| S-5 (Valerio Porfilio Javier Santiago) | 5 | Cargado, S/71 — falta condonar CORTE(40)+MULTA(30) |
| D-16 (Esteban Guerrero Chingel) | 6 | Cargado, S/85 — falta condonar MULTA(50) |
| F-9 (Rosa Lucia Coronado Luna) | 7 | Cargado, S/52, sin ajustes |
| Macarlopu → **C1-17** (no C1-9) | 8-9 | Cargado, S/18.50 agua + S/200 CONCEPTO=tanque — falta condonar MULTA(30) |
| D1-6 (Onita Ponte Eguizabal) | 10 | Cargado, S/33, pago simple |
| I-9 (Julia Cardenas Alvarado) | 11 | Cargado, S/86, pago simple |
| L-4 (Delia Doris Huamansupa Perez) | 12 | Cargado, S/58, pago simple |
| C-39 (Janet Romero Mayo) | — | **NO cargado** — duplicado confirmado (mismo yape ya matcheado 09/06) |

Formato de fecha corregido en esas 9 filas (`mm-dd-yy`, sin hora/minuto/segundo).

**Documentación creada:**
- `jass_system\LEER_ANTES.md` (raíz) — describe el evento, tabla fila-por-fila de `mesa_5.xlsx`.
- `jass_system - junio\LEER_ANTES.md` (raíz) — mismo evento visto desde el repo junio.
- `jass_system\PARA_AGOSTO.md` (raíz) — **las 4 condonaciones pendientes** (S-5, D-16, C1-17, Hernestina G-12) con motivo. No ejecutadas todavía, quedan para agosto.
- CLAUDE.md de **ambos repos**: nueva Regla 10 — `LEER_ANTES.md` en la raíz para eventos que rompen el flujo normal del pipeline.

---

## 2. Patrones nuevos encontrados en `arrastre_devolucion_2026-06.xlsx` (repo junio)

Se auditaron los 13 excesos de ese archivo contra memoria + investigación nueva:

| Caso | Resultado | Estado |
|---|---|---|
| **E-1** (Yolanda Montalvo, S/100) | Yape "tanque" nunca tuvo `CONCEPTO=tanque`, solo en el MENSAJE | ✅ **RESUELTO** — movido a `aportes_tanque.xlsx` (repo junio), sacado de `arrastre_devolucion_2026-06.xlsx` |
| **G-18** (Benita León Jara, S/25) | Pagó S/58 en junio (debía 33) — el sobrante cubría su CONVENIO/medidor (saldo 50, cuota ya cargó 25) | ✅ **RESUELTO** — ver §3 |
| **M-12** (Iglesia Evangélica Bautista, S/266) | Deuda histórica 2019-oct.2025, ya confirmado legítimo (sesión D4 anterior) | Sin acción — ya explicado |
| **P-6** (Flor Valdivia Milla, S/300) | "300 de convenio" — **el comentario NO tiene fuente real** (mesa/pagos_efectivo/trazabilidad muestran COMENTARIO vacío), se tipeó a mano en `arrastre_devolucion` sin verificar contra el registro de convenio real | ⚠️ Sin confirmar — pendiente cruzar contra `obligaciones/inputs/mayo-planilla...xlsx` hoja "Cobro medidores" |
| **D-9** (Rufina Cabello Ticlio, S/6) | Comentario "exonerado" en la propia mesa | Sin acción — ya explicado |
| **R-7** (Rosaura Oroya Sacre, S/24) | Comentario "Exoneracion" — es el mismo caso del pendiente D2 (blanco sin dueño, ya identificado) | Sin acción — ya explicado |
| **C1-15** (Ever Cervantes, S/8) / **D1-5** (Elmer Melendez, S/2) | Montos chicos, probable redondeo | Sin acción, baja prioridad |
| **I-7** (Dario Diego Rosales, S/45) | **NO es exceso de I-7** — es un pago mal atribuido, ver §4 (V-6/T-7) | ✅ **RESUELTO** |
| **V-16** (Sonia Masias Cusihuaman, S/44) | **NO es exceso de V-16** — pago mal atribuido, ver §4 | ✅ **RESUELTO** |
| **G1-1** (Gilberto Luis Meza, S/47) | Sin registro en `seguimiento_pueblo`, sin comentario | ❌ **Sin explicar** — investigar a mano |
| **M-1** (Segundina Eutopia Espinoza, S/33) | Sin registro en `seguimiento_pueblo`, sin comentario | ❌ **Sin explicar** — investigar a mano |

**Nota:** `arrastre_devolucion_2026-06.xlsx` (repo junio) todavía tiene las filas de E-1 (ya sacada ✓),
G-18, I-7 y V-16 — **falta limpiar esas 3 filas** (G-18/I-7/V-16 quedaron resueltas por otra vía,
ver §3 y §4, pero el archivo del repo junio no se tocó para reflejarlo). Cosmético, no bloquea nada.

---

## 3. Mecanismo nuevo — `devoluciones_aplicadas.xlsx` (créditos de exceso aplicados a deuda)

**Por qué existe:** `seguimiento_pueblo.xlsx` es un derivado (lo escribe `5_cobranza`) que el
roadmap del ledger va a reemplazar — parchearlo hoy habría creado un fix que se pierde en la
migración. En cambio, se creó una fuente propia, durable, candidata directa a fuente de siembra
del futuro `libro_mayor` (es una versión manual de `SALDO_A_FAVOR`/`Ajuste`, ver
`libro_mayor/dominio/README.md`).

```
shared/devoluciones_aplicadas.xlsx   (repo principal)
  MZ · LT · CONCEPTO · MES_ANO_ORIGEN · MES_ANO_APLICA · MONTO · MOTIVO · RECLAMO_REF
  G · 18 · CONVENIO · 2026-06 · 2026-07 · 25 · "..." · reclamos_2026-07.xlsx G-18 (EN_REVISION)
```

**Código nuevo en `5_cobranza/main.py`:**
- `DEVOLUCIONES_APLICADAS_PATH` (constante, cerca de `AUDIT_MULTA_PATH`)
- `_cargar_devoluciones_aplicadas(mes_ano)` — lee el xlsx, filtra por `MES_ANO_APLICA` (se aplica
  una sola vez, no en cada ciclo futuro)
- Overlay en `_cargar_planilla()`, después del overlay de penalidad — resta el MONTO del campo
  correspondiente (`_CONCEPTO_DEVOLUCION_A_CAMPO`: CONVENIO/MULTA/ACUERDOS/CORTE_RECONEXION)

**Resultado verificado:** G-18 pasó de SALDO=25 (PARCIAL) a **SALDO=0 (CANCELADO)**.

---

## 4. Mecanismo nuevo — `reidentificacion.xlsx` (pagos mal atribuidos a otro lote)

**Distinto de una devolución:** no es dinero de más que se decide aplicar a un concepto — es un
pago que **nunca fue de ese lote**, quedó mal anotado por un typo o confusión de nombre. Modela
por adelantado el `reasignar_abono` del ledger (decisión I1-I4, `libro_mayor/dominio/README.md`):
ancla la **transacción específica** (`origen_archivo + fila`), no un remapeo global de `(MZ,LT)`
como `correcciones_lote.xlsx` (que sí seguiría aplicando a pagos futuros — hubiera sido incorrecto
acá).

**Los 2 casos encontrados, ambos confirmados contra la fuente cruda:**

```
V-16 → V-6   (mesa_5.xlsx repo junio, fila 58, cobrador Maximo Encarnacion)
  Dos vecinas llamadas "Sonia". El pago de Sonia Arquino Melendez (V-6, debía
  EXACTO 44 = 11 agua+3 mant+30 multa, sin ningún pago) quedó anotado en el
  lote de Sonia Masias Cusihuaman (V-16, debía solo 10). Corrobora el reclamo
  de V-6 en julio: "Ya pagué multa 30" (EN_REVISION).

I-7 → T-7   (mesa_3.xlsx repo junio, fila 23, cobrador Janet Villanueva)
  La fila CRUDA de mesa_3 YA DICE "T · 7" (comentario "compromiso") — la
  resolución de discrepancia multi_mesa lo reasignó a I-7 por error. I-7 tenía
  su propia deuda de junio (47) cubierta por OTRO pago de la misma resolución
  (mesa_5, Maximo Encarnacion, 47). T-7 sigue debiendo después de esto (no es
  cancelación completa) — se aplicó a ACUERDOS_ASAMBLEA (50→5).
```

```
shared/reidentificacion.xlsx   (repo principal)
  ORIGEN_ARCHIVO · FILA · COBRADOR · MZ_ORIGEN · LT_ORIGEN · MZ_CORRECTO · LT_CORRECTO
  · MONTO · FECHA_PAGO · CONCEPTO_DESTINO · MES_ANO_APLICA · MOTIVO · RECLAMO_REF
```

**Código nuevo en `5_cobranza/main.py`:**
- `REIDENTIFICACION_PATH` (constante)
- `_CAMPOS_WATERFALL_REIDENTIFICACION = ("mes_anterior", "corte_reconexion", "multa",
  "acuerdos_asamblea", "convenio")` — **sin** `mes_actual`/`mantenimiento`: la deuda mal atribuida
  es de un ciclo YA CERRADO, no debe cancelar el consumo vigente del mes en curso
- `_cargar_reidentificaciones(mes_ano)` — si `CONCEPTO_DESTINO` viene vacío, reparte en cascada
  por esa lista; si viene lleno, va directo a ese campo (reusa `_CONCEPTO_DEVOLUCION_A_CAMPO`)
- Overlay en `_cargar_planilla()`, después del de devoluciones

**Resultado verificado:**
- V-6: SALDO 44 → **0 (CANCELADO)** — mes_anterior 14→0, multa 30→0, exacto
- T-7: SALDO 95 → **50 (PARCIAL)** — acuerdos_asamblea 50→5

---

## 5. Corrida de verificación — HECHA

`5_cobranza --force` corrido 2 veces (la primera falló por `PermissionError` — `planilla_cobrado.xlsx`
abierto en Excel; la segunda, con el archivo cerrado, completó limpio).

```
Overlay devoluciones aplicadas · 1 crédito(s) aplicado(s)
Overlay reidentificación · 2 crédito(s) aplicado(s)
Cobranza completada · ciclo 14 · 2026-07 · 565 usuarios
  arrastre_deuda_2026-07.xlsx      (281 pendientes)
  arrastre_devolucion_2026-07.xlsx (35 excesos)
  → 70 usuarios elegibles para corte (antes 71 — bajó por el fix)
```

Verificado con lectura directa de `planilla_cobrado.xlsx`: G-18/V-6/T-7 exactos a lo calculado.
`arrastre_devolucion_2026-07.xlsx` se regeneró DESPUÉS de los overlays — no debería tener ya
ninguno de estos 3 casos (no verificado fila por fila todavía, ver §7).

---

## 6. ⚠️ Deuda de metodología — NO se corrió la regla de regresión (Regla 5/CLAUDE.md)

**No se corrió nada aguas abajo de `5_cobranza` hoy.** Con SALDO/elegibilidad de corte cambiando
(71→70), corresponde correr y revisar, en este orden, ANTES de dar por cerrado el día:

1. `5b_validacion` — confirmar que sigue OK con los nuevos totales.
2. `6_corte/generar_lista.py` — la lista de corte cambió (70 vs 71 elegibles); si ya se publicó
   una lista con el 71 viejo, hay que decidir si se re-emite.
3. Revisar `arrastre_devolucion_2026-07.xlsx` fila por fila — confirmar que G-18/V-6/T-7 no
   aparecen ahí con algún residuo raro.

---

## 7. Pendientes explícitos para mañana

**Limpieza (cosmética, no bloquea):**
- Sacar las filas de G-18, I-7, V-16 de `jass_system - junio\5_cobranza\outputs\arrastre_devolucion_2026-06.xlsx`
  (ya resueltas por otra vía, el archivo del repo junio no se actualizó).
- Marcar reclamos como resueltos en `4b_reclamos/outputs/reclamos_2026-07.xlsx` (repo principal):
  - **V-6** "Ya pagué multa 30" → FUNDADO (la reidentificación lo confirma).
  - **T-7** "verificar medidores..." → sigue abierto, es un tema distinto (CONVENIO), no lo resuelve
    la reidentificación de los 45.

**Documentación (deuda, Regla 7 — README debe ir con el cambio, no se hizo hoy):**
- `5_cobranza/README.md` y/o `docs/diagrama_5_cobranza.html` no mencionan todavía
  `devoluciones_aplicadas.xlsx` ni `reidentificacion.xlsx` — actualizar antes de que otra sesión
  los encuentre sin contexto.

**Investigación sin pista (bajo esfuerzo, sin urgencia):**
- **G1-1** (Gilberto Luis Meza, S/47) y **M-1** (Segundina Eutopia Espinoza, S/33) — excesos sin
  ningún registro en `seguimiento_pueblo`, sin comentario en la mesa. Preguntar al cobrador
  (Wagner Trujillo / Maximo Encarnacion) o esperar reclamo.
- **P-6** (Flor Valdivia Milla, S/300) — el "300 de convenio" no tiene fuente verificable, cruzar
  contra `obligaciones/inputs/mayo-planilla...xlsx` antes de darlo por bueno.

**Del checklist de AYER, todavía sin tocar:**
- **§10** (RETOMAR de ayer) — cuál carpeta "copia (2)/(3)/-1/-2" corresponde a junio, sigue sin
  preguntarse/resolverse.
- **§3 del RETOMAR de ayer** — fix de `2_planilla`/`6_corte` (predios SIN_MEDIDOR) sigue **sin
  commitear** en el repo principal (4 archivos que quedaron staged ayer). Revisar `git status`
  antes de commitear — no arrastrar el borrado de `7b_historial_pagos/`.
- **Blancos §11 de ayer** (Janet Villanueva w-6/T-7, Yerald Romero C1-16) — siguen sin dueño.

**`PARA_AGOSTO.md` (raíz del repo principal) — 4 condonaciones, ninguna ejecutada:**
S-5 (CORTE 40 + MULTA 30) · D-16 (MULTA 50) · C1-17 Macarlopu (MULTA 30) · Hernestina G-12 (MULTA 30).

---

## No tocar

- `4_pagos/yape/validacion/main.py` (usuario resuelve a mano).
- `tests/test_integracion.py` de `4_pagos/efectivo` en NINGÚN repo (no está aislado, escribe sobre
  datos reales — incidente ya documentado en el RETOMAR de ayer, §9).
- No reabrir el ciclo de junio corriendo `main.py` completo en el repo junio.
- `correcciones_lote.xlsx` NO es el lugar para V-16/I-7 — es remapeo global de lote, se usó
  deliberadamente `reidentificacion.xlsx` en su lugar (ver §4, motivo explicado).

---

## Estado git

Nada de lo de hoy está commiteado (ni lo de ayer — el fix de §3 del RETOMAR anterior sigue
staged sin commit). Revisar `git status` completo en ambos repos antes de cualquier commit —
hay trabajo de varias sesiones mezclado en el working tree.
