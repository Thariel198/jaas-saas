# RETOMAR — Backfill: diseñar el flujo para pagos nunca procesados · Sesión 2026-07-20 (2ª parte)

Continúa `docs/retomar/RETOMAR_junio_restauracion_y_precursores_ledger_2026-07-20.md` (mismo día,
escrito a mitad de sesión) — esa parte ya se cerró casi entera. Esta es la 2ª mitad del
día: se terminó de vaciar `mesa_5` y quedó **una sola pregunta de diseño abierta**, la que
da nombre a este doc.

---

## 1. Lo que se CERRÓ desde el RETOMAR anterior de hoy

### `abonos_rezagados.xlsx` — creado + overlay codificado y verificado

`shared/abonos_rezagados.xlsx` (7 filas: T-12, S-5, D-16, F-9, D1-6, I-9, L-4 — los pagos
yape de junio que Wagner Trujillo retuvo y "devolvió" en efectivo en julio). Modelo cerrado
con el usuario: **efectivo/julio en la caja, pero salda junio** por aplicación al cargo,
no por fecha del abono — así el extracto del vecino (⑫) muestra "pagaste en junio", sin la
historia de la retención (eso vive en `RETENIDO_POR`/`CANAL_ORIGEN`, capa de arqueo, no en
la vista del vecino).

Código: `5_cobranza/main.py` — `ABONOS_REZAGADOS_PATH` + `_cargar_abonos_rezagados()` +
overlay en `_cargar_planilla()` (misma cascada que reidentificación, sin
mes_actual/mantenimiento).

### C1-9 → C1-17 resuelto con 2 precursores (no 1) — el balde no se puede partir con las tools de hoy

Se descartó tu primera idea (SALDO_A_FAVOR 218 → reidentificar → reclasificar balde) porque
`reasignar_abono` mueve el abono **completo** de predio, no separa por balde — no existe hoy
una tool de "reclasificar balde". La solución que sí usa las herramientas actuales:
**2 abonos, mismo ancla, cada uno reasignado por separado:**

```
shared/reidentificacion.xlsx        +1 fila: C1-9→C1-17 · 18.5 · agua
shared/aportes_tanque_manuales.xlsx  NUEVO: C1-9→C1-17 · 200 · tanque
  ambos anclados a: mesa_2.xlsx (repo junio), fila 36
```

`aportes_tanque_manuales.xlsx` es un 4º precursor nuevo — no existía. Se creó y se cableó:
`4_pagos/consolidar_tanque.py` ahora tiene `_mes_ciclo_actual()` (deriva el mes vigente
desde `pagos_efectivo.xlsx`, porque el script no recibe `--mes`) + `_leer_tanque_manuales()`
+ wiring en `main()`. Verificado: `aportes_tanque.xlsx` incluye la fila C1-17/200 con
`referencia="reasignado de C1-9"`.

### `mesa_5.xlsx` (repo principal) — VACIADA

Las 9 filas (7 rezagados + C1-17 ×2) ya tenían destino durable → se vaciaron (filas 4-12),
queda solo el template `A-8C` (María García, idéntico en las 5 mesas, no es dato real).

**Re-corridas y verificadas, en orden:**
1. `4_pagos/efectivo/main.py` → 357 cobros, **multi_mesa desapareció** (0 discrepancias,
   antes eran 5 grupos en conflicto por T-12/S-5/D-16/F-9/D1-6/L-4/I-9 chocando con sus
   pagos reales de julio en mesa_4).
2. `consolidar_tanque.py` → `aportes_tanque.xlsx` con C1-17/200 incluido.
3. `5_cobranza --force` → verificado matemáticamente para los 8 predios (D-16 cuadra
   exacto: `10 mes_ant + 50 multa + 25 acuerdos = 85` el rezagado). Tanque queda **fuera**
   de `SALDO` (correcto, no es deuda).

**Hallazgo colateral (bug, ya reparado):** al correr la suite de tests de `4b_reclamos`
esta mañana, **clobbeó** `pagos_efectivo.xlsx` real de julio con 1 fila de fixture — mismo
problema de aislamiento ya documentado para `4_pagos/efectivo/tests/test_integracion.py`.
Se reparó re-corriendo `4_pagos/efectivo/main.py` desde las mesas reales (intactas). **No
correr la suite de tests de `4b_reclamos` contra el estado real sin aislarla primero.**

### `backfill_ledger/` — carpeta nueva en el pipeline (Opción B, decisión del usuario)

No es `libro_mayor/` (eso es diseño del ledger) — es la **tarea de desarrollo de agosto**
que usa ese diseño para sembrar la historia vieja. Creada con:
```
backfill_ledger/
├── README.md                    qué es el backfill, mapa de los 4 precursores → eventos
│                                de ledger, casos abiertos (bloque C1-17 partido)
└── docs/cuaderno_backfill.html  3 láminas: qué es backfill · mapa 4 precursores ·
                                 diagrama del bloque C1-17 partido (con chequeo de suma)
```
**Pendiente:** agregar `backfill_ledger/` al pipeline del `README.md` raíz (Regla 7) — no
se hizo todavía, quedó ofrecido y no confirmado.

---

## 2. Mecánico sin cerrar (heredado, todavía sin ejecutar)

- **C1-9 → marcar `ESTADO=RESUELTO`** en
  `jass_system - junio/5_cobranza/outputs/arrastre_devolucion_2026-06.xlsx` (columna M).
  Se discutió, se confirmó el mecanismo (reidentificación, no devoluciones_aplicadas), pero
  la fila **nunca se marcó** — la conversación saltó a "eliminemos mesa_5" antes de ejecutar
  ese marcado puntual.
- **Regresión (Regla 5) — no se corrió todavía.** `5_cobranza --force` cambió los totales
  de julio (278 pendientes, 36 excesos, 63 elegibles a corte — antes eran otros números,
  esperado porque se estaba corrigiendo data mal cargada, no un regression real). Falta:
  1. Correr `5b_validacion` — confirmar que cuadra con los nuevos totales.
  2. Revisar `6_corte/generar_lista.py` — la lista de elegibles cambió (63).
  3. Nota técnica para cuando se corra `5_cobranza` en Windows: el `main.py` imprime
     `═` (box-drawing) y la consola `cp1252` lo rompe (`UnicodeEncodeError`). Correr con
     `PYTHONIOENCODING=utf-8 py main.py --force`. No es bug del código, es la consola.

---

## 3. ⚡ EL PENDIENTE DE DISEÑO — flujo de backfill para pagos que NUNCA se procesaron

**Esto es lo que el usuario pidió dejar explícitamente para mañana.** Es distinto de todo
lo que se cerró hoy — los 4 precursores (`abonos_rezagados`, `reidentificacion`,
`devoluciones_aplicadas`, `aportes_tanque_manuales`) modelan plata que **en algún momento
SÍ entró** a algún ciclo de `pagos_efectivo` (mal atribuida, retenida-y-luego-recuperada,
o con exceso) — todas terminan reflejadas en un output real, aunque haya que corregirlas.

**El caso de `mesa_6.xlsx` (repo junio) fila 7 es otra categoría: plata que NUNCA entró a
`pagos_efectivo` en ningún ciclo, ni junio ni julio.**

```
mesa_6 (repo junio) fila 7 — BLANCO · S/45 · "Hernestina Valladares?"
   │
   ├─ junio: bug B6 (4_pagos/efectivo/main.py:341) descarta en silencio los blancos de
   │         efectivo (a diferencia de yape, que rutea a blancos_acumulados.xlsx) →
   │         el S/45 NUNCA apareció en pagos_efectivo_2026-06, en ningún archivo
   │
   ├─ hoy: identificado conceptualmente como G-12 (Hernestina) — pero la identificación
   │       vive solo en la conversación / en PARA_AGOSTO.md, no en ningún archivo del
   │       sistema. La fila de mesa_6 sigue como BLANCO (restaurada así a propósito, ver
   │       RETOMAR anterior — NO se debe editar el crudo, mismo principio que C1-9)
   │
   └─ julio: NO se cargó en mesa_5 del repo principal (a diferencia de los 7 retenidos,
             que sí se cargaron y hoy ya están en abonos_rezagados)
```

**Por qué no encaja en `abonos_rezagados.xlsx` tal como está diseñado hoy:** ese archivo
asume que la plata **ya se recuperó** (hay un movimiento real, efectivo, en julio, que se
puede anclar). G-12 no tiene ningún movimiento real todavía — es una identificación de un
blanco que nunca se contó en ningún lado. Aplicar el mismo precursor sin más sería inventar
un "abono de julio" que no existe.

**La pregunta que el usuario dejó pendiente:** ¿cuál es el flujo correcto —

1. ¿Se necesita un **5º precursor** (algo como `blancos_identificados_no_procesados.xlsx`)
   que module "blanco identificado, plata nunca contada, pendiente de cobrarse/reconocerse"?
2. ¿O el flujo correcto es primero **cargar el pago a julio** (mesa_5 del repo principal,
   como se hizo con los 7 retenidos) y **recién ahí** aplica `abonos_rezagados` — es decir,
   G-12 necesita un paso previo que los otros 7 no necesitaban?
3. ¿Hay **"otros pagos más"** en esta misma situación (blancos de efectivo silenciados por
   B6, en junio o en otros meses) que haya que auditar antes de diseñar el flujo — para no
   diseñar para 1 caso y descubrir que el patrón es más grande?

**Nada de esto se resolvió ni se ejecutó hoy.** Es la primera tarea de mañana.

**Contexto relevante para retomar esto:** el bug B6 ya está documentado en
`docs/pendientes_plan.md` (tabla de Bugs) — la causa raíz (efectivo no rutea blancos, a
diferencia de yape) es la raíz de todo este caso. Vale la pena leerlo de nuevo antes de
diseñar, para no repetir el diagnóstico.

---

## No tocar (heredado + nuevo de hoy)

- Todo lo de las secciones "No tocar" de los 2 RETOMAR anteriores (07-19 y 07-20 1ª parte)
  sigue vigente.
- **Nuevo:** no correr la suite de tests de `4b_reclamos` (`4b_reclamos/tests/`) contra el
  estado real sin verificar aislamiento primero — hoy clobbeó `pagos_efectivo.xlsx` de
  julio (reparado, pero el riesgo sigue latente para la próxima vez).
- No diseñar el flujo de G-12 asumiendo que es igual a `abonos_rezagados` — ya se estableció
  que es un caso distinto (plata nunca contada vs. plata retenida-y-recuperada).

---

## Estado git — nada commiteado (sigue igual)

Repo principal: además de lo ya listado en el RETOMAR anterior, ahora también sin commitear
— `shared/abonos_rezagados.xlsx` (nuevo), `shared/aportes_tanque_manuales.xlsx` (nuevo),
`shared/reidentificacion.xlsx` (+1 fila C1-9→C1-17), `5_cobranza/main.py` (overlay
rezagados), `4_pagos/consolidar_tanque.py` (wiring tanque manual), `4_pagos/efectivo/inputs/mesa_5.xlsx`
(vaciada), `backfill_ledger/` (carpeta nueva completa), + todos los outputs regenerados
(`pagos_efectivo.xlsx`, `aportes_tanque.xlsx`, `planilla_cobrado.xlsx`,
`arrastre_deuda/devolucion/consolidado_2026-07.xlsx`, `trazabilidad_cobranza.xlsx`,
`vista_seguimiento_pueblo.xlsx/.pdf`).

Repo junio: sin cambios adicionales desde el RETOMAR anterior (la columna `ESTADO` sigue
sin el marcado de C1-9, ver sección 2).
