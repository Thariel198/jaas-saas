# RETOMAR — Restauración de junio + precursores manuales del ledger · Sesión 2026-07-20

Handoff preventivo (se escribió a mitad de sesión, no al cierre — leer completo antes de
tocar nada). Continúa `docs/RETOMAR_reconciliacion_junio_ledger_2026-07-19.md` (ayer) —
hoy se cerró parte de lo pendiente de ese doc y aparecieron 2 casos nuevos sin resolver.

---

## ⚡ Contexto — dos repos, no confundir (sigue aplicando)

```
C:\...\jass_system            → repo principal. Ciclo julio en curso.
C:\...\jass_system - junio    → repo aparte (git propio), foto de junio YA CERRADO.
```

**Hallazgo nuevo de hoy:** el repo `jass_system - junio` NO está realmente congelado.
Su último commit es del **2026-07-02**, pero el working tree tiene semanas de trabajo sin
commitear (casi todos los módulos modificados + `seguimiento_pueblo`, `arqueo`, `entregas`,
`7_cierre` implementado, etc. — el mismo tipo de desarrollo que en el repo principal).
Eso es "trabajo seguro" intencional, NO se tocó. Lo que sí se rompió y se arregló hoy fue
mucho más puntual — ver sección 1.

---

## 1. Restauración de `4_pagos/efectivo` en el repo junio — HECHA Y VERIFICADA

**Qué pasó:** en una sesión previa (mtimes muestran 2026-07-18), se editaron a mano
`mesa_2.xlsx` a `mesa_6.xlsx` del repo junio — moviendo 7 pagos de la columna YAPE a
EFECTIVO (para poder "cobrarlos en efectivo en julio", idea descartada) y
re-identificando 2 blancos (`G-12` S/45 "Hernestina Valladares?", `O-6` S/107). Después se
corrió `4_pagos/efectivo/main.py`, que regeneró `pagos_efectivo.xlsx` y
`trazabilidad_2026-06.xlsx` reflejando esos cambios. **El ciclo de junio no se re-corrió
completo** — se paró ahí a propósito.

**Verificado con timestamps (no solo con memoria):**
```
4_pagos/efectivo/outputs/pagos_efectivo.xlsx      2026-07-18 15:57  ← tocado
4_pagos/efectivo/trazabilidad_2026-06.xlsx        2026-07-18 15:57  ← tocado
5_cobranza/outputs/*  (todos)                     2026-07-02        ← SIN TOCAR
5b_validacion/outputs/validacion_diferencias.xlsx 2026-07-02        ← SIN TOCAR
6_corte/outputs/*     (todos)                     2026-06-23/24     ← SIN TOCAR
```
Confirmado: el daño se quedó contenido en `4_pagos/efectivo`, nunca llegó a
`5_cobranza`/`5b_validacion`/`6_corte`.

**Restauración ejecutada:**
1. `git checkout HEAD -- mesa_2.xlsx mesa_3.xlsx mesa_4.xlsx mesa_5.xlsx mesa_6.xlsx trazabilidad_2026-06.xlsx`
   (HEAD = commit 2026-07-02, previo a las ediciones). `mesa_6.xlsx` falló la primera vez
   (`PermissionError` — archivo abierto en Excel), se reintentó tras cerrarlo. OK.
2. Se corrió **solo** `4_pagos/efectivo/main.py` (no el pipeline completo) para
   regenerar `pagos_efectivo.xlsx`/`blancos_mes.xlsx` desde las mesas restauradas.
   Log: 0 discrepancias nuevas, 6 resoluciones multi_mesa re-incorporadas igual,
   `solo_un_cobrador` preservadas del Ciclo 1.

**Verificación final (doble-chequeada a pedido del usuario):**
- Las 7 mesas (`mesa_1` a `mesa_7`) están **byte-idénticas a HEAD** —
  `git status`/`git diff --stat` vacíos para las 7. `mesa_1` y `mesa_7` nunca se tocaron.
- `pagos_efectivo.xlsx`: T-12, S-5, D-16, D1-6, I-9, L-4, F-9 → AUSENTES (correctamente
  vuelven a ser YAPE, no efectivo). `G-12` → solo su pago legítimo de S/34 (compromiso);
  el S/45 "Hernestina Valladares?" volvió a `blancos_mes.xlsx` (5 blancos, como antes).

**Nota de método para la próxima vez que haga falta algo así:** los inputs de mesa NO
tienen backup automático (`4_pagos/efectivo` respalda `discrepancias` y
`pagos_efectivo_pre_nombres`, pero no las mesas crudas) — la única red de seguridad real
fue el commit de git. Si algún día se edita una mesa de un ciclo cerrado sin que el archivo
esté commiteado primero, no hay forma de reconstruirlo.

---

## 2. Precursores manuales del ledger — trabajo de hoy (repo principal)

### 2a. `shared/devoluciones_aplicadas.xlsx` — columna `SUB_CONCEPTO` agregada

El archivo (creado ayer 07-19, mecanismo de "exceso que se aplica a un concepto de deuda
en vez de esperar reclamo") no distinguía sub-concepto. Se agregó la columna
`SUB_CONCEPTO` (entre `CONCEPTO` y `MES_ANO_ORIGEN`) para que el ledger futuro
(`CONVENIO→medidor/instalación`, decisión ⑪ del contrato) no tenga que re-interpretar el
`MOTIVO` en texto libre.

Estado actual del archivo (2 filas):
```
G · 18 · CONVENIO · MEDIDOR      · 2026-06 · 2026-07 · 25  · (backfill de la fila que ya existía)
P · 6  · CONVENIO · INSTALACION  · 2026-06 · 2026-07 · 300 · confirmado por el dueño como pago de instalación de agua
```
No hizo falta tocar `5_cobranza/main.py` — `_cargar_devoluciones_aplicadas()` lee con
`pd.read_excel(header=1)`, la columna nueva pasa como campo extra sin romper nada; el
mapeo a `campo` sigue siendo por `CONCEPTO` (plano).

### 2b. `jass_system - junio/.../arrastre_devolucion_2026-06.xlsx` — columna `ESTADO` agregada

Nueva columna (grupo "¿Revisado?", junto a `REVISION`), dropdown `PENDIENTE`/`RESUELTO`.
Estado marcado hoy en las 12 filas:

| Predio | ESTADO | Vía |
|---|---|---|
| G-18 | **RESUELTO** | `devoluciones_aplicadas.xlsx` (CONVENIO/MEDIDOR) |
| I-7 | **RESUELTO** | `reidentificacion.xlsx` (reasignado a T-7) |
| V-16 | **RESUELTO** | `reidentificacion.xlsx` (reasignado a V-6) |
| P-6 | **RESUELTO** | `devoluciones_aplicadas.xlsx` (CONVENIO/INSTALACION, cerrado hoy) |
| C1-15, D-9, D1-5, G1-1, M-1, M-12, R-7 | PENDIENTE | sin cubrir todavía |
| **C1-9 (Roberto Macarlopu)** | PENDIENTE ⚠️ | **probablemente mal marcado — ver §3a** |

### 2c. `4b_reclamos` — nueva fuente de intake manual

Problema planteado: la secretaria recibe reclamos/tareas que NO pasan por un cobro
(verbal, WhatsApp) — el único intake de `4b_reclamos` hoy era automático
(`COMENTARIO≈"reclamo"` en `pagos_efectivo.xlsx`).

Solución implementada (sin tocar la maquinaria existente de preservación/arrastre/dedup/
trazabilidad, que está atada a la clave `MESA+MZ+LT+FECHA_COBRO`):

```
4b_reclamos/inputs/reclamos_manuales.xlsx   (NUEVO, tipeado a mano)
  FECHA · MZ · LT · RECLAMO · QUIEN_REPORTA

_cargar_manuales(mes)  en main.py  (NUEVA función)
  sintetiza MESA="MANUAL", COBRADOR=QUIEN_REPORTA, FECHA_COBRO=FECHA, MONTO=""
  → mismo shape que _cargar_detectados() → se concatena antes del resto del flujo
```
Verificado con un caso sintético aislado (no en la suite real): 1 fila válida entra
correctamente, 1 fila vacía se descarta, re-corrida no duplica.

**Los 9 fails de `4b_reclamos/tests/test_reclamos.py` son PRE-EXISTENTES** — confirmado
con `git stash` (mismos 9 fails con y sin el cambio de hoy). No están arreglados, pero
tampoco los causó esta sesión. Pendiente de investigar aparte (parece relacionado a que el
test no está aislado del estado real de `outputs/reclamos_2026-06.xlsx`, similar al
incidente ya documentado de `4_pagos/efectivo/tests/test_integracion.py`).

Docs actualizados: `4b_reclamos/README.md` (sección "Qué hace", "Estructura",
"Dependencias externas", "Reglas clave") + `4b_reclamos/docs/formato_reclamos.html`
(ejemplo de fila `MESA=MANUAL`, nota explicativa, versión 3.1).

---

## 3. Casos sin resolver — quedan para la próxima sesión

### 3a. ⚠️ C1-9 / C1-17 (Macarlopu) — un TERCER mecanismo de resolución, sin marcar

El exceso `C1-9` (Roberto Macarlopu Flores, S/218.50, "reclamo") en
`arrastre_devolucion_2026-06.xlsx` **ya está resuelto**, pero no por `devoluciones_aplicadas`
ni `reidentificacion` — se resolvió reasignándolo a `C1-17` directamente en
`4_pagos/efectivo/inputs/mesa_5.xlsx` **del repo principal** (filas 8-9, S/18.50 agua +
S/200 tanque, comentario "Pago de junio registrado por error en C1-9 (reclamo) -
reasignado a su lote real C1-17"). Quedó marcado `PENDIENTE` en la columna `ESTADO` por
error — nadie confirmó el cambio a `RESUELTO` todavía.

**Acción pendiente:** confirmar con el usuario y marcar `C1-9 → RESUELTO` en
`arrastre_devolucion_2026-06.xlsx` (columna M, fila del predio C1-9).

**Nota de diseño para cuando se toque esto:** hay ahora TRES mecanismos manuales
resolviendo excesos de junio (`devoluciones_aplicadas.xlsx`, `reidentificacion.xlsx`, y
"reasignación directa en mesa_5 de julio") — vale la pena, en algún momento, decidir si
C1-9/C1-17 debería haber sido una fila de `reidentificacion.xlsx` en vez de una edición
directa en mesa. No se resolvió hoy, solo se detectó.

### 3b. Pagos retenidos por el cobrador (yape no transferido) — diseño abierto, sin código

**El problema de fondo** (planteado hoy, con el lente de "¿cómo entra esto al ledger
mañana?"): 7 predios pagaron por Yape en junio a un cobrador (Wagner Trujillo, mayormente)
que no transfirió esa plata a la cuenta de la JASS. Son dos hechos distintos que hoy se
fusionan en uno falso si se anota como "efectivo cobrado en julio":

```
HECHO A · el vecino PAGÓ en junio, canal YAPE, a un teléfono que no es de la JASS
HECHO B · quien recibió RETUVO — la plata no llegó a la cuenta
```

**Ya identificado en `4_pagos/efectivo/inputs/mesa_5.xlsx` del REPO PRINCIPAL** (filas 4-12,
cargado el 07-19): T-12, S-5, D-16, F-9, D1-6, I-9, L-4 (+ C1-17 que es el caso 3a, distinto).
Ahí quedaron como pago de julio en efectivo — funcionalmente correcto para cobrar la
plata, pero pierde el hecho de que originalmente fue un yape de junio retenido.

**Propuesta que se discutió (Opus, NO implementada):**
```
shared/abonos_rezagados.xlsx  (archivo nuevo, no existe todavía)
  MZ · LT · CANAL(=yape) · FECHA_REAL(=junio) · MES_CICLO_REAL(=2026-06) · MONTO
  · BALDE(=agua) · EVIDENCIA · RETENIDO_POR · MES_ANO_APLICA(=2026-07)
```
Diseñado como pre-imagen 1:1 de la Entidad 1 (`MovimientoCaja`) del contrato del ledger —
se cargaría a `caja.registrar_movimiento()` renombrando columnas, sin interpretar nada.
`FECHA_REAL` ≠ `MES_ANO_APLICA` es literalmente `FECHA` vs `recorded_at` que el contrato
ya define. Overlay en `5_cobranza` sería ~1 función, mismo patrón que
`devoluciones_aplicadas`/`reidentificacion`.

**Pregunta sin responder — bloquea implementar esto:**
El comentario en mesa_5 dice *"cobrado en efectivo en julio"*. ¿Es:
- **(a)** el vecino pagó **dos veces** (yape de junio retenido + efectivo nuevo de julio) →
  son 2 movimientos, hay que devolverle uno, o
- **(b)** se **recuperó** la plata del cobrador que la retuvo → es 1 solo movimiento
  (el yape de junio), solo cambió de custodio.

Esto cambia todo el modelado. **No se puede avanzar en `abonos_rezagados.xlsx` sin
esta respuesta.**

**Hueco del contrato anotado en caliente (para cuando se retome):** `caja` asume 2 caminos
de evidencia (yape→export del banco, efectivo→mesa). Un yape a un teléfono personal de un
cobrador no es ninguno de los dos — anotar en la decisión ① del contrato
(`libro_mayor/caja/README.md` / `libro_mayor/estado_cuenta/README.md`) cuando se retome
el diseño.

### 3c. G-12 (S/45) — sigue en el limbo, no entró a julio

A diferencia de los otros 7 retenidos, **G-12 no se cargó en `mesa_5` del repo principal**.
Está correctamente restaurado como blanco en el cierre de junio (§1), pero todavía no
existe en ningún ciclo activo. `PARA_AGOSTO.md` lista una condonación de MULTA para
Hernestina G-12 — eso es un tema distinto (condonar una multa), no reconocerle este pago
de S/45.

**Dos tools del ledger futuro, no confundir (quedó claro en la conversación):**
```
identificar_abono   blanco (sin dueño)  → predio     ← este es el caso de G-12
reasignar_abono     predio A            → predio B   ← lo que ya modela reidentificacion.xlsx
```
No hay archivo hoy que modele "blanco → predio" limpiamente (pre-ledger). Podría ser una
columna `TIPO` en `reidentificacion.xlsx` (`IDENTIFICACION` vs `REASIGNACION`) o un archivo
propio — no se decidió, quedó como pregunta abierta.

---

## 4. Pendientes heredados de ayer (07-19), siguen sin tocar

- **§6 regresión del RETOMAR de ayer** — no se corrió `5b_validacion` ni se revisó
  `6_corte/generar_lista.py` después de que `5_cobranza` empezara a aplicar los overlays
  de `devoluciones_aplicadas`/`reidentificacion` (71→70 elegibles a corte).
- Marcar reclamos resueltos en `4b_reclamos/outputs/reclamos_2026-07.xlsx`: V-6 → FUNDADO,
  T-7 sigue abierto (tema distinto, CONVENIO).
- `PARA_AGOSTO.md` — 4 condonaciones sin ejecutar (S-5, D-16, C1-17, Hernestina G-12).
- Limpieza cosmética: sacar filas G-18/I-7/V-16 de
  `jass_system - junio/5_cobranza/outputs/arrastre_devolucion_2026-06.xlsx` — **ya no
  aplica igual**, ahora esas filas quedan y se marcan `RESUELTO` en vez de eliminarse
  (la columna `ESTADO` de hoy reemplaza esa limpieza pendiente).

---

## No tocar

- Todo lo de la sección "No tocar" del RETOMAR de ayer (07-19) sigue vigente:
  `4_pagos/yape/validacion/main.py` (usuario resuelve a mano) ·
  `tests/test_integracion.py` de `4_pagos/efectivo` en ningún repo ·
  `correcciones_lote.xlsx` no es el lugar para reidentificaciones puntuales.
- **Nuevo:** no volver a correr `4_pagos/efectivo/main.py` (ni nada) en el repo junio sin
  necesidad puntual y verificación de timestamps antes/después — hoy se hizo de forma
  controlada y quedó verificado, pero no es una operación "segura por defecto" en un
  repo que se supone cerrado.
- No decidir 3b (abonos rezagados) sin resolver la pregunta (a)/(b) de arriba — cualquier
  archivo o código que se escriba antes de esa respuesta hay que rehacerlo.

---

## Estado git — nada commiteado

**Repo principal (`jass_system`):** sin commitear —
`shared/devoluciones_aplicadas.xlsx` (columna + fila P-6), `4b_reclamos/main.py` +
`inputs/reclamos_manuales.xlsx` (nuevo) + `README.md` + `docs/formato_reclamos.html`.

**Repo junio (`jass_system - junio`):** sin commitear —
columna `ESTADO` en `arrastre_devolucion_2026-06.xlsx`; `mesa_2-6.xlsx` +
`trazabilidad_2026-06.xlsx` restaurados a HEAD (vuelven a estar "limpios" respecto a git,
sin diff); `pagos_efectivo.xlsx`/`blancos_mes.xlsx` regenerados (no trackeados en git, no
aplica commitear). El resto del repo sigue con semanas de trabajo sin commitear, sin tocar
(ver sección ⚡ Contexto arriba).

**Diff completo de mesas guardado** (por si hace falta revisar el detalle de lo que se
restauró): `scratchpad/diff_mesas.txt` de esta sesión — es temporal, no persiste entre
sesiones. Si se necesita de nuevo, se puede re-generar comparando `git show HEAD:<mesa>`
contra el estado en el momento.
