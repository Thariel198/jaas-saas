# RETOMAR — Cierre ledger julio + boletas agosto + reportes de referencia · Sesión 2026-07-31

Handoff detallado. Sesión larguísima (Sonnet). Arrancó cerrando el checklist del
RETOMAR anterior (`RETOMAR_limpieza_ledger_y_reasignaciones_2026-07-31.md`,
ahora se puede borrar) y terminó regenerando las boletas de agosto completas
más dos reportes nuevos de auditoría de pagos.

---

## ⚡ PRIMER PASO al retomar

1. **Nada bloqueante queda del día de hoy.** Las boletas de agosto ya se
   imprimieron (PDF) y se validaron. Lo único pendiente real es decidir qué
   hacer con los pendientes de la sección 6 (abajo) — ninguno urge.
2. Si se va a tocar `2_planilla`/`3_boletas` de nuevo: **recordar sincronizar
   `planilla_cobrado.xlsx`** de `jass_system - Julio` al activo además de
   `seguimiento_pueblo.xlsx` y `arrastre_consolidado_2026-07.xlsx` — hoy se
   encontró tarde que ese archivo también se desincroniza y generaba
   inconsistencias en A-6 y otros.
3. Si se agregan más predios nuevos al padrón (como A-5A hoy): correr
   `1_lecturas/proponer_sincronizacion.py` + `aplicar_sincronizacion.py`
   ANTES de `2_planilla`, no parchar `DATA_boletas.xlsx` a mano — la sync real
   arregla nombres/eliminaciones de raíz, el parche a mano se pierde en cada
   regeneración.

---

## 1. Checklist del ledger de julio — CERRADO

Los 8 casos del RETOMAR anterior (T-7, S-5, G-18, D-1, F-12, F1-10, C-19, R-5)
más D-16 e I-2B (que ya estaban resueltos) quedaron verificados tras correr
`5_cobranza --force` en `jass_system - Julio` (ciclo 17) tras corregir dos
errores propios en el camino:

- **T-7 MES bug**: al registrar el pago histórico de T-7 con `MES=2026-06`
  en vez de `2026-07`, se ignoró un evento de julio ya existente y el saldo
  salió mal calculado — restaurado desde backup y rehecho con `MES=2026-07`.
- **F-12 mal etiquetado**: el redirect MULTA→CONVENIO de F-12 se hizo primero
  con `MES=2026-06` (por error propio, "experimentando" sin que el usuario lo
  pidiera) — corregido re-etiquetando a julio, mismo patrón que el resto de
  la cola.

Después del rerun, 5 de los 8 quedaron con el "contador tuerto" esperado
(PAGO manual → AJUSTE automático negativo) y se estabilizaron con el AJUSTE
manual +X de siempre (mismo patrón F-12 ya documentado). **F-12 en sí nunca se
disparó via el precursor** (su multa pagada estaba tageada `MES=2026-06`, el
rerun de julio no la tocaba) — se aplicó directo a mano, mismo resultado.

`notas_2026-07.xlsx` fila 174 (T-7) y 167 (D-1) marcadas `RESUELTO`.

---

## 2. Duplicados de predios eliminados — deuda cancelada en el ledger

`C1-17`, `C-29A`, `Q-16`, `S-14` seguían con deuda real en
`seguimiento_pueblo.xlsx` (MULTA/ACUERDOS) aunque el predio ya no existe
(son el mismo predio/persona que C1-9, C-34, Q-13, S-13 respectivamente,
ya eliminados del padrón por override). Por instrucción directa del usuario
("esa deuda ya está incluida en el lote que sobrevive, elimínala") se
canceló con AJUSTE directo (source=manual, motivo documentando el duplicado):

```
C1-17 MULTA -30   |   C-29A MULTA -20   |   Q-16 ACUERDOS -75
S-14  MULTA -50 + ACUERDOS -75
```

También se descubrió y corrigió: `2_planilla` arma su lista de predios desde
`registro_operario_acumulado.xlsx` (lecturas), **no** desde el padrón — un
override `ELIMINAR` en el padrón no basta por sí solo para que un predio deje
de aparecer. La sync real de `1_lecturas` (sección 4) es la que lo saca de
raíz; antes de eso, solo aparecía si tenía deuda de pueblo pendiente
(`_extra_keys_deuda_pueblo` en `2_planilla/main.py`, no revisa si el predio
sigue existiendo — mismo patrón estructural que ya documentaba
`LEER_ANTES.md` para B-20/C-43).

---

## 3. W-4 y R-5 — parche manual de MES_ANTERIOR (solo agosto, julio sin tocar)

Vicki Masias Cusihuamán (W-4) dice que pagó en su mesa en julio; no hay
ningún rastro del pago en `pagos_efectivo.xlsx`, `trazabilidad_2026-07.xlsx`
ni `blancos_efectivo.xlsx`. Se parcheó `DEUDA_AGUA` a 0 en
`arrastre_consolidado_2026-07.xlsx` **solo para que la boleta de agosto salga
limpia** — julio no se tocó. Documentado en `LEER_ANTES.md` (sección nueva) y
en `shared/parches_manuales_pendientes_julio.xlsx`.

R-5 (Frank Kelvin Teran Masias) tenía el mismo síntoma y se parchó igual al
principio, pero el usuario confirmó que **no vino a pagar** — la deuda es
real. Revertido el mismo día (`DEUDA_AGUA` restaurado a 8). Ambos casos
quedaron documentados en el archivo de parches con su estado final
(PENDIENTE / REVERTIDO).

**Pendiente real:** encontrar el blanco/pago real de julio de W-4 (~S/17) —
sin eso, el parche de agosto es solo un tapón, no una corrección.

---

## 4. Sincronización real de `1_lecturas` — nombres y eliminaciones de raíz

Se encontró que varios overrides de padrón (nombres corregidos, predios
eliminados) nunca se propagaron a `registro_operario_acumulado.xlsx`, la
fuente real que usa `1_lecturas`/`2_planilla`. Se corrió el mecanismo real
en vez de seguir parchando `DATA_boletas.xlsx` a mano cada vez:

```
py 1_lecturas/proponer_sincronizacion.py   → detecta 9 deltas
py 1_lecturas/aplicar_sincronizacion.py    → los aplica
```

Resultado: 1 AGREGADO (A-5A, Isabel Puntillo — split de A-10), 5
SIN_SERVICIO (A-10A, C-29A, Q-16, Q-4, S-14), 3 RENAME (A-10↔A-5A, G-9
apellidos invertidos, Q-13 nombre correcto). Backups automáticos en
`1_lecturas/inputs/backups/`.

**A-5A sin lectura este ciclo** (el código no existía cuando el operario hizo
la ronda) — bloqueaba el cierre de `1_lecturas`. Resuelto por decisión del
usuario: "sin consumo este ciclo" (`MARC_ACT_corregido = MARC_ANT = 712`,
`resuelto_por = lote_nuevo_sin_lectura` en `correcciones_2026-08.xlsx`). Se
retoma su lectura real el próximo mes.

⚠ **Nota de diseño no resuelta:** `2_planilla` sigue con un caso residual —
`Q-13`/`G-9` habían quedado con el nombre viejo pese al override porque
`1_lecturas/main.py` vuelve a fusionar `registro_operario_mes.xlsx` (el
archivo CRUDO que anota el operario en papel, que no sabe del cambio) sobre
el acumulado ya corregido cada ciclo. La sync de hoy lo arregló para este
ciclo, pero **puede volver a pisarse el próximo mes** si el operario sigue
anotando el nombre viejo en papel. No hay fix de fondo para esto todavía —
requeriría que el papel/planilla de campo ya traiga el nombre corregido, o
que `1_lecturas` priorice el nombre del padrón sobre el de la hoja mensual.

---

## 5. Boletas de agosto — CERRADO, 556/556 impresas y validadas

Pipeline completo re-corrido varias veces por los hallazgos de arriba:
`2_planilla` → parche M-12 (MANTENIMIENTO=83, se pierde en cada regeneración,
**repetir siempre**) → `3_boletas/enriquecimiento` → `3_boletas/main.py` →
`validar_boletas.py`.

**Bug de secuencia propio, encontrado a mitad de sesión:** se imprimieron 560
boletas con `DATA_boletas.xlsx` de fecha 30/07 (un día viejo) porque no se
había re-corrido `enriquecimiento/main.py` después de regenerar
`planilla_2026-08.xlsx` con las correcciones de hoy — ninguna corrección del
día estaba reflejada. Se detectó por los timestamps, se corrigió, se volvió
a imprimir. Ocurrió 3 veces más por hallazgos sucesivos (duplicados
flotantes, W-4/R-5, nombres) — cada vez: limpiar `3_boletas/outputs`,
re-enriquecer, reaplicar M-12, volver a imprimir. **Resultado final: 556/556
correctos**, nada impreso físicamente todavía (solo generado/validado el PDF).

`3_boletas/inputs/backups/` tiene ~4 snapshots de `DATA_boletas.xlsx` de las
iteraciones — se pueden limpiar si se acumulan demasiado.

---

## 6. Reportes nuevos — auditoría de pagos con referencia

Dos scripts nuevos en `4b_reclamos/`:

- **`reporte_convenio_multa.py`** — audita predios con CONVENIO pendiente vs.
  cuánto ya pagaron de MULTA (si se redirige, ¿a cuántos les queda saldado?).
  `calcular_tabla()` da la lista, `corregir_tabla_por_redirects()` evita
  contar dos veces el dinero de un redirect (lee `reasignaciones_aplicacion.xlsx`).
- **`reporte_referencias_pago.py`** — agrega, debajo del historial mensual de
  cada predio, de dónde vino cada pago: Yape con fecha/hora exacta (oct-may
  desde la hoja "Reporte" de los archivos históricos; jun-jul cruzando
  `maestro_yape.xlsx` contra el banco crudo `shared/reporte_acumulado_procesado/`),
  o efectivo (con día/cobrador cuando se encuentra en
  `4_pagos/efectivo/trazabilidad/`). `verificar_predio()` chequea que el
  total de arriba cuadre con la suma de abajo, mes a mes.

**Output final:** `4b_reclamos/outputs/reporte_convenio_multa_referencias_2026-07.pdf`
— 66 páginas (portada + C-16 agregado a pedido + los 62/63 con convenio
pendiente). Verificado dos veces contra el PDF real (no solo el cálculo
Python) — el primer intento de verificación por regex tuvo falsos positivos
por el layout en columnas; el segundo, parseando por posición fija de línea,
confirmó que cuadra.

**Decisión de diseño tomada en el camino:** cuando arriba y abajo no cuadran
en junio/julio, el reporte **NO especula el origen** (se probó con notas tipo
"posiblemente exceso retenido" / "revisar blancos_efectivo" y el usuario las
sacó — generaban confusión, "eso es lo que la secretaria se inventó o
declaró luego, mejor no afirmar nada"). Ahora simplemente se muestra el pago
real encontrado y punto; si sobra o falta, no se explica.

**Bug de fondo encontrado y corregido en el camino:** `reporte_historico.py`
(`_filas_recientes()`) nunca calculaba consumo/mantenimiento/mes
anterior/corte para junio — solo para "el ciclo actual" (julio). Cualquier
predio que pagó en junio veía su consumo desaparecer de la tabla aunque sí lo
hubiera pagado (ej. C-16: pagó 33, la tabla solo mostraba 25 de multa).
Corregido con `_datos_ciclo()` genérico que ahora lee el `planilla_cobrado`
de CADA ciclo (junio vive en `jass_system - junio`, julio en el activo).
**Este fix vive en el módulo compartido — beneficia también al reporte
oficial `reporte_historico_CONFIRMACION_2026-07.pdf` (116 páginas), que
sigue generado con la versión vieja del código. No se regeneró hoy.**

---

## 7. Pendientes reales, sin resolver

- **W-4**: blanco de julio (~S/17) sin encontrar — parche de agosto es solo
  cosmético (ver sección 3).
- **F1-6 / F1-7**: el abono de S/30 (multa al día) sigue con nota
  contradictoria en `notas_2026-07.xlsx` ("corregir a F1-6" vs "este abono es
  para F1-7 tal cual está anotado") — nunca se resolvió cuál lote es. Ver
  `shared/abonos_rezagados.xlsx` fila 26.
- **B-29, C-45, E-14A**: filas huérfanas en `DATA_boletas.xlsx` sin `NOMBRES`
  (se caen solas por `dropna`, nunca se imprimen) — tienen deuda real
  flotante (283, 346, 75) bajo códigos que no están en el padrón actual.
  Nunca se investigó si son duplicados de otro predio (como C1-17/C-29A/etc)
  o casos genuinamente distintos. Ver conversación de hoy ~19:00.
- **13 predios** del `reporte_convenio_multa.py` que "cubrirían completo" si
  se redirige su multa a convenio, pero sin confirmación de que hayan
  asistido/no corresponda la multa (a diferencia de D-1/F-12 que sí se
  confirmaron y ejecutaron hoy) — quedan como candidatos, no tocados.
- **Q-4 (Grupo 3)**: pasos intermedios ruidosos (falló→revirtió→re-condonó)
  sin limpiar — mencionado por el usuario hace varias sesiones, nunca
  investigado a fondo.
- **B-19, F1-4, E-8, N-6, P-6, H1-2**: filas `RESUELTO` en `notas_2026-07.xlsx`
  sin ningún detalle de motivo — no se verificaron a fondo hoy (se listaron,
  no se investigó cada una).
- **S/550 Nivel 1a de `5b_validacion`**: gap conocido, aceptado hoy para
  poder avanzar (no bloqueaba lo urgente), causa raíz sigue sin encontrar.
- **`reporte_historico_CONFIRMACION_2026-07.pdf` (116 páginas)**: desactualizado
  respecto al fix de consumo de junio (sección 6) — no regenerado hoy.
- **Ampliar la revisión de "arriba/abajo cuadra" a TODO el sistema** (no solo
  los 62-66 de convenio) — el usuario lo pidió a mitad de sesión, se optimizó
  el código para que sea viable (cache de los 8 archivos históricos) pero no
  se llegó a correr sobre el universo completo de predios.
- **Nada de lo de hoy se imprimió físicamente** — boletas de agosto y los
  reportes de referencia están listos en PDF, falta la entrega real.

## Cómo cerrar este RETOMAR

Cuando se resuelvan los pendientes de la sección 7 que se decida priorizar
(ninguno es bloqueante hoy), y se confirme que las boletas de agosto se
entregaron, borrar este archivo.
