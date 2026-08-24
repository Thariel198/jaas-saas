# RETOMAR — yape anotado que no entró a la JASS + buscar_pago · 2026-08-12

Todo commiteado en los 3 repos, working tree limpio. 64 tests en verde.

---

## ⚡ PRIMER PASO — el hallazgo que hay que llevar a la gente

**S/191 en 7 filas: plata que un cobrador anotó como yape recibido y que NO está
en la cuenta de la JASS.** Ciclos cerrados (junio + julio), o sea concluyente.

```
Wagner Trujillo   S/158   W-5 · F1-8 · H1-15 · E1-9 · S-2 · P-12
Yerald Romero     S/ 33   P-3
```

Reporte: `4b_reclamos/outputs/verificacion_yape_todos.xlsx`, filtrar
`ESTADO = NO_EXISTE`.

**Qué hacer:** pedirle al vecino la captura del yape. Ahí sale el número de
destino, que es lo único que distingue las tres explicaciones posibles:

```
el vecino yapeó al teléfono personal del cobrador y este no rindió
el vecino yapeó a otro número y le dijo al cobrador que ya pagó
error de anotación del cobrador
```

El reporte NO elige entre las tres, a propósito.

Caso más claro para empezar — **W-5**:

```
Wagner anotó en mesa_4 (04/07):  MONTO=0 · MONTO_YAPE=15
el vecino reclama textual:       "le pagó a Wagner mes anterior 15 + convenio 3 + techado 37"
                                 → el monto que dice el vecino y el anotado coinciden exacto
                                 → ese yape no está en la cuenta
```

### NO reclamar todavía por lo de agosto

`NO_EXISTE_PROVISIONAL` — 5 filas · S/266 (Yerald S/243 · Wagner S/13 · Máximo S/10).
El ciclo de agosto está ABIERTO: el reporte del banco llega al 06/08 y
motor_matching puede volver a correr. Se confirma al cerrar el mes.

> Antes de separar cerrado/abierto, Yerald figuraba con S/276. Sus S/243 son
> todos de agosto. Señalarlo hoy hubiera sido injusto.

---

## Las 2 herramientas nuevas (`4b_reclamos/herramienta/`)

```
verificar_yape.py     barre TODO el pueblo: cada fila con MONTO_YAPE > 0 en las
                      mesas, cruzada contra el reporte del banco
                      py verificar_yape.py --todos
                      → outputs/verificacion_yape_todos.xlsx

buscar_pago.py        los reclamos "ya pagué mes anterior" (solo ese tipo),
                      con 8 verificaciones por reclamo
                      py buscar_pago.py
                      → outputs/busqueda_pago_mes_anterior_2026-08.xlsx
```

**Por qué son dos y no una:** `buscar_pago` arranca desde los reclamos, así que
nunca encuentra a quien no reclamó. De los afectados por el yape no rendido,
**5 de 12 nunca dijeron nada** — pagaron, no se les acreditó, y les llegó la
deuda igual. Esos solo aparecen barriendo el pueblo entero.

### Estados de verificacion_yape

```
YA_ACREDITADO           24  el yape entró y motor_matching lo levantó → cobró
YA_REGULARIZADO          7  abonos_rezagados ya lo aplicó → cobró
NO_EXISTE                7  ciclo cerrado, no está en la cuenta → concluyente
NO_EXISTE_PROVISIONAL    5  ciclo abierto → confirmar al cerrar
POSIBLE                  2  calza monto y fecha pero nada confirma que sea suyo
```

Los 2 `POSIBLE` son F1-13 y V-8, ambos de S/8 — el monto más repetido del pueblo
(36 transacciones). Mirar esas 2 transacciones puntuales antes de concluir: si
resultan ser de ellos, es problema de matching y no de rendición.

---

## Los 29 reclamos de mes anterior — dónde quedaron

`busqueda_pago_mes_anterior_2026-08.xlsx`, columna `CONCLUSION`:

```
10  RESUELTO                      la boleta vigente ya no les cobra
 9  PEDIR RECIBO O CAPTURA        se buscó en las 8 fuentes, no hay rastro
 4  varios candidatos             se listan, no se elige (2+ posibles)
 2  YAPE QUE NO ENTRO A LA JASS
 1  REVISAR CANDIDATO ÚNICO: O-25
 1  PAGÓ PARTE
 1  PEDIR RECIBO
 1  SIN BOLETA                    S-16 no existe en DATA_boletas — revisar MZ/LT
```

**Patrón del pueblo:** 13 de los 18 reclamantes con arrastre vivo tienen **julio
2026** entre sus meses sin pago. El resto: junio (E-8, F1-1, I-9) o mayo (H-13).

Ejemplo cerrado — **J-8, Julia Victoria Robles Castillo**: pagó los 11 meses
desde oct-2025 menos julio; su arrastre de S/16 nació exactamente ahí. Dice que
pagó en efectivo y mostró recibo con código de billete B2495. No aparece en
ninguna fuente → ese recibo es la única evidencia que queda.

---

## Lo que quedó protegido (era el riesgo real de la sesión)

```
crear_templates.py        se niega si las mesas tienen cobros escritos, exige --force
utils_templates.py        crear_mesa_vacio() respalda SIEMPRE a
                          backup/mesas_pre_reset_<ts>/ antes de pisar
4_pagos/efectivo/tests/   conftest.py con fixture autouse: corre el _setup() del
                          propio archivo Y verifica que las rutas no apunten al
                          repo real. Un test nuevo sin aislar falla antes de escribir.
```

Antes de esto, `pytest 4_pagos/efectivo/tests` sobrescribía las mesas del ciclo
en curso. Destruyó mesa_1 (59→2 filas) y mesa_2 (106→2); se recuperaron con git.
**Ahora: 64 tests, 0 archivos reales modificados** (verificado con hashes).

---

## Pendientes

```
① confirmar los provisionales de agosto al cerrar el mes
   py verificar_yape.py --todos  → los 5 se resuelven solos

② mirar las 2 transacciones de los POSIBLE (F1-13 y V-8, S/8 c/u)
   si son de ellos → bug de matching, no de rendición

③ los 9 "PEDIR RECIBO" de buscar_pago necesitan trabajo de campo
   J-8 tiene el código de billete B2495 anotado, es el más concreto

④ auditar el ORIGEN del arrastre de los 13 que no pagaron en julio
   ¿por qué julio concentra tanto? es la pregunta de fondo sin responder

⑤ el filtro de abonos_rezagados matchea por (MZ, LT, monto EXACTO)
   si alguien rindió un monto distinto al anotado, no lo descuenta.
   Hoy no pasa (ninguno de los 12 está en rezagados), pero si el número
   no cuadra en el futuro, mirar ahí primero
```

---

## Correcciones que me hiciste y que valen para la próxima

```
· la cascada cobrando consumo+mantenimiento primero es CORRECTA, no un bug.
  Lo raro sería que cobre multa/acuerdos/convenio antes del arrastre —
  eso ahora se detecta como CASCADA_FUERA_DE_ORDEN (0 casos hoy)

· buscar_pago es SOLO para mes_anterior. Convenio y cuota tienen otra causa
  (el orden de la cascada) y hay que arreglar eso antes

· el reporte del banco abarca 3 meses: hay que acotar por la ventana del
  ciclo (ancla de corte) o se cruzan pagos de meses distintos

· no contaminar datos reales con tests

· las mesas vacías de un ciclo CERRADO son normales: 7_cierre las archiva en
  7_cierre/archivo/<mes>/ antes de resetear. Ese es el primer lugar donde
  mirar, no las copias sueltas del repo
```

Referencia permanente de todo lo de respaldos: `docs/RESPALDOS_Y_RECUPERACION.md`

---

## Commits de la sesión

```
2ef55ab  ciclo ABIERTO no da veredicto definitivo (NO_EXISTE_PROVISIONAL)
6c670a9  un mes SIN FILA se contaba como pagado (caso J-8)
f64cc4a  limpia párrafo duplicado
eb592b6  riesgo de los tests documentado como resuelto
4a1940a  los tests dejan de escribir sobre el inputs del ciclo en curso
82ee80d  corrige RESPALDOS — 7_cierre archiva, no hubo pérdida de datos
4495224  ventana del ciclo + extractor de motor_matching
bda176d  guarda contra el borrado de mesas + verificar_yape
16dffbc  buscar_pago.py

repo Julio    36ac565 · 95e2d18   ciclo 07 + mesas del archivo oficial de 7_cierre
repo Junio    22285d7             ciclo 06 completo
```
