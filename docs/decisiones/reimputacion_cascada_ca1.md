# Re-imputación de pagos al orden CA1 — CONVENIO → ACUERDOS → MULTA

**Estado: plan cerrado, NADA escrito en el ledger todavía.** Decisiones tomadas con el
usuario el 2026-08-13 (sesión Opus). Lo que falta ejecutar está en "Los 6 pasos".

---

## La decisión

El tramo pueblo de la cascada pasa de `MULTA → ACUERDOS → CONVENIO` a
`CONVENIO → ACUERDOS → MULTA`, y la plata ya imputada con el orden viejo se re-acomoda
al orden nuevo.

```
ORDEN VIEJO (el que aplicó el código)        ORDEN NUEVO (CA1)
P1 agua (consumo+mant+mes anterior)          P1 agua (consumo+mant+mes anterior)
P2 corte                                     P2 corte
P3 MULTA          ← se llevaba la plata      P3 CONVENIO      ← solo se salda con plata
P4 ACUERDOS                                  P4 ACUERDOS      ← solo se salda con plata
P5 CONVENIO       ← quedaba al final         P5 MULTA         ← faena o exoneración
```

**Por qué, confirmado por la directiva:** *el dinero cubre primero lo que SOLO el dinero
puede saldar.* La multa se recupera trabajando en faena o se exonera por asistencia;
convenio (medidor) y acuerdos (techado/campo) son obligatoriamente plata. Al cobrarse
la multa primero, el sistema dejaba dos reclamos recurrentes que eran ciertos:
*"ya pagué mi medidor"* y *"ya pagué techado y campo"*.

Decisión de dominio previa: `libro_mayor/dominio` CA1, 2026-07-23. El dominio y 9 READMEs
ya están en el orden nuevo desde entonces; **el código vivo nunca se cambió** —
`5_cobranza/main.py::_descomponer_saldo` sigue con el orden viejo. Este documento cierra
esa brecha.

## Los dos movimientos

```
① solo si el convenio es MEDIDOR (deuda sembrada ≤ 100 y fuera de las listas de instalación)
     MULTA ──► CONVENIO         hasta que el convenio quede en 0
     si no alcanza:
     ACUERDOS ──► CONVENIO      hasta que el convenio quede en 0

② para TODOS los predios
     lo que sobre de MULTA ──► ACUERDOS   hasta que acuerdos quede en 0

La MULTA solo ABSORBE deuda; nunca recibe plata.
```

Clasificación del convenio (`reporte_reimputacion_cascada.py::clasificar_convenio`),
fuente: `obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx`

| clase | criterio | ¿recibe plata? |
|---|---|---|
| `MEDIDOR` | saldo ≤ 100 y fuera de las listas | **sí** |
| `SIN_CONVENIO` | no debe convenio | no (solo entra al paso ②) |
| `INSTALACION` | saldo > 100, o en NUEVAS INSTALACIONES / ANTERIOR DIRECTIVA | **no** — decidido intencional |
| `REACTIVACION` | en la hoja REACTIVACION | **no** — decidido intencional |

`INSTALACION` y `REACTIVACION` quedan afuera a propósito (decisión R5, 2026-08-13): el
alcance es la deuda de medidor sembrada de ≤100, no la instalación.

---

## Qué plata se puede mover, y qué plata no

Tres causas de descarte, las tres verificadas contra el ledger el 2026-08-13.

### ① Plata pre-génesis — ya hizo su trabajo

El ledger nace en junio 2026: todos los `CARGO` son `MES=2026-06 · SOURCE=sembrar_seguimiento_pueblo`,
y **el monto sembrado es lo que quedaba debiendo**, ya neto de lo pagado antes.

```
N-6     multa 50  − 30 pagado feb  =  20 sembrado en junio    ✔
C1-14   multa 30  − 10 pagado feb  =  20 sembrado             ✔
B-12A   multa 60  − 30 pagado feb  =  30 sembrado             ✔
K-9     multa 105 − 75 pagado may  =  30 sembrado             ✔
A-4     multa 60  − 30 pagado mar  =  30 sembrado             ✔
P-12    multa 30  − 30 pagado feb  =   0 → no se sembró nada  ✔  (6 de 6)
```

Un pago de feb-may ya bajó la deuda antes de la siembra. Volver a moverlo lo gasta dos
veces. **La ventana del reporte era `feb→jul` (`MESES_VENTANA`) y debe derivarse de los
meses en que el concepto tiene eventos en el ledger** — no una constante, para que un
backfill futuro la extienda solo.

De los 86 predios que se mueven, solo 6 tenían pago de multa pre-génesis (S/205). Los
otros 80 tienen todo su pago dentro del ledger.

### ② Crédito que no es plata

`4b_reclamos/reporte_historico.py:398` arma la columna que el reporte usa como pago:

```python
saldado = float(r["PAGO"]) + max(0.0, -float(r["DEBIA"]))
#          plata real        + crédito que vino de un AJUSTE
```

Está bien para lo que ese reporte hace (mostrarle al vecino qué quedó saldado), y mal
como fuente de "cuánta plata puedo mover". **El tope de cada movimiento debe ser
`Σ PAGO` de clases con plata real, por (predio, CONCEPTO que cede)** —
`COBRANZA`, `ABONO_REZAGADO`, `DECLARACION`, `DECLARACION_SECRETARIA`.

Decisión R2 (2026-08-13): `EXONERACION` **no es plata** — la directiva no cubrió nada,
eliminó la deuda; nadie va a pagar eso. `CORRECCION_SISTEMA` menos todavía: ese crédito
nació de arreglar un defecto nuestro (el bug de signo del 07/08 y el race condition del
13/07).

**El tope es por (predio, concepto que cede)**, no por predio: la plata que A-4 pagó a
`ACUERDOS` no puede financiar un movimiento que sale de su `MULTA`. Pero sí puede financiar
el movimiento `ACUERDOS → CONVENIO` del paso ① — que es distinto, y es lo que en la práctica
resolvió a 3 de los predios que parecían recortados (ver "Los 7 predios afectados").

### ③ Sin pago registrado

Predios cuyo concepto que cede no tiene ningún `PAGO` en el ledger. No hay nada que
reubicar.

---

## Los números — corrida del 2026-08-13 con el paso 0 ya aplicado

Las 14 validaciones en verde, incluidas las 3 nuevas de la capa 2.

```
DEUDA TOTAL              16,124.50 → 16,124.50   idéntica ✔
PLATA REAL TOTAL          7,845.50 →  7,845.50   idéntica ✔
deuda MULTA               3,640.00 →  6,063.50
deuda ACUERDOS            6,659.50 →  5,647.00
deuda CONVENIO            5,825.00 →  4,414.00
predios que mueven más que su plata real (por concepto) ....... 0 ✔
en el agregado: lo que sale ≤ plata real de multa+acuerdos .... OK ✔
movimientos vivos sin MES de origen (asiento inescribible) .... 0 ✔
```

```
88 predios se mueven · S/3,003.50

LIMPIO      187 predios   pedía 2,603.50 → mueve 2,603.50
AMPLIADO     10 predios   pedía    96.00 → mueve   302.00
RECORTADO     4 predios   pedía   195.00 → mueve    98.00
EXCLUIDO      3 predios   pedía   196.00 → mueve     0.00
```

`AMPLIADO` es el sentido que faltaba: la ventana vieja feb-jul **no veía la plata de
agosto**, que sí está en el ledger y sí se puede mover. Etiquetar solo el recorte lo
escondía. Los 10 son todos "plata de 2026-08 que la ventana vieja no contaba".

### Quién escribe qué

```
ASIENTO MANUAL      2026-06   S/1,326.00  ┐ ciclos CERRADOS
                    2026-07   S/1,282.50  ┘ S/2,608.50
LO HACE EL CÓDIGO   2026-08   S/  395.00    ciclo ABIERTO → pasos 4b + 5
                              ──────────
                              S/3,003.50
```

⚠ **Los de agosto NO se escriben a mano.** `_reconciliar_pagos_pueblo` recalcula el mes
entero y sabe revertir (`delta < 0` → AJUSTE), y su `SET_TIENE` cuenta solo eventos con
`source="5_cobranza"` — no vería un asiento manual y movería la plata otra vez.

### Los 7 predios afectados

| predio | nombre | pide | mueve | estado | causa |
|---|---|---|---|---|---|
| J-6 | Capilla del Pueblo | 100 | 0 | EXCLUIDO | `EXONERACION`, no es plata |
| J-1 | Comedor Popular | 75 | 0 | EXCLUIDO | `EXONERACION`, no es plata |
| N-6 | David Juárez Toledo | 21 | 0 | EXCLUIDO | pago pre-génesis S/30 |
| K-9 | Fortunato Vargas Cabello | 75 | 30 | RECORTADO | pago pre-génesis S/75 |
| B-12A | Rosmeri Sánchez Espada | 60 | 30 | RECORTADO | pago pre-génesis S/30 |
| C1-14 | Mari Celidonio Chávez | 30 | 20 | RECORTADO | pago pre-génesis S/30 |
| D1-6 | Onita Ponte Eguizábal | 30 | 18 | RECORTADO | `CORRECCION_SISTEMA` |

**Su reclamo queda abierto, total o parcialmente. Hay que hablar con ellos.**

Eran 10 en el análisis preliminar y bajaron a 7 **sin cambiar ninguna decisión**: al calcular
el reparto con la plata real, el paso ① (`ACUERDOS → CONVENIO`) cubre lo que a la multa no le
alcanza. Ese análisis preliminar capaba cada movimiento por separado en vez de re-correr la
cascada, y por eso los daba por recortados:

```
A-4   convenio 75 = 30 de multa + 45 de ACUERDOS   → cubierto completo · LIMPIO
P-12  convenio 50 = 50 de ACUERDOS                 → cubierto completo · LIMPIO
L-4   convenio 25 = 17 de multa +  8 de ACUERDOS   → cubierto completo · LIMPIO
```

### J-1 y J-6 — centros del pueblo · CERRADO 2026-08-13

Decisión del usuario: **el convenio se paga, se exonera la multa nomás.** El medidor es un
bien que recibieron; la multa y los acuerdos de asamblea no aplican a un local comunal, y
sembrárselos fue un error.

El ledger muestra que **nunca pagaron nada** — no tienen un solo evento `PAGO`. Por eso el
movimiento que el reporte proponía era 100% crédito de exoneración: estaba intentando
pagarles el medidor con la multa que la directiva les había perdonado.

```
J-1  COMEDOR POPULAR CLUB DE MADRES        clase de convenio: MEDIDOR
   MULTA      CARGO  20 (jun) → AJUSTE −20 EXONERACION (jul)   saldo 0
   ACUERDOS   CARGO  75 (jun) → AJUSTE −75 EXONERACION (jul)   saldo 0
   CONVENIO   CARGO  75 (jun) · sin pago · sin ajuste          saldo 75  ← SE COBRA

J-6  CAPILLA DEL PUEBLO                    clase de convenio: MEDIDOR
   MULTA      CARGO  50 (jun) → AJUSTE −50 EXONERACION (jul)   saldo 0
   ACUERDOS   CARGO  75 (jun) → AJUSTE −75 EXONERACION (jul)   saldo 0
   CONVENIO   CARGO 100 (jun) · sin pago · sin ajuste          saldo 100 ← SE COBRA
```

**A cobrar: S/175 de medidor entre los dos.** No hay nada que re-imputar en ellos.

### Las 3 exoneraciones de ACUERDOS — CERRADO 2026-08-13: quedan como están

Regla vigente **hacia adelante**: techado y campo (`ACUERDOS`) se pagan; solo la multa se
exonera. Pero las 3 exoneraciones de acuerdos que ya están en el ledger **se mantienen**,
porque cada una tiene su `MOTIVO` escrito y ese motivo es de la directiva.

> **El MOTIVO escrito en el ledger le gana a la memoria de cualquiera — incluida la del
> usuario.** Textual, 2026-08-13: *"si el ledger ya tiene un motivo de que acuerdos se
> condona entonces que se condone, el ledger le gana a mi memoria; obviamente en su momento
> la directiva habrá dicho condona acuerdos y a mí se me olvidó."*
>
> Por eso el campo `MOTIVO` no es decoración: es lo que impide que una decisión de la
> directiva se revierta seis meses después porque nadie se acuerda de por qué se tomó.
> Y por eso los **29 `AJUSTE` sin `MOTIVO`** dejan de ser deuda de documentación menor:
> son los únicos asientos que no podrían defenderse solos en esta misma discusión.

Estado final: **S/175 a cobrar** (J-1 convenio 75 + J-6 convenio 100). Los acuerdos de
J-1 y J-6 quedan condonados. Nada que revertir.

El detalle de las 7, para que quede el registro de qué respalda a cada una:

```
MULTA (4) — en regla, no se tocan
   J-1  −20   directiva 01/08 · faena y reunión
   O-2  −30   directiva 01/08 · Carmen Ingaruca
   C-21 −50   condonación institucional 09/08 · Puesto de Salud
   J-6  −50   condonación institucional 09/08 · Capilla del Pueblo

ACUERDOS (3) — hay que decidir
   J-1  −75   "exonerado por la directiva 01/08 — techado y campo"
              ← decisión explícita de la directiva sobre acuerdos
   J-6  −75   "condonación institucional — bienes del pueblo sin vecino que participe,
              no generan MULTA/ACUERDOS reales" · figura EXONERADO en registro_cortes
              desde 2026-02
   S-5  −40   NO es una exoneración real: el pago SÍ existía (Wagner lo retuvo en su
              yape personal, recuperado en efectivo, abonos_rezagados S/71). Al liberar
              los S/40 de corte, la plata real cubrió los acuerdos.
              → NO revertir. La CLASE está mal puesta (es plata real, no una exención):
                deuda de documentación, no una decisión.
```

```
a cobrar                             J-1        J-6      total
  convenio                          75.00     100.00    175.00
  acuerdos                       condonado  condonado        —
```

C-21 (Puesto de Salud), O-2 y S-5 tienen saldo 0 en los tres conceptos y **no están en el
reporte de re-imputación** — no los toca nada de esto.

**Lo único que queda de esto es documentación:** la `CLASE` de S-5 está mal puesta (es
plata real, no una exención) y los 29 `AJUSTE` sin `MOTIVO` siguen sin poder defenderse.

### P-12 y N-6 — NO se exoneran · verificado y cerrado 2026-08-13

⚠ Con el paso 0 aplicado, **P-12 ya no está en la lista**: su convenio de S/50 queda cubierto
entero con su plata de `ACUERDOS` (queda debiendo acuerdos 50 en su lugar — misma deuda total,
aplicada donde corresponde). El único que sigue excluido de los dos es **N-6**. La decisión de
abajo se tomó cuando los dos figuraban, y vale igual: ninguno se exonera.

Salieron en la misma lista de excluidos que J-1/J-6, pero **por una causa distinta**, y la
diferencia decide el caso:

```
J-1 · J-6     excluidos porque su crédito era EXONERACION de la directiva
              ← decisión documentada, con MOTIVO en el ledger

N-6 · P-12    excluidos porque el REPORTE contaba un pago de febrero que ya estaba
              descontado en el CARGO sembrado en junio
              ← artefacto del bug de la ventana, el que corrige el paso 0
```

Estado real en el ledger, y ninguno de los dos tiene reclamo registrado en los 5 archivos
(may · jun · jul · ago):

```
N-6  DAVID JUAREZ TOLEDO — vecino que paga, no un bien del pueblo
   ACUERDOS  CARGO 50 (jun) → PAGO 29 (jul, COBRANZA = plata real)   saldo 21  DEBE
   MULTA     CARGO 20 (jun) → nada más                              saldo 20  DEBE

P-12 JUDITH VENTURO ROSALES — no tiene multa que exonerar
   ACUERDOS  CARGO 50 → PAGO 50                                     saldo  0
   CONVENIO  CARGO 100 → PAGO 25 → AJUSTE −25 (corrección génesis)   saldo 50  DEBE
   MULTA     sin un solo evento                                     saldo  0
```

> **Aparecer en una lista de auditoría no es motivo para exonerar.** Decisión del usuario,
> 2026-08-13: *"si no está en el mismo saco que centros y tu auditoría fue la que los metió
> al mismo saco, entonces sí debe N-6."* El `MOTIVO` habría dicho "apareció en una lista de
> auditoría" — exactamente el tipo de motivo que no se puede defender después, al revés del
> principio de la sección anterior.
>
> Si algún día se exonera la multa de N-6, el respaldo tiene que ser la **asistencia**: sus
> S/20 son la tarifa de reunión, y la fuente son las hojas de asistencia que usa
> `obligaciones/`. Eso da un motivo real; la lista de excluidos no.

---

## El mecanismo de escritura

### Precursor + ledger, los dos

Regla vigente (`docs/diario/2026-08-03_solucion_precursor_mas_ledger.html`): se escriben
siempre los dos, atados por `AUDIT_REF`. El precursor existe para que **si mañana se
descarta el ledger y se rehace por backfill, los precursores cuenten la historia solos.**
Confirmado 2026-08-13: se mantiene el patrón, con el puntero de cada fila bien dirigido.

```
shared/reasignaciones_aplicacion.xlsx     84 filas, una por movimiento
    MZ · LT · CONCEPTO_ORIGEN · CONCEPTO_DESTINO · MONTO
    MES_ANO = (vacío)  → SOLO REGISTRO: ninguna corrida lo re-aplica
    MOTIVO  = re-imputación CA1 + mes de origen del pago + puntero al contrato
    REF_TRANSACCION = el AUDIT_REF del par de asientos

shared/seguimiento_pueblo.xlsx            2 asientos por movimiento
    AJUSTE +X en el concepto que CEDE     (sube su deuda)
    AJUSTE −X en el concepto que RECIBE   (baja su deuda)
    CLASE = REASIGNACION   ("la deuda se movió entre conceptos, sin plata nueva")
    MES   = mes del PAGO que se mueve (2026-06 o 2026-07)
    AUDIT_REF = reimputacion_ca1_<MZ>_<LT>_<ORIGEN>_<DESTINO>   ← determinista
```

`REASIGNACION` no está en `CLASES_SUMAN_CAJA`: **la caja no se mueve ni un sol.** Correcto,
porque no entra plata nueva — solo cambia de dueño dentro del mismo vecino.

### ⚠ El asiento se escribe ANTES de correr 5_cobranza — `SALDO` es una columna almacenada

`SALDO` no se recalcula: `_saldo_previo()` toma la última fila con `MES <= mes` ordenada
por `(MES, TIMESTAMP)` y le suma el monto; `get_saldo()` devuelve el `SALDO` de esa última
fila.

```
A-4 MULTA hoy
   MES=2026-06  CARGO 30          SALDO=30
   MES=2026-07  PAGO  30          SALDO= 0   ← get_saldo(2026-08) lee ESTA fila

si escribo un AJUSTE con MES=2026-06:
   MES=2026-06  AJUSTE +30 (hoy)  SALDO=60   ← ordena ANTES de julio
   MES=2026-07  PAGO  30          SALDO= 0   ← su SALDO no se recalcula
   → get_saldo(2026-08) sigue en 0: el asiento queda SEPULTADO, sin efecto y en silencio
```

Medido el 2026-08-13: **0 sepultados hoy** — en los 90 casos el último `PAGO` del concepto
que cede es también el último evento de ese (predio, concepto). Pero es una propiedad del
dato de hoy, no una garantía: si una corrida de `5_cobranza` escribe un evento de agosto
para uno de esos predios antes de que escribamos, ese movimiento pasa a estar sepultado.

**Regla:** escribir primero, y verificar predio por predio que el asiento quedó como
última fila (validación ⑩).

---

## ⚠ El sepultamiento tiene DOS lados — y por eso la re-imputación va después del cierre

Descubierto el 2026-08-13 **escribiendo**: la primera corrida de `--escribir` metió los 168
asientos, la validación falló, y se restauró el backup en el acto (ledger intacto, 1563
eventos, 0 asientos).

```
lo que se había medido   el lado que CEDE queda sepultado?  → 0 de 84   ✔
lo que faltaba medir     el lado que RECIBE?                → 24 de 84 · S/665  ✘

   B-5  MULTA→CONVENIO 50 · asiento en 2026-06
        su CONVENIO tiene un evento en 2026-08 → el −50 queda sepultado
        la multa subió 50 y el convenio no bajó → la deuda del predio subió 50

   resultado: deuda agregada 3,871.50 → 4,017.50  (+146, debía quedar igual)
```

El concepto que recibe tiene su **propia** cadena de eventos, y casi siempre uno posterior
(los pagos de agosto). Fechar los dos lados en el mes del pago solo funciona si **ninguno de
los dos** conceptos tiene eventos después.

Y no hay un mes que sirva para todo a la vez mientras agosto esté abierto:

```
asiento en 2026-06  sepultado si el concepto tiene fila de 07 u 08
asiento en 2026-07  sirve para planilla_2026-08 (lee hasta julio), pero NO para el
                    saldo de hoy si el concepto tiene fila de agosto
asiento en 2026-08  sirve para el saldo de hoy, pero planilla_2026-08 no lo ve
```

### La solución: timing, no un mes especial (decisión del usuario, 2026-08-13)

**Agosto se cierra primero; la re-imputación se escribe después del cierre y antes de generar
la planilla de septiembre.** En ese momento no hay ningún evento posterior al que enterrarse:
el asiento es la última fila de los dos conceptos, sin casos especiales.

```
AHORA          5_cobranza --force  → la cascada ya corregida re-imputa agosto SOLA
               6_corte → lista de corte normal → ... → 7_cierre

POST-CIERRE    re-correr el reporte (saldos post-cierre) → congelar → escribir → validar
               todos los asientos con MES = 2026-08

SEPTIEMBRE     2_planilla lee get_saldos_bulk(concepto, "2026-08") → toma todo
               5_cobranza corre con la cascada nueva sobre una planilla ya correcta
```

⚠ Los asientos van con **MES = 2026-08**, no 2026-09: `2_planilla` lee el mes ANTERIOR, así
que con 2026-09 la planilla de septiembre no los vería y la corrección se atrasaría un mes.

**Esto elimina el paso 4b**: nunca hace falta regenerar `planilla_2026-08`, porque la planilla
que se genera después de la re-imputación es la de septiembre. Y elimina los 14 movimientos
de agosto del contrato manual: los hace `5_cobranza --force` al correr con la cascada nueva.

> El contrato congelado hoy (`contrato_reimputacion_ca1_20260813_141120.xlsx`, 84 movimientos,
> S/2,608.50) queda como **ensayo**. Hay que re-congelar después del cierre: los saldos van a
> ser otros, y el tope de cada movimiento se calcula contra el saldo vigente del destino.

## Lo que sigue vale como historia — la planilla de agosto es una FOTO del ledger

Encontrado el 2026-08-13, después de codificar el paso 0. Es el riesgo más concreto de todo
el cambio y por eso existe el paso **4b**.

Las dos fuentes de deuda del sistema no son la misma, y solo una la arregla la re-imputación:

```
AGUA · CORTE              planilla → arrastre_consolidado → planilla     cadena propia
MULTA·ACUERDOS·CONVENIO   seguimiento_pueblo (ledger)     ← FUENTE ÚNICA

  2_planilla/main.py:144  _join_saldo_pueblo()
      repo.get_saldos_bulk(concepto, mes_ant)
      "MULTA/ACUERDOS_ASAMBLEA/CONVENIO vienen de seguimiento_pueblo, NO del
       consolidado — writer único es seguimiento_repo"
```

Buena noticia: re-imputar el ledger **sí** llega a la planilla. Mala: solo cuando
`2_planilla` vuelve a correr, y `planilla_2026-08.xlsx` es una foto tomada al cierre de
julio, con el split viejo.

```
5_cobranza lee la deuda por concepto DE LA PLANILLA, no del ledger (main.py:833-835)
        r["convenio"], r["multa"], r["acuerdos_asamblea"]
        ↓
si se corre --force con el orden nuevo sobre la foto vieja:

   A-4 · ledger post-re-imputación   convenio = 0   (cubierto con 30 multa + 45 acuerdos)
         planilla_2026-08 sigue en   convenio = 75
         la cascada nueva pone el convenio PRIMERO → le vuelve a cobrar 75
         → PAGO sobre saldo 0 → saldo NEGATIVO (salta el warning de main.py:2416)

   LE COBRAMOS EL MEDIDOR DOS VECES.
```

**El arreglo es el paso 4b: regenerar `planilla_2026-08` con `2_planilla` después de escribir
los asientos y antes de correr `5_cobranza`.** `get_saldos_bulk(c, "2026-07")` filtra
`MES <= 2026-07`, así que toma los asientos de jun/jul y **no** los pagos de agosto — la foto
arranca exactamente en el estado correcto.

```
⚠ 2_planilla NO tiene backup ni preservación (grep de backup/preserv: 0 coincidencias).
  Sobreescribe planilla_2026-08 a ciegas → respaldarla a mano ANTES, con fecha en el nombre.

⚠ Después de 4b, planilla_2026-08 y las boletas de agosto ya entregadas van a DECIR COSAS
  DISTINTAS. Es deliberado (el vecino ve el cambio en septiembre). Nadie debe "arreglarlo"
  regenerando boletas.

✔ 4b no es trabajo que invente este cambio: re-correr 2_planilla de agosto ya era el arreglo
  pendiente desde el 08/08 por los 11 predios con deuda negativa — cuya causa raíz era
  exactamente esta (planilla congelada antes de que el ledger cambiara).
```

## Validación en 4 capas

Las capas 1 y 3 son mecánicas. La capa 2 es la que faltaba y es la que atrapa los 12
movimientos sucios. La capa 4 no valida: mide, para decidir y comunicar.

```
CAPA 1 · CONSERVACIÓN — "¿la aritmética cierra?"
   ① plata total            antes = después
   ② deuda total            antes = después
   ③ por (CONCEPTO, MES)    Σ sale = Σ entra    ← por MES también, no solo por concepto:
                            si no, mover multa-junio a convenio-agosto "cuadra" y cambió
                            plata de un ciclo a otro
   ④ por predio             antes = después     ← el único que detecta plata movida ENTRE vecinos
   ⑤ ningún saldo negativo nuevo

CAPA 2 · SIGNIFICADO — "¿lo que movemos era movible?"
   ⑥ lo movido ≤ Σ PAGO de clase real, por (predio, concepto que cede)
   ⑦ el concepto origen no tiene EXONERACION viva
   ⑧ el crédito no viene de CORRECCION_SISTEMA

CAPA 3 · REALIDAD — "¿sigue siendo el mismo mundo?"
   ⑨ saldo de hoy == saldo del contrato congelado, predio por predio
   ⑩ resultado escrito == el DESPUES que predijo el contrato  ← atrapa el sepultamiento
   ⑪ AUDIT_REF determinista: el segundo intento no escribe nada

CAPA 4 · CONSECUENCIA — mide, no valida
   ⑫ cuántos pasan a deber SOLO multa (hoy: 65) → lista para comunicar
```

**Por qué la capa 1 sola no alcanzaba** — el caso J-6, Capilla del Pueblo: mover S/50 de
su multa (crédito de `EXONERACION`, plata real S/0) al convenio conserva el total, conserva
por concepto, conserva por predio y no deja negativos. Las 4 validaciones en verde, y el
resultado es pagarle el convenio con una exoneración. Lo mismo con la doble aplicación:
aplicar el movimiento N veces satisface la conservación siempre.

---

## La ola de exoneraciones que viene después — y por qué el orden importa

Confirmado por el usuario el 2026-08-13: después de este cambio **la mayoría del pueblo
queda debiendo multa, van a reclamar, y se va a exonerar a varios.** Eso no es un efecto
colateral: es el circuito que la directiva quiso al poner la multa última. La plata cubre
lo que solo la plata salda (convenio, acuerdos) y la multa se resuelve con faena o
exoneración.

```
⚠ REGLA DE ORDEN — re-imputar PRIMERO, exonerar DESPUÉS

  re-imputación → exoneración        ✔ correcto
     el AJUSTE +X devuelve la deuda de multa · después la exoneración la perdona
     el vecino ve su convenio saldado Y su multa en 0

  exoneración → re-imputación        ✘ rompe las dos cosas
     la exoneración deja la multa en 0 · el AJUSTE +X resucita una deuda ya perdonada
     y la validación ⑦ (excluir a los que tienen EXONERACION viva) va a sacar a ese
     predio del movimiento → su reclamo de convenio queda sin resolver
```

**El número que se le muestra a la directiva es un pico, no un estado final.** La deuda de
multa sube de 3,640 a ~6,109.50 al re-imputar, y **baja de nuevo** a medida que se
exonera. Conviene decirlo así en la reunión, o el salto se lee como morosidad nueva
cuando en realidad es multa que siempre estuvo impaga y estaba tapada por el orden viejo.

**Lo que mantiene el ledger honesto cuando esa lista crezca** (hoy hay 7 `EXONERACION`,
van a ser decenas):

```
cada exoneración va con CLASE = EXONERACION, nunca se confunde con REASIGNACION
su crédito NO es plata y NO es movible  → capas 2 ⑦ y ⑧ lo excluyen para siempre
   ↑ esto es lo que evita que en una re-imputación futura la exoneración de un vecino
     termine pagándole el convenio a otro concepto, que es el caso J-1/J-6 de hoy
```

Con ese volumen, exonerar deja de ser un ajuste a mano por caso y pasa a ser el flujo
principal del ciclo: **necesita su propia herramienta y su propio registro**, con la
asistencia a faena/reunión como respaldo. No está diseñado todavía — queda anotado acá
porque es lo primero que va a hacer falta después de este cambio.

## Lo que este cambio NO toca

```
BOLETAS DE AGOSTO       no se regeneran. Salieron el 31/07 (3_boletas/Outputs/
                        CONSOLIDADO_*.pdf) y se cobraron el 01/08, y DATA_boletas.xlsx
                        tiene ~25 correcciones hechas a mano ese día (5 backups en
                        3_boletas/inputs/backups/). Regenerar las pisaría — es la
                        Regla 9 de CLAUDE.md, el bug B5.
                        → el vecino ve el cambio en las boletas de SEPTIEMBRE
                        → el paso 5 NO corre 2_planilla ni 3_boletas

LISTA DE CORTE          no se afecta y no se congela. El único corte vigente es por no
                        pago de agua en 2 meses: 6_corte usa CONCEPTOS_SALDO =
                        [AGUA, MANTENIMIENTO] y 5_cobranza cuenta elegibles por
                        (SALDO>0 & MES_ANT>=8). La multa no entra en la decisión.
                        ⚠ El "corte por multa" se diseñó y NUNCA se ejecutó: los README
                        que lo describen (6b_corte_multas, el trigger en el manifiesto
                        del tenant) son diseño OBSOLETO, no comportamiento vigente.

MES ANTERIOR            no se toca: vive en el bloque P1 (agua) y la re-imputación solo
                        mueve MULTA/ACUERDOS/CONVENIO. Pero ojo con dos acoplamientos:
                        · cerrar reclamos de "ya pagué" con ABONO_REZAGADO/DECLARACION
                          ANTES del cambio de código vuelve a imputar multa-primero
                        · las 10 conclusiones "RESUELTO" de buscar_pago dicen "la boleta
                          vigente ya no le cobra" — hay que releerlas después

FEB A MAYO              no se escribe nada. Esa plata ya está dentro del CARGO sembrado.
```

---

## Los pasos, en el orden definitivo (revisado 2026-08-13 tarde)

```
[x] 0   CORREGIR el reporte      ventana derivada del ledger · tope = plata real
[x] 4   CORREGIR la cascada      P3 CONVENIO · P4 ACUERDOS · P5 MULTA (4 lugares)
[x] 2b  HERRAMIENTA              4b_reclamos/reimputar_cascada.py — dry run por
                                 defecto, backup, idempotente, --revertir MZ LT
                                 (ensayada: escribió, falló la validación, se
                                  restauró el backup · ledger intacto)

  ───── agosto sigue su curso, SIN tocar el ledger ─────
[ ] A1  5_cobranza --force       agosto se re-imputa solo con la cascada nueva
[ ] A2  6_corte                  lista de corte, como siempre
[ ] A3  7_cierre                 agosto cerrado

  ───── recién ahí, y ANTES de generar la planilla de septiembre ─────
[ ] B1  re-correr el reporte     con los saldos post-cierre
[ ] B2  congelar el contrato     el de hoy queda como ensayo
[ ] B3  escribir                 --escribir · todos los asientos con MES = 2026-08
[ ] B4  validar                  capas 1 + 2 + 3
[ ] B5  reporte para la directiva

  ───── septiembre corre solo ─────
        2_planilla toma la re-imputación · 5_cobranza usa la cascada nueva
```

Aparte, y **por un motivo previo a todo esto**: `planilla_2026-08` quedó congelada el 08/08,
antes de la limpieza del ledger del 06/08 — de ahí los 11 predios con deuda negativa.
Regenerarla antes de cerrar agosto sigue siendo lo correcto, pero por ese motivo viejo.

## Los 7 pasos (versión anterior — se conserva por el detalle de cada paso)

```
0  CORREGIR el reporte   ✔ HECHO 2026-08-13 · 4b_reclamos/reporte_reimputacion_cascada.py
                         ① ventana derivada del ledger (hoy 2026-06 .. 2026-08)
                         ② tope = Σ PAGO de clase real, por (predio, concepto que cede)
                         ③ columnas TOPE_* · MES_ORIGEN_* · PEDIDO_LEGACY · RECORTE ·
                            ESTADO · CAUSA_RECORTE · PRE_GENESIS · CREDITO_AJUSTE
                         ④ 3 validaciones de capa 2 + 2 bloques nuevos en la página 1
                         14/14 validaciones OK · sigue siendo simulación

1  CONGELAR              contrato fechado: los movimientos de origen 2026-06 y 2026-07
                         S/2,608.50 · es también el precursor que cuenta la historia

2  ESCRIBIR              por movimiento: 1 fila precursora + 2 asientos REASIGNACION
                         MES = mes del pago · fecha = hoy · AUDIT_REF determinista
                         + la herramienta de reversión por predio (ver LEER_ANTES.md)
                         ← ANTES de correr 5_cobranza

3  VALIDAR               capas 1 + 2 + 3. Si ⑩ falla en un predio, se revierte ESE predio

4  CÓDIGO                ✔ HECHO 2026-08-13 · el orden vivía en 4 lugares acoplados
                         POR POSICIÓN y se reordenaron los 4:
                            _descomponer_saldo  comps P3/P4/P5 → CONVENIO·ACUERDOS·MULTA
                            _AC_P               columnas de arrastre_consolidado
                            _CONCEPTOS_PUEBLO   índice → concepto del ledger
                            _CAMPOS_WATERFALL_REIDENTIFICACION
                         + anchos de columna 7 y 9 · comentario de cabecera
                         + formato_arrastre_consolidado.html y el README del módulo
                         (corte se queda en P2) · 3/3 tests del módulo en verde
                         ⚠ FALTA el detector CASCADA_FUERA_DE_ORDEN, codifica el viejo

4b PLANILLA              regenerar planilla_2026-08 con 2_planilla ← respaldarla antes
                         sin esto, la cascada nueva cobra el convenio dos veces
                         ⚠ SIN correr 3_boletas
                         🔴 desde que el paso 4 está hecho, cualquier corrida de
                            5_cobranza antes de 4b cobra el medidor dos veces

5  AGOSTO                5_cobranza --force → los S/395 de agosto se acomodan solos
                         → 5b_validacion

6  REPORTE               generar el reporte desde el ledger para la directiva:
                         casi todos con deuda de multa, y los reclamos de convenio y de
                         techado/campo que quedan anulados
```

## Archivos de esta sesión (2026-08-13, solo lectura, todos en `4b_reclamos/outputs/`)

| archivo | qué tiene |
|---|---|
| `cuadro_reimputacion_con_tope_2026-08-13.xlsx` | los 98 movimientos con su tope de plata real, `FINAL`, `RECORTE` y `ESTADO` (LIMPIO/RECORTADO/EXCLUIDO) |
| `auditoria_p2_plata_real_2026-08-13.xlsx` | por movimiento: pedido vs `PAGO_REAL` vs `CREDITO_AJUSTE`, con la clase del ajuste |
| `auditoria_pregenesis_2026-08-13.xlsx` | crédito partido en pre-génesis (feb-may) vs ledger (jun-jul), por concepto |
| `comparacion_reporte_vs_ledger_2026-08-13.xlsx` | 2 hojas para los 10 predios: "Reporte_dice" (con columna `EPOCA`) vs "Ledger_dice" (filas crudas con CLASE y MOTIVO) |

## Deuda de documentación que salió de acá (no bloquea)

- **29 `AJUSTE` sin `MOTIVO`.** Ninguno es un misterio: la explicación vive en el
  `AUDIT_REF` (16 `notas_2026-07|...` del 30-31/07, 7 `recon_...` del bug de signo,
  3 `fix_race_condition_yape_20260713_...`, 1 `correccion_lote_F3B_a_F3A`, 2
  estabilizadores de Q-4/Q-5). Quedaron sin llenar por la urgencia de sacar la lista de
  corte. Falta pasarlos a `MOTIVO`.
- Los README que describen **corte por multa** como vigente.
- Los 4 README del ledger mezclan diseño con notas de estado (pendiente desde 2026-07-23).
