# RETOMAR — re-imputación de la cascada a CA1 · paso 0 hecho, nada escrito en el ledger · 2026-08-13

Sesión de diseño (Opus). **No se escribió una sola fila en el ledger.** Se cerró el plan
completo, se auditó contra datos reales y se codificó el paso 0. Sin commitear.

---

## ⚡ PRIMER PASO al retomar

**Leé `docs/decisiones/reimputacion_cascada_ca1.md`** — es la verdad única de este cambio
(decisión, números, mecanismo, validaciones, lo que no se toca). Este RETOMAR solo dice
dónde quedamos y qué sigue.

Después: **cerrar agosto normal.** La cascada ya está corregida, así que al correr
`5_cobranza --force` los pagos de agosto se re-imputan solos; después la lista de corte y el
cierre siguen su curso. **No se escribe nada en el ledger hasta que agosto esté cerrado.**

La re-imputación (`--congelar` → `--escribir`) recién arranca **después del cierre y antes de
generar la planilla de septiembre**. El contrato de hoy es un ensayo: hay que re-congelarlo
con los saldos post-cierre.

---

## Qué es esto en 5 líneas

La cascada cobraba `MULTA → ACUERDOS → CONVENIO`. La directiva decidió lo contrario:
`CONVENIO → ACUERDOS → MULTA`, porque la multa se recupera con faena o se exonera, y el
convenio (medidor) y los acuerdos (techado/campo) solo se saldan con plata. El dominio y 9
READMEs están en el orden nuevo desde el 23/07; **el código vivo nunca se cambió.** Este
trabajo cierra esa brecha en dos mitades: re-acomodar la plata ya imputada, y cambiar el
código para que no vuelva a pasar.

---

## Estado

```
✔ PASO 0 HECHO     4b_reclamos/reporte_reimputacion_cascada.py corregido y re-corrido
                   14/14 validaciones OK · sigue siendo SIMULACIÓN, no escribe nada
✔ PASO 4 HECHO     la cascada ya está en CA1: P3 CONVENIO · P4 ACUERDOS · P5 MULTA
                   4 lugares acoplados por posición reordenados · 3/3 tests en verde

✘ NADA en el ledger
✘ sin commitear
```

```
✔ NO hay trampa armada. Se ensayó la escritura, la validación falló por sepultamiento
  del lado que RECIBE, y se restauró el backup en el acto.
  Verificado: 1563 eventos · 0 asientos de reimputacion_ca1 · precursor en 5 filas.

🟡 LA RE-IMPUTACIÓN VA DESPUÉS DE CERRAR AGOSTO (decisión del usuario, 13/08 tarde)

   agosto necesita su lista de corte, o sea que 5_cobranza VA a correr. Y un asiento
   con MES pasado queda sepultado si el concepto tiene un evento posterior.
   La solución no es un mes especial: es el TIMING.

   AHORA        5_cobranza --force → agosto se re-imputa SOLO (la cascada ya está en CA1)
                6_corte → lista de corte normal → ... → 7_cierre
   POST-CIERRE  re-correr el reporte → --congelar → --escribir → validar
                todos los asientos con MES = 2026-08
                (NO 2026-09: 2_planilla lee el mes anterior)
   SEPTIEMBRE   2_planilla toma la re-imputación · 5_cobranza usa la cascada nueva

   → el paso 4b desaparece: nunca hace falta regenerar planilla_2026-08
   → los 14 movimientos de agosto salen del contrato manual: los hace 5_cobranza
```

```
88 predios se mueven · S/3,003.50, en dos manos

   ASIENTO MANUAL      2026-06  S/1,326.00  ┐ ciclos CERRADOS · pasos 1-3
                       2026-07  S/1,282.50  ┘ S/2,608.50
   LO HACE EL CÓDIGO   2026-08  S/  395.00    ciclo ABIERTO · pasos 4b + 5
```

```
LIMPIO      187 predios      AMPLIADO   10 predios   (plata de agosto que la ventana
RECORTADO     4 predios      EXCLUIDO    3 predios    feb-jul no veía)
```

---

## Los 7 pasos y dónde estamos

```
[x] 0  CORREGIR el reporte    ventana derivada del ledger · tope = plata real por concepto
[ ] 1  CONGELAR               el contrato de jun/jul · S/2,608.50
[ ] 2  ESCRIBIR               1 fila precursora + 2 asientos REASIGNACION por movimiento
                              + la herramienta de reversión por predio
[ ] 3  VALIDAR                capas 1 + 2 + 3
[x] 4  CÓDIGO                 cascada en CA1 · falta el detector CASCADA_FUERA_DE_ORDEN
[ ] 4b PLANILLA               regenerar planilla_2026-08 (respaldarla antes · SIN 3_boletas)
                              🔴 bloquea correr 5_cobranza
[ ] 5  AGOSTO                 5_cobranza --force → 5b_validacion
[ ] 6  REPORTE                desde el ledger, para la directiva
```

El orden **no es negociable**, y está explicado en `LEER_ANTES.md`:

```
① escribir asientos      ANTES de correr 5_cobranza   (sepultamiento)
② cambiar el código      ANTES de cerrar reclamos de "ya pagué"
③ exonerar               DESPUÉS de re-imputar, nunca antes
④ regenerar la planilla  ENTRE el ledger y 5_cobranza  (cobro doble del medidor)
```

---

## Lo que hay que codificar en el paso 2

```
por cada movimiento del contrato congelado:

  shared/reasignaciones_aplicacion.xlsx    1 fila
      MZ · LT · CONCEPTO_ORIGEN · CONCEPTO_DESTINO · MONTO
      MES_ANO = (vacío)   → SOLO REGISTRO, ninguna corrida lo re-aplica
      MOTIVO  = re-imputación CA1 + mes de origen + puntero al contrato
      REF_TRANSACCION = el AUDIT_REF del par de asientos

  shared/seguimiento_pueblo.xlsx           2 asientos
      AJUSTE +X en el concepto que CEDE     (sube su deuda)
      AJUSTE −X en el concepto que RECIBE   (baja su deuda)
      CLASE = REASIGNACION      ← no está en CLASES_SUMAN_CAJA: la caja no se mueve
      MES   = mes del PAGO que se mueve (2026-06 o 2026-07)
      AUDIT_REF = reimputacion_ca1_<MZ>_<LT>_<ORIGEN>_<DESTINO>   determinista

  + la reversión:  --revertir <MZ> <LT>   (procedimiento completo en LEER_ANTES.md)
```

El precursor va **además** del asiento, no en lugar de: la regla del 03/08 es que si mañana
se descarta el ledger y se rehace por backfill, los precursores cuentan la historia solos.
Confirmado el 13/08 que se mantiene el patrón.

---

## El riesgo que se hizo real — `SALDO` es una columna almacenada, y tiene DOS lados

`_saldo_previo()` toma la última fila con `MES <= mes` ordenada por `(MES, TIMESTAMP)` y le
suma el monto. `get_saldo()` devuelve el `SALDO` de esa última fila. **No se recalcula.**

Se midió el sepultamiento del lado que CEDE (0 de 84) y **se olvidó el lado que RECIBE**. El
concepto que recibe tiene su propia cadena de eventos, y casi siempre uno posterior:

```
24 de 84 movimientos · S/665 con el lado que RECIBE sepultado

   B-5  MULTA→CONVENIO 50 · asiento en 2026-06
        su CONVENIO tiene un evento en 2026-08 → el −50 queda sepultado
        la multa subió 50 y el convenio no bajó → la deuda del predio subió 50

   la escritura de prueba dejó la deuda agregada en 3,871.50 → 4,017.50 (+146)
   la validación lo atrapó, se restauró el backup, ledger intacto
```

**Lección para cualquier corrección futura del ledger:** un asiento con `MES` pasado solo
tiene efecto si es la última fila de **su propio** (predio, concepto). Hay que verificarlo
para los DOS conceptos que toca, no solo para el que cede.

## Detalle del mecanismo

`_saldo_previo()` toma la última fila con `MES <= mes` ordenada por `(MES, TIMESTAMP)` y le
suma el monto. `get_saldo()` devuelve el `SALDO` de esa última fila. **No se recalcula.**

```
A-4 MULTA hoy
   MES=2026-06  CARGO 30          SALDO=30
   MES=2026-07  PAGO  30          SALDO= 0   ← get_saldo(2026-08) lee ESTA fila

si escribo un AJUSTE con MES=2026-06:
   MES=2026-06  AJUSTE +30 (hoy)  SALDO=60   ← ordena ANTES de julio
   MES=2026-07  PAGO  30          SALDO= 0   ← su SALDO no se recalcula
   → get_saldo(2026-08) sigue en 0: el asiento queda SEPULTADO, sin efecto y en silencio
```

Medido el 13/08: **0 sepultados**, porque en cada caso el último `PAGO` del concepto que cede
es también el último evento de ese (predio, concepto). Es una propiedad del dato de ese día,
**no una garantía**. La validación ⑩ (resultado == el DESPUES que predijo el contrato) es lo
que lo atrapa si cambió.

---

## Las 4 capas de validación (el detalle está en el doc de decisiones)

```
CAPA 1 · CONSERVACIÓN   ① plata total  ② deuda total  ③ por (CONCEPTO, MES)  ④ por predio
                        ⑤ ningún saldo negativo nuevo
CAPA 2 · SIGNIFICADO    ⑥ ≤ plata real por concepto  ⑦ sin EXONERACION  ⑧ sin CORRECCION_SISTEMA
CAPA 3 · REALIDAD       ⑨ saldo == contrato  ⑩ resultado == predicho  ⑪ AUDIT_REF idempotente
CAPA 4 · CONSECUENCIA   ⑫ cuántos pasan a deber SOLO multa (mide, no valida)
```

La capa 2 es la que faltaba. Prueba de por qué: **J-6, Capilla del Pueblo** — mover S/50 de
su multa (crédito de `EXONERACION`, plata real S/0) al convenio conserva el total, conserva
por concepto, conserva por predio y no deja negativos. Las 4 validaciones de conservación en
verde, y el resultado es pagarle el medidor con una exoneración. La conservación también pasa
si el movimiento se aplica dos veces.

---

## Decisiones cerradas en esta sesión (no re-litigar)

```
R2  exoneración y corrección de bug NO son plata; plata es COBRANZA · ABONO_REZAGADO ·
    DECLARACION (la declaración se corrige con otro evento si ese pago nunca entró)
R5  INSTALACION y REACTIVACION no reciben plata en el convenio — intencional
R6  el asiento va con MES = mes del pago (ciclo de origen), fecha = hoy
①   boletas de agosto NO se regeneran (salieron 31/07, cobradas 01/08, DATA_boletas con
    ~25 correcciones a mano) → el vecino ve el cambio en SEPTIEMBRE
②   la lista de corte NO se afecta ni se congela: el único corte vigente es por no pago de
    AGUA en 2 meses. El "corte por multa" se diseñó y nunca se ejecutó → los README que lo
    describen son diseño OBSOLETO
③   J-1 y J-6 (centros del pueblo): el convenio SE COBRA (S/175 entre los dos). Multa y
    acuerdos quedan condonados — tienen MOTIVO de la directiva en el ledger
④   el precursor se mantiene, con su puntero bien dirigido
⑤   la reversión por predio se construye desde el día 1 y está documentada en LEER_ANTES
⑥   N-6 DEBE (multa 20 + acuerdos 21). Aparecer en una lista de auditoría no es motivo
    para exonerar
```

Dos principios que salieron de acá y valen para todo el ledger:

```
el MOTIVO escrito le gana a la memoria de cualquiera — una regla nueva aplica hacia
adelante, no revierte un asiento que ya tiene motivo
    → los 29 AJUSTE sin MOTIVO son los únicos indefendibles: llenarlos dejó de ser cosmética

aparecer en una lista de auditoría NO es motivo para exonerar — el respaldo tiene que venir
de una fuente independiente (para multa de reunión/faena: las hojas de asistencia)
```

---

## Lo que viene después de los 7 pasos

**La ola de exoneraciones.** Con casi todo el pueblo debiendo solo multa, van a reclamar y se
va a exonerar a varios — es el circuito que la directiva quiso. Hoy hay 7 `EXONERACION` en el
ledger; van a ser decenas. Exonerar deja de ser un ajuste por caso y pasa a ser el flujo
principal del ciclo: **necesita su propia herramienta y su propio registro**, con la asistencia
a faena/reunión como respaldo. Sin diseñar.

**Para la reunión con la directiva:** la deuda de multa sube de 3,640 a 6,063.50 al re-imputar
y **baja de nuevo** a medida que se exonera. Es un pico, no un estado final — decirlo así, o el
salto se lee como morosidad nueva cuando es multa que siempre estuvo impaga y estaba tapada
por el orden viejo.

---

## Archivos de esta sesión

```
MODIFICADO
  4b_reclamos/reporte_reimputacion_cascada.py     paso 0 (ventana + tope + columnas +
                                                  3 validaciones + 2 bloques en la pág. 1)
  5_cobranza/main.py                              paso 4 (los 4 lugares del orden)
  5_cobranza/docs/formato_arrastre_consolidado.html   columnas P3/P5 intercambiadas
  5_cobranza/README.md                            la cascada documentada en CA1
  5_cobranza/tests/test_cobranza.py               expectativa stale de arrastre_devolucion
                                                  (1 → 2: incluye el no identificado;
                                                   el FAIL era previo, no del cambio)
  LEER_ANTES.md                                   sección nueva al tope
NUEVO
  docs/decisiones/reimputacion_cascada_ca1.md     ← la verdad única
  docs/retomar/RETOMAR_reimputacion_cascada_ca1_2026-08-13.md   (este archivo)
REGENERADO
  4b_reclamos/outputs/reporte_reimputacion_cascada_2026-07.pdf + .xlsx
  4b_reclamos/outputs/reporte_reimputacion_cascada_2026-07_reclamos.pdf + .xlsx
AUDITORÍAS (solo lectura, respaldan los números)
  4b_reclamos/outputs/cuadro_reimputacion_con_tope_2026-08-13.xlsx
  4b_reclamos/outputs/auditoria_p2_plata_real_2026-08-13.xlsx
  4b_reclamos/outputs/auditoria_pregenesis_2026-08-13.xlsx
  4b_reclamos/outputs/comparacion_reporte_vs_ledger_2026-08-13.xlsx
```

## Deuda de documentación que salió de acá (no bloquea)

```
· los 29 AJUSTE sin MOTIVO (la explicación está en el AUDIT_REF, hay que pasarla a MOTIVO)
· la CLASE de S-5 ACUERDOS −40 dice EXONERACION y era plata real
· los README que describen "corte por multa" como vigente
· los 4 README del ledger mezclan diseño con notas de estado (pendiente desde 23/07)
```

## Otra cosa de esta sesión, sin relación con la cascada

Se borraron 10 carpetas copia de `jass_system` (~40.8 GB). Quedan 3: el repo activo,
`Julio\jass_system - Julio` y `Junio\jass_system - junio`. Lo irrecuperable se rescató a
`PycharmProjects\_archivo_rescate_2026-08-13\` (con `LEEME.md`). Método y detalle en
`docs/diario/2026-08-13_limpieza_de_copias.html`.
