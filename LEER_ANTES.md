# LEER ANTES — estado_cuenta completo comienza en agosto, sin backfill (21/08/2026)

```text
arrastre cerrado julio + lecturas/planilla agosto + audit de corte + pagos
                                  |
                                  v
                  snapshot completo de 5_cobranza
                                  |
                                  v
                     commit único de 7_cierre
```

- Agosto abre `AGUA` y `CORTE_RECONEXION` desde el arrastre cerrado de julio; no reconstruye meses anteriores.
- `AGUA` y `MANTENIMIENTO` nuevos salen de lecturas + `2_planilla`; los cargos nuevos de corte salen del audit de `6_corte`.
- Desde septiembre, `2_planilla` lee agua, mantenimiento y corte del ledger comprometido, no de `arrastre_consolidado`.
- El nombre físico `seguimiento_pueblo.xlsx` se conserva durante la transición, pero su taxonomía ya es la de `estado_cuenta`.
- `7_cierre` sigue siendo el único commit oficial. No ejecutar el cierre real sin completar la lista de corte vigente.

# LEER ANTES — la cascada se reordena a CA1 y se re-imputan S/2,545.50 de pagos viejos (13/08/2026) — PLAN CERRADO, SIN EJECUTAR

**Si venís a tocar `5_cobranza`, `seguimiento_pueblo.xlsx` o cualquier reporte de deuda de
multa/acuerdos/convenio: leé esto primero.** Hay un cambio grande decidido y todavía no
aplicado, y el orden en que se hagan las cosas importa.

Decisión completa, números y validaciones: **`docs/decisiones/reimputacion_cascada_ca1.md`**.
Acá solo el estado y lo que no se puede improvisar.

## Estado — qué está hecho y qué no

```
✔ decidido       la cascada del tramo pueblo pasa a CONVENIO → ACUERDOS → MULTA
                 (el dinero cubre primero lo que SOLO el dinero salda; la multa se
                 recupera con faena o se exonera)
✔ PASO 0 HECHO   4b_reclamos/reporte_reimputacion_cascada.py corregido y re-corrido:
                 ventana derivada del ledger + tope = plata real por concepto.
                 14/14 validaciones OK. Sigue siendo simulación, no escribe nada.
✔ medido         88 predios · S/3,003.50, en dos manos:
                    S/2,608.50  origen jun/jul → ASIENTO MANUAL (ciclos cerrados)
                    S/  395.00  origen agosto  → lo hace el código (pasos 4b + 5)
✔ auditado       7 predios afectados, cada uno con su causa escrita en el reporte
                 (3 EXCLUIDO · 4 RECORTADO)

✔ PASO 4 HECHO   la cascada ya está en CA1: 5_cobranza/main.py::_descomponer_saldo
                 tiene P3 CONVENIO · P4 ACUERDOS · P5 MULTA (corte sigue en P2).
                 Se reordenaron los 4 lugares acoplados por posición y los 3 tests
                 del módulo pasan.

✘ NADA escrito en el ledger — verificado: 1563 eventos, 0 asientos de reimputacion_ca1

✔ NO hay trampa armada. Se ensayó la escritura, la validación falló por sepultamiento
  y se restauró el backup en el acto. El ledger y la planilla vuelven a estar de acuerdo,
  así que **agosto puede seguir su curso normal**: 5_cobranza, lista de corte y cierre.

🟡 LA RE-IMPUTACIÓN VA DESPUÉS DE CERRAR AGOSTO — no antes
   Escribir un asiento con MES pasado no sirve si el concepto tiene un evento
   posterior: queda SEPULTADO (SALDO es un total corrido guardado, no se recalcula).
   Al escribir después del cierre, el asiento es la última fila de los dos conceptos
   y no hay nada que lo entierre.

   AHORA        5_cobranza --force → agosto se re-imputa solo (la cascada ya está en CA1)
                6_corte → lista de corte normal → ... → 7_cierre
   POST-CIERRE  re-correr el reporte → congelar → --escribir → validar
                todos los asientos con MES = 2026-08 (NO 2026-09: 2_planilla lee el
                mes anterior, y con 09 la planilla de septiembre no los vería)
   SEPTIEMBRE   2_planilla toma la re-imputación · 5_cobranza usa la cascada nueva

   → el paso 4b (regenerar planilla_2026-08) YA NO HACE FALTA por este cambio
```

## El orden no es negociable

```
① re-imputar (asientos en el ledger)      ANTES de correr 5_cobranza otra vez
   por qué: SALDO es una columna ALMACENADA. Si una corrida escribe un evento de
   agosto para un predio, el asiento con MES=2026-06 ordena antes y queda
   SEPULTADO — sin efecto y en silencio. Hoy hay 0 sepultados; no es garantía.

② cambiar el código                        ANTES de cerrar reclamos de "ya pagué"
   por qué: si se escribe un ABONO_REZAGADO/DECLARACION con el orden viejo vivo,
   esa plata nueva se vuelve a imputar multa-primero y se recrea el problema.

③ exonerar                                 DESPUÉS de re-imputar, nunca antes
   por qué: la exoneración deja la multa en 0; el AJUSTE de la re-imputación
   devuelve deuda de multa. Al revés, resucita una deuda ya perdonada y el predio
   queda excluido del movimiento con su reclamo de convenio sin resolver.
   ⚠ Esta lista VA A CRECER: con casi todo el pueblo debiendo solo multa, van a
   reclamar y se va a exonerar a varios. Es el circuito buscado.

④ los 8 movimientos de origen 2026-08      NO se escriben a mano, nunca
   5_cobranza recalcula el mes entero y sabe revertir, pero su SET_TIENE cuenta
   solo eventos con source="5_cobranza": no vería el asiento manual y movería la
   plata DOS VECES.
```

## Lo que NO se toca (y por qué, para no re-discutirlo)

```
boletas de agosto     NO se regeneran. Salieron el 31/07 y se cobraron el 01/08, y
                      DATA_boletas.xlsx tiene ~25 correcciones a mano de ese día
                      (backups en 3_boletas/inputs/backups/). Regenerarlas las pisa
                      → Regla 9 de CLAUDE.md. El vecino ve el cambio en SEPTIEMBRE.
                      El paso final NO corre 3_boletas.

⚠ planilla_2026-08    SÍ se regenera (paso 4b), y es obligatorio. 2_planilla lee
                      MULTA/ACUERDOS/CONVENIO del LEDGER (main.py:144
                      _join_saldo_pueblo → get_saldos_bulk), pero 5_cobranza lee la
                      deuda por concepto DE LA PLANILLA (main.py:833-835). La planilla
                      de agosto es una foto del ledger al cierre de julio: si la
                      cascada nueva corre sobre esa foto, vuelve a cobrar el convenio
                      que la re-imputación ya cubrió → COBRO DOBLE del medidor.
                      2_planilla no tiene backup ni preservación → respaldar a mano antes.
                      Después de 4b la planilla y las boletas de agosto ya entregadas
                      dicen cosas distintas: es deliberado, no lo "arregles".

lista de corte        NO se afecta ni se congela. El único corte vigente es por no pago
                      de AGUA en 2 meses (6_corte usa CONCEPTOS_SALDO=[AGUA,MANTENIMIENTO];
                      5_cobranza cuenta elegibles por SALDO>0 & MES_ANT>=8). La multa no
                      entra en la decisión de corte.
                      ⚠ El "corte por multa" se diseñó y NUNCA se ejecutó. Los README que
                      lo describen (6b_corte_multas, el trigger del manifiesto) son diseño
                      OBSOLETO, no comportamiento vigente. No los tomes como estado real.

feb a mayo            NO se escribe nada. El ledger nace en junio y la deuda se sembró YA
                      NETA de lo pagado antes: mover un pago de feb-may lo gasta dos veces.

mes anterior          NO se toca. Vive en el bloque P1 (agua); la re-imputación solo mueve
                      MULTA/ACUERDOS/CONVENIO.
```

## Cómo revertir un predio si un vecino aparece con un recibo

El ledger es append-only: revertir = escribir el par inverso, nunca borrar filas.

```
herramienta (se crea en el paso 2, junto con la escritura):
    py 4b_reclamos/reimputar_cascada.py --revertir <MZ> <LT>

qué hace, y qué hay que hacer a mano si la herramienta todavía no existe:
    ① busca los asientos con AUDIT_REF = reimputacion_ca1_<MZ>_<LT>_<ORIGEN>_<DESTINO>
    ② escribe el par INVERSO con el mismo MES y CLASE=REASIGNACION
       AJUSTE −X en el que había cedido · AJUSTE +X en el que había recibido
       AUDIT_REF = reimputacion_ca1_REVERTIDO_<MZ>_<LT>_<ORIGEN>_<DESTINO>
       MOTIVO = por qué se revierte + el recibo/evidencia que apareció
    ③ apaga la fila del precursor en shared/reasignaciones_aplicacion.xlsx
       (columna ESTADO = REVERTIDO, no se borra la fila)
    ④ re-corre la validación: el predio tiene que volver al saldo de antes

NO hace falta revertir el lote entero por un predio: cada movimiento es independiente
y su AUDIT_REF es determinista.
```

## Los 10 predios cuyo reclamo queda abierto

```
J-1 Comedor Popular · J-6 Capilla del Pueblo    el convenio SE COBRA (cerrado 13/08)
    nunca pagaron nada: cero eventos PAGO. El medidor es un bien que recibieron.
        J-1 convenio S/ 75   ·   J-6 convenio S/100      → S/175 a cobrar
    no hay nada que re-imputar en ellos
    sus ACUERDOS quedan CONDONADOS (S/75 + S/75) — cerrado 13/08. La regla hacia
    adelante es "techado y campo se pagan, solo la multa se exonera", pero esas dos
    exoneraciones tienen su MOTIVO escrito y viene de la directiva, así que se mantienen.
    ⚠ S-5 ACUERDOS −40 figura como EXONERACION y NO lo es: era plata real (el yape
      que Wagner retuvo, recuperado en efectivo). NO revertir — la CLASE está mal
      puesta, es deuda de documentación.

## El MOTIVO escrito le gana a la memoria de cualquiera

Regla fijada el 13/08/2026, y vale para todo el ledger, no solo para este evento:

```
si un asiento tiene MOTIVO escrito, ese motivo MANDA — por encima de lo que
cualquiera recuerde hoy, incluido el usuario

  caso que la fijó: la regla es "solo la multa se exonera", pero el ledger tenía
  2 exoneraciones de ACUERDOS con motivo de la directiva. Ganó el ledger:
  "obviamente en su momento la directiva habrá dicho condona acuerdos y a mí se me olvidó"
```

Consecuencia directa: **los 29 `AJUSTE` sin `MOTIVO` son los únicos asientos que no
podrían defenderse solos en una discusión así.** Su explicación está en el `AUDIT_REF`
(`notas_2026-07|...`, `recon_...`, `fix_race_condition_...`), pero mientras no esté en
`MOTIVO`, cualquiera puede revertirlos de buena fe. Llenarlos dejó de ser cosmética.

P-12 Judith Venturo · N-6 David Juárez          NO se exoneran (cerrado 13/08)
    están en la lista de excluidos por el bug de la ventana del reporte, NO por una
    decisión de la directiva. Son vecinos, no bienes del pueblo, y ninguno tiene
    reclamo registrado en los 5 archivos de reclamos.
        N-6   multa 20 DEBE · acuerdos 21 DEBE (ya pagó 29 reales en julio)
        P-12  convenio 50 DEBE · multa 0 (no tiene ni un evento de multa)
    ⚠ aparecer en una lista de auditoría NO es motivo para exonerar. Si algún día se
      exonera la multa de N-6, el respaldo son las hojas de ASISTENCIA (sus S/20 son
      tarifa de reunión) — no la lista.

K-9 · A-4 · B-12A · C1-14 · D1-6 · L-4          se mueve solo lo que tienen de plata real
```

## Cómo se cierra este evento

Cuando los 6 pasos de `docs/decisiones/reimputacion_cascada_ca1.md` estén ejecutados y
validados, y el reporte para la directiva esté entregado. Mientras alguna de estas filas
siga sin ejecutar, esta sección se queda:

```
[x] 0   corregir el reporte (ventana derivada del ledger · tope = plata real)
[x] 4   _descomponer_saldo al orden CA1 · falta el detector CASCADA_FUERA_DE_ORDEN
[x] 2b  herramienta 4b_reclamos/reimputar_cascada.py (dry run · backup · --revertir)

  ── agosto sigue su curso, sin tocar el ledger ──
[ ] A1  5_cobranza --force   (agosto se re-imputa solo con la cascada nueva)
[ ] A2  6_corte              lista de corte
[ ] A3  7_cierre             agosto cerrado

  ── recién ahí, y ANTES de generar la planilla de septiembre ──
[ ] B1  re-correr el reporte con los saldos post-cierre
[ ] B2  --congelar           (el contrato de hoy es un ensayo, no se usa)
[ ] B3  --escribir           todos los asientos con MES = 2026-08
[ ] B4  validar capas 1+2+3
[ ] B5  reporte desde el ledger para la directiva
```

---

# LEER ANTES — correcciones_lote.xlsx resucitó una regla al revés (Roberto C1-9) + 2 lotes mal escritos + 1 pendiente (10/08/2026)

**Si `correcciones_lote.xlsx` tiene una fila `C1-9 → C1-9` (misma MZ/LT origen y
destino): es a propósito, no la borres.** Es un candado contra el auto-sanador
(`_recuperar_correcciones_trazabilidad`), que si no está esa fila vuelve a
resucitar la regla mala `C1-9 → C1-17`.

## Qué pasó — la plata de Roberto (C1-9) se perdía sola

```
09/08  verificar_lotes.py corrige el typo en la FUENTE: mesa_2 ya escribe C1-9
       (antes escribía C1-17 — Roberto Macarlopu, ver project_override_c1-9_roberto)
              │
10/08 07:45   `5_cobranza --force` corre. `_recuperar_correcciones_trazabilidad`
              escanea trazabilidad_cobranza.xlsx buscando correcciones viejas para
              "auto-sanar" correcciones_lote.xlsx si se perdió alguna (ver docstring
              en main.py). Encuentra una fila de CICLO 17 — de ANTES del fix de
              arriba, cuando la creencia era la contraria — que dice
              "C1-9 se corrigió a C1-17" y la reinserta como regla activa.
              │
              ▼  el pago de Yerald (S/8, mesa_2, Roberto) entra correctamente
                 como C1-9 → la regla lo desvía a C1-17 → C1-17 no existe
                 → cae en discrepancias_cobranza.xlsx, Roberto queda sin ese pago
```

El mecanismo de auto-sanado asume que las correcciones son monótonas (nunca se
revierten). Acá el sentido correcto se invirtió (antes se creía C1-17, ahora se
sabe que es C1-9) y el auto-sanador no tiene forma de saber cuál versión es la
vigente — solo mira "¿ya existe una regla para este origen?".

## El fix — 2026-08-10T08:52-09:27, 3 corridas de `5_cobranza --force`

```
intento 1   correcciones_lote: C1-9→C1-17 se vuelve C1-17→C1-9 (flip simple)
            resultado: el auto-sanador la vuelve a resucitar en la MISMA corrida
            (existentes ya no tiene la clave (C1,9) → la re-agrega desde
            trazabilidad) → bug sigue igual, plata de Roberto sigue sin aplicar

intento 2   se agrega ADEMÁS una fila identidad  C1-9 → C1-9
            (motivo: bloquea el auto-sanador — con (C1,9) ya presente en
            existentes, deja de re-agregar la regla vieja)
            resultado: "corrección aplicada: C1-9 → C1-9" (no-op) · Roberto
            recibe su MONTO_EFECTIVO=8 · sin AJUSTE en seguimiento_pueblo
```

`correcciones_lote.xlsx` queda así para C1: `C1-17 → C1-9` (por si el typo
viejo reaparece en algún crudo) + `C1-9 → C1-9` (candado). Verificado:
`planilla_cobrado.xlsx` → C1-9 ROBERTO MACARLOPU FLORES, MONTO_EFECTIVO=8.

## De paso, 2 lotes más mal escritos por el cobrador (mismo ciclo, verificados)

```
G-36  (S/39, mesa_4, Wagner Trujillo)  →  el lote real es C-36 (MARIA NORABUENA ESPADA)
                                            corregido en correcciones_lote.xlsx, aplicado
S-16  (S/22, mesa_1, Wilder Trujillo)  →  candidato: S-6 (VICTORIA COSME CERNA) — NO
                                            confirmado, se deja SIN corregir a propósito
```

## S-16 — pendiente de reidentificar, ahora visible donde corresponde

S-16 no es un usuario real (no existe ese lote), así que nunca podía aparecer
en `arrastre_devolucion_2026-08.xlsx` (esa lista sale de `SALDO<0` de un
usuario real). Quedaba solo en `discrepancias_cobranza.xlsx`, que nadie revisa
buscando plata para devolver/reidentificar. **Cambio de código:**
`_exportar_arrastre_devolucion()` (`5_cobranza/main.py`) ahora recibe también
`disc_yape`/`disc_efec` y agrega una sección "no identificado" (fondo rojo,
mismo color que discrepancias) debajo de los excesos reales. Verificado:
`arrastre_devolucion_2026-08.xlsx → 13 usuarios con SALDO<0 · 1 no
identificado(s)`, fila S-16 visible con MOTIVO="predio no encontrado en
planilla". Mañana: confirmar si S-16 es de verdad S-6 y corregir en
`correcciones_lote.xlsx` igual que G-36→C-36.

## L-9 → J-9: la misma regla resucitada tenía OTRO problema — no tenía confirmación humana

`correcciones_lote.xlsx` también traía, desde el mismo ciclo 17, una regla
`L-9 → J-9` (Denis Riv\*, S/500, 03/08/2026 18:43:26, mensaje "L-9"). A
diferencia de C1-9 (que sí tiene respaldo — memoria COFOPRI + override
`ELIMINAR` de C1-17 en el padrón, 28/07), esta regla solo decía "Recuperado
desde trazabilidad_cobranza.xlsx" — **sin ningún rastro de quién confirmó que
"L-9" era un error del pagante y no lo que de verdad quiso decir.**

```
CRITERIO aplicado para distinguir una corrección real de una adivinanza:
   ¿hay una persona con NOMBRE que lo confirmó (WhatsApp, cuaderno, reclamo)?
      SÍ → confiar (ej. D1-1→D1-3, Janet Villanueva Alegre, WhatsApp 14/07)
      NO → revertir, la plata se queda donde el mensaje real la puso

L-9 → J-9 no pasó la prueba. Revertido: fila borrada de correcciones_lote.xlsx.
La plata de Denis Riv* (S/500) vuelve a contarse como pago de L-9 (Marianela
Elizabet Rivera Vitate, que debía MULTA 30 + ACUERDOS 50 = 80 — el pago la
cubre entera y deja exceso).
```

**Verificado antes de revertir que no dispara el bug de compensación:** ni L-9
ni J-9 tienen MULTA/ACUERDOS/CONVENIO activos que dependan de CUÁNTO recibió
cada uno más allá de lo que ya tenían — el monto total de cada predio no
cambiaba entre "antes" y "después" del revert en ninguno de los dos casos
salvo el agua/exceso puro, así que `_reconciliar_pagos_pueblo` no generó
ningún AJUSTE nuevo al aplicar esto.

## Lección aparte — `forzar_mzlt.xlsx` no es lo mismo que `pendientes.xlsx`

Al revisar los otros 3 registros de
`4_pagos/yape/motor_matching/correcciones/forzar_mzlt.xlsx` (Giovanna San\*,
PLIN Juan Carlos Graciano, Patricia Tar\*) se encontró que **2 de los 3 usaban
la herramienta equivocada** — no es un problema de plata (los 3 destinos
verificados coinciden exacto con el titular real del padrón), es un problema
de proceso:

```
forzar_mzlt.xlsx    → SOLO para "error de digitación del propio pagante"
                       (el pagante escribió algo, pero él mismo se equivocó)
pendientes.xlsx      → para AMBIGÜEDAD DE MATCHING (el sistema no tiene
(Ambiguos/             info suficiente — sin mensaje, mensaje sin espacios
Maestro_inexacto)      que no parsea, etc.) — un humano llena MZ/LOTE + OK=SI

Giovanna San*/K-17      mensaje "mzklt17..." sin espacios → AMBIGÜEDAD
                        (el pagante no se equivocó, el parser no pudo leerlo)
                        → debió ir a pendientes.xlsx, no a forzar_mzlt
PLIN Juan Carlos/H1-13  SIN mensaje (PLIN no manda texto) → AMBIGÜEDAD
Patricia Tar*/H1-15     SIN mensaje → AMBIGÜEDAD
                        → los 2 debieron ir a pendientes.xlsx
```

Ya corrieron en julio (ciclo cerrado), no se movió nada — queda como lección
para no repetir: ante un caso ambiguo nuevo, usar `pendientes.xlsx`, reservar
`forzar_mzlt.xlsx` solo para cuando el pagante mismo se equivocó al escribir.

## Hallazgo colateral — E-14B (Juan Saavedra) duplicado en planilla, genera ruido en cada corrida

`shared/planilla_mes/planilla_2026-08.xlsx` tiene **2 filas** para E-14B (Juan
Saavedra Saavedra). Cada re-corrida de `5_cobranza --force` procesa las dos
filas contra el MISMO evento de `seguimiento_pueblo` (ACUERDOS) y deja un par
`AJUSTE +20 / PAGO +20` (saldo neto no cambia, 75→55, pero quedan 2 eventos de
ledger que no deberían existir). No es efecto de las correcciones de arriba —
ya se verificó que en la corrida anterior a este evento (ciclo 18, antes de
tocar nada) también estaba duplicado en planilla, solo que no había generado
todavía el AJUSTE. Pendiente: borrar la fila duplicada de E-14B en
`planilla_2026-08.xlsx` (queda una sola) — hasta entonces, cada `--force`
sigue agregando ese par de eventos.

## Cómo se cierra este evento

Cuando: (1) S-16 se confirme o descarte contra S-6, (2) la fila duplicada de
E-14B se borre de `planilla_2026-08.xlsx` y una corrida limpia no vuelva a
generar el par AJUSTE/PAGO, (3) `5b_validacion` corra sin la alerta de C1/G-36
(las otras alertas — L→J falso positivo, Nivel 1a −50 — son de sesiones
anteriores, no de este evento).

---

# LEER ANTES — lote de 8 predios de julio con abono rezagado que nunca se aplicó (09/08/2026) — CERRADO

**Si buscás por qué `shared/abonos_rezagados.xlsx` tiene una fila con
`MES_ANO_APLICA=2026-07` para un predio y el saldo de julio en
`seguimiento_pueblo.xlsx` no la refleja: es porque literalmente no pudo aplicarse.**
`abonos_rezagados.xlsx` no existía en esta forma hasta el 03-06/08/2026 (ver
`git log` del archivo) — las 3 corridas de `5_cobranza` de julio (08/07, 20/07,
31/07) ya habían pasado antes de que el archivo tuviera esas filas.
`_cargar_abonos_rezagados()` solo se consulta cuando `5_cobranza` corre para ese
`mes_ano`, y julio no se volvió a correr (está prohibido, ver sección de abajo).

## Los 3 casos verificados y corregidos — D-16, D1-6, L-4

Los tres tenían, además del abono rezagado sin aplicar, un `AJUSTE` con el bug de
signo (sección de abajo) sobre el mismo predio — coincidencia real: los tres son
del mismo lote de cobradores que retuvieron yapes (Wagner Trujillo), investigado y
sembrado en `abonos_rezagados.xlsx` a principios de agosto.

```
Método: reconstruir el CARGO real de julio (planilla_2026-07.xlsx del repo
"jass_system - Julio", que es el correcto — ver LEER_ANTES anterior sobre por
qué el planilla_cobrado archivado en ESTE repo está desincronizado) y correr a
mano la cascada P1→P5 (agua → corte → multa → acuerdos → convenio) con
total_pagado = abono rezagado + efectivo de mesa (los dos son plata real, sin
CONCEPTO en la mesa → cuentan como agua/libre, igual que hace 5_cobranza).

D-16   abono 85 + efectivo 34 = 119   → cubre agua(19)+acuerdos(50)+convenio(50)
       exacto. ACUERDOS y CONVENIO quedan en 0 (antes: 25 y 25 con el bug).
D1-6   abono 33 + efectivo 33 = 66    → cubre agua(16)+multa(30) completos, dedja
       20 de los 50 de acuerdos. MULTA corregida a 0 (antes: 1, bug de signo).
       ACUERDOS ya estaba en el valor correcto (30), no se tocó.
L-4    abono 58 + efectivo 41 = 99    → cubre agua(24)+multa(20) completos, deja
       55 de los 75 de acuerdos aplicados. ACUERDOS corregido a 20 (antes: 45).
       MULTA y CONVENIO ya estaban correctos (0 y 25), no se tocaron.
```

Backup antes de tocar: `shared/backups_ledger/seguimiento_pueblo_pre_cascada_completa_D16_D1_6_L4_20260809_150214.xlsx`.
Cada AJUSTE quedó con `CLASE=CORRECCION_SISTEMA` y motivo completo en el ledger
(visible en la hoja `Ajustes` de `vista_seguimiento_pueblo.xlsx`).

## Resto del lote de Wagner — verificado y corregido (09/08/2026)

`abonos_rezagados.xlsx` tiene 8 filas de este mismo lote (agua, retenido por
cobrador, `MES_ANO_APLICA=2026-07`) — 7 de Wagner Trujillo, 1 de Yanet Villanueva.

```
T-12  Samuel Samaritano Romero          155   → MULTA y ACUERDOS quedan pagados
                                                 completo (125 de deuda, sobran 30
                                                 sin aplicar — exceso de agua, no
                                                 vive en este ledger)
I-9   Julia Cardenas Alvarado            86   → MULTA pagada completo, ACUERDOS
                                                 pasa de 75 a 58
F-9   Rosa Lucia Coronado Luna           52   → CONVENIO pagado completo (antes 25,
                                                 solo el efectivo se habia acreditado)
F1-4  Antonio Gutierrez Pachamango      101   → SIN CAMBIOS — agua+corte (130) ya
                                                 consumen todo el abono, nunca llega
                                                 a MULTA (Yanet Villanueva, no Wagner)
```

Backup: `shared/backups_ledger/seguimiento_pueblo_pre_wagner_T12_I9_F9_20260809_*.xlsx`.
CLASE=`ABONO_REZAGADO` (plata real de un ciclo viejo que la caja recién ve), no
`CORRECCION_SISTEMA` — a diferencia de D-16/D1-6/L-4, acá no había un bug de signo
de por medio, era simplemente plata nunca acreditada.

## S-5 (Valerio Porfilio Javier Santiago), Wagner, S/71 — cerrado (09/08/2026)

El "30 declarado" en ACUERDOS (31/07, `CLASE=DECLARACION_SECRETARIA`) **no es la
misma plata que el abono** — se investigó el reclamo original
(`notas_2026-07.xlsx`, GRUPO 2): *"Verificar su campo, techado y multas — reclamo
porque salió en corte, dice que está al día y ya pagó."* Resolución (28/07): la
multa de julio (30) ya estaba pagada, y además se reconoció un **crédito
histórico** — pagó multa en ene/feb 2026 (S/15+S/15=30, de antes de que existiera
el ledger, nunca reconocido) — aplicado contra ACUERDOS. Es plata real, distinta,
sin relación con el abono rezagado.

**Segunda parte del reclamo, la del corte, se resolvió hoy:** confirmado que el
corte de S-5 (`registro_cortes.xlsx`, CORTADO desde 2026-06) fue un error — el
pago sí existía, Wagner lo tenía retenido (el mismo abono de S/71). Se exoneró:

```
registro_cortes.xlsx    ESTADO CORTADO → EXONERADO, motivo escrito
ajustes_cargo.xlsx      CORTE_RECONEXION 40, MES_ANO_APLICA=2026-08 (julio cerrado,
                         mismo patron que F1-4/V-14), CLASE=EXONERACION
seguimiento_pueblo      ACUERDOS AJUSTE −40 (con el corte exonerado, el abono(71)+
                         efectivo(74)=145 ya alcanza para cubrir agua+multa+acuerdos
                         completo: 70+30+45=145 exacto) → SALDO 0
```

S-5 queda "al día" en pueblo (MULTA=0, ACUERDOS=0, CONVENIO=0). Backups:
`seguimiento_pueblo_pre_S5_corte_20260809_*.xlsx` y `registro_cortes_pre_S5_20260809_*.xlsx`.

## Las otras 22 filas de `abonos_rezagados.xlsx` (declaraciones de la secretaria)

Estas NO pasan por `_cargar_abonos_rezagados`+cascada — se aplicaron directo como
`registrar_pago` con `CLASE=DECLARACION_SECRETARIA` (ver sección "los pagos que
declaró la secretaria" más abajo). Ese mecanismo es independiente de la fecha del
archivo — no tienen el mismo riesgo que las 8 de arriba. No se re-verificaron acá.

## Fix de código — `4b_reclamos/reporte_historico.py` leía agua/corte de una fuente stale

Al revisar S-5 en el reporte, la fila de julio mostraba Corte=15 y Mes ant.
vacío en vez de 40 y 46. Causa: `_datos_ciclo()` sacaba esas 4 columnas
(consumo/mant/mes_ant/corte) de `5_cobranza/outputs/planilla_cobrado.xlsx` —
que quedó congelado de una corrida vieja de julio, de ANTES de una corrección
de `2_planilla` (18/07/2026, mismo día que la nota de D-16 en
`ajustes_cargo.xlsx`) que nunca se volvió a propagar porque `5_cobranza` no se
re-corre sobre un ciclo cerrado. **Esto no es un problema de repos separados
(Julio vs agosto) — el mismo archivo archivado del propio repo "jass_system -
Julio" tiene el mismo dato viejo.** `shared/planilla_mes/planilla_2026-07.xlsx`
(2_planilla) sí tiene el CARGO correcto, verificado exacto contra
`DATA_boletas.xlsx`.

Fix aplicado: `_datos_ciclo()` ahora lee mes_actual/mantenimiento/mes_anterior/
corte_reconexion de `shared/planilla_mes/planilla_<mes>.xlsx` en vez de
`planilla_cobrado.xlsx`, y suma `abonos_rezagados.xlsx` (filtrado por predio +
`MES_ANO_APLICA`) al `total_pagado` de la cascada agua→corte, para que el
`TOTAL PAGADO` de la tabla ya no ignore esa plata. Con esto arreglado,
`planilla_2026-06.xlsx` resultó ser un archivo de 15 bytes (inválido) — se
agregó manejo defensivo (`try/except`) para que un mes sin copia válida caiga
al comportamiento anterior en vez de romper.

**Segundo fix, mismo archivo:** `_datos_ciclo()` sumaba MES_ACTUAL+MANTENIMIENTO
en un solo número y lo mandaba todo a la columna "Consumo", dejando "Mant."
siempre en 0 para junio/julio en adelante (los meses históricos anteriores sí
los mostraban separados, por venir de otra fuente). Separado: ahora "Mant."
se llena igual que "Mes ant."/"Corte" (su propio paso de la cascada agua→mant→
mes_ant→corte), verificado en S-5 (Consumo 21 + Mant 3, antes aparecía 24+0).

**Tercer fix, mismo archivo — el que de verdad importaba:** MULTA/ACUERDOS/
CONVENIO del historial solo sumaban la columna `PAGO` del ledger, ignorando lo
saldado vía `AJUSTE` (las condonaciones/reasignaciones de D-16/D1-6/L-4/S-5 de
hoy) — L-4 mostraba TOTAL PAGADO=71 cuando la plata real era 99. Corregido:
`fila_d[concepto] = PAGO + max(0, −DEBIA)` (DEBIA = CARGO+AJUSTE del mes, ver
`reporte_seguimiento._resumen_y_historial` — negativo cuando un AJUSTE perdona
deuda). Verificado exacto contra la plata real en los 4 predios corregidos:
L-4 99 (=41+58) · D-16 119 (=34+85) · D1-6 66 (=33+33) · T-12/I-9/F-9 también
cuadran (con su sobrante sin exponer cuando la plata alcanzaba de más).

**Cuarto fix:** el corte exonerado (S-5, F1-4, V-14 en `ajustes_cargo.xlsx`)
se seguía mostrando como "pagado" en la columna Corte de la cascada agua,
doble-contando la misma plata que ya se había acreditado a ACUERDOS por
`AJUSTE` — `_datos_ciclo()` ahora consulta `ajustes_cargo.xlsx` y pone
`corte_debido=0` si hay una fila `CORTE_RECONEXION` para ese predio (sin exigir
`CLASE=EXONERACION` — F1-4 no la tiene escrita, el motivo dice "anula" igual).
S-5 quedó en TOTAL=175 (145 real + 30 crédito histórico), exacto.

Con los 4 fixes, `4b_reclamos/reporte_historico.py` y
`reporte_reimputacion_cascada.py` ya reflejan fielmente el ledger — reporte
completo (208 predios) re-corrido y verificado a las 16:2x del 09/08/2026.

## Cómo se cierra este evento

Ya cerrado (09/08/2026) — los 8 predios del lote de Wagner/Yanet quedaron
verificados: D-16, D1-6, L-4, T-12, I-9, F-9, S-5 corregidos; F1-4 verificado sin
cambios necesarios. El fix de `reporte_historico.py` ya está aplicado y
verificado contra los 6 predios corregidos (arriba). `reporte_reimputacion_cascada_2026-07.pdf/.xlsx`
todavía refleja el estado de las 11:52 (antes de todo esto) — falta re-correrlo
completo (208 predios) para que quede al día; se dejó pendiente a propósito
para el final de la sesión. Borrar esta sección después de esa corrida, si no
aparece nada más.

---

# LEER ANTES — el AJUSTE de reversión de `5_cobranza` cambió de signo (07/08/2026)

**Si vas a leer un AJUSTE con `SOURCE=5_cobranza`, el signo significa cosas distintas
según la fecha.** Desde hoy un AJUSTE **positivo** de 5_cobranza quiere decir *"se
revirtió un pago, la deuda vuelve"*. Los 10 anteriores al 07/08/2026 dicen lo mismo con
el signo al revés — y por eso están mal.

## Qué estaba roto

La columna `AJUSTE` del ledger está en unidades de **DEUDA** (`seguimiento_repo._registrar`:
`saldo += monto`). `5_cobranza` la escribía y la leía en unidades de **CRÉDITO**, sin
traducir en ninguna de las dos puntas:

```
delta = pagado_fresco − ya          delta < 0  =  "acredité de más"
   escribía   registrar_ajuste(delta)   → −75  → el saldo BAJA otra vez
   correcto   registrar_ajuste(−delta)  → +75  → el saldo SUBE, la deuda vuelve

cargo 200 · run 1 acredita 75  → saldo 125
run 2, el pago ya no está      → saldo  50     ✗ la verdad es 200
                                  2 × el monto por debajo, y POSITIVO:
                                  el chequeo de "0 saldos negativos" no lo ve
```

No dispara en la primera corrida de un mes (`ya = 0` ⇒ `delta ≥ 0` ⇒ solo PAGO). Dispara
en re-corridas donde el insumo encogió — un `--force` tras corregir el crudo, o un humano
que sacó esa deuda de la planilla.

**Por qué importa más que un reporte:** `2_planilla/main.py:148 _join_saldo_pueblo` lee
esos saldos para MULTA · ACUERDOS · CONVENIO del mes siguiente. Un saldo torcido se
factura en la boleta, de menos, y nadie reclama por una boleta más barata.

## Qué se cambió (código, no datos — el ledger NO se tocó)

```
5_cobranza/main.py  · las dos mitades del mismo cambio de unidad, van juntas
   :2349  ya = pago_registrado − ajuste_reconciliado      (antes: +)
   :2374  registrar_ajuste(..., −delta, ...)              (antes: delta)

test nuevo  5_cobranza/tests/test_reversion_signo.py
   la secuencia entera: corrida → el insumo encoge → re-corrida → el pago reaparece
   antes del fix FALLA en el paso 2 (AJUSTE −75, saldo 50); después pasa 11/11
   invertir UNA sola de las dos líneas rompe la idempotencia — el test lo cubre (paso 3)

tests alineados a la convención nueva (mismo hecho, distinto signo esperado)
   5_cobranza/tests/test_reconciliacion_pueblo.py  caso 4: AJUSTE +25, saldo 30
   shared/tests/test_seguimiento_repo.py           bloque 12: ajuste +25, saldo 50
```

Los AJUSTE **manuales** no cambian: siguen en unidades de deuda (−75 = la directiva
perdona 75). `ajuste_reconciliado` filtra por `SOURCE`, así que el fix no los toca.

## ⚠ Lo que sigue vivo

```
NO CORRER  5_cobranza --force sobre 2026-07
   (la prohibición sigue en pie por disciplina de ciclo cerrado — pero las 10 filas
   de abajo ya están todas corregidas a mano, con AJUSTE fechado hoy, no con --force)

AGOSTO ARRANCA LIMPIO
   0 AJUSTE con source=5_cobranza y MES=2026-08. El chequeo del ciclo sigue siendo
   el mismo: contarlos después de correr. Ahora, si aparece alguno, es legible.

LOS 10 AJUSTE DE JULIO — TODOS CERRADOS (09/08/2026)
   P-6 CONVENIO · D1-3 MULTA · Q-4 MULTA/ACUERDOS · L-4 MULTA · C-21 MULTA ·
   J-6 MULTA/ACUERDOS: ya estaban en su valor correcto (0) o se restauraron al
   CARGO real / se condonaron por ser bien institucional. Ver sección
   "condonación institucional" y "abono rezagado" más abajo.
   D-16 ACUERDOS y D1-6 MULTA necesitaron algo más — ver esa misma sección de
   abajo, "3 predios con abono rezagado sin aplicar".

HALLAZGO COLATERAL, YA CERRADO
   5_cobranza/tests/test_cobranza.py redirigía el ledger a su carpeta temporal pero
   NO la vista: escribía shared/vista_seguimiento_pueblo.xlsx y .pdf REALES con datos
   sintéticos. El 07/08 solo se salvó porque el .xlsx estaba abierto en Excel (el .pdf
   sí se pisó y se restauró con git). Redirigidos VISTA_PATH y VISTA_PDF_PATH.
   Falla pre-existente que queda abierta en ese test, ajena a esto y verificada contra
   HEAD: «trazabilidad · filas (incluye huérfano) esperado=6 obtenido=5».
```

## Cómo se cierra este evento

Cuando las 10 filas de julio estén reescritas con la convención nueva (o el ciclo se
archive de forma que nadie pueda re-correrlo) y D-16 / D1-6 tengan su MOTIVO escrito.
Hasta entonces, la prohibición del `--force` sobre 2026-07 sigue en pie.

Backup previo al fix: `shared/backups_ledger/seguimiento_pueblo_pre_fix_signo_20260807_085808.xlsx`
(1486 eventos, idéntico al archivo vivo — el fix es de código).

---

# LEER ANTES — A-4 vuelve a deber S/75 de CONVENIO: su aporte al tanque se lo había pagado (06/08/2026)

**Si A-4 (Yolanda Espinoza Jaimes) figuraba al día en CONVENIO y ahora debe S/75: es
correcto, y hay que avisarle.** El ledger la daba por pagada con plata que nunca fue
para su deuda. La secretaria ya había reclamado que A-4 debía; se le ignoró porque el
sistema mostraba un pago de S/236 que parecía cubrir todo.

## Qué pasó

```
07/07 07:20:05  yape S/136  mensaje "mz A lt 4"          → deuda normal
07/07 07:20:55  yape S/100  mensaje "mz A lt 4 tanque"   → aporte al tanque, NO es deuda
        │
        └─ motor_matching identificó el lote pero dejó CONCEPTO vacío
           (ESTADO_PAGO = "sin deuda en planilla")
                │
                └─ los 236 entraron juntos y la cascada de 5_cobranza se
                   los comió como pago de deuda → CONVENIO quedó "pagado"
```

Es el mismo bug de `CONCEPTO` vacío que documentan más abajo las auditorías de junio y
julio (E-1 con S/100, y los S/550 de C-15 · P-7 · A-4 · P-17).

## La cuenta que lo prueba

```
deuda de julio (planilla_cobrado)          plata REAL de deuda: 136
   agua 5 + mant. 3 + mes ant. 23 =  31    136 − 31 (agua)        = 105
   MULTA                             30          − 30 MULTA       =  75
   ACUERDOS                          75          − 75 ACUERDOS    =   0
   CONVENIO                          75          CONVENIO 75 sin pagar
                                    ───
                                    211    planilla dice SALDO 75 · PARCIAL  ✔
```

`planilla_cobrado` **siempre tuvo razón**: `5_cobranza/main.py:1422` ya lee
`shared/aportes_tanque_manuales.xlsx` y descuenta el aporte antes de la cascada. Quien
nunca se enteró fue el ledger, que se había escrito antes de esa segregación.

## Qué se borró

```
SE BORRÓ (5 filas de A-4 · CONVENIO)              criterio: corrige un error del
   08/07 16:22  PAGO   +75  5_cobranza            sistema sobre sí mismo
   13/07 05:51  AJUSTE −75  5_cobranza            (el mismo del 06/08)
   13/07 06:16  PAGO   +75  5_cobranza   ← el mismo pago contado de nuevo
   25/07 14:51  AJUSTE −75  5_cobranza
   27/07 17:07  AJUSTE +225 correccion_manual (fix_race_condition_yape_20260713)

SE CONSERVÓ
   el CARGO de la siembra (75)  → saldo CONVENIO = 75, la verdad
   MULTA 30 y ACUERDOS 75, que SÍ entran dentro de los 136 reales
```

Ledger: 1491 → 1486 eventos · **0 saldos negativos en todo el pueblo**.
`vista_seguimiento_pueblo` regenerada.

## Backups

**El error está en TODOS los backups** — se escribió el 08/07 y el backup más viejo es
del 03/08:

```
seguimiento_pueblo_pre_clase_20260803_112318.xlsx      ← el primero que existe, ya con el error
seguimiento_pueblo_pre_A4_convenio_20260806_153740.xlsx ← justo antes de esta corrección
```

Las 5 filas, con sus `AUDIT_REF`, timestamps y saldos, están en el segundo. Para
reconstruir la historia de cómo llegó a estar mal, sirve cualquiera de los dos.

## Lo que sigue vivo

```
① el arreglo de fondo NO está hecho: motor_matching sigue sin llenar CONCEPTO=tanque
   cuando el mensaje lo dice y el lote matchea directo. Ver la sección "auditoría del
   ciclo julio" más abajo, § "Arreglo de fondo".

② los otros 3 del mismo lote de tanque (C-15, P-7, P-17) NO se revisaron en el ledger.
   Sus saldos en planilla_cobrado dieron 0, así que probablemente no tengan este
   problema — pero no está verificado contra seguimiento_pueblo.xlsx.

③ AVISAR A A-4: debe S/75 de convenio. Si ya se le dijo que estaba al día, es un
   cambio que ella va a notar.
```

## Cómo se cierra este evento

Cuando A-4 esté avisada, los otros 3 predios de tanque estén verificados en el ledger
y `motor_matching` marque `CONCEPTO` por mensaje, borrar esta sección.

---

# LEER ANTES — los pagos que declaró la secretaria y el ruido que generaron (06/08/2026)

**Si buscás en `shared/seguimiento_pueblo.xlsx` los ajustes de julio de B1-12, D-1, F1-5,
G-18, G1-14, L-16, N-5, Q-5, Q-10, Q-11, R-4, S-1, S-5, S-8, S-9, S-12, T-7 o V-6 y no
están: se borraron a propósito el 06/08/2026.** Los pagos de esos predios SIGUEN ahí; lo
que se fue es el ruido que el sistema generó alrededor de ellos.

## Qué pasó

```
28/07  la secretaria confirma que 28 predios-concepto ya pagaron.
       Los precursores no servían: abonos_rezagados/ajustes_cargo NO tocan el ledger de
       MULTA/ACUERDOS cuando no hay plata real del ciclo (5_cobranza solo escribe si hay
       diferencia contra el total_pagado REAL) — ver
       docs/retomar/RETOMAR_notas_secretaria_julio_grupo2_2026-07-29.md § 3.1
       ⇒ se escribió registrar_pago DIRECTO en el ledger
          TIPO_EVENTO=PAGO · SOURCE=manual · AUDIT_REF=notas_2026-07|<predio>-<CONCEPTO>
              │
30-31/07      ▼  corre 5_cobranza. Su reconciliación pregunta "¿cuánto ya registré?" con
                 pago_registrado(), que suma TODOS los PAGO del mes SIN filtrar por SOURCE
                 (seguimiento_repo.py:361 — su hermana ajuste_reconciliado SÍ filtra).
                 Cuenta el pago de la secretaria como propio ⇒ delta negativo
                 ⇒ AJUSTE −X a ciegas ⇒ saldo torcido
              │
30-31/07      ▼  se escribe a mano un "estabilizador" +X para enderezarlo
                 (AUDIT_REF …-condonacion-estable-… / …-estabilizador-…)
```

Las dos últimas filas de cada caso suman cero entre sí: son el sistema equivocándose y el
parche de esa equivocación. Nada de eso ocurrió en el mundo real.

## Qué se borró y qué se conservó

```
SE BORRÓ   44 filas = 22 pares (AJUSTE de 5_cobranza + su estabilizador manual)
           invariante verificado: NINGÚN saldo cambió · 0 saldos negativos

SE CONSERVÓ  los 28 PAGO de la secretaria   ← el hecho real
             los CARGO de la siembra y los PAGO reales de 5_cobranza
```

**Backup con el ruido intacto:**
`shared/backups_ledger/seguimiento_pueblo_pre_pares_condonacion_20260806_093522.xlsx`
(1535 filas; el ledger vivo quedó en 1491). Backup de la clasificación posterior:
`shared/backups_ledger/seguimiento_pueblo_pre_clase_declaracion_20260806_093613.xlsx`.

## CLASE nueva: `DECLARACION_SECRETARIA` — es una marca de trabajo pendiente

Los 28 pagos estaban en `SIN_CLASIFICAR`. Ahora llevan `DECLARACION_SECRETARIA`
(S/ 1,257.00 en total), clase agregada a `shared/seguimiento_repo.py:CLASES_VALIDAS`.

```
significa   la secretaria dijo que el vecino YA PAGÓ · el pago vale y salda su deuda
NO significa que esté resuelto de dónde salió esa plata

⚠ PENDIENTE — falta investigar, uno por uno, si cada uno es:
     (a) un EXCESO que ya estaba en caja          → clase final DECLARACION
                                                     no suma a caja: la plata ya se contó
     (b) un pago nuevo, nunca registrado          → sembrar en abonos_rezagados.xlsx
                                                     y clase final ABONO_REZAGADO: SÍ suma

mientras tanto NO suma a caja (no está en CLASES_SUMAN_CAJA), que es la posición
conservadora. Es la misma decisión abierta desde el 04/08 para los 8 casos de "ya pagué".
```

---

# LEER ANTES — el ledger perdió 52 filas a propósito (06/08/2026)

**Si estás buscando en `shared/seguimiento_pueblo.xlsx` los pagos de julio de A-8, B-5,
C-1, C-7, E-12, F-4, H-16, I-11, I-16, J-3, K-17, K-2, L-5, P-12 o W-5 y no los
encontrás: se borraron a propósito el 06/08/2026.** No eran hechos del negocio: eran el
rastro de un bug del sistema. Todo está preservado en el backup, y acá está la historia.

## Qué pasó

```
06/07/2026 14:08-14:13   UNA corrida de 5_cobranza cobró el ciclo 2026-07 leyendo
                         4_pagos/yape/motor_matching/outputs/pagos_yape_tepago.xlsx
                         que TODAVÍA era el de junio (4_pagos aún no se había
                         re-corrido para julio; el de efectivo sí estaba al día)
                              │
                              ▼
                         los pagos de JUNIO se aplicaron contra la deuda de JULIO
                         → 19 PAGO sin respaldo en 15 predios
                         → los 15 pagaron por yape en junio · ninguno por efectivo
                         → el monto de cada fantasma = pago de junio − agua de julio,
                           topado por la deuda del predio (14 de 15 cuadran exacto)

31/07/2026 18:08-18:20   5_cobranza recalcula, ya no encuentra esos pagos y los
                         revierte… con el signo invertido (main.py:2320): el AJUSTE
                         negativo RESTA deuda en vez de devolverla
                         → saldos negativos o mordidos

03/08/2026               se restauran 12 a mano (+2 × el monto revertido) — lote bug_signo

06/08/2026               se BORRAN las 52 filas del bug (ver abajo)
```

## Qué se borró y qué NO

```
SE BORRÓ (52 filas)                          criterio: corrige un error del sistema
   19  PAGO fantasma del 06/07                       sobre sí mismo, efecto neto 0
   18  AJUSTE de reversión mal firmada               (o el neto ES el error)
   15  AJUSTE manual del 03/08 (bug_signo)

SE CONSERVÓ                                  criterio: corrige un dato del NEGOCIO,
   los 59 pagos restantes del 06/07 que sí          su efecto es parte del saldo real
   coincidían con un pago real de julio
   los 109 AJUSTE de correccion_genesis_formula (04/07) — bug de fórmula de abril:
   el Saldo de génesis omitía los pagos de abril y el convenio salió sobrecobrado.
   Ya se les escribió el MOTIVO (06/08) porque estaban en blanco.
```

**Backup con todo el rastro del bug, intacto:**
`shared/backups_ledger/seguimiento_pueblo_pre_eliminar_bug0607_20260806_081824.xlsx`
(1589 filas; el ledger vivo quedó en 1537). Ahí están las 52 filas con sus AUDIT_REF,
timestamps y saldos, por si hay que reconstruir la historia o explicar una boleta vieja.

**Saldos después del borrado** — cero negativos, cero desajustes. Cuatro predios
cambiaron porque su bug nunca se había compensado:

```
E-12 ACUERDOS   0 → 25       L-5 MULTA   34 → 50
F-4  MULTA     48 → 50       W-5 ACUERDOS 37 → 47
```

⚠ **Sus boletas del 01/08 salieron con el saldo mordido**, así que en septiembre esos
cuatro vecinos ven subir la deuda sin haber consumido nada. Avisarles.

## Lo que sigue vivo y hay que arreglar

```
5_cobranza/main.py:2320   el AJUSTE de reversión sigue saliendo con el signo invertido.
                          Va a volver a fabricar saldos negativos en cuanto corra agosto.
                          Dos caminos, ninguno decidido — ver
                          3_boletas/inputs/reclamos_2026-08-01/README.md § BUG_SIGNO.

ya mitigado (06/08)       el mes del ciclo ahora lo declara 1_lecturas en
                          shared/ciclo_activo.json, y los outputs de 4_pagos llevan el
                          periodo en el nombre (pagos_yape_tepago_2026-08.xlsx). Si
                          4_pagos no corrió para el ciclo, 5_cobranza corta en vez de
                          leer el archivo del mes anterior. Ver shared/ciclo.py.
```

---

# LEER ANTES — auditoría del ciclo junio 2026 (04/08/2026) · sirve para armar el balance

**Para qué existe esta sección:** el 04/08/2026 se corrió `5b_validacion` sobre el
ciclo de junio (ya cerrado) para responder dos preguntas: *(a)* ¿junio quedó
realmente cuadrado?, y *(b)* ¿los reclamos de "ya pagué el mes anterior" vienen de
un error de cálculo de junio? Las respuestas están abajo. **Cuando toque preparar un
balance o explicar la recaudación de junio, leer esto primero** — evita re-investigar
las mismas 4 diferencias, y una de ellas cambia el número de junio.

## Resultado en una imagen

```
5b_validacion · junio 2026 · ventana 19/05 → 15/06

NIVEL 1a  banco TE PAGÓ 4,164 vs ∑partes 4,264   dif +100  ✗ ERROR REAL
NIVEL 1b  banco PAGASTE 2,278.90 vs procesado    dif    0  ✔
NIVEL 2   agua yape → planilla                   dif    0  ✔
EFECTIVO  procesado 19,129.50 vs planilla 19,211.50  +82  ✗ pero NO es plata
```

## El único error real de junio: S/100 contados dos veces (E-1)

```
HECHO    10/06/2026 · Yolanda Montalvo (E-1) · S/100 · yape
         mensaje del yape: "E1 1 tanque"
              │
INTERPRETACIÓN — quedó doble
         ├─► aportes_tanque.xlsx      CONCEPTO=tanque   ✔ correcto
         └─► pagos_yape_tepago        CONCEPTO=VACIO    ✗ el motor no lo marcó
                   │
                   └─► 5_cobranza lo aplicó como AGUA a E-1
                       TOTAL_A_PAGAR=8 · MONTO_YAPE=100 · SALDO=−100 · EXCESO

CONSECUENCIA
   · agua de junio inflada en 100 (3,567 real → 3,467)
   · tanque de junio correcto (500) pero 100 de eso ya se contó como agua
   · E-1 figura con un crédito de 100 que no le corresponde
   · su agua (S/8) ya la había pagado en EFECTIVO → el yape es 100% tanque
```

**Impacto acotado, verificado:** E-1 **no** aparece en `arrastre_consolidado_2026-06`
ni en `arrastre_devolucion_2026-06`, así que el −100 nunca se propagó a julio. Es un
error contenido dentro de junio.

**Arreglo decidido (PENDIENTE DE EJECUTAR, 04/08/2026):** una fila en
`shared/reasignaciones_aplicacion.xlsx`:

```
MZ=E · LT=1 · CONCEPTO_ORIGEN=AGUA · CONCEPTO_DESTINO=TANQUE · MONTO=100
MES_ANO = (vacío)  → SOLO REGISTRO, no lo aplica ninguna corrida
MOTIVO  = el yape decía "tanque"; tepago lo dejó sin concepto y 5_cobranza lo
          aplicó como agua. Su agua (S/8) ya estaba pagada en efectivo.
REF_TRANSACCION = pagos_yape_tepago 10/06/2026 19:08:13
```

**NO se re-corre junio.** Un ciclo cerrado se lee, no se reescribe; se corrige con un
asiento fechado hoy y efecto declarado en el ciclo de origen. Para cualquier reporte:

```
recaudación de junio  =  junio congelado  +  asientos con ORIGEN=2026-06
```

## Las otras 3 diferencias NO son plata — son el reporte mirando mal

Distinción que hay que tener clara al usar estos números:

```
EL LIBRO                              EL LENTE
mesas · tepago · pagos_efectivo       5b_validacion
planilla_cobrado · arrastre_*
                                      no escribe nada del libro
INMUTABLE una vez cerrado             RECOMPUTABLE sin límite
se corrige con asiento nuevo          mejorarlo NO es restatement
```

| dif | predio | qué es | clase |
|---|---|---|---|
| +33 | C1-6 | el N° de recibo (2167) quedó tipeado en la columna `CONCEPTO` de `mesa_6` | lente |
| +25 | O-2 | idem (1206) | lente |
| +24 | R-7 | blanco de efectivo asignado a un lote; `pagos_efectivo` no contiene blancos | lente |

- **C1-6 y O-2:** `5b_validacion/main.py:358` excluye toda fila con `CONCEPTO` no
  vacío. Un número de recibo ahí la saca del conteo. El pago entró y se aplicó bien.
  Arreglo: que 5b excluya solo una **lista blanca** de conceptos
  (`tanque · honorario · gasto …`) y reporte como «CONCEPTO desconocido» lo que no
  reconozca, en vez de excluirlo en silencio.
- **R-7:** el pago existe — `mesa_2 · Yerald Romero · 05/06 · S/24 · "Exoneracion"`
  sin MZ/LT → salió a `blancos_mes.xlsx`, no a `pagos_efectivo.xlsx`. Después se
  atribuyó a R-7 (30 de mesa_5 + 24 del blanco = 54), generó EXCESO −24 y **en julio
  ya se corrigió**: `arrastre_devolucion_2026-06`, R-7, 24, `ESTADO=RESUELTO`,
  *"Este blanco era de M-7. Se le cambio en Julio."* Arreglo del lente: la sección
  EFECTIVO de 5b debe sumar los blancos asignados, igual que el Nivel 1a ya suma los
  blancos de yape (197).

Junio tuvo **S/282 de blancos en efectivo** (69 · 37 · 24 · 45 · 107). Solo uno se
asignó a un lote dentro de junio; por eso la diferencia es 24 y no 282.

## Reconciliación crudo ↔ procesado — verificada, no se perdió ningún pago

```
YAPE                                                              1:1 exacto ✔
  crudo banco TE PAGÓ   4,164.00  =  pagos_yape_tepago   4,164.00
  crudo banco PAGASTE   2,278.90  =  pagaste procesado   2,278.90

EFECTIVO — mesas vs pagos_efectivo: 6 lotes difieren, todos con causa ✔
  A-8C  140  →  es la FILA DE EJEMPLO de la plantilla (María García, 03/06)
                repetida en las 7 mesas: 7 × 20 = 140. No es plata.
  BLANCO 131 →  van a blancos_mes.xlsx (24 + 107)
  T-7 → I-7   45  →  corrección de lote, ciclo 3
  C1-16 → B1-12 16 →  corrección de lote, ciclo 4
```

⚠ **Al sumar las mesas a mano:** traen **filas de totales del cobrador al pie de la
hoja**. Sumar la columna entera da ~25,342 en vez de ~19,400. El módulo las descarta
bien; la comparación válida es lote por lote, no por total de columna.

## `trazabilidad_2026_06.xlsx` tiene 2 pagos que junio no tiene — y está bien

Esto asustó en su momento; queda cerrado:

```
trazabilidad_2026_06 — 9 identificaciones manuales · S/1,037
   7 están en pagos_yape_tepago de junio                    ✔
   2 NO están:
        ANALY QUINECHE MORANTE  S/300  18/06/2026  C-35
        MARIA GODO SIFUENTES    S/ 41  17/06/2026  F-1

   ventana junio ────┤ 15/06 21:13        los 2 cayeron DESPUÉS del corte
                     │  17/06 · 18/06      → pertenecen a JULIO
```

Los dos están cobrados en `7_cierre/archivo/2026-07/planilla_cobrado.xlsx`
(C-35: yape 300, EXCESO −206 · F-1: yape 41, PARCIAL, saldo 20). **Ningún pago
perdido.** `trazabilidad_<mes>` es un registro de *identificaciones manuales*, no un
libro de pagos del mes: guarda la identificación aunque el pago pertenezca al ciclo
siguiente. No usarlo como fuente de "lo que se cobró en el mes".

## S/200 sin ubicar en mesa_2 — única cifra sin rastro en archivos

```
mesa_2.xlsx · pie de la hoja de Yerald Romero (05/06)
   fila 41   1,699.50 │ 151 │ 1,850.50    ← su total
   fila 42   "falta ubicar lote"  │ 200   ← S/200 cobrados, sin dueño anotado
   fila 43   1,899.50                      ← total CON los 200

   pagos_efectivo.xlsx   ✗      blancos_mes.xlsx      ✗
   blancos_acumulados    ✗      arqueo_2026-06 tomó 1,699.50 (los dejó afuera)
```

El usuario los da por explicados por conocimiento propio (04/08/2026); **no hay
evidencia en ningún archivo del repo**. Si al armar el balance el total no cierra por
~200, empezar por acá. Una nota escrita al pie de una hoja no es una fila: ninguna
validación la ve.

## Conclusión sobre los reclamos "ya pagué el mes anterior"

**No vienen de un error de cálculo de junio.** Junio reconcilia contra el banco al
céntimo y contra las mesas lote por lote. También se descartó un hueco de ventana:

```
junio cierra 15/06 21:13   ·   julio abre 17/06 20:32
crudo del banco el 16/06: CERO movimientos  ⇒ no hay pagos huérfanos entre ciclos
```

Quedan dos causas posibles, y son distintas — conviene separarlas al atender a cada
vecino:

```
① EL PAGO SÍ SE ANOTÓ, pero la cascada lo consumió en otra cosa
   ejemplo verificado — F-1 MARIA GODO SIFUENTES:
      yape S/41 del 17/06 → SÍ se aplicó
      agua 17 + mantenimiento 3 + mes_anterior 21 = 41 exacto
      MULTA 20 quedó impaga → ella cree que pagó la multa
   ⇒ discrepancia de ORDEN DE IMPUTACIÓN, no de plata

② EL PAGO NO SE ANOTÓ NUNCA
   los casos YA_PAGO de reclamos_2026-08-01 (K-8 · D-6 · F-1 multa · F-7 · F-10 …)
   se buscaron el 03/08 contra los ciclos congelados de 7_cierre/archivo/ y
   NINGUNO tiene un pago que coincida con el monto declarado.
   ⇒ si el vecino tiene recibo, el pago existió y nunca entró al sistema
```

**Acción sugerida por el usuario (04/08/2026): consultar las hojas físicas de
registro / los recibos.** Es el único lugar donde puede estar la evidencia del caso ②
— el sistema ya se descartó como origen del problema.

## Cómo se cierra este evento

1. Escrita la fila de E-1 en `shared/reasignaciones_aplicacion.xlsx`.
2. 5b corregido: lista blanca de `CONCEPTO` + suma de blancos asignados en la sección
   EFECTIVO + lectura de asientos para marcar `EXPLICADA` en vez de `ALERTA`.
3. Re-corrido `5b_validacion` de junio dando **0 diferencias sin explicar**.
4. Resueltos o descartados los S/200 de mesa_2.

Mientras tanto, para cualquier balance: **junio = archivos congelados + los S/100 de
E-1 reclasificados de agua a tanque.** Todo lo demás de junio está sano.

---

# LEER ANTES — auditoría del ciclo julio 2026 (04/08/2026) · sirve para armar el balance

Mismo ejercicio que la sección de junio, corrido el mismo día sobre
`C:\Users\wilde\PycharmProjects\Julio\jass_system - Julio`. **Julio salió mejor que
junio: el efectivo cuadra exacto y la única alerta tiene una causa de una sola
línea.** Leer junto con la sección de junio antes de armar cualquier balance.

## Resultado en una imagen

```
5b_validacion · julio 2026 · ventana 17/06 → 20/07

NIVEL 1a  banco TE PAGÓ 4,603 vs ∑partes 5,153     dif +550  ✗
NIVEL 1b  banco PAGASTE 5,405.90 vs procesado      dif    0  ✔
NIVEL 2   agua yape → planilla                     dif    0  ✔
EFECTIVO  procesado 10,718.50 vs planilla          dif    0  ✔  ← junio tenía 3 difs
```

Es el `gap_conocido` que `shared/reporte_acumulado_procesado/estado_ciclo.json` ya
registraba para 2026-07 como *"causa raiz sin resolver"*. **Queda resuelto acá.**

## Los +550 son 4 filas, exactas al sol

```
shared/aportes_tanque_manuales.xlsx · filas CANAL=yape cuyo pago YA vino del banco

   C-15  JULIO RIOS BLAS         200   mensaje del yape: "mz c lt 15 tanque"
   P-7   HIPOLITO MELGAREJO      200   confirmado verbalmente (el mensaje no lo decía)
   A-4   YOLANDA ESPINOZA        100   mensaje del yape: "mz A lt 4 tanque"
   P-17  MAURICIA DURAND          50   mensaje del yape: "tanke adelanto mz p lote 17"
                                 ───
                                 550   = la diferencia, sin resto
```

## Por qué el reporte los cuenta dos veces

```
un solo yape de C-15 · S/200
        │
        ├─► pagos_yape_tepago    CONCEPTO = VACIO  ──► 5b lo mete en «agua»   3,841
        └─► aportes_tanque.xlsx  CONCEPTO = tanque ──► 5b lo mete en «tanque» 1,250

   ∑partes suma los dos ⇒ 550 de más contra el crudo del banco
```

Causa raíz, ya escrita en el `MOTIVO` del propio archivo:

> *"el hueco es que `CONCEPTO` nunca se llena para pagos matcheados directo (solo se
> llena a mano en `pendientes.xlsx` para casos ambiguos)"*

Por eso M-18, Q-3 y G1-4 —que pasaron por `pendientes.xlsx` hoja `Segregacion`— **sí**
llevan `CONCEPTO=tanque` en tepago y no descuadran. Los 4 de arriba matchearon directo
y quedaron sin marcar.

## La reconciliación a mano — cuadra exacto

```
  banco crudo TE PAGÓ                            4,603.00
  ────────────────────────────────────────────────────────
  agua identificada (tepago CONCEPTO vacío)      3,841.00
  − aportes tanque manuales, canal yape            −550.00
  = agua real                                    3,291.00
  + blancos                                           0.00
  + tanque (aportes_tanque.xlsx, canal yape)     1,250.00
  + otros conceptos (deuda_directiva)               62.00
  ────────────────────────────────────────────────────────
  ∑ partes corregida                             4,603.00   = banco  ✔  dif 0
```

**Para el balance de julio:** la recaudación de agua por yape es **3,291**, no 3,795.
`planilla_cobrado.MONTO_YAPE` (3,795) incluye los 550 de tanque. El total de plata
está bien; lo que está mal es el *reparto* entre agua y tanque.

## En julio la plata SÍ quedó bien segregada — nadie recibió deuda perdonada

Diferencia clave con junio: `5_cobranza/main.py:1422` **lee**
`aportes_tanque_manuales.xlsx` y descuenta el aporte del total pagado **antes** de la
cascada P1-P6.

```
   C-15  deuda  13 · yape 200 + efectivo 13 → SALDO  0   (aplicó 13, no 213)
   P-7   deuda  19 · yape 219               → SALDO  0   (aplicó 19)
   A-4   deuda 211 · yape 236 (136+100)     → SALDO 75   (aplicó 136, no 236)
   P-17  deuda  10 · yape  50 + efectivo 10 → SALDO  0   (aplicó 10)

   ninguno aparece en arrastre_devolucion_2026-07
   ⇒ el exceso tampoco se arrastró como crédito al mes siguiente
```

Quién lee el precursor y quién no:

```
                aportes_tanque_manuales.xlsx
                          │
     ┌────────────────────┼────────────────────┐
     ▼                    ▼                    ▼
 5_cobranza          consolidar_tanque    5b_validacion
   LO LEE ✔             LO LEE ✔            NO LO LEE ✗   ← acá nace el +550
```

**El +550 es un artefacto de reporte, no un descuadre de caja.**

## Dos pendientes que deja julio

```
① O-16 (S/100) y H1-3 (S/100), cargados el 25/07 en aportes_tanque_manuales.xlsx,
   NO están en 4_pagos/outputs/aportes_tanque.xlsx
   ⇒ consolidar_tanque.py no se re-corrió desde el 25/07
   ⇒ cuando corra, la dif de Nivel 1a pasa de 550 a 750 (mismo origen, no es nuevo)

② junio NO tenía este overlay: shared/ del repo junio no tiene
   aportes_tanque_manuales.xlsx. Por eso el E-1 de junio quedó con SALDO −100 y los
   4 de julio no. Mismo bug de CONCEPTO; junio lo sufrió sin red, julio ya la tiene.
```

## Arreglo de fondo

```
① motor_matching: llenar CONCEPTO=tanque cuando el MENSAJE diga "tanque"/"tanke",
   aunque el lote matchee directo (hoy solo se llena vía pendientes.xlsx).
   Cubre C-15, A-4 y P-17. NO cubre P-7 — ese se confirmó de palabra y siempre
   va a necesitar el precursor manual.

② 5b_validacion: restar aportes_tanque_manuales.xlsx (canal yape, del mes) del
   bucket «agua» del Nivel 1a. Es el mismo tipo de arreglo que la sección de junio
   pide para blancos de efectivo y para la lista blanca de CONCEPTO:
   el lente tiene que leer los mismos precursores que ya lee el libro.
```

## Cómo se cierra este evento

1. `motor_matching` marca `CONCEPTO` por mensaje; `5b` resta el overlay de tanque.
2. Re-corrido `5b_validacion` de julio dando **0 diferencias sin explicar**.
3. Retirado el `gap_conocido` de 2026-07 en `estado_ciclo.json` (ya no es "causa raíz
   sin resolver": es esta sección).
4. Corrido `consolidar_tanque.py` para que O-16 y H1-3 entren a `aportes_tanque.xlsx`.

---

# CONCLUSIÓN CONJUNTA — junio + julio (04/08/2026)

Las dos auditorías de arriba responden la misma pregunta de fondo. **El resultado es
que el sistema no perdió pagos.**

```
                         junio              julio
banco crudo ↔ procesado   1:1 exacto        1:1 exacto
mesas ↔ pagos_efectivo    6 difs, todas     0 difs
                          con causa
ventana entre ciclos      sin hueco (cero movimientos el 16/06)
descuadres encontrados    todos con causa identificada y acotada
```

```
Lo único que las dos auditorías encontraron mal:
   junio  S/100  (E-1)   ─┐
   julio  S/550  (4 lotes)─┴─ el MISMO bug: CONCEPTO vacío en pagos_yape_tepago
                             para aportes al tanque que matchearon lote directo.
                             Afecta el REPARTO agua/tanque en los reportes,
                             no el total de plata recibida.
```

**Entonces los reclamos de "ya pagué el mes anterior" no vienen de un error del
sistema.** El flujo de datos hizo lo que tenía que hacer en los dos ciclos. Quedan dos
causas posibles, y hay que distinguirlas al atender a cada vecino:

```
① EL PAGO SÍ SE ANOTÓ, y la cascada lo consumió en otra cosa
   caso verificado — F-1 MARIA GODO SIFUENTES: yape S/41 del 17/06 se aplicó, pero
   agua 17 + mantenimiento 3 + mes_anterior 21 = 41 exacto, y la MULTA de 20 quedó
   viva. Ella cree que pagó la multa.
   → el recibo físico no aporta nada acá: es orden de imputación, no plata faltante.

② EL PAGO NUNCA SE ANOTÓ  ← la causa dominante
   K-8 · D-6 · F-1 (multa) · F-7 · F-10 · K-9 · T-14 y el resto de los casos YA_PAGO
   de reclamos_2026-08-01: se buscaron el 03/08 contra los ciclos congelados de
   7_cierre/archivo/ y NINGUNO tiene un pago que coincida con el monto declarado.
   → si el vecino tiene recibo, el pago existió y nunca entró al sistema.
```

**Acción definida (usuario, 04/08/2026): consultar el pago físico en las hojas de
registro / los recibos de los vecinos.** Es el único lugar donde puede estar la
evidencia del caso ② — el sistema ya quedó descartado como origen del problema en los
dos ciclos auditados.

---

# LEER ANTES — 20 boletas parchadas a mano por pedido de la directiva (01/08/2026)

## Qué pasó

La directiva pidió cambiar los datos de 20 boletas por reclamos, con cobro
el mismo día (01/08/2026) — no había tiempo de pasar por el flujo normal de
`4b_reclamos` antes de imprimir. Fotos/evidencia de cada reclamo en
`3_boletas/inputs/reclamos_2026-08-01/`.

## Qué se hizo (01/08/2026)

Parche directo en `3_boletas/inputs/DATA_boletas.xlsx` para los 20 predios
(no en `seguimiento_pueblo.xlsx` ni en ningún ledger real) — cosmético,
solo para que la boleta impresa hoy salga con el monto correcto. Detalle
de cada predio (campo cambiado, valor viejo→nuevo) queda en
`3_boletas/inputs/reclamos_2026-08-01.xlsx`.

## Por qué esto puede volver a pasar / por qué no está cerrado

El ledger real (`seguimiento_pueblo.xlsx`, `arrastre_consolidado`) sigue
con los valores viejos. Si `2_planilla`/`5_cobranza` se vuelven a correr
para este ciclo, `DATA_boletas.xlsx` se regenera y el parche se pierde.

## Arreglo de fondo (pendiente — cuadrar mañana)

Está planeado cuadrar esto contra el sistema real el 02/08/2026: para
cada uno de los 20, decidir si el reclamo es válido y aplicar la
corrección real en `seguimiento_pueblo.xlsx` (`registrar_ajuste`) o donde
corresponda según el concepto, siguiendo el mismo criterio de las
correcciones de `notas_2026-07.xlsx`.

## Cómo se cierra este evento

Cuando los 20 predios estén corregidos en el ledger real (no solo en
`DATA_boletas.xlsx`) y se confirme que un ciclo futuro ya no necesita
este parche, borrar esta sección.

---

# LEER ANTES — W-4 con MES_ANTERIOR parchado a mano (posible blanco de julio sin identificar)

## Qué pasó

Vicki Masias Cusihuamán (W-4) dice que pagó en su mesa en julio, pero el sistema
no tiene ningún registro de ese pago — busqué en `pagos_efectivo.xlsx`,
`trazabilidad_2026-07.xlsx` (las 3 hojas) y `blancos_efectivo.xlsx`, sin
resultado. El sistema mostraba SALDO=17 (MES_ACTUAL 14 + Mantenimiento 3) sin
ningún yape/efectivo aplicado.

R-5 (Frank Kelvin Teran Masias) tenía el mismo síntoma (S/8 de arrastre) y se
parchó igual al principio, pero el usuario confirmó que R-5 **no vino a pagar**
el mes anterior — la deuda es real. Revertido el mismo día (31/07/2026):
`DEUDA_AGUA` restaurado a 8 en `arrastre_consolidado_2026-07.xlsx`. Solo W-4
queda con el parche activo.

## Qué se hizo (31/07/2026)

Parche directo en `5_cobranza/outputs/arrastre_consolidado_2026-07.xlsx`
(DEUDA_AGUA: W-4 17→0) — **solo en agosto**, para que la boleta de agosto
salga limpia mientras se investiga julio. Documentado en
`shared/parches_manuales_pendientes_julio.xlsx`. El ledger/arrastre real de
**julio no se tocó** — este parche es cosmético para la boleta de agosto, no
una corrección del registro histórico.

## Por qué esto puede volver a pasar

Si `5_cobranza --force` se vuelve a correr para julio (ej. por otro precursor
pendiente), `arrastre_consolidado_2026-07.xlsx` se regenera de cero y este
parche se pierde — hay que reaplicarlo, o mejor, resolverlo de raíz antes de
que eso pase.

## Arreglo de fondo (pendiente, NO decidido)

Encontrar el blanco/pago real de julio para W-4 (~S/17) — revisar
`shared/blancos_efectivo.xlsx`, `4_pagos/efectivo/inputs/mesa_*.xlsx` de
julio (backups), o preguntar directo a la secretaria/cobrador. Si aparece,
reidentificarlo correctamente contra julio (no solo tapar agosto). Si NO
aparece tras buscar, decidir si es una condonación real o si hay que cobrarlo
igual.

## Cómo se cierra este evento

Cuando se confirme el origen real del pago de julio (o se decida que no
corresponde) y se corrija en el ledger de julio (no solo el parche de
agosto), borrar esta sección y la fila correspondiente en
`shared/parches_manuales_pendientes_julio.xlsx`.

---

# LEER ANTES — PREDIOS_INSTALACION_EXCLUIDOS (B-20, C-43, C-35, F1-11, G-21, W-2) perdieron su carga directa de CONVENIO

## Qué pasó

`shared/seguimiento_repo.py` excluye 6 predios (`PREDIOS_INSTALACION_EXCLUIDOS`) de
CONVENIO en `seguimiento_pueblo.xlsx`, asumiendo que su deuda de instalación "ya
está completa en arrastre_consolidado (junio la cargó full desde DATA_boletas)".
Verificado 30/07/2026: **eso ya no es cierto para B-20 y C-43** — ninguno de los
dos tiene fila en `arrastre_consolidado_2026-07.xlsx`, y `DATA_boletas.xlsx` de
julio les imprime "USTED ESTÁ AL DÍA" en vez de su deuda real (263 y 334, según
`obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx`). Se
perdió en alguna regeneración de `2_planilla`/`3_boletas` sin preservación
cableada — mismo patrón que rompió el MANTENIMIENTO de M-12.

## Qué se hizo (30/07/2026)

Sembrado en `shared/genesis_tardia.xlsx` (no en `seguimiento_pueblo.xlsx`, que
sigue excluyéndolos a propósito) — **los 6 predios de la lista, confirmado el
mismo problema en todos**: B-20 CONVENIO=263, C-43 CONVENIO=334, C-35
CONVENIO=94, F1-11 CONVENIO=676, G-21 CONVENIO=50, W-2 CONVENIO=50, todos
`MES_ANO_APLICA=2026-07`. F1-11 (María Elizabeth Silva Sosa) estaba en la hoja
`INSTALACIONES ANTERIOR DIRECTIV` (no `NUEVAS INSTALACIONES` — la búsqueda
inicial se la había perdido por un desfase de fila de encabezado en esa hoja).
`_cargar_genesis_tardia()` no consulta `PREDIOS_INSTALACION_EXCLUIDOS` (esa
lista solo la usan `sembrar_seguimiento_pueblo.py` y
`_reconciliar_pagos_pueblo`), así que no hay riesgo de duplicar contra
`seguimiento_pueblo.xlsx`.

## Arreglo de fondo (decisión de diseño, NO tomada todavía)

La premisa original de `PREDIOS_INSTALACION_EXCLUIDOS` (carga directa en
DATA_boletas que sobrevive regeneraciones) ya no se sostiene. O se implementa
preservación cableada de verdad para estos 6 predios, o se abandona el diseño
"carga directa" y se migran los 6 a genesis_tardia/seguimiento_pueblo como todos
los demás (quitándolos de la lista de exclusión). No decidir esto sin pasar por
Fase 1.

## Cómo se cierra este evento

Cuando se revisen los otros 4 predios y se tome la decisión de fondo (arreglar
preservación vs. migrar a genesis_tardia para todos), borrar esta sección.

---

# LEER ANTES — bienes del pueblo (Posta, Capilla, etc.) siguen generando MULTA/ACUERDOS pese a estar EXONERADO

## Qué pasó

`shared/registro_cortes.xlsx` tiene 4 predios institucionales marcados
`ESTADO=EXONERADO` desde 2026-02 (J-6 Capilla del Pueblo, J-14 Centro Educativo
20902, Z-6 Área Comunal del Pueblo, C-21 Puesto de Salud Tupac Amaru) — pero ese
flag **solo lo lee `6_corte`/`6b_corte_multas`** (`grep EXONERADO *.py` confirma:
ningún otro módulo lo consulta). La génesis mensual de MULTA (faena/reunión) y
ACUERDOS_ASAMBLEA (techado/campo) no sabe que estos predios son bienes
comunales sin vecino que participe — les sigue generando cargo cada mes.

## Qué se hizo (2026-07-30) — parche puntual, no arreglo de fondo

Condonado directo en `shared/seguimiento_pueblo.xlsx` vía `registrar_pago`
(SOURCE=manual, AUDIT_REF=notas_2026-07|LOTE-CONCEPTO):
- C-21: MULTA 50 → 0
- J-6: MULTA 50 → 0, ACUERDOS 75 → 0 (CONVENIO=100 de J-6 NO se tocó — podría
  ser instalación real de medidor, no cargo de participación social)

## Corregido (2026-08-09) — el parche del 30/07 se había roto y estaba mal modelado

Dos problemas encontrados al auditar los 10 `AJUSTE` con signo viejo de julio
(ver sección "el AJUSTE de reversión cambió de signo" más arriba):

1. **El 30/07, una corrida de `5_cobranza` con el bug de signo escribió un
   `AJUSTE` sobre la condonación de C-21/J-6**, y el fix posterior
   ("revertir-condonacion-fallida") deshizo TODO — pago + ajuste — dejando la
   deuda completa otra vez (C-21 MULTA=50, J-6 MULTA=50, J-6 ACUERDOS=75).
   Q-4 (mismo día, mismo patrón, condonación por otro motivo) sí recibió un
   segundo paso que lo devolvió a 0; C-21 y J-6 se quedaron así hasta hoy.
2. **Estaba mal modelado desde el origen:** se había registrado como
   `TIPO_EVENTO=PAGO, CLASE=DECLARACION_SECRETARIA` (el mecanismo de "la
   secretaria dijo que el vecino ya pagó", que arrastra un pendiente de
   investigar si es EXCESO o ABONO_REZAGADO — ver sección de abajo). Para un
   bien exonerado nunca hubo plata que investigar.

Se borraron los 12 eventos de esa cadena (PAGO + 3 AJUSTE × 3 predio-concepto;
backup `shared/backups_ledger/seguimiento_pueblo_pre_condonacion_institucional_C21_J6_20260809_095814.xlsx`)
y se reemplazaron por un solo `AJUSTE` limpio cada uno, `CLASE=EXONERACION`
("la directiva decidió no cobrar — nunca hubo plata"), con motivo explícito en
el ledger. Saldo actual: C-21 MULTA=0, J-6 MULTA=0, J-6 ACUERDOS=0.
Verificado: ninguno de los dos aparece en los precursores
(`abonos_rezagados.xlsx`, `ajustes_cargo.xlsx`, `genesis_tardia.xlsx`,
`reasignaciones_aplicacion.xlsx`, `blancos_efectivo.xlsx`).

## Por qué esto va a volver a pasar

Este parche solo limpia julio. Si la génesis de agosto (o cualquier ciclo
futuro) vuelve a generar MULTA/ACUERDOS para estos 4 predios, hay que repetir
el mismo `registrar_pago` a mano cada vez — no hay filtro en el origen.

## Arreglo de fondo (decisión de diseño, NO tomada todavía)

Que la génesis de MULTA/ACUERDOS filtre por el mismo `EXONERADO` que ya
respeta `6_corte`, para que estos predios dejen de generar el cargo desde la
fuente. Es una decisión de diseño (Fase 1, no un fix de Sonnet) — no
implementar sin pasar por ahí primero.

## Cómo se cierra este evento

Cuando se implemente el filtro de origen y se confirme que un ciclo futuro ya
no genera MULTA/ACUERDOS para estos 4 predios, borrar esta sección.

---

# LEER ANTES — "saldo a favor no reclamado" detrás de las declaraciones de la secretaria en abonos_rezagados

## Qué pasó

La secretaria revisa su cuaderno y declara que ciertos predios "están al día"
(no deben multa/reunión/faena/techado y campo), aunque el ledger (`planilla_cobrado.xlsx`)
muestre esos montos como pendientes. Verificado predio por predio (2026-07-30,
notas_2026-07.xlsx GRUPO 3): de 9 declaraciones, solo 3 coincidían exactas con el
ledger (SALDO=0 en MULTA/ACUERDOS); las otras 6 no.

## Por qué se cargan igual como abono rezagado, sin exigir comprobante

La secretaria explica el mecanismo: al declarar estos pagos se genera una
deuda/crédito en el sistema (el abono rezagado), pero eso tiene contrapartida —
existen saldos a favor (EXCESO) en otros meses que **nadie ha reclamado todavía**
(pagos de más por error de tipeo u otra causa en las mesas). Con el tiempo, ella
conecta cada declaración de "ya pagó" con alguno de esos saldos a favor sueltos.
Por eso se acepta la declaración sola (`RESPALDO=solo_declaracion`) sin exigir
comprobante bancario — el respaldo real es el pool de EXCESO no reclamado, no
un yape/depósito puntual.

## Riesgo real: doble conteo

Si en el futuro alguien identifica el EXCESO específico que efectivamente cubre
una de estas declaraciones (vía `reidentificacion.xlsx` o `blancos_efectivo.xlsx`),
**no debe aplicar un crédito adicional** — la fila de `abonos_rezagados.xlsx` ya
canceló esa deuda. Antes de reconciliar cualquier EXCESO contra un predio que
aparece en `abonos_rezagados.xlsx`, revisar si ya tiene una fila ahí para ese
mismo MZ/LT y período — si la tiene, marcarla como cubierta/anulada en vez de
sumar un crédito nuevo.

## Caso que reabrió esto (2026-07-30)

`I-9` tenía una decisión previa (2026-07-18, `reconciliacion_junio_a_julio.xlsx`)
que dejó su ACUERDOS=50 como deuda REAL, no condonable, tras aplicar un crédito
de junio de 86. La nueva declaración de la secretaria (GRUPO 3) contradice esa
decisión — se cargó igual como abono rezagado nuevo (fila aparte, no se tocó la
de 86), bajo este mismo mecanismo de saldo a favor pendiente de conectar.

## Cómo se cierra este evento

No se cierra — es un patrón recurrente, no un incidente puntual. Mantener esta
sección como referencia permanente mientras la secretaria siga reconciliando así.

---

# LEER ANTES — pagos de junio cargados en julio vía mesa_5.xlsx

## Qué pasó

Varios pagos cobrados en junio (por Wagner Trujillo y Yerald Romero) nunca
llegaron a `pagos_efectivo.xlsx` de junio — quedaron anotados en las mesas de
papel pero no se procesaron a tiempo (ciclo de junio ya cerrado, ver repo
`jass_system - junio`). Se recuperan cargándolos como pagos en efectivo del
ciclo de julio.

## Dónde están cargados

```
C:\Users\wilde\PycharmProjects\jass_system\4_pagos\efectivo\inputs\mesa_5.xlsx
```

Filas 4-12 (cobrador Wagner Trujillo salvo donde se indica):

| Fila | Lote | Usuario | Monto | Motivo de no-carga en junio |
|---|---|---|---|---|
| 4 | T-12 | Samuel Samaritano Romero | 155 | Yape cobrado al vecino, nunca transferido a la cuenta Yape de la JAAS |
| 5 | S-5 | Valerio Porfilio Javier Santiago | 71 | Idem — yape no transferido |
| 6 | D-16 | Esteban Guerrero Chingel | 85 | Idem — yape no transferido |
| 7 | F-9 | Rosa Lucia Coronado Luna | 52 | Anotado en mesa de papel, no llegó al output de junio |
| 8 | C1-17 (Macarlopu, cobrador Yerald Romero) | — | 18.50 | Registrado por error en C1-9 (lote sin deuda) — reasignado a su lote real C1-17, cubre agua |
| 9 | C1-17 (Macarlopu, cobrador Yerald Romero) | — | 200 | Idem — CONCEPTO=tanque |
| 10 | D1-6 | Onita Ponte Eguizabal | 33 | Anotado en mesa de papel, no llegó al output de junio |
| 11 | I-9 | Julia Cardenas Alvarado | 86 | Idem |
| 12 | L-4 | Delia Doris Huamansupa Perez | 58 | Idem |

**No incluido a propósito:** C-39 (Janet Romero Mayo) — es plata duplicada, ya
cobrada por un pago Yape real que canceló su deuda de junio. No se carga.

## Fuente / detalle completo

- `jass_system - junio\4_pagos\efectivo\outputs\reconciliacion_junio_a_julio.xlsx` — expediente completo (los 8 casos, con decisión y motivo por caso).
- `jass_system - junio\LEER_ANTES.md` — mismo evento visto desde el repo junio.
- `jass_system - junio\docs\RETOMAR_junio_cierre_reconciliacion_2026-07-18.md` — handoff original.

## Pendiente relacionado (NO es parte de esta carga)

4 condonaciones de MULTA/CORTE derivadas de esta reconciliación (S-5, D-16,
C1-17, G-12) quedan para agosto — ver `PARA_AGOSTO.md` en la raíz de este
repo. No tocar esas columnas ahora.

---

# LEER ANTES — 6_corte no se entera cuando 4b_reclamos resuelve un reclamo

## Qué pasó

El 26/07 se descubrió (caso E-8 y F1-4, ciclo 2026-07) que un reclamo resuelto
en `4b_reclamos` **nunca llega** al archivo que `6_corte` realmente lee para
bloquear el corte de un usuario. Son dos archivos distintos que no están
conectados:

```
4b_reclamos/trazabilidad/trazabilidad_reclamos.xlsx
   ESTADO_FINAL=RESUELTO  ← acá SÍ queda registrada la resolución
        │
        ✗  nadie copia esto al archivo de abajo
        ▼
6_corte/outputs/resolucion_reclamos_YYYY-MM.xlsx
   ← generar_lista.py lee ESTE archivo (no trazabilidad_reclamos.xlsx)
     para decidir si un reclamo bloquea el corte de un usuario
     → HOY no existe ningún resolucion_reclamos_2026-07.xlsx
```

## Consecuencia concreta

Un usuario cuyo reclamo ya se resolvió (ej. E-8: "está al día, solo debe
consumo de este mes"; F1-4: recibo mal atribuido) **vuelve a aparecer como
elegible para corte** cada vez que se corre `generar_lista.py` de nuevo,
porque el módulo no tiene forma de saber que ya se resolvió.

## Qué hacer mientras tanto

Antes de correr `6_corte/generar_lista.py` (con `--force`, ver candado
agregado el 26/07) en cualquier ciclo con reclamos activos: revisar a mano
`4b_reclamos/trazabilidad/trazabilidad_reclamos.xlsx` filtrando
`MES_CIERRE` = ciclo actual y `ESTADO_FINAL=RESUELTO`, y excluir esos
usuarios manualmente del resultado si vuelven a aparecer como
`EJECUTAR_CORTE=SI`.

## Pendiente — arreglo real (no decidido todavía)

Dos caminos posibles, sin resolver:
1. `4b_reclamos` escribe/actualiza `resolucion_reclamos_YYYY-MM.xlsx` cada vez
   que una fila de `trazabilidad_reclamos.xlsx` pasa a RESUELTO.
2. `6_corte/generar_lista.py` lee directamente `trazabilidad_reclamos.xlsx`
   en vez de (o además de) `resolucion_reclamos_YYYY-MM.xlsx`.

No implementar ninguno sin antes decidir cuál — es un cambio de diseño entre
dos módulos, no un fix de una línea.

## Cómo se cierra este evento

Cuando se implemente el arreglo real (uno de los dos caminos de arriba) y se
confirme que un reclamo resuelto deja de reaparecer en `lista_corte.xlsx`,
borrar esta sección de `LEER_ANTES.md`.

---

# LEER ANTES — C-29A (Eddy Santiago Garcilazo Trujillo) retirado del padrón

## Qué pasó

La secretaria confirmó (cuaderno papel, 28/07) que el lote **C-29A** ya no
existe como predio separado — es el mismo predio físico que **C-34** (Eber
Agüero Trujillo), quien es hoy el ocupante real y activo (tiene lecturas de
medidor este ciclo). C-29A seguía generando una boleta duplicada (S/20 este
ciclo, "NO ESTÁ AL DÍA") a nombre de Eddy, aunque él ya no vive ahí.

## Por qué no se había resuelto solo

`C-29A` ya estaba marcado `SIN_MEDIDOR` en
`1_lecturas/sin_servicio/inputs/lista_sin_servicio.xlsx` desde 2026-06-21 —
pero **ese flag no frena la facturación**, solo evita la alerta interna
`SIN_LECTURA` de `1_lecturas`. El flag que sí frena la facturación es
`SIN_SERVICIO=Si` en `registro_operario_acumulado.xlsx`, que gobierna
únicamente `1_lecturas/aplicar_sincronizacion.py` comparando contra
`0_padron/02_matching/outputs/padron_reconciliado.xlsx`. Como C-29A seguía
existiendo como fila en el padrón, nunca se detectó como
`SIN_SERVICIO` y `2_planilla` lo siguió facturando.

## Decisión final (28/07/2026) — override ELIMINAR, no SIN_SERVICIO

El camino de `SIN_SERVICIO` (vía `aplicar_sincronizacion.py`) se evaluó pero
se descartó por ahora: es correcto pero indirecto (depende de que el ciclo
completo de 1_lecturas/2_planilla vuelva a correr). Se optó por lo mismo que
se decidió para C1-17 (ver sección de abajo): un override `ACCION=ELIMINAR`
en `0_padron/overrides_padron.xlsx`, que borra la fila directo de
`padron_reconciliado.xlsx` (hoja `cobranza`) y se re-aplica solo en cada
regeneración futura. Además, se borró a mano la fila C-29A de
`3_boletas/inputs/DATA_boletas.xlsx` para que no se imprima la boleta ESTE
ciclo (el override del padrón por sí solo no afecta un `DATA_boletas.xlsx`
ya generado).

## Qué se hizo (28/07/2026)

1. Se implementó `ACCION=ELIMINAR` en `0_padron/aplicar_overrides.py`
   (antes solo existía `CORREGIR_CAMPO` y `REASIGNAR_LOTE`).
2. Se agregó fila en `overrides_padron.xlsx`: `C-29A ELIMINAR` (motivo: mismo
   predio que C-34, confirmado por la secretaria).
3. Se corrió `aplicar_overrides.py --mes 2026-08` → borra la fila de
   `padron_reconciliado.xlsx`.
4. Se borró a mano la fila C-29A de `3_boletas/inputs/DATA_boletas.xlsx`
   (este ciclo, para que no se imprima la boleta ya).

## Cómo se cierra este evento

Cuando se confirme que `2_planilla`/`3_boletas` de un ciclo futuro ya no
generan ninguna fila para C-29A (porque no está en el padrón), borrar esta
sección.

---

# LEER ANTES — C1-17 (Roberto Macarlopu) eliminado — mismo predio que C1-9

## Qué pasó

C1-17 y C1-9 son el mismo predio (ver `docs/retomar/RETOMAR_agosto_override_C1-9_y_deuda_directiva.md`,
error de `02_matching` que no detecta mismo titular en 2 lotes). El 27/07/2026
se aplicó un override `CORREGIR_CAMPO NOMBRE` que solo le puso la etiqueta
"CERRADO - VER C1-9 (mismo predio)" a C1-17 — **nunca se borró la fila**, lo
que seguía generando confusión (C1-17 aparecía con boleta propia, S/39, en
`DATA_boletas.xlsx` del ciclo 2026-07/08).

## Qué componía el Total=39 de C1-17 y qué se decidió

```
Total C1-17 = 39 = MES_ANTERIOR (9) + MULTA faena+reunión (30)

MULTA (30)        → condonación YA DECIDIDA en PARA_AGOSTO.md desde el
                     18/07/2026 ("dice haber asistido a faena, la directiva
                     no lo registró, mismo criterio que D-16"), nunca
                     ejecutada. Se ejecuta ahora (28/07/2026).

MES_ANTERIOR (9)  → arrastre ya cubierto por el pago de S/18.50 registrado
                     en junio bajo C1-17 (ver LEER_ANTES.md, sección
                     "pagos de junio cargados en julio vía mesa_5.xlsx",
                     fila 8: "cubre agua"). No es deuda real pendiente —
                     se pone en 0, no se traslada a C1-9.
```

C1-9 (Roberto, activo, Total=8, AL DÍA) no tenía nada real pendiente que
heredar — queda sin cambios.

## Qué se hizo (28/07/2026)

1. Se agregó fila en `overrides_padron.xlsx`: `C1-17 ELIMINAR` (reemplaza en
   la práctica al `CORREGIR_CAMPO NOMBRE` del 27/07, que queda obsoleto —
   no se borra esa fila vieja, se marca `ESTADO=CERRADO` para trazabilidad).
2. Se corrió `aplicar_overrides.py --mes 2026-08` → borra la fila de
   `padron_reconciliado.xlsx`.
3. Se puso en 0 `MULTA (faena+reunión)` y `MES ANTERIOR` de C1-17 en
   `DATA_boletas.xlsx` (documentando la condonación + el arrastre ya
   cubierto) y luego se borró la fila completa (este ciclo).
4. Pendiente relacionado que sigue sin tocar: las otras 3 condonaciones de
   `PARA_AGOSTO.md` (S-5, D-16, G-12) — no son parte de este evento.

## Cómo se cierra este evento

Cuando se confirme que un ciclo futuro de `2_planilla`/`3_boletas` ya no
genera ninguna fila para C1-17, borrar esta sección. Marcar también el
ítem de `PARA_AGOSTO.md` correspondiente a C1-17 como ejecutado.

---

# LEER ANTES — correcciones notas_2026-07 (GRUPO 2) aplicadas a agosto — parche manual de M-12 pendiente de repetir si se regenera planilla

## ⚠ CORRECCIÓN 01/08/2026 — el parche de M-12 estaba en el campo equivocado

El parche original (ver "Qué se hizo" abajo) puso los S/80 de M-12 en
MANTENIMIENTO (3→83). La directiva confirmó que eso está **mal** — los S/80
van en CONVENIO, no en mantenimiento. Corregido en
`3_boletas/inputs/DATA_boletas.xlsx` (01/08/2026): Mantenimiento vuelve a 3,
Convenio queda en 80 (Consumo/Multa sin tocar). Total de la boleta no
cambia (118), solo se reubicó el monto al campo correcto.

**Sigue pendiente:** `2_planilla/outputs/planilla_2026-08.xlsx` y
`shared/planilla_mes/planilla_2026-08.xlsx` todavía tienen el parche viejo
(MANTENIMIENTO=83) — si se vuelven a leer para regenerar `DATA_boletas.xlsx`
o para un ciclo futuro, hay que repetir la corrección ahí también: fila
M-12, poner Mantenimiento=3 y Convenio=80.

## Qué pasó

La secretaria confirmó en `4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx`
(GRUPO 2) que varios predios están al día / ya pagaron MULTA y/o ACUERDOS
(techado, campo, faena, reunión), que P-9 y M-12 necesitaban un cargo nuevo
(sembrado, no estaba), y que G1-3/K-8/G1-14/B1-12 tenían "Verificando"
impreso en la boleta sin necesidad. Todo esto ya se aplicó y se verificó
contra `planilla_2026-08.xlsx` (28/07/2026).

El trabajo se hizo primero en una carpeta separada,
`C:\Users\wilde\PycharmProjects\jass_system - Julio` (snapshot de julio,
para poder re-correr `5_cobranza` apuntando a julio sin tocar el repo activo),
y después se portó a este repo. En el camino se encontraron y corrigieron
2 errores reales — ver detalle abajo.

## Qué se hizo — resumen final

**17 predios con deuda cancelada** (K-8, L-16, N-5, O-16, Q-5, Q-10, Q-11,
R-4, S-1, S-8, S-9, S-12, V-6, B1-12, D1-3, F1-5, G1-14): MULTA/ACUERDOS
llevados a 0 en `shared/seguimiento_pueblo.xlsx` vía `registrar_pago` directo
(`SOURCE=manual`, `AUDIT_REF=notas_2026-07|MZ-LT-CONCEPTO`) — el overlay
`abonos_rezagados.xlsx` por sí solo NO alcanza para tocar MULTA/ACUERDOS
(solo corrige el reporte del mes, no el ledger, salvo que haya plata real de
por medio).

**2 cargos nuevos sembrados:**
- P-9: ACUERDOS 75 (techado+campo) — `CARGO` directo en `seguimiento_pueblo.xlsx`
  + fila en `shared/genesis_tardia.xlsx`. Ya visible en `planilla_2026-08.xlsx`.
- M-12: MANTENIMIENTO +80 — MANTENIMIENTO no tiene ledger ni override por
  predio en `2_planilla` (es un valor fijo para todos), así que se aplicó
  **parche manual directo** en `2_planilla/outputs/planilla_2026-08.xlsx`
  Y en `shared/planilla_mes/planilla_2026-08.xlsx` (celda MANTENIMIENTO de
  M-12, fila con MZ=M/LT=12: 3 → 83). También hay fila en
  `shared/genesis_tardia.xlsx` para que quede documentado, pero esa fila
  por sí sola NO alcanza la boleta (solo se aplicaría si `5_cobranza`
  corre para agosto, no en el archivo impreso).

**4 "Verificando" sacados** de `3_boletas/inputs/pendientes_convenio_multas.xlsx`
(K-8, G1-3, G1-14, B1-12) — ya mostraban el monto real (0 en todos esos casos).

**2 marcados para confirmar, no ejecutados:** P-6 (convenio nuevo de 350,
instalación de Flor Valdivia) y C1-3 (montos con "?", ella misma pidió
verificar) — `TIPO_RECLAMO=CONFIRMACION` en `notas_2026-07.xlsx`, sin tocar
nada más.

## 2 errores encontrados y corregidos en el camino

1. **Q-5**: el monto que se le iba a cancelar incluía por error 20 de
   "MES_ANTERIOR" que en realidad era su consumo real de julio sin pagar
   (la nota decía textual "ponle su consumo de este mes" — no tocar). Ese
   20 se coló a CONVENIO (medidor) vía la cascada del overlay. Revertido en
   `5_cobranza/outputs/arrastre_consolidado_2026-07.xlsx` (CONVENIO 5→25)
   — nunca llegó a tocar el ledger real (`seguimiento_pueblo.xlsx` intacto).
   Monto corregido en `abonos_rezagados.xlsx`: 89→69.
2. **O-16 y Q-11**: sus filas de agua vieja se agregaron a
   `abonos_rezagados.xlsx` DESPUÉS de correr `5_cobranza --force`, así que
   nunca se procesaron por la cascada normal. Se aplicaron a mano, directo
   en `arrastre_consolidado_2026-07.xlsx` (DEUDA_AGUA a 0 en ambas filas).

## ⚠ Si se vuelve a correr 2_planilla para agosto

El parche manual de M-12 (MANTENIMIENTO 83) **se pierde** si `2_planilla`
regenera `planilla_2026-08.xlsx` de nuevo — no hay preservación automática
para este campo. Repetir el parche: fila M-12, columna MANTENIMIENTO,
poner 83 (no sumar 80 de nuevo si ya dice 83).

## Verificado (28/07/2026)

`planilla_2026-08.xlsx` revisado predio por predio para los 19 casos — todo
cuadra: Q-5 con su consumo real intacto (agua 20, convenio 25) y
multa/acuerdos en 0; el resto de los 17 en 0 donde correspondía; K-8 con su
ACUERDOS=30 sin tocar (nunca se pidió); P-9 con ACUERDOS=75; M-12 con
MANTENIMIENTO=83.

## Cómo se cierra este evento

Cuando `3_boletas` imprima agosto y se confirme que salió con estos valores
(sin necesidad de reimprimir nada), borrar esta sección.

---

# LEER ANTES — Q-13/Q-16 y S-13/S-14 unificados (mismo predio, 2 lotes) — traspaso de lectura en Q-13

## Qué pasó

La secretaria confirmó en `notas_2026-07.xlsx` (GRUPO 2, ORDEN 2/3 y 5/6) que
Q-13/Q-16 son el mismo lote (Judith Margot Luis Peña) y que S-13/S-14 también
(Ana Garro Rojas / Ana Isabel Garro Rojas). Mismo patrón que C1-9/C1-17: un
predio quedó duplicado en 2 códigos de lote por error de `02_matching`.

## Qué se hizo (29/07/2026)

1. 3 filas nuevas en `0_padron/overrides_padron.xlsx`: `Q-13 CORREGIR_CAMPO
   NOMBRE→JUDITH MARGOT LUIS PEÑA`, `Q-16 ELIMINAR`, `S-14 ELIMINAR`.
2. Corrido `aplicar_overrides.py --mes 2026-08` → Q-16 y S-14 ya no existen en
   `padron_reconciliado.xlsx`, Q-13 con el nombre corregido, S-13 sin cambios.
3. Q-13, Q-16, S-13, S-14 marcadas RESUELTO en `notas_2026-07.xlsx`.

## Traspaso de lectura — mismo problema que C1-9/C1-17, resuelto distinto en cada par

`registro_operario_acumulado.xlsx` (que gobierna `1_lecturas`) tiene una
serie de lecturas por cada lote, independiente del padrón. Al eliminar un
lote duplicado hay que fijarse **cuál de los dos tenía la serie real del
medidor** — no siempre es el que sobrevive:

```
Q-13 (sobrevive) → solo Jun=10 (M3=5), Jul/Ago vacío -- lectura suelta, NO es
                    el medidor real
Q-16 (eliminado) → May=165, Jun=172, Jul=178, Ago=184 -- serie REAL continua
   → SE TRASPASÓ: Q-13 fila 341, columna 2026-08 MARCACION puesta en 184
     (sin M3 -- ese consumo de agosto ya se facturó bajo Q-16 este ciclo).
     Así setiembre calcula el consumo real (lectura nueva − 184), no un
     salto falso contra el "10" viejo de Q-13.

S-13 (sobrevive) → May=4516, Jun=4534, Jul=4546, Ago=4553 -- serie REAL, ya
                    correcta
S-14 (eliminado) → solo Jun=10 (M3=5) -- lectura suelta
   → NO necesitó traspaso, S-13 ya tenía la serie correcta.
```

Precedente: C1-9/C1-17 tuvo el mismo caso (C1-17 tenía la serie real
May-Jul: 411→418→424; C1-9 la continuó en agosto con 429 = 424+5). Ahí el
traspaso ocurrió solo porque el operario ya tomó la lectura de agosto bajo el
código correcto; acá se anticipó a mano porque agosto ya está cerrado para
Q-16.

**Pendiente:** `NOMBRE` de Q-13 en `registro_operario_acumulado.xlsx` sigue
diciendo "YONY CELEDONIO CERVANTES" (no se tocó a mano) — se corrige solo
cuando `1_lecturas/proponer_sincronizacion.py` corra y detecte el delta
RENAME contra `padron_reconciliado.xlsx`. Igual para el delta SIN_SERVICIO
de Q-16/S-14 (siguen como filas en `registro_operario_acumulado.xlsx`,
pendientes de que el sync las marque).

## Cómo se cierra este evento

Cuando `1_lecturas` corra su próxima sincronización (ciclo setiembre) y
confirme el RENAME de Q-13 + los SIN_SERVICIO de Q-16/S-14 aplicados, borrar
esta sección.

---

# LEER ANTES — regeneración de boletas de agosto (3_boletas) EN CURSO, cortada a media noche del 30/07 — retomar acá mañana

## Estado exacto al cortar (30/07/2026, ~20:40)

`3_boletas/outputs/` está **INCOMPLETO**. Se borraron TODOS los archivos
`RECIBO_*.pdf`/`.docx` del rango de agosto (17970-18540, ~1132 archivos) por el
motivo explicado abajo, y la regeneración con `py main.py` se cortó en
**~45 de 560 recibos generados**. Antes de imprimir o entregar nada de agosto,
hay que terminar de correr esto.

**Primer paso al retomar:**
```
cd 3_boletas
set PYTHONIOENCODING=utf-8 && py main.py
```
Es **idempotente** — salta los recibos que ya existen y solo genera los que
faltan. Va a cortarse solo varias veces por dos motivos conocidos (ver
"Problemas conocidos" abajo) — simplemente volver a correr el mismo comando
hasta que el resumen final diga `[OK] Boletas generadas en PDF correctamente`
y `561/561` (o el número que corresponda) recibos. Recién ahí correr:
```
py validar_boletas.py
```
para confirmar. Los que salgan `[ERROR]` con "esperado X" en un concepto
donde `3_boletas/inputs/pendientes_convenio_multas.xlsx` todavía dice
`VERIFICANDO` **no son errores reales** — es el validador que no sabe del
mecanismo "Verificando" (ver sección de abajo, lista de los 11 predios que
siguen así a propósito).

## Por qué se borró todo y se está regenerando de cero (dos rondas, dos motivos distintos)

**Ronda 1 (más temprano, 30/07 tarde):** los recibos habían quedado con
datos del 27/07 (3 días viejos) porque `main.py` solo comprueba si el
archivo YA EXISTE para saltearlo — no si los datos cambiaron. Entre el 27/07
y hoy se re-corrió julio con arrastre fresco, se sembró `genesis_tardia`
para 6 predios, y se pasaron 27 predios de VERIFICANDO a RESUELTO en
`pendientes_convenio_multas.xlsx` — nada de eso se reflejaba en los PDFs
viejos. Se borraron 1756 archivos y se regeneró completo. Validado
561/561 correctos (los 11 pendientes eran VERIFICANDO legítimo).

**Ronda 2 (esta noche):** el usuario pidió pasar los 17 predios que
seguían VERIFICANDO a CONFIRMACION (ver sección siguiente) — eso reveló
que **Q-4 tenía una condonación rota** (ver sección "Q-4" más abajo) que
había que corregir. Al corregir Q-4 (eliminarlo de `planilla_2026-08.xlsx`)
se corrió la numeración de recibo de TODOS los predios alfabéticamente
posteriores (~260 predios, de Q en adelante) — sus archivos viejos quedan
con un número de recibo que ya no les corresponde. Se borró todo el rango
otra vez (1132 archivos, pdf+docx) y se está regenerando desde cero.

## Los 17 predios pasados de VERIFICANDO a CONFIRMACION hoy (30/07, distinto del lote de 28/07)

E-12, C-20, P-8, G-12, G-17, B-6, W-5, D-5, Q-3, F1-10, D-1, Z-13, Q-1,
B-11, E-14B, H-13, F1-1 — actualizados en `notas_2026-07.xlsx`
(`TIPO_RECLAMO=CONFIRMACION`) y en `pendientes_convenio_multas.xlsx`
(`ESTADO=RESUELTO`), **sin verificación adicional de la secretaria** —
el usuario pidió explícitamente pasarlos todos así, avisado del riesgo
(el monto que imprime cada uno es el que ya estaba en el sistema, no uno
re-confirmado). El reporte de historial (`4b_reclamos/reporte_historico.py`)
ya se regeneró con estos 13 (de los 17, 4 no tienen saldo>0 en
`seguimiento_pueblo` — E-12, E-14B, H-13, F1-1 — no salen en el reporte
porque no tienen deuda de MULTA/ACUERDOS/CONVENIO que mostrar ahí, su
reclamo era de otro concepto como MES_ANTERIOR).

**11 predios siguen genuinamente VERIFICANDO** (no se tocaron, tienen
reclamo real sin confirmar): B-6, B-11, D-5, G-12, G-17, H-13, P-8, Q-1,
W-5, Z-13, F1-10. El motivo puntual de cada uno está en la columna
`MOTIVO` de `pendientes_convenio_multas.xlsx`.

## Q-4 — condonación rota + override nunca aplicado a la planilla (corregido hoy)

Q-4 (Beatriz Romero Valladares) está unificado con Q-5 desde el 30/07 —
override `ELIMINAR` en `overrides_padron.xlsx`, MES_VIGENCIA=2026-08,
ACTIVO, aplicado correctamente a `padron_reconciliado.xlsx`. Pero:

1. La condonación original de su MULTA(50)+ACUERDOS(75) se hizo con
   `registrar_pago` (no `registrar_ajuste`) — el mismo bug de siempre: la
   corrida de `5_cobranza --force` de julio (hoy, 12:03) la revirtió sola.
2. Alguien "arregló" eso a las 13:21 pero en la dirección equivocada:
   restauró la deuda original completa en vez de re-condonar.
3. Aparte, el override `ELIMINAR FILA` sí sacó a Q-4 de
   `padron_reconciliado.xlsx`, pero **`planilla_2026-08.xlsx` seguía
   arrastrando a Q-4 con MULTA=50+ACUERDOS=75** porque esa cifra viene de
   `arrastre_consolidado_2026-07.xlsx` (salida de `5_cobranza`, no del
   padrón) — el override nunca toca ese archivo.

**Arreglado hoy:**
- `registrar_ajuste` (no PAGO) de -50 MULTA y -75 ACUERDOS en
  `seguimiento_pueblo.xlsx` (AUDIT_REF `Q-4-...-recondonacion-definitiva-30072026`)
  — con AJUSTE no se revierte en la próxima reconciliación.
- Fila Q-4 borrada a mano de `shared/planilla_mes/planilla_2026-08.xlsx`
  **y** de `2_planilla/outputs/planilla_2026-08.xlsx` (son dos copias
  distintas, no el mismo archivo — hay que tocar las dos).
- `3_boletas/enriquecimiento/main.py` re-corrido → `DATA_boletas.xlsx`
  regenerado sin Q-4 (560 recibos, antes 563).
- Esto disparó la Ronda 2 de regeneración completa de `3_boletas/outputs`
  (ver arriba) por el corrimiento de numeración.

**Si en algún ciclo futuro Q-4 reaparece con deuda:** revisar primero si
el override `ELIMINAR` (fila Q-4 en `overrides_padron.xlsx`) sigue activo
y si `2_planilla` se corrió DESPUÉS de aplicar cualquier cambio al padrón
— `arrastre_consolidado` de 5_cobranza no sabe nada del padrón ni de los
overrides, así que un predio eliminado del padrón puede seguir arrastrando
deuda vieja si `5_cobranza`/`2_planilla` no se vuelven a correr en orden.

## Problemas conocidos que van a cortar la corrida de `main.py` (ya pasó varias veces hoy)

1. **Encoding de consola**: sin `PYTHONIOENCODING=utf-8`, el script revienta
   al imprimir un carácter "═" — pasa ANTES de escribir nada, no pierde
   trabajo, solo hay que agregar la variable de entorno.
2. **Word (COM automation) se cuelga o queda con un archivo bloqueado**:
   el script usa Word para convertir cada `.docx` a `.pdf`. Si se corta la
   corrida (timeout, dos corridas en paralelo por error), puede quedar un
   proceso `WINWORD.exe` zombie con un archivo `.docx` bloqueado
   (`PermissionError` al guardar). Cerrar el proceso a mano:
   ```
   Stop-Process -Name WINWORD -Force
   ```
   y volver a correr `main.py` (el archivo bloqueado puede necesitar
   borrarse a mano primero si el error ya ocurrió, `rm` en Git Bash da
   "Device or resource busy" si sigue bloqueado — cerrar Word primero).
3. **NO correr dos instancias de `main.py` en paralelo** — fue exactamente
   lo que causó el problema del punto 2 esta noche (se lanzó una corrida
   nueva antes de confirmar que la anterior había terminado/cortado).

## Qué falta después de que 3_boletas termine

1. `py validar_boletas.py` → confirmar que da limpio salvo los 11
   VERIFICANDO esperados.
2. **Nada se imprimió físicamente todavía** en toda la sesión — el usuario
   no ha dado la orden final de imprimir, solo de generar/validar los PDF.
3. Si se confirma o descarta alguno de los 11 predios en VERIFICANDO,
   repetir el mismo patrón de hoy (notas_2026-07.xlsx TIPO_RECLAMO,
   pendientes_convenio_multas.xlsx ESTADO, borrar su recibo viejo,
   `main.py` de nuevo) — ojo que si alguno de esos 11 se elimina/reasigna
   como pasó con Q-4, va a volver a correr la numeración y hay que repetir
   la limpieza completa del rango.

## Cómo se cierra este evento

Cuando `3_boletas/outputs/` tenga las 560+ boletas de agosto generadas y
validadas, y el usuario confirme que ya se imprimieron/entregaron, borrar
esta sección.
