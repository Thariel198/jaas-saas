# Reclamos 01/08/2026 — parches cosmeticos, pendientes de reconciliar el mes que viene

Boletas cobradas hoy (01/08/2026) con datos corregidos a mano por pedido de la
directiva, sin pasar por el flujo normal de reclamos (no había tiempo). Cada
corrección de acá solo tocó `3_boletas/inputs/DATA_boletas.xlsx` (impresión),
**no el ledger real** (`shared/seguimiento_pueblo.xlsx` / `arrastre_consolidado`).
El mes que viene hay que revisar esta lista y aplicar cada corrección en el
ledger real (`registrar_ajuste` / `registrar_pago` / reasignación según
corresponda), predio por predio.

Detalle completo de las 26 fotos (transcripción + estado) en
`3_boletas/inputs/reclamos_2026-08-01.xlsx`. Acá solo quedan las que YA se
confirmaron y parcharon.

Backup de `DATA_boletas.xlsx` antes de cualquier parche:
`3_boletas/inputs/backups/DATA_boletas_pre_correccion_20260801.xlsx`.

---

# CHECKLIST — estado de cada corrección en el registro real

**Actualizado: 2026-08-03.** Una fila por predio+concepto: **37 correcciones sobre
28 predios** (un predio puede tener 2-3 conceptos tocados). La verdad de qué se
parchó salió del diff `DATA_boletas_pre_correccion_20260801.xlsx` ⟷
`DATA_boletas.xlsx`, no de la lista de fotos (26 fotos, 12 quedaron en
"VERIFICAR, no cambiar aún").

**Cómo se marca `[x]`:** no a mano — se verifica contra el saldo real de
`shared/seguimiento_pueblo.xlsx` (conceptos de pueblo) o contra la existencia de
la fila en el precursor (agua/corte). Si el saldo de hoy no coincide con el
objetivo, la fila queda `[ ]`.

**Regla que se está aplicando** (ver `docs/diario/2026-08-03_solucion_precursor_mas_ledger.html`):
siempre se escriben **los dos** — el precursor guarda la historia, el ledger hace
el efecto, atados por `AUDIT_REF`. Para conceptos de pueblo el precursor va con
`MES_ANO_APLICA` **vacío** (no lo aplica ninguna corrida de `5_cobranza`); para
agua/corte va con `MES_ANO_APLICA` lleno, porque ahí no existe ledger y el
precursor es el que actúa.

| ✓ | grupo | predio | concepto | saldo hoy | objetivo | precursor | nota |
|---|---|---|---|---|---|---|---|
| [x] | EXONERACION | J-1 | MULTA | 0 | 0 | ajustes_cargo | directiva 01/08 — faena y reunion |
| [x] | EXONERACION | J-1 | ACUERDOS | 0 | 0 | ajustes_cargo | directiva 01/08 — techado y campo |
| [x] | EXONERACION | O-2 | MULTA | 0 | 0 | ajustes_cargo | directiva 01/08 |
| [x] | EXONERACION | V-14 | CORTE_RECONEXION | n/a | 0 | ajustes_cargo | secretaria 01/08 — agua, el precursor hace el efecto |
| [ ] | YA_PAGO | K-9 | CONVENIO | 75 | 0 | — | S/100 efectivo en el local — exceso 25 sin resolver |
| [ ] | YA_PAGO | T-14 | CONVENIO | 50 | 0 | — | S/75 efectivo — exceso 25 sin resolver |
| [ ] | YA_PAGO | K-8 | ACUERDOS | 30 | 0 | — | pago todo campo en junio (secretaria) |
| [ ] | YA_PAGO | F-10 | ACUERDOS | 50 | 0 | — | mostro recibos a la secretaria |
| [ ] | YA_PAGO | F-1 | MULTA | 20 | 0 | — | pago la multa el mes pasado |
| [ ] | YA_PAGO | F-7 | ACUERDOS | 25 | 0 | — | yape antiguo no anotado |
| [ ] | YA_PAGO | D-6 | ACUERDOS | 50 | 0 | — | yapeo a Wagner Trujillo hace anos |
| [ ] | YA_PAGO | B-8 | ACUERDOS | 5 | 0 | — | al dia (secretaria) |
| [ ] | YA_PAGO | B-8 | CONVENIO | 75 | 0 | — | al dia (secretaria) |
| [ ] | YA_PAGO | B-8 | MES_ANTERIOR | n/a | 0 | — | agua — va a ajustes_cargo, no al ledger |
| [x] | REASIGNACION | G-4 | CONVENIO | 0 | 0 | abonos_rezagados + reasignaciones_aplicacion | S/50 por fuera (ABONO_REZAGADO, suma a caja) + 25 desde MULTA |
| [x] | REASIGNACION | G-4 | MULTA | 50 | 50 | reasignaciones_aplicacion | recibe los 25 que salen de CONVENIO |
| [ ] | REASIGNACION | G-14 | CONVENIO | 38 | 0 | — | 25 desde MULTA + S/50 efectivo a la secretaria |
| [ ] | REASIGNACION | G-14 | MULTA | 0 | 50 | — | origen NO explicado por la directiva |
| [ ] | REASIGNACION | G-14 | ACUERDOS | 21 | 50 | — | origen NO explicado (21 -> 50) |
| [x] | CARGO_NUEVO | M-12 | CONVENIO | 80 | 80 | genesis_tardia | ⚠ NO es de este expediente — ver nota abajo |
| [x] | REASIGNACION | E-14A | ACUERDOS | 0 | 0 | genesis_tardia | lote fantasma: no existe en el padron — cargo movido a E-14B |
| [x] | REASIGNACION | E-14B | ACUERDOS | 75 | 75 | reidentificacion_cargo | recibe el cargo de E-14A (Juan Saavedra). NO era cargo nuevo |
| [x] | BUG_SIGNO | A-8 | CONVENIO | 50 | 50 | — | pago fantasma 06/07 + reversion con signo invertido |
| [x] | BUG_SIGNO | B-5 | ACUERDOS | 25 | 25 | — | idem — ver nota de Pompeyo abajo |
| [x] | BUG_SIGNO | B-5 | CONVENIO | 50 | 50 | — | idem — ver nota de Pompeyo abajo |
| [x] | BUG_SIGNO | C-1 | ACUERDOS | 25 | 25 | — | idem |
| [x] | BUG_SIGNO | C-1 | CONVENIO | 50 | 50 | — | idem |
| [x] | BUG_SIGNO | C-7 | CONVENIO | 25 | 25 | — | idem |
| [x] | BUG_SIGNO | E-12 | CONVENIO | 26 | 26 | — | idem |
| [x] | BUG_SIGNO | I-11 | CONVENIO | 25 | 25 | — | idem |
| [x] | BUG_SIGNO | I-16 | MULTA | 18 | 18 | — | idem |
| [x] | BUG_SIGNO | I-16 | ACUERDOS | 75 | 75 | — | no era negativo (47) pero igual incorrecto |
| [x] | BUG_SIGNO | J-3 | CONVENIO | 50 | 50 | — | idem |
| [x] | BUG_SIGNO | K-17 | CONVENIO | 25 | 25 | — | idem |
| [x] | BUG_SIGNO | K-2 | CONVENIO | 25 | 25 | — | idem |
| [x] | BUG_SIGNO | P-12 | CONVENIO | 50 | 50 | — | idem |
| [x] | BUG_SIGNO | H-16 | ACUERDOS | 75 | 75 | — | mismo bug, se escapo del filtro (47 no es negativo) |

**TOTAL 37 · hechas 24 · pendientes 13** (2026-08-03, tras cerrar G-4)

## ⚠ E-14B no era una contradicción — era un lote fantasma

El 30/07 se resolvió "E-14B ya pagó" y el 01/08 la directiva pidió cargarle 75.
Las dos cosas eran ciertas: hablaban de **códigos de lote distintos**.

```
   DEUDORES Y PAGOS DEL TECHADO Y CAMPO (hoja Corregido)
        JUAN SAAVEDRA SAAVEDRA  →  E · 14A     ← la lista dice 14A
   padron_reconciliado
        JUAN SAAVEDRA SAAVEDRA  →  E · 14 B    ← el padrón dice 14 B (con espacio)

   ⇒ el CARGO de 75 se sembró en E-14A, lote que NO EXISTE en el padrón
   ⇒ E-14A imprimió boleta SIN NOMBRE (recibo 18083, S/75)
   ⇒ E-14B quedó sin cargo → por eso el 30/07 figuraba en cero
   ⇒ cargar 75 a E-14B sin sacar E-14A habría cobrado los mismos 75 DOS VECES
```

**Trampa adicional encontrada:** `E · 14 B` era el **único** predio de ~570 con un
espacio en el LT, y los módulos no lo normalizan igual —
`seguimiento_repo`/`5_cobranza` borran el espacio (`14B`),
`2_planilla`/`3_boletas`/`0_padron` lo conservan (`14 B`). Un cargo escrito en el
ledger habría quedado **invisible** para la planilla, que además habría generado
otra fila fantasma. Canonizado a mano el 03/08 (`14 B` → `14B`) en
`padron_reconciliado`, `registro_operario_acumulado` y las dos copias de
`planilla_2026-08`. **`DATA_boletas.xlsx` se dejó con `14 B` a propósito**: es la
referencia de lo que se imprimió y se cobró.

Deuda pendiente que esto deja abierta: hay ~29 normalizadores de LT en el
pipeline con al menos dos comportamientos opuestos para espacios. Se corrigió el
dato (1 predio), no el código. Va a volver con el próximo lote con letra o espacio.

## ⚠ M-12 no pertenece a este expediente

Está en el checklist porque se parchó en `DATA_boletas.xlsx` el mismo 01/08 y
aparece en el diff, pero **no lo dictó la directiva ese día**: viene de
`notas_2026-07.xlsx` (GRUPO 2) y su corrección está documentada en
`LEER_ANTES.md` (§ "correcciones notas_2026-07"). Los S/80 se habían sembrado en
MANTENIMIENTO el 28/07 y la directiva confirmó el 01/08 que van en CONVENIO.
Hecho el 03/08: fila de `genesis_tardia` apagada (`MES_ANO_APLICA` → vacío, nunca
se había aplicado) + fila nueva de CONVENIO inerte + `registrar_cargo` en el ledger.

## BUG_SIGNO (15) — HECHO el 2026-08-03, pero el codigo sigue roto

Los 15 saldos se restauraron al CARGO real (S/914 de deuda que vuelve a existir —
no es plata que entra). **Cero saldos negativos en todo el ledger.**
`CLASE=CORRECCION_SISTEMA`, `source=manual`.

La restauracion es estable: `ajuste_reconciliado` filtra por `SOURCE`, asi que un
AJUSTE con `source=manual` no entra en el `ya` de la reconciliacion y una
re-corrida de julio no lo deshace.

**Lo que NO se arreglo:** `5_cobranza/main.py:2320` sigue emitiendo el ajuste de
reversion con el signo invertido, porque la columna `AJUSTE` carga dos
convenciones opuestas — `ajuste_reconciliado` (:2313) la lee como CREDITO,
`_registrar` (`seguimiento_repo:280`) la aplica como DEUDA. El bug solo amenaza
ciclos futuros (julio esta congelado), pero va a volver a fabricar negativos
cuando corra `5_cobranza` de agosto. Dos caminos, ninguno decidido:

```
 ① _registrar deja AJUSTE con signo de CREDITO (saldo -= monto)
      + coherente con PAGO y con ajuste_reconciliado
      - invierte el sentido de los AJUSTE manuales ya escritos
        (los -20/-75/-30 de exoneracion pasarian a SUMAR deuda)

 ② ajuste_reconciliado invierte el signo al acumular, _registrar no cambia
      + no toca ninguna fila de las 1.584 existentes
      - la correccion queda escondida en la funcion de idempotencia
```

## Qué bloquea a los 13 pendientes

```
YA_PAGO (10)      la BUSQUEDA ya se hizo (03/08, contra los ciclos congelados
                  en 7_cierre/archivo/). Resultado: NINGUNO de los 8 casos
                  tiene un pago que coincida con el monto declarado; los pagos
                  que aparecen son de agua del ciclo, ya aplicados.
                  PERO 4 tienen nota de mesa contemporanea que corrobora:
                    F-7  "Cancelo total"                (2026-07, Wagner)
                    B-8  "Reclamo. Ya pague mes anterior, techado y ca..."
                    F-10 "Multa exonerada, faena permanente"
                    T-14 "Mes pasado pago 100 restaban 28..."
                    K-9  "Pago medidor total"
                  Sin corroborar: K-8 · D-6 · F-1.
                  Propuesta: los 9 de pueblo con CLASE=DECLARACION (no suman a
                  caja: la plata ya entro y ya se conto como exceso sin
                  atribuir) + B-8 MES_ANTERIOR a ajustes_cargo.
                  FALTA DECIDIR: si se le pregunta a Wagner por D-6 antes.

REASIGNACION (3)  G-4 CERRADO el 03/08 (los S/50 por fuera entraron como
                  ABONO_REZAGADO -- plata real que nunca se registro -- y los
                  25 como reasignacion MULTA->CONVENIO).
                  Queda solo G-14 (3 filas): la directiva NO explico el origen
                  de MULTA 0->50 ni de ACUERDOS 21->50, solo el resultado final.
                  Confirmar con ella antes de escribir. Su ledger sigue intacto
                  (CONVENIO 38 - MULTA 0 - ACUERDOS 21, ultimas filas del 04-08/07).
```

---

## J-1 — COMEDOR POPULAR CLUB DE MADRES (recibo 18170)

**Motivo (dictado por la directiva, 01/08/2026):** exonerar MULTA (faena y
reunión) y TECHADO Y CAMPO (acuerdos).

**Valores originales impresos:** Convenio=75, Multa=20, Techado y campo=75,
Total=216.

**Valores corregidos (boleta de hoy):** Multa=0, Techado y campo=0,
Convenio=75 (sin tocar), Total=121.

**Pendiente:** CONVENIO (75) sigue con la nota de la foto original
("revisar convenio, se exoneró 4") sin resolver — no se tocó, falta
confirmar con la directiva. También quedó mencionado un posible cambio de
nombre a "Graciela Bastidas" en la foto — no aplicado, confirmar si
corresponde. El mes que viene: `registrar_ajuste` en `seguimiento_pueblo.xlsx`
para MULTA -20 y ACUERDOS -75 de J-1 (`source=manual`, motivo="exonerado por
la directiva 01/08/2026").

---

## K-9 — FORTUNATO VARGAS CABELLO (recibo 18192)

**Motivo (dictado por la directiva, 01/08/2026):** ya pagó CONVENIO (medidor)
en efectivo, S/100, en el local.

**Valores originales impresos:** Convenio=75, Total=87.

**Valores corregidos (boleta de hoy):** Convenio=0, Total=12.

**Pendiente:** el pago fue S/100 en efectivo pero la deuda de CONVENIO era
75 — confirmar dónde quedan los S/25 de exceso (¿otro concepto, a favor, o
el monto real era distinto?). El mes que viene: buscar este pago de S/100
en `4_pagos/efectivo/inputs/mesa_*.xlsx` o `pagos_efectivo.xlsx` de julio (si
no aparece, cargarlo como `abonos_rezagados`) y `registrar_ajuste`/
`registrar_pago` de CONVENIO en `seguimiento_pueblo.xlsx` para K-9.

---

## K-8 — VICTOR TEODORO FLORES DURAND (recibo 18191)

**Motivo (confirmado por la secretaria, 01/08/2026):** en junio ya pagó todo
campo — borrar la deuda de TECHADO Y CAMPO.

**Valores originales impresos:** Techado y campo=30, Total=45.

**Valores corregidos (boleta de hoy):** Techado y campo=0, Total=15.

**⚠ Nota — la decisión del 28/07 nunca se aplicó al ledger real:**
`notas_2026-07.xlsx` lista a K-8 entre los 17 predios con deuda cancelada el
28/07/2026, pero verifiqué `seguimiento_pueblo.xlsx` y ACUERDOS de K-8 sigue
con **SALDO=30** (CARGO 50, PAGO 20 en junio, nunca se canceló el resto) —
no es solo una boleta desactualizada, falta el `registrar_pago`/
`registrar_ajuste` real. El mes que viene: `registrar_ajuste` ACUERDOS -30
en `seguimiento_pueblo.xlsx` para K-8 (`source=manual`, motivo="pagó todo
campo en junio, confirmado por la secretaria 01/08/2026").

---

## T-14 — PEDRO CANDACHO HUARAC (recibo 18342)

**Motivo (confirmado por la secretaria, 01/08/2026):** no debe CONVENIO, ya
pagó S/75 en efectivo.

**Valores originales impresos:** Convenio=50, Total=187.

**Valores corregidos (boleta de hoy):** Convenio=0, Total=137.

**Verificado en el ledger:** `seguimiento_pueblo.xlsx` tiene CONVENIO de
T-14 con CARGO=50, sin ningún PAGO ni AJUSTE registrado — coincide con la
boleta (no es un caso ya resuelto que quedó desactualizado, es una
corrección nueva).

**Pendiente:** el pago fue S/75 pero la deuda de CONVENIO era 50 — igual que
K-9, confirmar dónde quedan los S/25 de exceso. El mes que viene: buscar
este pago de S/75 en efectivo de julio (mesa_*.xlsx / pagos_efectivo.xlsx —
no encontrado en la investigación de hoy) y `registrar_pago`/
`registrar_ajuste` CONVENIO en `seguimiento_pueblo.xlsx` para T-14.

---

## V-14 — LEONARDO HUAMANI SOTELO (recibo 18368)

**Motivo (pedido por la secretaria, 01/08/2026):** borrar CORTE Y
RECONEXIÓN — no corresponde.

**Valores originales impresos:** Corte y reconexión=40, Total=84.

**Valores corregidos (boleta de hoy):** Corte y reconexión=0, Total=44.

**Pendiente:** Corte y reconexión no vive en `seguimiento_pueblo.xlsx` (ese
solo trackea MULTA/ACUERDOS/CONVENIO) — es de `5_cobranza`/
`arrastre_consolidado_2026-07.xlsx` y `6_corte/outputs/audit_penalidad.xlsx`.
El mes que viene: revisar `audit_penalidad.xlsx` por si V-14 salió en una
lista de corte por error (mismo patrón que F1-4, ver `shared/ajustes_cargo.xlsx`)
y aplicar la corrección real ahí, no solo en la boleta.

---

## B-8 — ROSALINA OLIMPIA CIRIACO SOTELO (recibo 17992)

**Motivo (confirmado por la secretaria, 01/08/2026):** está al día en
TECHADO Y CAMPO, y CONVENIO (75) sí lo pagó.

**Valores originales impresos:** Convenio=75, Techado y campo=5, Total=150.

**Valores corregidos (boleta de hoy):** Convenio=0, Techado y campo=0,
Total=70.

**Verificado en el ledger:** `seguimiento_pueblo.xlsx` tiene ACUERDOS
SALDO=5 y CONVENIO SALDO=75 — coincide con la boleta, no es un caso
desactualizado, es corrección nueva.

**Actualización 01/08/2026 — MES_ANTERIOR también corregido:** el usuario
atendió personalmente a B-8 (un chico vino a pagar) y confirmó que no debe
MES_ANTERIOR — corregido de 46 a 0. Total final: 24 (antes 70). Esto
coincide con la nota ya existente de MES_ANTERIOR resuelta el 30/07/2026 en
`notas_2026-07.xlsx`, que había quedado desactualizada en el arrastre — acá
se corrigió también en la boleta.

**Pendiente:** el mes que viene, en `seguimiento_pueblo.xlsx`: `registrar_ajuste`
ACUERDOS -5 y CONVENIO -75 para B-8 (`source=manual`, motivo="confirmado al
día por la secretaria 01/08/2026"); MES_ANTERIOR (46→0) vive en
`arrastre_consolidado`/`5_cobranza`, no en `seguimiento_pueblo` — corregir
ahí también (mismo patrón que W-4/R-5 en `LEER_ANTES.md`).

---

## F-10 — HERMINIO LUCERO TRUJILLO (recibo 18096)

**Motivo (dictado por la directiva, 01/08/2026):** dice que pagó el mes
pasado, le mostró los recibos a la secretaria.

**Valores originales impresos:** Techado y campo=50, Total=80.

**Valores corregidos (boleta de hoy):** Techado y campo=0, Total=30.

**Verificado en el ledger:** `seguimiento_pueblo.xlsx` tiene ACUERDOS
SALDO=50 (CARGO 75, PAGO 25 en junio) — coincide con la boleta, corrección
nueva. El mes que viene: `registrar_ajuste` ACUERDOS -50 en
`seguimiento_pueblo.xlsx` para F-10 (`source=manual`, motivo="mostró
recibos a la secretaria, confirmado pagado 01/08/2026").

---

## F-1 — MARIA GODO SIFUENTES (recibo 18086)

**Motivo (confirmado por la secretaria, 01/08/2026):** pagó su MULTA el mes
pasado.

**Valores originales impresos:** Multa=20, Total=37.

**Valores corregidos (boleta de hoy):** Multa=0, Total=17.

**Verificado en el ledger:** `seguimiento_pueblo.xlsx` tiene MULTA SALDO=20
(CARGO 20, sin ningún PAGO registrado) — coincide con la boleta, corrección
nueva. El mes que viene: `registrar_ajuste` MULTA -20 en
`seguimiento_pueblo.xlsx` para F-1 (`source=manual`, motivo="pagó multa el
mes pasado, confirmado por la secretaria 01/08/2026").

---

## F-7 — VICTOR LAURENCIO VALLADARES (recibo 18093)

**Motivo (confirmado por la secretaria, 01/08/2026):** está al día en
TECHADO Y CAMPO — se le olvidó anotar un pago yape de hace años.

**Valores originales impresos:** Techado y campo=25, Total=126.

**Valores corregidos (boleta de hoy):** Techado y campo=0, Total=101.

**⚠ Nota — igual que K-8, la decisión previa nunca se aplicó al ledger:**
esto ya estaba `RESUELTO` en `notas_2026-07.xlsx` desde el 30/07/2026
("no debe techado ni campo -- sacar esos S/25"), pero `seguimiento_pueblo.xlsx`
sigue con ACUERDOS SALDO=25 (CARGO 50, PAGO 25 en junio, nunca se canceló el
resto) — no se aplicó el `registrar_pago`/`registrar_ajuste` real en su
momento. El mes que viene: `registrar_ajuste` ACUERDOS -25 en
`seguimiento_pueblo.xlsx` para F-7 (`source=manual`, motivo="pago yape
antiguo no anotado, confirmado por la secretaria").

---

## D-6 — HERMELINDA JARA TRUJILLO (recibo 18056)

**Motivo (dictado por la directiva, 01/08/2026):** dice que yapeó hace años
a Wagner Trujillo (cobrador) — solo el campo de techado y campo (50).

**Valores originales impresos:** Techado y campo=50, Convenio=25 (sin
tocar), Total=239.

**Valores corregidos (boleta de hoy):** Techado y campo=0, Total=189.

**Verificado en el ledger:** `seguimiento_pueblo.xlsx` tiene ACUERDOS
SALDO=50 (CARGO 50, sin ningún PAGO) — coincide con la boleta, corrección
nueva. Convenio (25, con AJUSTE de la corrección de fórmula de abril) no se
tocó — la nota original de la foto era ambigua ("no debe nada" sin
desglose) y solo se confirmó el campo de techado y campo. El mes que viene:
`registrar_ajuste` ACUERDOS -50 en `seguimiento_pueblo.xlsx` para D-6
(`source=manual`, motivo="yapeó a Wagner Trujillo hace años, no anotado").

---

## O-2 — CARMEN INGARUCA JULCA (recibo 18245)

**Motivo (dictado por la directiva, 01/08/2026):** se exonera la MULTA.

**Valores originales impresos:** Multa=30, Total=38.

**Valores corregidos (boleta de hoy):** Multa=0, Total=8.

**Verificado en el ledger:** `seguimiento_pueblo.xlsx` tiene MULTA SALDO=30
(CARGO 30, sin ningún PAGO) — coincide con la boleta, corrección nueva. El
mes que viene: `registrar_ajuste` MULTA -30 en `seguimiento_pueblo.xlsx`
para O-2 (`source=manual`, motivo="exonerado por la directiva 01/08/2026").

---

## H-16 — GREGORIO TOLENTINO SANCHEZ (recibo 18144)

**Motivo (decisión de la directiva, 01/08/2026):** poner la deuda de
TECHADO Y CAMPO en 75, no 47. La directiva dice que el yape está
"congelado" aunque le entró un pago — decidieron el 75 igual, pese a la
inconsistencia.

**Valores originales impresos:** Techado y campo=47, Total=66.

**Valores corregidos (boleta de hoy):** Techado y campo=75, Total=94.

**⚠ Mismo patrón que el lote de saldo negativo (ver `CONSOLIDADO.md`
Bloque B) — se escapó del filtro porque 47 no es negativo:**
`seguimiento_pueblo.xlsx` tiene ACUERDOS de H-16 con CARGO=75, un PAGO
fantasma de S/14 registrado el 06/07 (sin respaldo real, mismo tipo que
I-16/B-5/etc.) y el AJUSTE -14 del 31/07 que lo revirtió mal, dejando
SALDO=47 en vez de volver a 75. El 75 que pide la directiva **coincide
exacto** con el método usado en el resto del lote (restaurar al CARGO real,
sin el pago fantasma) — no es una decisión arbitraria, es la cifra correcta.
El mes que viene: `registrar_ajuste` ACUERDOS +28 en `seguimiento_pueblo.xlsx`
para H-16 (`source=manual`, mismo motivo que el lote de 11: pago fantasma
sin respaldo).

---

## G-4 — NATALIA CHINCHAY COLLAS (recibo 18104)

**Motivo (dictado por la directiva, 01/08/2026):** la secretaria dice que
pagó CONVENIO por fuera, S/50. Su pago de MULTA (25) también se pasa a
CONVENIO.

**Valores originales impresos:** Convenio=75, Multa=25, Techado y campo=50,
Total=158.

**Valores corregidos (boleta de hoy):** Convenio=0, Multa=50, Techado y
campo=50, Total=108.

**✔ HECHO el 2026-08-03** — CONVENIO=0, MULTA=50, ACUERDOS=50 (ACUERDOS nunca se
tocó, ya estaba en 50). Se escribieron los dos lados:

```
precursor                                    ledger (seguimiento_pueblo)
reasignaciones_aplicacion  MULTA→CONVENIO 25  AJUSTE MULTA    +25 → 50   REASIGNACION
                           (MES_ANO vacío)    AJUSTE CONVENIO −25 → 50   REASIGNACION
abonos_rezagados  G-4 50 convenio efectivo    AJUSTE CONVENIO −50 →  0   ABONO_REZAGADO
```

Los S/50 van como **ABONO_REZAGADO y no DECLARACION**: se buscaron el 03/08 contra
los ciclos congelados de `7_cierre/archivo/` y no aparecen en mesas, yape, blancos
ni en el pool de exceso — es plata real de la JASS que nunca entró al registro, así
que sí suma a caja.

---

## G-14 — MARGARITA GOMEZ BONIFACIO (recibo 18114)

**Motivo (dictado por la directiva, 01/08/2026):** tiene un pago de MULTA de
S/25 que en realidad debe contarse contra CONVENIO, y un pago aparte de S/50
en efectivo que le hizo directamente a la secretaria, que también va a
CONVENIO.

**Valores originales impresos:** Convenio=38, Multa=0, Techado y campo=21,
Total=83.

**Valores corregidos (boleta de hoy):** Convenio=0, Multa=50, Techado y
campo=50, Total=124.

**Pendiente para el mes que viene:** en `seguimiento_pueblo.xlsx`, registrar
contra G-14: PAGO CONVENIO por los S/25 (reasignados desde MULTA) + S/50
(efectivo a la secretaria, verificar que esté en `pagos_efectivo.xlsx` o
cargarlo como `abonos_rezagados`) hasta cubrir el CONVENIO=38 (queda exceso
de 37 — confirmar si va a otro concepto o queda a favor). Además, confirmar
por qué MULTA pasa de 0 a 50 y TECHADO Y CAMPO de 21 a 50 — la directiva no
detalló el origen de esos dos montos nuevos, solo el resultado final. Revisar
con ella antes de tocar el ledger real.

---

## E-14B — JUAN SAAVEDRA SAAVEDRA (recibo 18084)

**Motivo (dictado por la directiva, 01/08/2026):** cargar TECHADO Y CAMPO =
75.

**Valores originales impresos:** Techado y campo=0, Total=11 ("USTED ESTÁ AL
DÍA").

**Valores corregidos (boleta de hoy):** Techado y campo=75, Total=86 ("NO
ESTÁ AL DÍA").

**⚠ Revierte una decisión previa ya cerrada:** el 30/07/2026
(`4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx`, GRUPO 3) se había
resuelto que E-14B **ya pagó campo** y no debía nada de ese concepto —
`pendientes_convenio_multas.xlsx` lo tenía en `RESUELTO` sin monto. Hoy la
directiva pide cargarlo de nuevo por 75. **El mes que viene, antes de tocar
el ledger real: confirmar con la directiva si esto reemplaza esa decisión de
30/07 (se determinó que el pago que se creía hecho no existió) o si es un
cargo nuevo y distinto de ese mismo concepto** — no asumir, son S/75 y hay
una contradicción directa entre las dos fuentes.

---

# Lote de 11 predios con SALDO negativo en seguimiento_pueblo — 01/08/2026

Distinto de los reclamos de arriba: esto NO lo pidió la directiva, salió de
investigar por qué varias boletas de hoy mostraban Convenio/Techado y
campo/Multa en **negativo** (un "crédito" imposible). Causa encontrada:
`5_cobranza` registró un PAGO de julio para estos 11 predios (evento
`SOURCE=5_cobranza`, `recon_2026-07_...`, 06/07/2026) que no tiene respaldo
en ningún archivo de pagos real (`pagos_yape_tepago.xlsx`,
`pagos_efectivo.xlsx`, ni ninguno de los 8 overlays de corrección, ni los 28
pagos manuales de la secretaria) — un "pago fantasma". El 31/07/2026,
`5_cobranza` volvió a correr, no encontró respaldo, y revirtió ese pago con
un `AJUSTE` negativo — pero el mecanismo de reversión deja el SALDO negativo
en vez de restaurar la deuda real (mismo patrón ya documentado para F-12/D-1
en `docs/RETOMAR_limpieza_ledger_y_reasignaciones_2026-07-31.md`, salvo que
para estos 11 nadie hizo el segundo paso de estabilización).

Investigación completa, con la tabla de "Referencia de pago" real de cada
uno: `4b_reclamos/outputs/reporte_lote_saldo_negativo_2026-07.pdf`.

**Pendiente para el mes que viene, para cada fila:** un `registrar_ajuste`
manual en `shared/seguimiento_pueblo.xlsx` (mismo patrón que F-12/D-1,
`source=manual`) que lleve el SALDO al valor de la tabla de abajo — NO se
aplicó todavía, solo se parchó `DATA_boletas.xlsx` para que la boleta de hoy
salga con el monto correcto.

| Predio | Concepto | SALDO hoy (negativo, bug) | Restaurar a | Boleta de hoy (parchado) |
|---|---|---|---|---|
| A-8 | CONVENIO | -50 | 50 | Convenio=50, Total=90 |
| B-5 | ACUERDOS (techado y campo) | -25 | 25 | Cuota directa=25 |
| B-5 | CONVENIO | -50 | 50 | Convenio=50, Total=162 |
| C-1 | ACUERDOS (techado y campo) | -25 | 25 | Cuota directa=25 |
| C-1 | CONVENIO | -50 | 50 | Convenio=50, Total=94 |
| C-7 | CONVENIO | -25 | 25 | Convenio=25, Total=66 |
| E-12 | CONVENIO | -16 | 26 | Convenio=26, Total=47 (ACUERDOS ya estaba bien, no se tocó) |
| I-11 | CONVENIO | -25 | 25 | Convenio=25, Total=51 |
| I-16 | MULTA | -18 | 18 | Multa=18 |
| I-16 | ACUERDOS (techado y campo) | 47 (mal, no negativo pero igual incorrecto) | 75 | Cuota directa=75, Total=201 |
| J-3 | CONVENIO | -30 | 50 | Convenio=50, Total=66 |
| K-17 | CONVENIO | -25 | 25 | Convenio=25, Total=40 |
| K-2 | CONVENIO | -25 | 25 | Convenio=25, Total=39 |
| P-12 | CONVENIO | -50 | 50 | Convenio=50, Total=86 |

**Caso especial B-5 (Pompeyo Celestino Lliuya):** él le dijo a la secretaria
que CONVENIO está al día y que TECHADO Y CAMPO sí debe 50 — lo contrario de
esta tabla (que restaura ACUERDOS=25 y CONVENIO=50, sin reasignar nada). Se
decidió restaurar tal cual (versión A) porque es lo verificado en el ledger
de junio, sin depender de confirmar con él. **Si al hablar con Pompeyo
confirma su versión, la corrección de fondo sería una REASIGNACIÓN
(ACUERDOS=50, CONVENIO=25, mismo mecanismo que F1-11/A-6/F-12/D-1), no el
`registrar_ajuste` simple de la tabla de arriba** — decidir con él antes de
tocar el ledger real.
