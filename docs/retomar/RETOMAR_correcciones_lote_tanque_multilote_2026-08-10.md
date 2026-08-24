# RETOMAR — correcciones_lote, tanque no detectado, multi-lote · 2026-08-10

Sesión larga. Detalle narrativo completo en la conversación de hoy; este doc es
el resumen operativo. Ver también `LEER_ANTES.md` (2 secciones nuevas al tope)
y `bugs_sistema.json` (raíz del repo, 15 bugs catalogados).

---

## ⚡ PRIMER PASO al retomar — C1-2/K-2, el S/400 de tanque está mal repartido

`shared/aportes_tanque_manuales.xlsx` tiene una fila que dice que el S/400 de
Antonio Espinoza (mensaje "mz k Lt 2,mz C1 Lt 2 tanque", 22/07) es **100% de
C1-2** — está MAL. El mensaje nombra los 2 lotes (K-2 y C1-2): es un aporte
conjunto para dos casas, no solo para C1-2.

```
K-2  Antonio Espinoza Sifuentes    SALDO=28 (CONVENIO 25 sin cubrir)
                                    → nunca recibió su parte del aporte
C1-2 Aubertina Trujillo Tolentino  ya recibió el S/400 completo
```

**Falta decidir el reparto** (¿mitad y mitad, S/200 cada uno? ¿confirmar con
Antonio Espinoza o la secretaria?) y corregir la fila en
`aportes_tanque_manuales.xlsx` (hoy dice `MZ=C1, LT=2, MONTO=400` — hay que
partirla en 2 filas, una por lote, o ajustar el monto). Después correr
`5_cobranza --force`.

---

## 2. Lo que quedó CERRADO hoy (no revisar de nuevo)

```
✔ C1-9 (Roberto Macarlopu)   auto-sanador resucitaba C1-9→C1-17 (predio que
                               ya no existe, eliminado del padrón 28/07).
                               Candado permanente: correcciones_lote.xlsx
                               tiene C1-9→C1-9 (identidad) + fix en
                               _recuperar_correcciones_trazabilidad (ya no
                               resucita reglas cuyo ORIGEN es un predio real)

✔ G-36 → C-36                 typo del cobrador Wagner, corregido en
                               correcciones_lote.xlsx

✔ L-9 → J-9 REVERTIDO          la regla no tenía confirmación humana (a
                               diferencia de otros casos) — la plata volvió
                               a L-9 (Marianela Rivera Vitate), que con eso
                               canceló MULTA 30 + ACUERDOS 50

✔ V-14 (Leonardo Huamani)      S/100 de Janet Vil* (26/07, mismo batch que
                               otros 4 aportes tanque de la misma persona)
                               → aportes_tanque_manuales.xlsx. Su MULTA de 20
                               volvió a aparecer como deuda real (correcto,
                               nunca estuvo pagada) — borré el par PAGO+AJUSTE
                               fantasma del ledger a pedido tuyo (backup en
                               shared/backups_ledger/)

✔ I-13 (Julio Solorzano)       S/200 de Walter Sol* (23/07) → tanque,
                               100% limpio, sin resto

✔ H1-36 (Tito Cosme)           segundo depósito de Nancy Mon* (S/50, 24/07,
                               sin la palabra "tanque" a diferencia del
                               primero) → confirmado tanque también

✔ D1-1→D1-3 (Margarita         confirmado con evidencia real (WhatsApp,
Criollo, caso de julio)        Janet Villanueva Alegre, 14/07) — NO se tocó,
                               es distinto a L-9 (que no tenía esa evidencia)

✔ arrastre_devolucion_2026-08  cambio de código: ahora también lista los "no
  .xlsx — código                identificados" (huérfanos sin lote válido),
                               antes solo vivían en discrepancias_cobranza

✔ K-3/K-4 — FIX DE CÓDIGO       extraer_multiples() no detectaba mensajes
  escrito, falta aplicar        tipo "K-3, K-4." (sin la palabra "mz"/"lt").
                               Agregado PATRON_MULTIPLE_GUION en
                               4_pagos/yape/motor_matching/main.py, probado
                               contra 8 mensajes reales sin regresión.
                               ⚠ FALTA correr motor_matching + 5_cobranza
                               para que se aplique — el archivo de pagos
                               todavía tiene a K-3 con el pago entero de
                               Yuly Lisbeth Trujillo Cruz (S/34)

✔ forzar_mzlt.xlsx             quitada 1 fila leftover de julio (Janet Vil*,
                               inofensiva pero desprolija). Revisados los
                               otros 3 (Giovanna San*, PLIN Juan Carlos,
                               Patricia Tar*) — 2 de esos 3 usaban la
                               herramienta equivocada (eran ambigüedad de
                               matching, no digitación del pagante — debieron
                               ir a pendientes.xlsx). Ya corrieron en julio,
                               no se tocó nada, queda como lección.
```

---

## 3. Lo que queda PENDIENTE, sin resolver

```
S-16 (mesa_1, Wilder Trujillo, S/22)
   candidato: S-6 (VICTORIA COSME CERNA) — NO confirmado. Visible en
   arrastre_devolucion_2026-08.xlsx como "no identificado".

X-11 (mensaje "8, 3:36, 2/8, se contradijo en poner mz y lt")
   NO se encontró ninguna transacción que calce (ni yape ni efectivo, ni
   03:36 ni 15:36). Falta el ORIGEN (nombre bancario de quien pagó) para
   poder buscar. Sigue en reclamos_manuales.xlsx sin resolver.

F-4 vs F1-4 ("mes pasado pagó 101, quedó de deuda 56")
   F-4 (Tolentino Julca Moreno) no tiene ningún pago de S/101 en julio.
   F1-4 (Antonio Gutierrez Pachamango) SÍ tuvo un abono rezagado de S/101,
   pero ya se investigó y cerró el 09/08 ("SIN CAMBIOS necesarios") — esa
   conclusión no menciona ningún "queda debiendo 56". Falta preguntarle a
   la secretaria cuál de los dos predios es.

E-14B (Juan Saavedra Saavedra) — DUPLICADO en planilla
   2 filas para el mismo (MZ,LT) en shared/planilla_mes/planilla_2026-08.xlsx
   — genera un par AJUSTE+PAGO fantasma en CADA corrida de 5_cobranza
   --force (no afecta el saldo neto, pero ensucia el ledger). Pendiente:
   borrar la fila duplicada y confirmar que una corrida limpia no vuelve a
   generar el par.

4 excesos sin explicación de fondo (matemática ya verificada, falta
confirmar con la secretaria si es tanque/reembolso o error):
   A-5A  Isabel Puntillo Vega     exceso 31 (Jaime Huerta pagó 39, debía 8)
   H-9   Clemente Rodriguez Rojas exceso 37 (pagó 67 efectivo, debía 30)
   U-3   Marcelina Gomez Bonif.   exceso 8  (2 pagos idénticos de S/8, mismo
                                   cobrador, 2 días seguidos — ¿duplicado?)
   D1-3  Margarita Criollo Gaspar exceso 11 (alguien más pagó su agua sin
                                   saber que ya estaba cubierta)

D-16 (Esteban Guerrero Chingel) — YA EXPLICADO, no es bug
   exceso 17 = arrastre de julio (abono rezagado S/85 + condonación de
   MULTA de faena S/50 este mes) — caso ya cerrado, no re-investigar.

8 reclamos nuevos cargados hoy en 4b_reclamos/inputs/reclamos_manuales.xlsx
(de la foto del cuaderno, FECHA=2026-08): K-11 (resuelto vía
reidentificación), O-16 (ya estaba resuelto, confirmado), G-14 (ya estaba
resuelto, confirmado), I-14, X-14, K-17, X-11 (sin resolver, ver arriba),
F-4 (ambiguo, ver arriba) — faltan pasar por el flujo normal de
4b_reclamos/main.py para que queden marcados RESUELTO donde corresponda.
```

---

## 4. Módulos que NO se corrieron después de los últimos cambios

```
motor_matching (4_pagos/yape)   falta correr — tiene el fix de K-3/K-4 sin
                                  aplicar
5_cobranza                       correr DESPUÉS de motor_matching y de
                                  decidir el reparto de C1-2/K-2
5b_validacion                    no se corrió desde los últimos cambios —
                                  correr al final para confirmar que no
                                  quedó ninguna alerta nueva
6_corte/generar_lista.py         63 usuarios elegibles a corte, no generado
                                  todavía este ciclo
```

## 5. Orden sugerido para mañana

1. Decidir reparto C1-2/K-2, corregir `aportes_tanque_manuales.xlsx`.
2. Correr `motor_matching` (aplica el fix de K-3/K-4).
3. Correr `5_cobranza --force`.
4. Revisar `arrastre_devolucion_2026-08.xlsx` — debería quedar con menos
   filas (K-2 recibe su parte, K-3/K-4 se separan).
5. Correr `5b_validacion` — confirmar 0 alertas nuevas.
6. Retomar los pendientes de la sección 3 (S-16, X-11, F-4/F1-4, E-14B
   duplicado, los 4 excesos sin explicar) con la secretaria.
