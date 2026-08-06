# RETOMAR — Reclamos de julio + patrón de "exceso que en realidad es otro concepto" · Sesión 2026-07-23 (tarde)

## Punto de entrada

Sesión larga de triage de reclamos de julio (`4b_reclamos/outputs/reclamos_2026-07.xlsx`) y de
excesos de `5_cobranza/outputs/arrastre_devolucion_2026-07.xlsx`. Aparece un patrón repetido: pagos
que en realidad son tanque/instalación/medidor/cuota-de-asamblea, pero que el sistema flaggea como
"exceso de agua" porque nadie le dijo el concepto real. Esta sesión resolvió varios casos puntuales
a mano y detectó que el patrón ya se repitió demasiadas veces para seguir parchando uno por uno.

Nota: esta sesión es posterior y distinta a `RETOMAR_cascada_ca1_y_pendientes_2026-07-23.md` (la de
la mañana, sobre CA1). Los pendientes de esa sesión (M-19, Z-14, resolucion_reclamos_2026-07,
5_cobranza --force) siguen abiertos, ver sección 4.

---

## 1. Resuelto hoy (verificado, guardado)

| Caso | Qué era | Dónde quedó |
|---|---|---|
| **Q-12** (Teodora Meza Garcilazo) | Blanco de S/50 (mesa_5, Máximo, junio) = aporte tanque, NO deuda. Distinto de Q-6 (Dedicación Garcilazo, caso ya cerrado 21/7) — mismo apellido causó confusión en una nota de sesión anterior. | `shared/aportes_tanque_manuales.xlsx` (fila nueva) + `reclamos_2026-07.xlsx` RESOLUCION+ESTADO=RESUELTO |
| **C-15** (Julio Rios Blas) | Yape S/200, mensaje literal "mz c lt 15 tanque" — motor de matching identificó bien el lote pero nadie llenó CONCEPTO (solo se llena a mano en `pendientes.xlsx` para casos ambiguos; este fue matcheado directo). | `shared/aportes_tanque_manuales.xlsx` + `arrastre_devolucion_2026-07.xlsx` REVISION+ESTADO=resuelto |
| **P-7** (Hipolito Melgarejo Mendoza) | Yape S/200, mensaje solo decía "P-7" (sin "tanque" explícito) — confirmado como tanque por el usuario. | idem C-15 |
| **C-35** (Analy Quineche Morante) | PLIN S/300 (18/06) = **100% CONVENIO INSTALACIÓN**, coincide exacto con `DEUDA POR INSTALACION=300 / SALDO=300` en hoja `INSTALACIONES ANTERIOR DIRECTIV` de `obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx`. No es exceso — su agua+acuerdos de julio (94) sigue debiendo aparte. | `arrastre_devolucion_2026-07.xlsx` REVISION+ESTADO=resuelto |
| **F1-1** (Nemecio Avelino Nolasco) | Arrastre MES_ANTERIOR=9 de junio en disputa. | `pendientes_convenio_multas.xlsx` (VERIFICANDO, aplica boleta agosto) + reclamos EN_REVISION |
| **T-20** (Maria Julca Rios) | Consumo alto julio (59 m³), pagó 19 de 62 — el saldo (43) pasará a MES_ANTERIOR en agosto, se pre-cargó VERIFICANDO para entonces. | idem |
| **16 reclamos con TIPO_RECLAMO ya clasificado** (convenio/multa/cuota/mes_anterior) | Se cargaron VERIFICANDO en `pendientes_convenio_multas.xlsx` para que sus boletas de **agosto** (no julio, que ya se imprimió el 4/7) muestren "Verificando" en vez del monto. | `pendientes_convenio_multas.xlsx` (23→41 filas) + reclamos ESTADO=EN_REVISION |
| **C1-17** | Ya estaba resuelto — fila 3 de `shared/reidentificacion.xlsx` (C1-9→C1-17, 18.5 agua, MES_ANO_APLICA=2026-07). Solo se confirmó, no se tocó nada. | — |
| **V-1** | Confirmado **EXONERADO** en `6_corte/inputs/registro_cortes.xlsx` (13/7, "Sí pagó, verificando pago comunitario"). **Falta marcar RESUELTO en `reclamos_2026-07.xlsx`** — se identificó pero no se ejecutó el update. | ⚠️ pendiente de escribir |
| **G-23** (Hipolito Ramirez Centeno) | El reclamo "el hijo fue a pagar" era cierto: `4_pagos/efectivo/inputs/mesa_3.xlsx` tenía una fila (71 total = 49 yape + 22 efectivo) que nunca se procesó porque `pagos_efectivo.xlsx` estaba desactualizado (generado 20/07, mesa_3 modificado 21/07). Se re-corrió `4_pagos/efectivo/main.py` (ciclo 3) — **G-23 ya aparece con 22 en `pagos_efectivo.xlsx`**. NO se corrió `5_cobranza` todavía (pedido explícito del usuario). | `4_pagos/efectivo/outputs/pagos_efectivo.xlsx` regenerado |

**Dedicación Garcilazo (Q-6) — segregación comunitaria mapeada** (dos depósitos, ya resueltos de antes, solo se documentó el detalle hoy):
- 05/07 20:42 (S/49) → G-23 completo (es el mismo yape que ya conocíamos, no plata nueva).
- 04/07 20:44 (S/361) → Q-3=200(tanque) · D-5=20 · F-14=68 · K-7=23 · V-1=33 · H-4=17.

---

## 2. Pendiente — necesita la corrida grande (~20 min, diferida a pedido del usuario)

Falta correr, en algún momento con más tiempo:
1. `5_cobranza --force` (para que G-23 baje su SALDO 22→0, y para reflejar la penalidad de corte de 7 usuarios aplicada el 22/7 — pendiente de antes)
2. `5b_validacion`
3. `4_pagos/consolidar_tanque.py` (para que los aportes de Q-12/C-15/P-7 pasen de `aportes_tanque_manuales.xlsx` al consolidado)

## 3. Pendiente — falta dato o decisión antes de tocar código/archivos

- **F1-11** (Maria Elizabeth Silva Sosa): pagó S/100 "medidor", pero **no hay ningún registro de su TOTAL de medidor en ningún lado** (ni `NUEVAS INSTALACIONES`, ni `MEDIDORES`/`MEDIDORES 2026`, ni el Drive del usuario). Se propuso cargarla en `MEDIDORES 2026` con TOTAL en blanco + PAGADO=100 — **no se ejecutó**, falta el TOTAL o la decisión de cargarla igual con el dato pendiente.
- **S/75 "techado y campo" × 5 lotes** (A1-12, C1-13, C1-15, H1-13, H1-36): todos pagaron exacto S/75 la misma semana (04-07/07), pero **ninguno tiene deuda de ACUERDOS_ASAMBLEA en ningún lado** (`seguimiento_pueblo` en 0 o inexistente, planilla en 0). Hipótesis: es una cuota nueva de asamblea que nunca se sembró como CARGO. **Falta confirmar con el usuario qué es este cargo** antes de sembrar nada.
- **P-6** (Flor Valdivia, convenio instalación S/1250, resta 350): el diseño del paso que falta (matchear pagos de `NUEVAS INSTALACIONES` contra excesos ya existentes en `arrastre_devolucion`, sin duplicar plata) **no se construyó** — quedó solo diagnosticado. Ver también el patrón general en sección 5.

## 4. Pendiente — de sesiones anteriores, sin cambios hoy

- B-19, F1-4: multas sin ningún respaldo de pago — preguntar directo a Yanet/secretaria (ver `project_b19_f1-4_multas_sin_respaldo.md`).
- G-23 MULTA: con lo de hoy, su lado de AGUA queda resuelto (49+22=71=deuda total incl. multa) — pero el reclamo original decía que pagó "la multa completa" y solo hay 49+22 registrados contra una multa de 50 dentro de esos 71. Cuando corra `5_cobranza`, revisar si la multa queda en 0 exacto o si sigue faltando algo.
- `resolucion_reclamos_2026-07.xlsx` sigue sin generar (29/76 reclamos clasificados).
- Z-14 en riesgo activo de corte, sin acción tomada.
- M-19 con guardado sin confirmar (de la sesión de la mañana).

## 5. Patrón detectado — candidato de diseño (NO implementado, solo señalado)

Van **3 conceptos distintos** cayendo en el mismo hueco esta sesión: instalación (P-6), tanque
dentro de mensajes de yape sin CONCEPTO (C-15/P-7), medidor (F1-11), y posiblemente una 5ª
categoría (cuota de asamblea nueva, sección 3). Todos terminan como "exceso" en
`arrastre_devolucion` porque `5_cobranza` no tiene forma de saber que esa plata era para otra cosa.

`seguimiento_pueblo.xlsx` ya resuelve exactamente este problema para MULTA/ACUERDOS/CONVENIO
(event-sourced, con génesis trazable desde `shared/genesis_inputs/` + reconciliación automática
contra pagos). La extensión natural sería sumar `MEDIDOR` e `INSTALACION` como conceptos del mismo
ledger, en vez de seguir anotando REVISION caso por caso cada mes. **Esto es diseño (Opus), no se
tocó código** — se preguntó y el usuario prefirió seguir parchando manual por ahora.

## 6. Archivos nuevos/tocados hoy (para saber qué mirar mañana)

- `3_boletas/inputs/pendientes_convenio_multas.xlsx` — **untracked en git**, 23→41 filas.
- `shared/aportes_tanque_manuales.xlsx` — **untracked en git**, 1→4 filas (C1-17 de antes + Q-12 + C-15 + P-7).
- `4b_reclamos/outputs/reclamos_2026-07.xlsx` — gitignored (outputs/), varios ESTADO/RESOLUCION actualizados.
- `5_cobranza/outputs/arrastre_devolucion_2026-07.xlsx` y `arrastre_devolucion_2026-06.xlsx` — gitignored, columna ESTADO nueva (dropdown resuelto/pendiente) + REVISION anotada en varios.
- `4_pagos/efectivo/outputs/pagos_efectivo.xlsx` — gitignored, regenerado (ciclo 3, 359 cobros).
- `4_pagos/efectivo/trazabilidad/trazabilidad_2026-07.xlsx` — untracked en git, regenerado.

Todos los backups de cada cambio quedaron en sus respectivas carpetas `backups/` con timestamp.
