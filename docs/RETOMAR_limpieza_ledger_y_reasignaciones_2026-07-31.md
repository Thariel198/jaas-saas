# RETOMAR — Limpieza del ledger, corrección de arquitectura y reasignaciones pendientes · Sesión 2026-07-31

Handoff detallado. Sesión larga (Sonnet, con un breve cambio a Opus revertido a Sonnet).
Empezó como "revisar por qué A-6 no aparece en el reporte histórico" y terminó en una
auditoría completa de `seguimiento_pueblo.xlsx` + corrección de un malentendido de
arquitectura + varias reasignaciones de pago en cola (sin ejecutar).

---

## ⚡ PRIMER PASO al retomar

1. **Reconciliar T-7**: el usuario dice que CONVENIO es **100**, pero verifiqué en
   `seguimiento_pueblo.xlsx` Y en `planilla_cobrado.xlsx` que muestran **50** en los dos
   lados. Antes de tocar nada de T-7, buscar de dónde sale el 100 (¿otro archivo? ¿una
   fuente externa como `SEGUIMIENTO INSTALACIONES...`? ¿confusión con otro predio?).
2. **T-7 — aclarar el mecanismo**: el usuario dijo "la solución es crearle su pago" —
   esto es DISTINTO de cancelar el CARGO (que es lo que veníamos planeando con
   `ajustes_cargo.xlsx`). "Crear un pago" sugiere registrar un PAGO real (lo que la
   secretaria dijo que recibió), no anular la deuda. Aclarar con el usuario cuál de los
   dos quiere antes de tocar el archivo.
3. Una vez resuelto T-7, **correr `5_cobranza --force` en `jass_system - Julio`** (el
   usuario fue explícito: NO correrlo sin que él lo diga) con todo lo que ya está en
   cola (ver sección "Cola de reasignaciones" más abajo).
4. Después de correr: copiar `shared/seguimiento_pueblo.xlsx` y
   `5_cobranza/outputs/arrastre_consolidado_2026-07.xlsx` de `jass_system - Julio` de
   vuelta al repo activo (`jass_system`, sin sufijo) — mismo patrón que el fix de
   race condition de yape del 27/07.
5. Verificar los resultados (D-16, A-6, F-12, I-2B, T-7) contra lo esperado, y agregar
   el AJUSTE de estabilización manual que le falta a F-12 (ver abajo, sección
   "mecanismo de reasignación").

---

## 1. Corrección de arquitectura — YA NO hay "dos fuentes de verdad"

Durante la sesión se armó (y luego se corrigió) `docs/diario/situacion_actual_sistema.html`.
El hallazgo importante, **verificado en código**:

```
2_planilla, boletas_sin_servicio.py, 4b_reclamos (reportes)
  → los TRES leen repo.get_saldos_bulk(concepto, mes) DIRECTO de
    seguimiento_pueblo.xlsx para MULTA/ACUERDOS/CONVENIO
  → NINGUNO lee esas columnas de arrastre_consolidado

arrastre_consolidado SÍ calcula su propia copia de MULTA/ACUERDOS/
CONVENIO (adentro de 5_cobranza, _exportar_arrastre_consolidado) —
pero es CÓDIGO MUERTO, nadie la consume. No causa desync real hoy.

DEUDA_AGUA y CORTE_RECONEXION SÍ dependen 100% de arrastre_consolidado
— ahí no hay problema, esos 2 conceptos no viven en ningún otro lado.
```

**El riesgo real no es "dos archivos que compiten"** — es que DENTRO de
`seguimiento_pueblo.xlsx`, un PAGO escrito a mano puede revertirse solo en la
siguiente corrida de `5_cobranza` (el "contador tuerto"): `pago_registrado()` suma
TODOS los PAGO sin importar el `source`, pero `ajuste_reconciliado(...,"5_cobranza")`
solo cuenta AJUSTE con ese source exacto. Por eso:
- Un **PAGO manual** → vulnerable, se puede revertir solo.
- Un **AJUSTE con `source` distinto de `"5_cobranza"`** → invisible para la
  reconciliación, sobrevive cualquier corrida futura.

Verificado empíricamente: 22 PAGO manuales del 28/07 SÍ se revirtieron el 30/07, y
alguien (antes de esta sesión) ya lo arregló agregando una tercera fila AJUSTE
estabilizadora — ese patrón (`condonacion-estable-30072026`) es robusto y NO hay que
tocarlo.

Actualizado: `docs/diario/situacion_actual_sistema.html` (diagrama principal,
secciones ③④⑦) — ya refleja esto correcto.

---

## 2. Limpieza del bug de yape (13/07) — CERRADA

Bug real: `5_cobranza` leyó `pagos_yape_tepago.xlsx` a medio guardar (13/07,
05:51-06:00), generó un AJUSTE negativo falso + luego re-registró el mismo pago como
PAGO duplicado — 36 pares (predio, concepto) con saldo negativo falso.

- **Bucket 1 (13 casos, 29 filas)**: patrón limpio (AJUSTE+PAGO del bug + compensación
  del 27/07), **BORRADO directo** (bug+compensación, sin dejar rastro). Saldo final
  verificado = 0 en los 13. Backup: `seguimiento_pueblo_pre_limpieza_bucket1_*.xlsx`.
- **Bucket 2 (4 casos)**: E-12, J-3, Z-12, Z-15(ACUERDOS) — la compensación del 27/07
  NO deshacía exacto el bug, "zanjaba" a 0 de más, borrando deuda real preexistente.
  **CORREGIDOS a su deuda real**: E-12=5, J-3=10, Z-12=30, Z-15(ACUERDOS)=30.
- **5 casos nuevos encontrados** (nunca llegaron a tener compensación, seguían rotos):
  F-4(MULTA), H-16(ACUERDOS), I-16(ACUERDOS), L-5(MULTA), W-5(ACUERDOS) —
  **restaurados** a su deuda real (49, 61, 61, 42, 42 respectivamente).
- **Fila fantasma "NAN/NAN"** (CONVENIO, -1334, sin predio real) — **borrada**, era
  ruido puro de una fila de totales que se coló al ledger.
- **A-4 CONVENIO**: investigado a fondo (3 golpes apilados: bug 13/07 + un tercer
  AJUSTE del 25/07 que resultó ser el residuo de un intento abortado de "cargar 75 de
  techado/campo" que ya existía desde junio). **Confirmado que el 0 actual es
  correcto** — coincide con la nota de la secretaria ("ya pagado, saldo 0"). No se
  tocó.
- **L-4 MULTA** (-17, 20/07): NO es el bug de yape (fecha distinta, corrida normal de
  5_cobranza). Sin nota de secretaria que lo respalde. **Se dejó sin tocar** — su
  deuda real sería S/3 si se limpiara, pero no hay decisión tomada.
- **P-6 CONVENIO**: el -300/+300 del 20-27/07 quedó **obsoleto pero inofensivo** —una
  reconciliación posterior (30/07, con nota de secretaria) ya resolvió el número real
  (S/58, instalación) correctamente. No se tocó, no hace falta.

Todos los backups de esta limpieza están en `shared/backups/seguimiento_pueblo_pre_*`.

---

## 3. D-16 — decisión de junio ejecutada, LUEGO reabierta para redo

**Decisión real** (`jass_system - junio/docs/RETOMAR_junio_cierre_reconciliacion_2026-07-18.md`,
sección 6): D-16 (Esteban Guerrero Chingel) sí asistió a faena, la directiva no
registró su asistencia → **eliminar la MULTA completa (S/50)**. Convenio/acuerdos
son deuda real, no se tocan. Esta decisión llevaba **13 días sin ejecutar**.

Primero se ejecutó como **borrado directo** en el ledger (CARGO+PAGO+AJUSTE de MULTA,
3 filas, en el repo activo) + un precursor en `ajustes_cargo.xlsx`
(`MES_ANO_APLICA=2026-08`, ya que julio no se puede re-correr en el repo activo).

**Luego se decidió rehacerlo bien**: el borrado directo perdió sin querer el rastro de
S/15 que D-16 sí pagó realmente (estaba aplicado a MULTA) — sin verificar a dónde
debía fluir esa plata real en la cascada (probablemente ACUERDOS). Se agregó una
**segunda fila de precursor** en `ajustes_cargo.xlsx` de `jass_system - Julio`:
`MULTA -50, MES_ANO_APLICA=2026-07` — para que al correr `5_cobranza --force` en esa
copia, el motor recalcule bien y mande los S/15 reales a donde corresponda.

⚠ **Pendiente**: verificar, después de correr, que D-16 termine con MULTA=0 y que los
S/15 reales aparezcan correctamente aplicados en otro concepto (probablemente
ACUERDOS) — no asumir, confirmar con los números reales post-corrida.

---

## 4. Grupo 1 — 15 condonaciones de secretaria, YA ROBUSTAS (no tocar)

Predios: B1-12(MULTA·ACUERDOS), F1-5, G1-14, L-16, N-5, Q-10, Q-11(MULTA·ACUERDOS),
Q-5(MULTA·ACUERDOS), R-4, S-1, S-12, S-8, S-9, V-6, D1-3.

Verificado: todos terminan en SALDO=0, con el patrón robusto
`AJUSTE negativo (source=5_cobranza) + AJUSTE positivo (source=manual,
"condonacion-estable-30072026")` — **NO necesitan ningún precursor** (2_planilla lee
`seguimiento_pueblo` directo, `arrastre_consolidado`/`ajustes_cargo` no influyen en
estos 3 conceptos). Confirmado estable ante una futura corrida de `5_cobranza`.

**Grupo 2 relacionado** — C-21 y J-6(MULTA·ACUERDOS): "revertir condonación fallida",
**sin nota de secretaria que lo respalde** — quedaron con deuda real reinstalada
(C-21 MULTA=50, J-6 MULTA=50+ACUERDOS=75). Correcto tal como está, no tocar sin nueva
evidencia.

**Grupo 3 (pendiente, sin resolver)** — Q-4: tiene pasos intermedios ruidosos (falló
→ revirtió → re-condonó, con "recondonacion-definitiva-30072026"). El usuario
mencionó que probablemente haya que simplificar/limpiar esas filas intermedias, pero
**no se investigó a fondo ni se decidió nada** — queda como tarea abierta.

`correccion_genesis_formula` (109 filas, bug de fórmula de abril, YA comunicado vía
109 boletas reimpresas en julio) y `correccion_lote_F3B_a_F3A` (1 fila, relabeling de
predio) — **confirmados legítimos, no tocar**.

---

## 5. reporte_historico.py — FIX DE CÓDIGO aplicado y verificado

Archivo: `4b_reclamos/reporte_historico.py`, función `_filas_recientes()`.

**Problema original**: la fila de un mes solo se creaba si existía un evento
MULTA/ACUERDOS/CONVENIO en `seguimiento_pueblo` ese mes — un pago de SOLO consumo
(como A-6 en julio) nunca generaba fila, aunque sí hubiera pagado.

**Fix aplicado**: el disparador ahora es "¿hubo pago real (yape+efectivo) en
`planilla_cobrado` ese mes?" — generaliza el antiguo hardcode de `"2026-07"` a
`mes_ciclo_actual` (leído dinámico de `MES_ANO`).

**Bug adicional encontrado y corregido en el camino**: `float(v or 0)` en Python NO
maneja NaN (`nan or 0` devuelve `nan`, no `0`, porque NaN es truthy) — esto hacía que
el chequeo de "hubo pago" nunca disparara cuando `MONTO_YAPE` era NaN. Se agregó
`_numf(v)` (NaN-safe) y se reemplazaron los usos de `or 0`.

**Segundo bug encontrado**: la columna CONSUMO mostraba el monto **debido** (MES_ACTUAL
+ MANTENIMIENTO), no el monto **pagado** — para A-6 (pago parcial) mostraba "46"
cuando ella solo pagó 43. Corregido: `consumo_ciclo_actual = min(consumo_debido,
total_pagado_ciclo)`.

**Verificado**: corrida contra los 116 predios de `_predios_confirmacion()`, 0
errores, 90 ganaron fila de julio (tenían pago real), 26 correctamente sin fila (sin
pago real, verificado caso por caso incluyendo un caso de YAPE negativo/retorno en
B-8). PDF regenerado: `4b_reclamos/outputs/reporte_historico_CONFIRMACION_2026-07.pdf`
(116 páginas) — A-6 (pág. 2) y T-7 (pág. 109) verificadas visualmente.

⚠ Este cambio de código vive SOLO en el repo activo — no se sincronizó a
`jass_system - Julio` (no hacía falta para lo que se corrió ahí).

---

## 6. Cola de reasignaciones — preparadas en `jass_system - Julio`, SIN CORRER

**Regla del usuario: NO correr `5_cobranza --force` en `jass_system - Julio` sin que
él lo diga explícitamente.**

### Sincronizado hacia `jass_system - Julio`
- `shared/seguimiento_pueblo.xlsx` (copiado del repo activo, incluye TODA la limpieza
  de las secciones 2 y 3 de arriba).
- Código: `_CONCEPTO_DEVOLUCION_A_CAMPO` en `5_cobranza/main.py` — se agregó
  `"MES_ANTERIOR": "mes_anterior"` (antes no existía ese mapeo, era necesario para
  I-2B). Aplicado en AMBOS repos (activo y Julio).

### `shared/ajustes_cargo.xlsx` (Julio) — 4 filas nuevas
| Predio | Concepto | Monto | Motivo |
|---|---|---|---|
| A-6 | AGUA | -17 | consumo 43→26 (reclamo consumo alto, su promedio es 26.13) |
| I-2B | MES_ANTERIOR | -6 | eliminar arrastre — **sin motivo detallado**, instrucción directa del usuario |
| D-16 | MULTA | -50 | asistió a faena, redo vía precursor (ver sección 3) |

### `shared/reasignaciones_aplicacion.xlsx` (Julio) — 2 filas nuevas
| Predio | Origen → Destino | Monto | Motivo |
|---|---|---|---|
| A-6 | ACUERDOS → CONVENIO | 14 | sobrante tras bajar consumo a 26 (46 pagó − 29 debía en DEUDA_AGUA − mantenimiento, redirige a medidor) |
| F-12 | MULTA → CONVENIO | 50 | asistió a faena/reunión, redirige el pago de multa hacia convenio/medidor |

### Mecanismo importante a recordar (para no confundirse post-corrida)
Cuando `reasignaciones_aplicacion` mueve un PAGO real de un concepto a otro, el
motor de reconciliación va a generar **automáticamente** un AJUSTE negativo en el
concepto ORIGEN (esto es esperado, verificado en código — NO es un bug). Ese ajuste
deja el saldo del origen en negativo (falso crédito), no en el valor final deseado.
**Hace falta un AJUSTE manual adicional después de correr**, para estabilizar:
- **F-12 MULTA**: después de correr, esperar saldo ≈ -50 → agregar AJUSTE manual +50
  ("esta multa no corresponde, se cancela") para dejarlo en 0.
- Revisar si A-6/ACUERDOS necesita el mismo tratamiento (probablemente si — el
  reasignaciones ahí también mueve plata desde ACUERDOS).

### T-7 — INVESTIGACIÓN HECHA, SIN EJECUTAR (ver "primer paso" arriba)
- Ledger hoy: CONVENIO CARGO=50/SALDO=50 (usuario dice que es 100 — **reconciliar
  antes de tocar**), ACUERDOS CARGO=50/PAGO=5/SALDO=45.
- Hallazgo fuerte: CONVENIO ya lo pagó completo ANTES de julio (marzo 25 + mayo 25 =
  50) en el sistema viejo — la génesis del ledger nuevo (02/07) lo volvió a sembrar
  como deuda fresca sin reconocer ese pago histórico. Esto respalda su reclamo con
  datos, no solo su palabra.
- ACUERDOS (45) NO tiene ese mismo respaldo histórico — nunca pagó acuerdos antes de
  julio. Solo tiene la palabra de la esposa (confirmó "está al día" en el cobro).
- Se agregó nota en `4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx` (fila 174,
  GRUPO=5, TIPO_RECLAMO=CONFIRMACION, ESTADO=PENDIENTE) documentando ambos hallazgos.
- PDF regenerado, T-7 visible en página 109 con su historial completo.
- **El usuario pidió "crearle su pago"** en vez de (o adicional a) cancelar el
  cargo — aclarar exactamente qué quiere decir antes de tocar nada (ver punto 2 del
  "primer paso").

---

## 7. Pendientes generales, sin resolver

- **Q-4 (Grupo 3)**: limpiar el ruido de pasos intermedios (falló→revirtió→re-condonó)
  — mencionado por el usuario, no investigado a fondo.
- **2_planilla de agosto**: sigue sin re-correr desde que se hicieron todas estas
  correcciones al ledger (D-16, Grupo 1, etc. — aunque Grupo 1 ya estaba bien desde
  antes). Al re-correr, **se pierde el parche manual de M-12** (mantenimiento 3→83,
  ver RETOMAR de `notas_2026-07`) — hay que repetirlo después.
- **3_boletas y reportes de agosto**: no se han generado/regenerado con los datos
  finales — pendiente hasta que se cierre la cola de reasignaciones de la sección 6.
- Backups generados esta sesión (todos en `shared/backups/` del repo correspondiente,
  con timestamp): múltiples para `seguimiento_pueblo.xlsx`, `ajustes_cargo.xlsx`,
  `arrastre_consolidado_2026-07.xlsx`, `notas_2026-07.xlsx` — no hace falta limpiarlos,
  pero si se acumulan demasiado en el futuro, revisar cuáles ya no hacen falta.

## Cómo cerrar este RETOMAR

Cuando T-7 esté resuelto, se corra `5_cobranza --force` en `jass_system - Julio`, se
copien los 2 archivos de vuelta al repo activo, se verifiquen los 5 casos (D-16, A-6,
F-12, I-2B, T-7) contra lo esperado, y se decida qué hacer con Q-4 — borrar este
archivo.
