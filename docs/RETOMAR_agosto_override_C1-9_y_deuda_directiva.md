# RETOMAR — Inicio ciclo AGOSTO 2026
### Handoff sesión 2026-07-08 · dos pendientes que se aplican ANTES de correr agosto

---

## TL;DR — lo que hay que hacer al arrancar agosto (en este orden)

1. **Override de padrón C1-17 → C1-9** (Roberto Macarlopu) — el único error de padrón real y
   confirmado de este ciclo. Aplicar en `0_padron` **antes** de que corra `2_planilla` de agosto.
2. **Crear la base de datos de "deuda de la directiva"** — para registrar pagos como el de
   Ronel (ex-tesorero), que no son agua ni tanque. Incluye arreglar `5b_validacion` para que
   cuadre esos conceptos.

Ninguno de los dos bloquea el corte de julio (ver más abajo por qué). Son para agosto.

---

## PENDIENTE 1 — Override padrón: Roberto Macarlopu C1-17 → C1-9

### Qué pasó (causa raíz)
El matching de padrón (`0_padron/02_matching/padron_matching.py`) **empareja por LOTE, no por
nombre** (es COFOPRI-driven: `padron_principal` = verdad). Por eso **no puede detectar la misma
persona en dos lotes**.

```
COFOPRI (padron_principal):   Roberto Macarlopu → C1-9   (una sola vez, = su título)
padron_secundario:  cobranza → C1-17
                    faena    → C1-9 y C1-17

Matching: COFOPRI dice C1-9. En cobranza C1-9 estaba LIBRE (solo estaba C1-17).
  → "lote libre → agregar nombre COFOPRI"  → mete a Roberto en C1-9
  → la fila vieja C1-17 (Roberto) queda ahí, nadie la borra
  → padron_reconciliado: DOS filas, mismo nombre, C1-9 + C1-17
```

En `padron_reconciliado`:
- **C1-9** → fila VACÍA (sin medidor, sin marcación) — vestigio de COFOPRI.
- **C1-17** → fila OPERATIVA (marcación 22, la que tiene el medidor). La planilla usó esta.

`usuarios_id.xlsx` también lo tiene mal: **C1-17** (MZ2/LOTE2 vacío → no son 2 lotes).

### Confirmado: es UN solo predio mal numerado, NO dos lotes
- Un solo medidor (marcación 22), bajo C1-17. C1-9 no tiene medidor ni consumo.
- Su reclamo (mesa_1) dice textual: *"Ya pague mes anterior 18.5. Pague faena.
  **Cambiar mz: C1-17 a C1-9**"* — pide cambiar, no "tengo dos".
- Lote real = **C1-9** (COFOPRI + su título).

### Mapa verificado — dónde vive Roberto (auditado 2026-07-09, NO re-auditar)
```
DÓNDE                                        LOTE    DATO
padron_reconciliado/cobranza  r510           C1-9    VACÍA (fantasma COFOPRI)
padron_reconciliado/cobranza  r515           C1-17   marcación 22 (operativa)
corregido/cobranza            r455           C1-17   convenio25 · mant3 · faena30 · total70
corregido/faena               r250           C1-9    MONTO30 SALDO30  ← real
corregido/faena               r252           C1-17   MONTO30 SALDO30  ← DUPLICADO
corregido/techado             r185           C1-9    saldo0 (pagado, ok)
usuarios_id                   U0457          C1-17   (registro de identidad)
arrastre_consolidado 2026-06  r53            C1-17   junio 18.5+30 = 48.5
planilla_cobrado    2026-07   r458           C1-17   julio 57.5 PENDIENTE
```
- **multa / acuerdos: NO tiene ninguno** (seguimiento_pueblo vacío para él). Ese riesgo no aplica.
- Nombre con 2 grafías en el crudo: `MACARLOPU` y `MACARLUPU` (+ variante `MARTIN`) → el match
  por nombre es frágil, usar LOTE.

### El fix — 3 lugares por CADA regeneración de 0_padron
```
corregido/cobranza :  C1-17 → C1-9      (no colisiona: no hay C1-9 en cobranza corregido)
corregido/faena    :  BORRAR fila C1-17  (C1-9 ya existe → es dedup, NO rename)
usuarios_id U0457  :  C1-17 → C1-9
```
Aplicar **antes de `2_planilla` de agosto** para que la planilla nazca C1-9.

### ⚠ El override.xlsx actual NO alcanza para este caso (verificado leyendo aplicar_override.py)
El mecanismo `override.xlsx` → `aplicar_override.py` **no sirve as-is** — produciría el lote
fantasma y la deuda doble que hay que evitar. Tres hoyos:
```
HAZARD 1 · aplicar_override solo RENOMBRA (primer match). faena ya tiene C1-9 + C1-17
           → renombrar crea DOS filas C1-9 = fantasma + deuda doble S/60. Hay que BORRAR.
HAZARD 2 · la hoja `cobranza` NO está en HOJA_COLS → el override no toca la deuda operativa
           (la que alimenta la planilla). Sin esto agosto vuelve a nacer C1-17.
HAZARD 3 · usuarios_id.xlsx tampoco lo toca el override.
```

### DECISIÓN PENDIENTE para agosto (A o B — el usuario aún no eligió)
- **A. Extender el override** (agregar hoja `cobranza` + acción `BORRAR`/dedup + toque a
  `usuarios_id`), cargar la fila `ACTIVO=Si` → declarativo, se re-aplica solo cada ciclo,
  **reutilizable para F1-6** (mismo patrón, ver más abajo). Recomendado: durable.
- **B. Corrección quirúrgica manual** en agosto (rename cobranza, borrar faena dup, fix
  usuarios_id). Rápida pero **0_padron regenera `corregido` desde el crudo cada ciclo → se
  deshace**. Frágil salvo que además se arregle el padrón crudo de raíz.

### Por qué NO se re-corre julio por esto
- Roberto ya está **protegido del corte** de julio (su fila es `[reclamo]`).
- Sus S/9 de julio (cobro huérfano registrado como C1-9) arrastran; se acreditan en agosto bajo C1-9.
- Re-correr todo julio (0_padron → 2_planilla → 5_cobranza → 6_corte) por S/9 es desproporcionado.

### Mejora estructural (opcional, anotar)
Falta en `02_matching` un paso de **dedup por nombre** (detectar mismo titular en 2+ lotes).
Hoy no existe → por eso Roberto pasó silencioso. Sería bueno revisar si hay más casos así en
`padron_reconciliado` antes de agosto.

---

## PENDIENTE 2 — Base de datos "deuda de la directiva" (caso Ronel) — ✅ CERRADO 2026-07-09

### Estado final (implementado siguiendo metodología)
```
(1) balde genérico 5b        ✅ _cargar_otros_conceptos() → Nivel 1a en verde (captura los 62)
(2) concepto controlado      ✅ token `deuda_directiva` + color en efectivo _CONCEPTO_BG/_TXT
(3) ledger append-only       ✅ 4_pagos/consolidar_deuda_directiva.py → shared/deuda_directiva.xlsx
                                 (BALDE 1 permanente · CICLO + dedup por (canal,ref,monto,fecha)
                                 · wired en 4_pagos/main.py paso 7 · contrato formato_deuda_directiva.html)
```
- **Fix del Gap del tanque:** el ledger NO regenera desde el mes (como aportes_tanque) → es
  append-only con columna CICLO, vive en shared/, 7_cierre no lo resetea. Sobrevive el cierre.
- **Ronel (caso julio, único):** fila escrita directo en el ledger (NOMBRE=Ronel, S/62, 2026-06)
  — decisión del usuario: no re-correr el motor por un caso que pasó 1 vez (no sobreingeniería).
  El pago yape sigue con CONCEPTO="saldo Ronel" en pagos_yape_tepago; el balde de 5b igual lo cuadra.
- **Captura futura yape:** para tagear deuda_directiva desde la fuente haría falta un campo de
  deudor en la resolución de pendientes del motor (hoy no existe). Diferido — no vale el cambio
  de schema hasta que haya ≥3 casos (Regla del Tres).

<details><summary>Diseño original (histórico)</summary>

### Qué pasó
Ronel (ex-tesorero de la directiva) tenía una deuda vieja y hoy la pagó. Entró como TE PAGÓ:
```
ORIGEN: Paul Tru*   MONTO: S/62   MENSAJE: "saldo Ronel"   FECHA: 28/06/2026 17:17:42
```
Al resolver el pendiente, se le puso CONCEPTO libre **"Saldo Ronel"** (se copió el mensaje).
No es agua, no es tanque → es una **deuda institucional de la directiva anterior**.

### El problema que causa hoy
`5b_validacion` Nivel 1a solo suma `agua + blancos + tanque`. No tiene balde para conceptos
libres → marca **falso descuadre de -62**. La plata SÍ cuadra (tepago total = 3461 = banco crudo
exacto); es solo que 5b no lo contabiliza.

### Lo que hay que hacer (DECISIÓN del usuario: crear la base de datos)
1. **Concepto controlado**: usar una etiqueta fija (ej. `deuda_directiva`) en vez de texto libre.
2. **Base de datos de deuda directiva**: crear el registro/ledger de estos pagos (patrón espejo
   del de tanque: `4_pagos/consolidar_tanque.py` → `outputs/aportes_tanque.xlsx`). Es decir un
   `consolidar_deuda_directiva.py` → `outputs/deuda_directiva.xlsx` que cosecha CONCEPTO=deuda_directiva
   desde las vistas de yape/efectivo. **Ojo: mismo Gap que tanque** — hay que darle acumulado
   entre meses (no efímero) desde el diseño.
3. **Arreglar `5b_validacion` Nivel 1a**: agregar el balde para que cuadre. Recomendación:
   balde **genérico** `otros_conceptos` = Σ tepago con concepto que no sea tanque, así cuadra
   con CUALQUIER concepto (deuda_directiva, y futuros) sin parchar 5b cada vez.
   Fórmula nueva: `banco TE PAGÓ = agua + blancos + tanque + otros_conceptos`.

> Nota de diseño: yo había recomendado NO crear módulo dedicado (regla de no sobre-ingeniería,
> pasó 1 vez). El usuario decidió crear la base de datos igual. Al implementar, evaluar si vale
> un módulo propio o basta el concepto controlado + balde genérico en 5b + registro simple.

</details>

---

## Contexto del ciclo julio — dónde quedó todo (2026-07-08)

### Estado del pipeline (julio)
```
motor_matching yape   ✅ pendientes=0 · ciclo cerrado y archivado (2026-07_procesado)
efectivo              ✅ Q-8 (María Martina Peña Vega) autorizado doble pago · discrepancias.xlsx eliminado
tanque               ✅ aportes_tanque.xlsx = 9 aportes S/1200 (Rossana G1-4 · ver caveat)
5_cobranza           ✅ corrió · seguimiento_pueblo.xlsx actualizado · planilla_cobrado.xlsx regenerada
5b_validacion        ⚠ NO pasó · 2 descuadres (ver abajo)
6_corte              ⚠ lista_corte.xlsx = BORRADOR (58 elegibles · 18 EJECUTAR=SI) — NO publicar
                        (se generó sobre números que 5b aún no validó)
```

### Descuadres que reportó 5b (a resolver antes de publicar/comprometer el corte)
```
-62 yape     → "saldo Ronel" (PENDIENTE 2). Plata cuadra; falta balde en 5b.
-33 efectivo → 3 cobros huérfanos (lote no existe en planilla):
   · C1-9  S/9   → Roberto Macarlopu (PENDIENTE 1). Real C1-9. Arrastra a agosto.
   · F-3B  S/8   → Abigail Gaspar Vega = F-3A. NO es error de padrón (reconciliado ya dice F-3A);
                   solo el cobro de mesa quedó mal tipeado. → re-imputar F-3B → F-3A.
                   (Comprobado: hoja PLANILLA de padron_secundario rotula "F-3-B"=Abigail;
                    monto S/8 calza con su deuda; no tiene otro pago. Salvedad: por monto solo
                    era ambiguo con Camila F-3, también debe 8 — desempata el rótulo.)
   · L-9   S/16  → mesa_2, Yreald, 04/07, comentario "Modificar". SIN resolver.
                   ❗ FALTA DATO del tesorero: ¿a qué lote real fue?
```

### Otros hallazgos (para no perder)
- **F1-6** → comentario en mesa_1: *"Mi lote es F1-7. Todos mis títulos están con F1-7."*
  Otra corrección de lote reportada, NO resuelta. Revisar si es otro caso de padrón como C1-9.
- **Rossana tanque (G1-4)**: los 2 aportes (S/100 + S/200) se atribuyeron a Rossana Samaritano
  (U0484, G1-4). El usuario avisó que **quizá uno sea de su hermano Jhon (U0485, G1-5)**. Se puede
  corregir editando `LOTE_FINAL` en `trazabilidad_2026_07.xlsx` hoja Ambiguos (4→5) + re-correr
  motor + consolidar_tanque — **solo mientras el ciclo julio siga vivo** (después pega el Gap de
  acumulado del tanque). El lote NO afecta el corte (tanque no toca deuda de agua).

### Regla de corte — confirmada leyendo el código (para que quede claro)
`generar_lista.py` + `config.py`: elegible = **`SALDO > 0` AND `MES_ANTERIOR >= 8`**.
`MES_ANTERIOR` es un **monto en soles** (deuda arrastrada del período anterior), no un contador;
8 ≈ una boleta mensual. O sea la regla real = **no pagó 2 meses seguidos** (debe del mes pasado
≥ 1 boleta Y debe este mes). `SALDO>0` solo NO alcanza (218 lotes con SALDO>0 → 58 con la 2ª
condición → 18 a ejecutar tras restar reclamos y pagos parciales).

---

## Cambios de código YA hechos esta sesión (no re-romper)

1. **`motor_matching` — primitivo `Segregacion`** (generaliza el viejo `Pagos_comunitarios`):
   - `CONCEPTO=comunitario | multiple` en Sin_identificar o Ambiguos → hoja `Segregacion` con
     columna `TIPO`. Persiste entre corridas. Resolvió el caso Alfredo Grados (multiple no
     detectado por regex). Docs y tests actualizados (9 tests nuevos pasan).
2. **Bug 1 fix (`motor_matching`)**: cuando CONCEPTO tiene valor + MZ/LOTE reales, ahora se
   **preserva** MZ/LOTE/USER_ID/NOMBRE (antes los borraba). Necesario para que consolidar_tanque
   sepa quién aportó. `estado_pago="concepto"` → sigue sin tocar deuda de agua.
3. **`5_cobranza` ignora concepto**: `_cargar_pagos_yape` y `_cargar_pagos_efectivo` saltan filas
   con CONCEPTO (no son agua). Cierra el riesgo del validador BLOQUE 2. (Este cambio es el que
   deja "saldo Ronel" fuera de la cobranza de agua, correcto.)
4. **`discrepancias.xlsx` efectivo**: nueva sección "¿Quién es?" (ID + NOMBRE) por lookup contra
   `usuarios_id.xlsx`, en ambas hojas. Contrato `formato_discrepancias.html` actualizado.

---

## Checklist para mañana (agosto)
- [ ] Override C1-9 en `0_padron` (Roberto) + corregir `usuarios_id.xlsx` → aplicar ANTES de 2_planilla agosto
- [ ] Revisar más duplicados mismo-nombre-2-lotes en `padron_reconciliado` (F1-6 sospechoso)
- [ ] Crear base/registro de `deuda_directiva` + concepto controlado
- [ ] Arreglar `5b_validacion` Nivel 1a (balde `otros_conceptos`)
- [ ] Resolver huérfano L-9 (falta dato del tesorero) y re-imputar F-3B → F-3A
- [ ] (Opcional julio) corregir atribución tanque Rossana/Jhon si un pago era de Jhon
- [ ] Recién con 5b en verde: revisar y publicar `lista_corte.xlsx` (BORRADOR → PUBLICADA)
