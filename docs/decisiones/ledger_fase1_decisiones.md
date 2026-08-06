# Ledger Fase 1 — Verdad única de decisiones

**Qué es esto:** el lugar ÚNICO donde vive la respuesta correcta de cada decisión de diseño
del ledger (Fase 1). Hoy las mismas preguntas tienen **varias verdades** dispersas en distintos
README/RETOMAR y se contradicen. Acá se colapsa cada una en **una sola verdad específica y sin
contradicción**.

**Cómo se usa:**
1. Bloque por bloque: se listan las "varias verdades" que hoy conviven + la decisión.
2. El usuario decide → se anota acá como VERDAD ÚNICA.
3. Con eso fijado, cada README se reconcilia para ser **específico y coincidir** con esta verdad
   (no la referencia con un puntero — la repite completa y consistente).

**Estado de cada bloque:** 🟢 decidido · 🟡 en discusión · ⚪ sin abrir

---

## ⚡ DÓNDE NOS QUEDAMOS — retomar acá (2026-07-17)

### PROPUESTA sobre los 4 pendientes de `tools/` — PRESENTADA, SIN APROBAR (retomar acá)
El 2026-07-17 se presentaron en consola las 4 decisiones con POR QUÉ. El usuario las
dejó para mañana (sin decidir). Recomendación puesta sobre la mesa:
- **④ ubicación:** `libro_mayor/tools/` **top-level** (no anidada en estado_cuenta). Las 4
  tools son read-only y cruzan los 2 agregados → superficie del contexto, no de un agregado.
  Los writers (`registrar_*`, `aplicar`) se quedan en su repo. Corregir la anidación del
  README de estado_cuenta.
- **① `estado_cuenta(mz,lt)`:** mantener separada de extracto = **snapshot AHORA** (máquina,
  hot-path para 2_planilla/6_corte/boleta): saldo_total, deuda_abierta[], saldo_a_favor,
  `meses_impagos` (lo que pide `politica_corte.evaluar`).
- **② `explicar_reclamo(ABONO_ID)`:** mantener, **llave=abono** (no predio) → sirve blancos
  sin predio que el extracto (predio-keyed) no alcanza. Devuelve la cascada de UN abono + conserva.
- **③ `auditoria_conservacion(jass_id, abono_id?/mz?/lt?)`:** **una tool, scope por parámetro**
  (global/abono/predio); **retorna violaciones como data, NUNCA raise** (self-check del agente).

Si el usuario dice "sí a las 4" → bajar a `estado_cuenta/README.md` (contrato) + lámina `tools`
del cuaderno. Si ajusta alguna → replantear esa sola.

### PLAN DE LOS 5 DÍAS (54 reclamos → reportes) — esbozado 2026-07-17, afinar mañana
Ver el detalle al pie de este archivo, sección "PLAN 5 DÍAS". Gate crítico = **la siembra**;
pregunta que mueve ±1-2 días = ¿el reporte solo MUESTRA (extracto) o además RESUELVE el reclamo
(identificar/reasignar/ajuste)?

---

Estamos construyendo **`libro_mayor/dominio/` firma por firma**, extrayendo cada una del
**código real** (no de los docs contradictorios) y fijando las decisiones acá como verdad única.

```
1  taxonomia.py     🟢 CERRADA   T1-T5 · 7 conceptos · cascada P1-P6 · sub · céntimos int
2  entidades.py     🟢 CERRADA   E1-E3 · CARGO_ID · frozen+validan(__post_init__) · Identif. propia
3  cascada.py       🟢 CERRADA   CA1-CA4 · aplicar(cargos,abono)→Reparto · SALDO_A_FAVOR · devol fuera
4  politica_corte   🟢 CERRADA   PC1-PC7 · evaluar(...,cfg)→(motivo,penalidad,salvado) · conductual N meses · config por trigger
5  saldo.py         🟢 CERRADA   S1-S5 · mes=agua+mant · impago=S/0 imputado · pura re-derivada · total · "seguidos" gratis por FIFO
6  identidad.py     🟢 CERRADA   I1-I4 · id=prefijo+shorthash · canon() central · abono predio-agnóstico · re-identif=evento (no re-mint)

✅ dominio/ COMPLETO — las 6 firmas cerradas (2026-07-17).
```

**RECONCILIACIÓN DE READMEs — HECHA (2026-07-17, Sonnet).** `dominio/` completo → se bajó a los
READMEs vivos:
- `libro_mayor/dominio/README.md` **creado** (no existía, solo `.gitkeep`) — spec consolidado de
  las 6 firmas (T·E·CA·PC·S·I), fuente para todo lo demás.
- `caja/README.md` + `estado_cuenta/README.md` (contrato byte-idéntico, verificado con `diff`):
  llave de ABONO efectivo corregida (`fecha` tipeada → `origen_archivo+fila` + gatillo de
  casi-duplicado con autorización) · cascada P1 aclara que `arrastre` no es concepto (T2).
- `6_corte/README.md` sección "diseño destino" reescrita: firma real de `politica_corte`
  (`evaluar(saldo, meses_impagos, ...)`, no la vieja `evaluar(saldo_cent, arrastre_cent, ...)`),
  umbral conductual (no `MES_ANTERIOR≥8`), 3 estados de PC5, salvado=regla universal PC2.
- `2_planilla/README.md` y `4b_reclamos/README.md` revisados — ya estaban consistentes, sin cambios.
- Pendiente, deliberadamente diferido: `formato_aplicacion.html` (HTML se toca después de los
  READMEs, decisión del usuario) · re-sembrar `SUB_CONCEPTO` histórico (Fase 2).

**Cuaderno `docs/cuaderno/libro_mayor.html` — 6 láminas completas** (dominio · caja · estado_cuenta
· motor · stores · tools), formato Altura 2 (cuadro) + Altura 3 (pregunta→cajas→POR QUÉ) en todas.
Validación estructural OK (`py html.parser`, 0 anclas rotas) en cada corte. Contenido de aprendizaje
real (no solo transcripción): ejemplo numérico de la cascada (S/45 repartido P1→P3, checksum),
trazado de 2 abonos en secuencia en `motor` (por qué el bucle no puede vivir en `cascada`, que es
pura por CA3), corrección explícita en `stores` (patrón puerto/adapter 🟢 cerrado, heredado del
lente — no "todo abierto") y en `tools` (madurez dispareja por tool, no un solo 🟡 para el módulo).

**PRIMER PASO al retomar — cerrar los pendientes de `tools/` (Opus, mismo tipo de sesión que ⑫).**
Hallazgo de esta sesión: a diferencia de `stores/` (bloqueada de verdad — depende de código de
`caja_repo`/`cuenta_repo` que no existe), los pendientes de `tools/` **no tienen bloqueo
estructural** — el contrato de `Cargo`/`Aplicacion`/`MovimientoCaja` ya está cerrado, alcanza para
decidir sin escribir código, igual que se cerró `⑫ extracto_predio` (5 decisiones de negocio,
2026-07-13, sin código). Faltan 4 cosas, mismo método (opciones A/B → decisión → POR QUÉ):
1. **`estado_cuenta.py`** (tool "panorama completo de un predio") — hoy solo el nombre en el
   README, sin campos ni formato de salida.
2. **`explicar_reclamo.py`** ("la historia de un pago: a dónde fue cada sol") — solo el nombre,
   sin formato.
3. **`auditoria_conservacion.py`** — el invariante que chequea YA está cerrado
   (`MONTO(abono) = Σ aplicado + Σ saldo_a_favor`); falta la interfaz de tool (¿por predio o
   global? ¿qué devuelve si falla la conservación?).
4. **Ubicación física de `tools/`** — `libro_mayor/README.md` la lista top-level (hermana de
   `caja`/`estado_cuenta`); `estado_cuenta/README.md` describe su propio `tools/` anidado adentro,
   con los mismos 4 archivos. Hoy en disco solo existe la carpeta top-level, vacía (`.gitkeep`).
   Pregunta real: si una tool cruza agregados (como `extracto_predio`), ¿vive arriba o dentro de
   quien "es dueña" del dato principal?

**Método (repetir en cada firma/decisión):** extraer del contexto real (contrato ya cerrado, no
inventar) → opciones A/B → decisión con POR QUÉ → fijar acá como verdad única → bajar a
`estado_cuenta/README.md` (contrato) y a la lámina `tools` del cuaderno.

**Sistema de docs armado hoy** (contexto, no parte de dominio):
- `docs/cuaderno/` = mini-sitio de aprendizaje (portada → módulo → submódulo colapsable → lámina).
- `docs/arquitectura/index.html` = mapa maestro (flujo + porqué + tipos motor/herramienta/config).

**Sesión 2026-07-16 (cont.) — método de diagramación de 3 alturas + lámina `libro_mayor`:**
- **Método fijado (3 alturas de zoom):** ① MAPA del sistema (módulos=cajas, flechas=archivos) ·
  ② FLUJO de un módulo (molde de 5: entra·hace·sale·quién lee·decisiones) · ③ DECISIÓN (code→destino).
  **Regla de oro:** Altura 2 y 3 van en la MISMA lámina de módulo (no aislar "qué hace" del "por qué").
  **Formato que enseña:** cada decisión se muestra como PREGUNTA→flujo de cajas (no un before/after seco)
  — las preguntas se recuperan del RETOMAR/backup, no se inventan.
- **`docs/cuaderno/cimientos.html`:** +1 lámina "Método de diseño" (las 3 alturas, con cascada CA1-CA4
  como ejemplo) · `id="modelado-tipos"` en la lámina de entidades · sección "Aprendizaje" (preguntas→láminas).
- **`docs/cuaderno/libro_mayor.html` (NUEVO):** Altura 1 (mapa del contexto: feeders→5 piezas→consumidores,
  motor=★, 3 puntos ciegos que cierra) + lámina `dominio/` con Altura 2 (cuadro de 5) y Altura 3 en orden
  de dependencia: **taxonomía ①-⑤** (formato pregunta→flujo, preguntas de `RETOMAR_dominio` "Bloques por
  abrir") · **entidades** = puntero a cimientos#modelado-tipos (no se duplica) · **cascada CA1-CA4** ·
  politica_corte/saldo/identidad ⚪ listadas para sumar al cerrar. Tarjeta `libro_mayor` activada en index.
- **Validación:** solo estructural (tags balanceados, anclas/links resueltos con `py html.parser`). NO se
  vio renderizado en navegador — pendiente confirmar visual con `start docs/cuaderno/libro_mayor.html`.
- **Al retomar dominio:** cerrar `politica_corte.py` por preguntas → fijar en Bloque A → bajar como bloque
  Altura 3 nuevo a la lámina `dominio/` de `libro_mayor.html` (misma página, se va llenando firma por firma).

**Sesión 2026-07-17 — `politica_corte.py` CERRADA (PC1-PC7) + regla del lente en CLAUDE.md:**
- **`politica_corte.py` cerrada por 7 preguntas** (extraídas del código real de `6_corte`:
  `config.py` + `generar_lista.py` + `aplicar_penalidad.py` + `seguimiento.py`). Verdad única en la
  tabla PC1-PC7 del Bloque A. Firma:
  `evaluar(saldo, meses_impagos, ya_cortado, en_revision, pago_ventana, cfg) → (motivo, penalidad, salvado)`.
- **Decidido con el lente de escala** (`docs/lente_escala.md`): política (números: umbral, penalidad) ≠
  mecanismo (la función). Los valores entran por manifiesto de tenant; agua y multa son **dos filas de
  config**, no dos ramas de código → `6b_corte_multas` NO es módulo (cut-trigger `multa` en el manifiesto).
- **Regla nueva en `CLAUDE.md` del proyecto:** al iniciar sesión, leer `docs/lente_escala.md` y confirmar
  con la frase *"ya leí el lente y lo voy a usar para mis recomendaciones"* antes de recomendar arquitectura.
- **Bajado a `docs/cuaderno/libro_mayor.html`:** bloque Altura 3 de `politica_corte` (7 preguntas con
  diagramas de cajas + POR QUÉ + ejemplos concretos, reescrito para que se entienda solo). Cuadro Altura 2
  → `politica 🟢`. "Por abrir" → solo `saldo` + `identidad`. Validado estructural (`py html.parser`).
- **PRECONDICIÓN abierta hacia `saldo.py`:** PC1 (umbral conductual) exige que `saldo.py` derive
  `meses_impagos` por predio (FIFO `MES_CARGO`) y lo mantenga al día. Sin eso, PC1 no se sostiene.
- **Siguiente sesión = `saldo.py`.**

**Sesión 2026-07-17 (cont.) — `saldo.py` (S1-S5) e `identidad.py` (I1-I4) CERRADAS → `dominio/` COMPLETO:**
- `saldo.py`: mes=AGUA+MANTENIMIENTO juntos (S1) · impago=S/0 imputado, parcial-frontera no cuenta
  (S2) · re-derivar siempre, cero estado — el pago solo apila, el conteo se calcula al leer, no en
  el pago (S3, el punto que hizo clic) · solo el total, sin desglose (S4) · "N seguidos" gratis por
  FIFO (S5). `identidad.py`: prefijo legible+shorthash (I1) · `canon()` centralizado (I2) · abono
  predio-agnóstico (I3) · re-identificar no re-mintea el id, es evento aparte (I4) — más el hallazgo
  colateral: la llave de efectivo original (con `fecha`) es frágil, corregida a procedencia.
- **`dominio/` completo (6/6 firmas).** Verdad única en Bloque A. Bajado al cuaderno
  (`docs/cuaderno/libro_mayor.html`, láminas `saldo` e `identidad`, con los keypoints de lo que
  hizo clic al usuario durante la sesión de preguntas).

**Sesión 2026-07-17 (cont., Sonnet) — reconciliación de READMEs + cuaderno `libro_mayor` completo (6 láminas):**
- Reconciliación mecánica de `caja/estado_cuenta/6_corte/dominio` READMEs contra el `dominio`
  cerrado (detalle arriba, en "⚡ DÓNDE NOS QUEDAMOS").
- Cuaderno: láminas `dominio · caja · estado_cuenta · motor · stores · tools` completas.
- **Hallazgo — `tools/` tiene pendientes cerrables SIN código** (a diferencia de `stores/`, que sí
  está bloqueada por falta de código de los repos). Ver detalle de las 4 cosas en "⚡ DÓNDE NOS
  QUEDAMOS" arriba — **siguiente sesión = Opus, mismo método que cerró `⑫ extracto_predio`.**
- **Freno de modelo respetado:** el usuario pidió cerrar esos 4 pendientes de `tools/` en esta
  misma sesión Sonnet; se paró antes de decidir nada de negocio nuevo (a diferencia de reconciliar,
  que es mecánico) y se dejó explícito acá para Opus.

---

## Bloque A · Spec de `dominio/` — 🟡 en discusión

### A1 · La verdad NO era "§6 vs §10" — es "código vs rediseño"

La fuente real es el **código** (`5_cobranza/main.py::_descomponer_saldo` línea 1741). El
destino es un **rediseño**, no una extracción literal. Verdad del código hoy:
```
P1 DEUDA_AGUA = mes_actual + mantenimiento + mes_anterior(arrastre) + blanco + devolucion
P2 CORTE_RECONEXION · P3 MULTA · P4 ACUERDOS · P5 CONVENIO
sin sub_concepto · sin OTROS · TOL=0.005 · store pueblo = {MULTA, ACUERDOS, CONVENIO}
corte SIEMPRE fue P2 (verificado)
```

### A2 · Dónde vive el spec — 🟢 DECIDIDO

Se consolida en **`libro_mayor/dominio/README.md`** nuevo (detallado, coincidente con esta
verdad). El RETOMAR queda como log histórico.

### `taxonomia.py` — decisiones (código → destino)

| # | Decisión | Verdad única |
|---|---|---|
| T1 | AGUA vs MANTENIMIENTO | 🟢 **SEPARADOS** — `AGUA` = consumo actual (mes_actual) · `MANTENIMIENTO` = concepto propio |
| T2 | "arrastre" (mes_anterior) | 🟢 **NO es concepto** — es AGUA de meses previos, resuelta por FIFO (`MES_CARGO`) |
| T3 | P6 OTROS | 🟢 **SE AGREGA** (aunque hoy no tenga emisor) |
| T4 | sub_concepto | 🟢 **SE AGREGA** — multa: reunión/faena · acuerdos: techado/campo · convenio: medidor/instalación |
| T5 | dinero céntimos int, sin TOL | 🟢 **SÍ** — plata en céntimos enteros, comparación exacta, TOL eliminado |

**Conceptos resultantes (7):** `AGUA` · `MANTENIMIENTO` · `CORTE_RECONEXION` · `MULTA` ·
`ACUERDOS` · `CONVENIO` · `OTROS`. **Cascada P1-P6.** Nombre canónico `ACUERDOS` (la planilla
lo llama `ACUERDOS_ASAMBLEA` — el feeder traduce al emitir).

### Candidatas de lámina (cuaderno) — escribir cuando `dominio` cierre

- **"Qué es una entidad"** (Cimientos, transversal · escribir cuando `entidades` cierre):
  tipo con forma fija + validación vs dict suelto; entidad = sustantivo (cosa) vs motor/tools =
  verbos (acciones); el flujo (feeders crean Cargo/MovimientoCaja → motor lee → crea Aplicacion);
  qué es `__post_init__` (guardia automático al construir). Los 3 diagramas que le sirvieron al
  usuario. Concepto reusable para las entidades de cualquier módulo — NO una lámina por `.py`.
- **"El código no es el diseño destino"** (método de modelado, transversal/dominio · 1 lámina, no 5):
  extraer código a un modelo obliga a **separar lo fusionado y decidir**, no transcribir.
  Evidencia de taxonomía: agua+mant venían sumados → se separan · `arrastre` parecía concepto →
  es FIFO de agua · `P6 OTROS` no existe en el código, es especulativo. Se reusa al extraer
  `cascada`/`saldo`/`politica_corte`. La taxonomía en sí (conceptos/orden) es SPEC → dominio/README, NO lámina.

### `entidades.py` — decisiones (🟢 CERRADA)

5 entidades frozen: `MovimientoCaja` · `Cargo` · `Aplicacion` · `Ajuste` · `IdentificacionAbono`.

| # | Decisión | Verdad única |
|---|---|---|
| E1 | Ajuste apunta al cargo por | 🟢 **`CARGO_ID`** (único/determinista). La tool recibe la llave humana (mz,lt,concepto,sub,mes) y calcula el id adentro |
| E2 | frozen + validan al construirse | 🟢 **SÍ** — `__post_init__` valida vs taxonomía (concepto/sub) + `monto_centimos>0` en Cargo y MovimientoCaja (Ajuste/devolución van con signo); frozen = append-only real |
| E3 | IdentificacionAbono | 🟢 **entidad propia** (no mueve plata; es corrección de identidad `abono_id → mz·lt·reclamo_id`) |

Todas llevan `recorded_at` (⑧ del contrato). Céntimos int en todo monto.

### `cascada.py` — decisiones (🟢 CERRADA 2026-07-16)

Firma: `aplicar(cargos_abiertos_de_1_predio, abono) → Reparto(aplicaciones, saldo_a_favor_centimos)`
· `clave_orden(c) = (prioridad, indice_sub, mes_cargo)`. Rediseño de
`5_cobranza/main.py::_descomponer_saldo` (L1741): 5 comps floats → 6 prioridades con sub + FIFO.

| # | Decisión | Verdad única |
|---|---|---|
| CA1 | orden del reparto | 🟢 **P1-P6 → sub-orden → FIFO por `MES_CARGO`**. sub: multa reunión→faena · acuerdos techado→campo · convenio medidor→instalación. `clave_orden=(prioridad, indice_sub, mes_cargo)` |
| CA2 | sobrante del pago | 🟢 **`SALDO_A_FAVOR`** (concepto explícito, céntimos — contrato ③). El código viejo lo tiraba (`restante` descartado) |
| CA3 | granularidad | 🟢 **1 abono, función pura**; el **motor** ordena por fecha e itera N abonos. cascada no sabe de tiempo ni I/O |
| CA4 | devolución | 🟢 **fuera de la cascada** — es egreso de caja vs `SALDO_A_FAVOR`, lo maneja el motor. El código viejo la sumaba dentro de P1 agua |

### `politica_corte.py` — decisiones (🟢 CERRADA 2026-07-17)

Firma: `evaluar(saldo, meses_impagos, ya_cortado, en_revision, pago_ventana, cfg) → (motivo, penalidad, salvado)`
· `cfg` = manifiesto del tenant, por trigger: `{umbral_meses, penalidad_base, permite_salvarse}`.
Rediseño de `6_corte`: `generar_lista` (elegibilidad Día 0) + `seguimiento` (salvado Día 2) colapsan en
**1 función pura**; el motor la llama en Día 0 y Día 2. Decidido con el **lente de escala** (`docs/lente_escala.md`):
los valores no se hardcodean, entran por config; agua y multa son dos filas de manifiesto, no dos ramas.

| # | Decisión | Verdad única |
|---|---|---|
| PC1 | umbral de corte | 🟢 **conductual "N meses impagos seguidos"**, N=config(tenant,trigger) (agua N=2). El monto `MES_ANTERIOR≥8` era proxy (planilla independiente solo veía 1 mes; deuda 1-7=parcial→no cortar). El ledger ahora da el conteo real. **PRECONDICIÓN:** `meses_impagos` derivado y mantenido por `saldo.py` (FIFO por `MES_CARGO`) |
| PC1b | qué es "mes impago" | 🟢 **pagó S/0 ese mes**. Un pago parcial NO cuenta como impago para el umbral (regla universal, no varía por JASS) |
| PC2 | qué salva del corte | 🟢 **cubrir la penalidad** (`pago_ventana ≥ penalidad`). Un parcial menor NO salva. Se retira el gate Día-0 "cualquier pago salva". `permite_salvarse` = bool de config por trigger |
| PC3 | granularidad | 🟢 **función pura de 1 instante**; el MOTOR la llama Día 0 y Día 2 (misma fn, otro input). No sabe de tiempo ni I/O — análogo a cascada CA3 |
| PC4 | penalidad base vs escalada | 🟢 **base = config(trigger)**; el motor la escala tras la ventana de gracia. `politica_corte` no sabe de días (consistente con PC3) |
| PC5 | ya_cortado / exoneración | 🟢 **3 estados**: `activo` \| `cortado` \| `exonerado`. EXONERADO con subtipo: **mensual** (CADUCA al mes, motivo obligatorio: enfermedad/reclamo/verificación) · **permanente** (junta, vejez). Guarda `{tipo, motivo, periodo}`, con `JASS_ID` (lente §5) |
| PC6 | multa | 🟢 **mismo motor, otra cfg**: `umbral_meses=0` (cualquier multa impaga corta), `permite_salvarse=no`, `penalidad_base=20` que **escala a 40 igual que agua**. `6b_corte_multas` = cut-trigger `multa` en el manifiesto, NO módulo aparte (lente, precedente ya fijado) |
| PC7 | forma de salida | 🟢 **motivo = solo el trigger** ∈ {`agua`, `multa`, `""`}; `salvado`=bool. La razón del no-corte (reclamo/pago) la reconstruye el motor de los inputs `en_revision`/`pago_ventana` — NO va en el retorno (un campo, un trabajo) |

**Contexto de las decisiones no obvias** (lo que preguntó el usuario y hay que recordar):

- **PC1 · por qué `MES_ANTERIOR ≥ 8` no era arbitrario.** Cada planilla era **independiente** y solo
  guardaba el pago del mes anterior — no había historia de pagos. Como la deuda mínima de consumo es S/8,
  una deuda de 1-7 delataba que la persona **pagó algo** (parcial) → a esos no se los cortaba. El S/8 era
  un **proxy** de "no pagó nada". El ledger ahora sí tiene el pago de cada mes → se cuenta conducta directa,
  el proxy se retira.

- **PC1b vs PC2 · no se contradicen, son DOS momentos distintos.** PC1b (entrar a la lista): un pago
  parcial te **frena** — ese mes no cuenta como impago, no acumulás hacia el umbral. PC2 (ya estás en la
  lista): para **salvarte ahí** sí o sí cubrís la penalidad; un parcial menor NO salva. Un momento premia
  pagar algo; el otro exige pagar la penalidad completa.

- **PC3 · "función pura de 1 instante" = calculadora que no sabe qué día es.** `evaluar()` recibe el estado
  de un predio AHORA y devuelve el veredicto AHORA; no sabe de Día 0/Día 2, no toca archivos. El **motor**
  (que sí sabe fechas) la llama 2 veces: Día 0 `evaluar(impagos=2, pago_ventana=0)` → entra a la lista;
  Día 2 `evaluar(impagos=2, pago_ventana=25)` → salvado / `pago_ventana=5` → cortado. Misma función, otro
  input. Mismo patrón que cascada CA3.

- **PC6 · agua y multa NO son dos códigos.** Una sola `evaluar()`; lo que cambia es la `cfg`:
  `cfg_agua={umbral 2, pen 20→40, salva sí}` vs `cfg_multa={umbral 0, pen 20→40, salva no}`. `evaluar(impagos=0,
  cfg_multa)` con umbral 0 = "cualquier multa impaga ya corta". La multa **comparte** la penalidad 20→40 del
  agua (no trae número propio).

- **PC7 · la tupla se lee sin interpretar texto.** `("agua",20,False)`=corta por agua · `("agua",20,True)`=iba
  a cortar pero se salvó · `("",0,False)`=motivo vacío=no hay disparador=no corta. El "por qué se salvó"
  (reclamo vs pago) lo reconstruye el motor de los inputs, no viaja en el retorno.

**Fuente:** los diagramas de cajas + ejemplos completos de cada decisión viven en
`docs/cuaderno/libro_mayor.html` (lámina `dominio/`, bloque `politica_corte`). Esta tabla es la verdad única;
el cuaderno es el material que la enseña.

### `saldo.py` — decisiones (🟢 CERRADA 2026-07-17)

Dos firmas puras: `saldo(cargos, aplicaciones) → int céntimos` (deuda total del predio) ·
`meses_impagos(cargos_agua, aplicaciones) → int` (**precondición de PC1**). Rediseño de `6_corte`,
que hoy adivina la conducta con el proxy `MES_ANTERIOR ≥ 8`. `saldo.py` **no re-corre la cascada** —
lee las `Aplicacion` que el motor ya creó y las suma por cargo. Cinco preguntas → S1-S5.

| # | Decisión | Verdad única |
|---|---|---|
| S1 | qué es "un mes" para el conteo | 🟢 **AGUA + MANTENIMIENTO** (la boleta del `MES_CARGO`, no un concepto suelto). impago ⟺ S/0 imputado a ambos cargos del mes |
| S2 | qué es "mes impago" bajo FIFO | 🟢 **el mes que recibió S/0 imputado**; el mes-frontera parcial NO cuenta (puso algo → enganchado, PC1b exacto). NO es "mes con saldo abierto" |
| S3 | ¿guardar el conteo o re-derivar? | 🟢 **re-derivar** (función pura, cero estado). En el pago solo se APILA el evento; el número se calcula al LEER (cuando PC lo pide). Igual que el saldo = deuda − Σpagos. Coherente con CA3/PC3 |
| S4 | `saldo()` total o desglose | 🟢 **solo el total** (int céntimos = Σcargos − Σaplicaciones, piso 0 por cargo). El desglose por concepto/mes lo arma `estado_cuenta` de los mismos datos — no se duplica |
| S5 | "N meses seguidos" | 🟢 **gratis por FIFO** — imputar al más viejo primero hace imposible un mes pagado en medio de impagos; los S/0 son la cola contigua. `meses_impagos` cuenta desde el mes actual hacia atrás, para en el 1er mes con pago. Sin código de consecutividad |

**Contexto de las decisiones no obvias** (lo que preguntó el usuario y hay que recordar):

- **S3 · "re-derivar" ≠ "recalcular en cada pago".** El pago NO dispara un recuento; solo apila el hecho
  en el ledger (append, cero cálculo). El conteo se deriva **al leer**, cuando `politica_corte` lo pregunta.
  Por eso no se desincroniza: *escribir no calcula, leer sí*. Es idéntico a cómo el ledger nunca guarda el
  saldo — lo re-deriva de `deuda − pago1 − pago2`. Un contador guardado sería el bug de writer único otra vez.
- **S5 · FIFO = priorizar lo antiguo.** Como cada abono cierra el mes más viejo primero, nunca puede haber un
  mes pagado "en medio" de impagos → "seguidos" no necesita lógica aparte, es una propiedad heredada del orden.

**Fuente:** los diagramas de cajas (opciones A/B + POR QUÉ) viven en `docs/cuaderno/libro_mayor.html`
(lámina `dominio/`, bloque `saldo`). Esta tabla es la verdad única.

### `identidad.py` — decisiones (🟢 CERRADA 2026-07-17)

Acuña el nombre único de cada hecho. Firmas: `cargo_id(cargo) → str` · `abono_id(mov) → str` ·
`canon(mz, lt, concepto, sub, mes) → tuple` (normaliza antes de generar el id). Consolida en la capa pura
las reglas ya fijadas en el contrato (① `ABONO_ID` · `CARGO_ID`) y resuelve sus contradicciones. **Qué tiene
id propio:** solo las COSAS que nacen — `Cargo` (deuda) y `MovimientoCaja` (abono). Los vínculos (`Aplicacion`,
`Ajuste`, `IdentificacionAbono`) heredan el id de lo que conectan, no acuñan uno. Cuatro preguntas → I1-I4.

| # | Decisión | Verdad única |
|---|---|---|
| I1 | formato del id | 🟢 **prefijo legible + shorthash** (`{jass}-{mz}-{lt}-{concepto}-#hash` cargo · `{jass}-{mes}-{canal}-#hash` abono). Resuelve la contradicción del contrato: el `[:8]` es del SUFIJO hash, no del id entero. El prefijo particiona el espacio → hash corto sigue seguro (hash puro de 32 bits colisiona al tope) · legible y debuggable |
| I2 | canonicalización | 🟢 **`identidad.py` dueño de un solo `canon()`** (MZ→upper · LT→int · concepto canónico · mes→YYYY-MM). El determinismo VIVE acá: si dos entradas "iguales" no canonizan igual, el id se duplica. Unifica los `_norm_mz/_norm_lt` hoy repetidos por módulo |
| I3 | ¿el abono lleva el predio? | 🟢 **no — abono predio-agnóstico** (yape: clave `(jass, origen, timestamp)`; Hueco 1: 1 depósito=1 abono aunque pague N predios, el reparto vive en las APLICACIONES). Efectivo SÍ lleva mz/lt (capturado en mesa, sin motor_matching) |
| I4 | ¿re-identificar cambia el id? | 🟢 **no** — el `ABONO_ID` sale de canal+ref, no del predio → nunca cambia. La identidad del DUEÑO se corrige con un HECHO append-only (`identificar_abono`/`reasignar_abono` → entidad `IdentificacionAbono`); el motor lo lee y re-aplica. Identidad del dinero ≠ identidad del dueño |

**Contexto de las decisiones no obvias** (lo que preguntó el usuario y hay que recordar):

- **Qué es una identidad · quién tiene id propio.** Identidad = nombre estable y determinista → re-sembrar el
  mismo hecho real saca el mismo id → reconozco "ya lo tenía" (idempotencia). Solo las COSAS que nacen
  (cargo=deuda, abono=pago) acuñan id propio; los vínculos (aplicación, ajuste, identificación) se identifican
  por el PAR que conectan (`(ABONO_ID, CARGO_ID)`, etc.) — inventarles un id sería un nombre que nadie referencia.
- **CARGO_ID ≠ ABONO_ID en formato.** El cargo lleva `mz-lt` (una deuda siempre es de un predio); el abono NUNCA
  (`{jass}-{mes}-{canal}-#hash`). No existe "cargo blanco"; sí "abono blanco" = `DESTINO=PENDIENTE` (un VALOR de
  campo, no un id incompleto). El id del abono está completo desde que entra la plata.
- **El hash no salva un campo editable (bug real de efectivo).** La llave de efectivo incluye `fecha`; un typo
  re-corrido duplica el pago aunque haya hash (hash distinto = "otro pago"). Regla: la identidad se arma SOLO con
  campos estables; lo editable es atributo. Fix: anclar efectivo a la PROCEDENCIA (`origen_archivo + fila`) +
  gatillo de casi-duplicado con autorización humana. → deuda a reconciliar en decisión ①.
- **A (mutar) vs B (sumar).** Corregir el dueño de un blanco: A recalcularía el id (rompe FKs, no idempotente);
  B deja el id estable y AGREGA `identificar_abono(...)`. Un ledger append-only nunca edita/borra el pasado,
  corrige agregando — para eso existe. Mismo principio que el saldo (`deuda − pago1 − pago2`).

**Fuente:** los diagramas de cajas (A/B + POR QUÉ) viven en `docs/cuaderno/libro_mayor.html`
(lámina `dominio/`, bloque `identidad`). Esta tabla es la verdad única.

---

## Bloques por abrir (orden tentativo)

- **B · Cascada** — ¿`corte_reconexion` fue siempre P2? ¿el único cambio P1-P5→P1-P6 es agregar P6 OTROS? ⚪
- **C · "arrastre"** — ¿concepto/token de taxonomía, o solo AGUA/MANT de meses previos por FIFO? (7 u 8 conceptos) ⚪
- **D · SUB_CONCEPTO genérico** — ¿existe sub genérico o las fuentes traen el split para todos? ⚪
- **E · `seguimiento_repo`** — ¿status = "existe y se disuelve en estado_cuenta"? ⚪
- **F · Planilla en el destino** — ¿conserva columnas MULTA/CONVENIO/ACUERDOS_ASAMBLEA o se derivan del ledger? ⚪
- **G · Proceso** — ¿se pueden editar los RETOMAR (logs) o son historia inmutable? ⚪
