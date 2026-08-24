# RETOMAR — corrección dominio + Extracto de cuenta + cascada P1-P5 · 2026-07-13

Sesión larga, dos bloques: **(A) diseño del ledger** (Opus → transcrito en Sonnet,
esta misma sesión) y **(B) trabajo operativo de corte** (Haiku/Sonnet, mismo día,
detalle al final — dejó 2 bugs de código corregidos y verificados).

Amplía `docs/retomar/RETOMAR_ledger_contrato_final_2026-07-11.md` y
`docs/retomar/RETOMAR_libro_mayor_2026-07-12.md` (siguen vigentes como historia).

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **Bloque A (diseño) — TRANSCRITO A LOS README, verificado byte-idéntico.** No
   queda nada pendiente de escribir de lo que se cerró hoy. Lo que falta es
   **diseñar más capacidades** (backlog abajo, sección "Pendiente Fase 1 — Opus").
2. **Bloque B (operativo) — RESUELTO.** Lista de corte de julio corregida y
   generada: 2 bugs de código encontrados y arreglados (ancla de corte, canal
   yape no protegía de corte). 14 predios a cortar, validado. Ver detalle abajo.
3. **Siguiente paso recomendado:** volver a Opus y diseñar las capacidades del
   backlog (`morosidad_total`, `riesgo_corte`, etc.) — especialmente
   `riesgo_corte`, porque el bug de hoy mostró que "2 meses de deuda" ≠
   `MES_ANTERIOR>=8` (el umbral actual es ~1 mes, no 2 — sin resolver).
4. **Nada de código nuevo del ledger todavía.** Sigue Fase 1 (spec). Falta crear
   `formato_extracto.html` (Sonnet) y actualizar 4 HTML de estado_cuenta que
   quedaron con la cascada vieja (Sonnet).

---

## BLOQUE A — Diseño del ledger (Opus → transcrito Sonnet, 2026-07-13)

### Cómo se llegó a esto

El usuario pidió enumerar todas las preguntas que el ledger debe responder
("te pagué en marzo, busca mi pago" · "cuánto entró y salió" · "mi reporte de
seguimiento en PDF"). Al diseñar la 3ª (el **Extracto de cuenta**), el usuario
corrigió dos veces al diseño:

1. Sobre el concepto `comunitario` (que el handoff previo había clasificado mal).
2. Sobre el detalle del Extracto: pidió mostrar la deuda **desglosada por
   sub-concepto** (multa→faena/reunión, convenio→medidor/instalación,
   acuerdos→campo/techado) y **cómo se distribuyó cada pago** — no solo el total.

Eso obligó a diseñar 4 cosas antes de poder cerrar el Extracto: (1) quién imprime
(arquitectura de render), (2) la taxonomía completa + orden de prelación, (3) el
modelo de CARGO con `SUB_CONCEPTO`, (4) el layout final.

### 1 · Corrección de dominio — dos errores en la decisión ⑩ del contrato anterior

**Error 1 — `comunitario` NO es un balde.** Verificado contra código:
```
"comunitario" en 5b (main.py:560)     → ADJETIVO: "tanque comunitario"
                                         (el tanque es propiedad de la comunidad)
"comunitario" en motor_matching       → MECANISMO DE SEGREGACIÓN
  (readme_motor líneas 312-315)         cobrador presta su Yape, agrega N vecinos,
                                         envía 1 depósito → se desgloza por lote
                                         (hoja Segregacion: PADRE_SEGREGADO→N×HIJO_SEGREGADO)
```
El handoff del 12/07 confundió el adjetivo de 5b con el mecanismo de
motor_matching y lo clasificó como balde de egreso voluntario. Es incorrecto: un
depósito `comunitario` es un INGRESO normal (`BALDE=agua` mayormente) que el
motor reparte en N aplicaciones — ya lo cubre la decisión ① (`ABONO_ID` sin
mz/lt). **Efecto:** `buscar_abono` debe buscar también en los `HIJO_SEGREGADO`
de motor_matching, no solo en depósitos de primer nivel — el pago de agua de un
vecino puede estar enterrado dentro del depósito de su cobrador.

**Error 2 — `deuda_directiva` NO cruza a `estado_cuenta`.** El contrato anterior
la listaba como `SOURCE` de un CARGO (`otros`). Es un caso específico: **dos
miembros de la directiva anterior repagando un faltante de caja detectado** — no
es deuda de ningún predio. Mismo tratamiento que `tanque`: balde de INGRESO
caja-only, ya reconciliado en `5b_validacion` como "otros conceptos"
(`_cargar_otros_conceptos`, Nivel 1a). Nunca genera CARGO, nunca entra a la
cascada de prioridad.

### 2 · Decisión ⑪ — `SUB_CONCEPTO` en el CARGO + cascada de prioridad P1-P5

**Árbol de conceptos**, verificado contra `shared/seguimiento_pueblo.xlsx`
(3 conceptos reales: MULTA 582 eventos · ACUERDOS 463 · CONVENIO 374) +
`5_cobranza/main.py::_descomponer_saldo()` (línea 1741 — la cascada P1→P5 **ya
está codificada**, el motor de aplicación del ledger la replica exacto):

```
P1  AGUA (por MES_CARGO) · MANTENIMIENTO · arrastre        FIFO por mes, sin sub
P2  CORTE_RECONEXION                                        sin sub
P3  MULTA          → sub: REUNION primero, luego FAENA
P4  ACUERDOS       → sub: TECHADO primero, luego CAMPO
P5  CONVENIO       → sub: MEDIDOR primero, luego INSTALACION

DEUDA_DIRECTIVA — fuera de la cascada (ver corrección arriba, es balde no cargo)
```

**Por qué ese sub-orden** (regla de negocio dada por el usuario):
| Concepto | 1º | 2º | Por qué |
|---|---|---|---|
| MULTA | reunión | faena | la faena se puede pagar con trabajo (doblar a 8h en vez de 4h); la reunión **nunca** se paga con trabajo — el dinero cubre primero lo que solo el dinero puede cubrir |
| ACUERDOS | techado | campo | techado es el monto más bajo |
| CONVENIO | medidor | instalación | medidor es deuda chica; instalación es grande, se paga de a pocos |

El CARGO pasa de `(CONCEPTO, MES_CARGO)` a `(CONCEPTO, SUB_CONCEPTO, MES_CARGO)`
— cambia la decisión ② del contrato. **Data histórica sembrada sin desglose**
(faena+reunión sumados, medidor+instalación sumados) queda con `SUB_CONCEPTO`
genérico hasta re-sembrar — la separación se hace **desde el diseño** (el CARGO
ya soporta el campo), la siembra la completa después, en Fase 2 (confirmado por
el usuario: "esa separación lo haremos más adelante a la hora de sembrar").

### 3 · Decisión ⑫ — Extracto de cuenta (diseño CERRADO)

Responde el pedido directo: *"quiero mi reporte de seguimiento de todos los
meses / últimos N, con todos mis pagos, en PDF o impreso."*

```
BOLETA      = 1 mes, "esto debes ahora"          → cobrar
EXTRACTO    = N meses, "esto pasó en tu cuenta"   → auditar / resolver reclamo
```

**Las 5 decisiones:**
| # | Decisión | Razón |
|---|---|---|
| 1 | Clave = por PREDIO (MZ-LT), no por persona | sistema predio-céntrico; identidad-persona no se construye por caso raro |
| 2 | Tanque en sección aparte, informativo, no afecta SALDO | aporte voluntario, no obligación del predio |
| 3 | Rango default = TODO el histórico; desde/hasta opcional | uso real = reclamos, no se sabe el mes disputado |
| 4 | Trigger = usuario on-demand, 1 predio, con o sin reclamo | no es batch de asamblea |
| 5 | Template nuevo (no reusa PLANTILLA_boletas) | layout distinto (tabla/libreta vs boleta) |

**Arquitectura:** `extracto_predio(mz, lt, desde, hasta)` — tool read-only en
`estado_cuenta/tools/`, cruza cargos+aplicaciones de estado_cuenta con abonos de
caja. Devuelve filas, no un PDF.

**Layout — 3 secciones** (pendiente crear `formato_extracto.html`, Sonnet):
```
① DEUDA POR CONCEPTO→SUB (estado actual, con prioridad P1-P5)
   Pri | CONCEPTO | SUB | CARGADO | PAGADO | DEBE

② PAGOS RECIBIDOS Y CÓMO SE APLICARON (traza la cascada, por pago)
   FECHA | MONTO | "se aplicó a: agua mar, agua abr, mant, corte,
                    multa reunión, multa faena (parcial, faltó X)"

③ APORTES VOLUNTARIOS (tanque) — informativo
   SALDO ACTUAL DE DEUDA: <derivado>
```

### 4 · Arquitectura de render (decisión de infraestructura, no de contrato)

Pregunta que abrió esto: *¿quién imprime el extracto? ¿3_boletas renombrado a
3_impresor, un módulo nuevo, o algo distinto?*

**Decisión: ningún módulo de negocio imprime.** Cada bounded context arma sus
propias filas (sabe qué significa el documento); un servicio compartido y
**stateless** convierte `(plantilla, filas) → PDF`:

```
estado_cuenta              3_boletas                (cada uno arma SUS filas)
extracto_predio()          data_boletas()
      │                         │
      └───────────┬─────────────┘
                   ▼
   RENDER  (hoy: shared/utils_render.py — no existe aún, crear en Fase 2)
   NO sabe qué es boleta ni extracto — solo (plantilla, filas) → PDF
                   │
                   ▼
                  PDF
```

**Por qué NO `3_impresor` único** (lo que el usuario imaginaba primero): mezclaría
2 mundos de datos (planilla + ledger) en un solo dueño, y los motores de render
probablemente difieren (boleta = sustitución de campos en .docx; extracto = tabla
con saldo corriente). El render se separa **porque es infraestructura
cross-cutting** (como una DB), no por código duplicado — no aplica la Regla del
Tres. Es candidato a escalar en su propio contenedor Docker mañana (PDF es
CPU-pesado) sin que 3_boletas ni estado_cuenta cambien cómo lo invocan.
`3_boletas` se queda donde está — no se renombra ni se mueve.

### 5 · Qué se transcribió hoy (verificado)

```
libro_mayor/caja/README.md                    ← corrección ⑩ + ⑪ + ⑫ (contrato)
libro_mayor/estado_cuenta/README.md           ← IDÉNTICO byte a byte (verificado con script)
libro_mayor/README.md                         ← estado + resumen de sesión + pendientes
libro_mayor/caja/docs/formato_evento_caja.html ← tabla de baldes corregida
```
Verificación de byte-identidad corrida con script Python — `True`, 21920
caracteres en ambos desde el marcador `CONTRATO DE INTERFAZ`.

### 6 · Pendiente Fase 1 — falta transcribir (Sonnet, diseño ya cerrado)

- [ ] `estado_cuenta/docs/formato_extracto.html` — crear (layout de arriba)
- [ ] `estado_cuenta/docs/diagrama_flujo_estado_cuenta.html` — actualizar a
      cascada P1-P5 + SUB_CONCEPTO (quedó con "CONSUMO→CORTE→...→OTROS" viejo)
- [ ] `estado_cuenta/docs/diagrama_estado_cuenta.html` — ídem
- [ ] `estado_cuenta/docs/formato_cargo.html` — agregar columna SUB_CONCEPTO
- [ ] `estado_cuenta/docs/formato_aplicacion.html` — CARGO ahora es
      `(CONCEPTO, SUB_CONCEPTO, MES_CARGO)`

### 7 · Pendiente Fase 1 — falta DISEÑAR (Opus, no transcripción)

Backlog de capacidades enumeradas el 13/07 pero **sin spec** (I/O, columnas,
reglas). No confundir "enumerada" con "cerrada" — esto quedó claro tras el
usuario corregir que "nuevo" sí requiere diseño Opus:

| Capacidad | Para quién | Peso de diseño |
|---|---|---|
| `morosidad_total()` | tesorero | liviano |
| `lista_morosos()` | tesorero | liviano |
| `historial_predio(mz,lt)` | auditor | liviano |
| `trazar_abono(abono_id)` | auditor | liviano |
| `explicar_saldo(mz,lt)` | vecino | medio |
| `recaudado_por_concepto(desde,hasta)` | tesorero | medio |
| `estado_cuenta(mz,lt,fecha)` — query temporal | auditor | medio |
| **`riesgo_corte`** | vecino/tesorero | **real** — reemplaza parte de `6_corte`; la definición de "2 meses de deuda" quedó abierta en el bloque B (ver abajo), es la misma ambigüedad |
| **`arqueo_caja(mes)`** | tesorero | **real** — reemplazaría `5b_validacion` |

**Recomendación para la próxima sesión:** empezar por `riesgo_corte`, porque el
bloque B de hoy mostró en producción que la regla actual (`MES_ANTERIOR>=8` en
`6_corte/config.py`) es ambigua — 8 es ≈1 mes, no 2. Diseñar esta capacidad
**obliga** a cerrar esa definición de una vez, con beneficio directo para
`6_corte` incluso antes de que el ledger esté implementado.

---

## BLOQUE B — Trabajo operativo: lista de corte de julio (mismo día, resuelto)

Contexto: se pidió generar la lista de corte de julio (penalidad S/20, no S/40 —
ya estaba bien en `config.py`). Aparecieron 2 bugs reales de código al auditar
por qué predios que habían pagado seguían apareciendo como deudores.

### Bug 1 — ancla de corte mal calculada (motor_matching)

```
CAUSA: obtener_ancla() (4_pagos/yape/motor_matching/main.py) elegía el
       *_procesado.xlsx MÁS RECIENTE en disco, no el del ÚLTIMO CICLO CERRADO.
       Julio (ABIERTO) ya tenía su propio 2026-07_procesado → la ancla se
       auto-avanzaba a 07/07 → todo pago del 16/06 al 07/07 quedaba excluido
       (fecha > ancla, filtro de cargar_reportes).

SÍNTOMA: K3 y K4 habían pagado su deuda completa el 06/07 y seguían con
         SALDO=46 y SALDO=135 — el pago estaba en el crudo pero no en el
         output pagos_yape_tepago.xlsx (solo 5 pagos sobrevivían, todos del 08/07).

FIX: obtener_ancla() ahora lee estado_ciclo.json y usa el _procesado del
     ÚLTIMO CICLO CERRADO (junio → 15/06), con fallback al comportamiento
     viejo si falta el JSON. Función nueva _ciclos_cerrados().

RESULTADO: ancla 07/07 → 15/06. Identificados en motor_matching: 5 → 63.
           K3/K4 acreditados (SALDO → 0). Lista de corte: 62 → 58 usuarios,
           22 → 18 a cortar.
```

### Bug 2 — pago parcial por yape no protegía del corte (generar_lista.py)

```
CAUSA: generar_lista.py (6_corte) solo miraba pagos_efectivo.xlsx para la
       regla "pagó algo este mes → no cortar". Un pago parcial por YAPE no
       contaba — inconsistencia de canal (efectivo sí salva, yape no).

SÍNTOMA: B-17 (Luz Dina Miraval) pagó S/68 por yape el 07/07, quedaba con
         SALDO=25 (PARCIAL), pero salía EJECUTAR_CORTE=SI de todos modos.
         Se encontraron 4 casos iguales: A-7, B-17, F-1, L-10.

FIX: _filtrar_corte() ahora chequea `mesa or monto_yape > TOL` (antes solo
     `mesa`). Nueva columna leída: MONTO_YAPE de planilla_cobrado.xlsx.

RESULTADO: 18 → 14 a cortar. Los 4 quedan con motivo "Pago parcial por yape".
```

### T-9 / T-17 — investigados, NO son anomalía

MULTA=120 es única en las 499 filas del padrón (solo estos 2 predios) pero está
trazada: sembrada a mano el 2026-07-02 18:19 (`sembrar_seguimiento_pueblo`,
`AUDIT_REF=siembra_2026-06_MULTA_T_9/17`). El usuario confirmó: "probablemente
ya han sido cortados antes" — coherente con `CORTE_RECONEXION=40` ya cargado en
ambos. **Nota menor sin resolver:** no figuran en `registro_cortes.xlsx` — si ya
están físicamente cortados, falta registrarlos ahí para que no reaparezcan en
listas futuras. Queda para quien opere el corte físico, no bloquea nada.

### Estado final de la lista de corte julio 2026

```
6_corte/outputs/lista_corte.xlsx — 58 en lista, 14 EJECUTAR_CORTE=SI
   · 15 protegidos por reclamo EN_REVISION (50 PENDIENTE → EN_REVISION al inicio de sesión)
   · 25 protegidos por pago parcial en mesa (efectivo)
   · 4 protegidos por pago parcial yape (fix de hoy)
   · 13 excluidos (ya CORTADO/EXONERADO en registro_cortes)
Validado: PENALIDAD=20 en las 58 filas · 0 reclamos EN_REVISION con EJECUTAR=SI
Siguiente paso (no ejecutado, pendiente de decisión del usuario):
   aplicar_penalidad.py — procesa los 14 EJECUTAR=SI
```

Cadena completa re-corrida hoy, en orden: `4b_reclamos` (marcar EN_REVISION) →
`4_pagos/yape/motor_matching` (fix ancla) → `5_cobranza --force` (2 veces, tras
cada fix) → `6_corte/generar_lista.py` (2 veces).

---

## Orden sugerido al retomar

1. Si se va a cortar de verdad: revisar los 14 nombres de `lista_corte.xlsx` y
   correr `aplicar_penalidad.py` cuando el usuario confirme.
2. (Sonnet) Transcribir lo que falta del Bloque A sección 6 (formato_extracto +
   4 HTML de estado_cuenta) — rápido, diseño ya cerrado.
3. (Opus) Diseñar `riesgo_corte` primero (resuelve la ambigüedad de "2 meses" que
   quedó abierta) — después el resto del backlog de la sección 7.
4. Recién con el backlog de capacidades cerrado, evaluar si se pasa a Fase 2
   (implementar `caja_repo.py` / `cuenta_repo.py` / `motor_aplicacion.py`) o se
   sigue puliendo spec.
