# RETOMAR — `obligaciones/` CERRADO · Fase 1 de diseño COMPLETA (salvo 6b) · 2026-07-14→15, revisado 07-16 (Opus)

Continúa `docs/retomar/RETOMAR_dominio_saldo_unico_2026-07-13.md` (§10 = spec de `dominio/`,
§11 = cierre de `5_cobranza`/`5b`/`6_corte`). Este doc es el punto de entrada de la
próxima sesión.

> **Nota de historial:** este doc se abrió el 07-14 (obligaciones a mitad, frenado en la
> "pregunta i"). La sesión del **07-15** cerró el diseño detallado de `obligaciones/` pero
> **no se documentó** (se olvidó actualizar el RETOMAR). Esta versión (07-16) absorbe el
> trabajo del 07-15 y refleja el estado real.

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **La Fase 1 de diseño está COMPLETA.** `obligaciones/` cerrado (07-15) y `6b_corte_multas`
   cerrado (07-16, ver abajo). Ya no queda ningún hueco de diseño.
2. **`6b_corte_multas` SE DISUELVE en `6_corte`** (decisión 07-16, bajo el lente de escala).
   No es un módulo espejo: "corte por multa" es el concepto `multa` marcado como *cut-trigger*
   en el manifiesto del tenant, sobre un **motor de corte único** que emite `corte_reconexion`
   (mismo concepto que el corte de agua — reconectar cuesta igual; el "por qué" vive en
   `registro_cortes.MOTIVO`). Detalle en la sección "Decisión 6b" abajo.
3. **El spec destino de 6b YA ESTÁ ESCRITO** (07-16) y byte-compatible en todos los README
   afectados — NO era "cierre mecánico", era diseño (ver [[feedback_no_cierre_mecanico]]):
   - `6_corte/README.md` = **autoridad**: sección POST-LEDGER multi-trigger con la tabla
     agua|multa (CONCEPTOS_SALDO · UMBRAL · PROTEGE_PAGO_PARCIAL · PENALIDAD 20→40 ·
     SALVADO_CUANDO), columna `MOTIVO` en `registro_cortes`, CARGO `corte_reconexion` único.
   - `6b_corte_multas/README.md` = banner "SE DISUELVE en 6_corte" + su aporte como trigger
     `multa`; el código pre-ledger queda descrito abajo (sigue corriendo).
   - `README.md` root + `docs/arquitectura_pipeline_futuro.html` + `docs/diagrama_flujo_sistema.html`
     = reconciliados (6b se disuelve).
   - **`dominio/politica_corte.evaluar()` NO cambia:** multa = `evaluar(saldo, 0, 0, ya_cortado,
     en_revision, False)`, caso degenerado. Los params por trigger viven en el manifiesto del tenant.
   - PENDIENTE (Fase 2, código, NO ahora): retirar el módulo `6b_corte_multas/` físico — hoy su
     código pre-ledger sigue corriendo; se elimina cuando 6_corte implemente el multi-trigger.
4. **SIGUIENTE = Fase 2 (código), roadmap B** empezando por `dominio/` (Opus para el arranque de
   la lógica pura; Sonnet para lo mecánico). Ver [[feedback_no_codificar_diseno_no_cerrado]] y
   [[feedback_readmes_100_compatibles_sin_punteros]]: el spec de CADA pieza firme y compatible
   antes de codificar (Modo B, sin deuda de integración).
5. **Lente de escala ampliado (07-16):** `docs/lente_escala.md` es el doc standalone (antes
   disperso). Se agregó el principio **general→específico por config, no por módulo**:
   onboarding de JASS = agregar/quitar CONCEPTOS en un manifiesto de tenant, no enchufar
   módulos. Un motor idéntico en 25k JASS; la variación de capacidad es 100% data. Es el
   fundamento de por qué 6b se disuelve. Ver [[reference_lente_escala]].

---

## Qué cerró la sesión 07-15 (lo que faltaba documentar)

### `obligaciones/` — diseño detallado CERRADO (las 6 preguntas del 07-14, resueltas)

Detalle completo en `obligaciones/README.md` (sección "Diseño detallado — CERRADO 2026-07-15").
Resumen:

```
(i)   split SUB_CONCEPTO      RESUELTO: sale de las FUENTES CRUDAS de la secretaria, no de
        MULTA/ACUERDOS               DATA_boletas (derivado). Cada concepto viene ya partido:
                                     · MULTA    → 2 hojas: reunión (tarifa 20) · faena (tarifa 30)
                                     · ACUERDOS → hoja "Corregido": cols TECHADO · CAMPO
                                     · CONVENIO → 2 archivos: medidor · instalación
                                     Ninguno queda con sub genérico. Ver mapeo de archivos↓.
(ii)  backfill MES_CARGO      RESUELTO: MES_CARGO va en la llave del CARGO_ID (mes real del
                                     evento). Carry (histórico) y nuevas (evento futuro) no se
                                     solapan → anti-doble-conteo automático por la llave.
(iii) anti-doble-conteo       RESUELTO: idem — la unicidad del CARGO_ID lo garantiza.
(iv)  trigger real            RESUELTO: la fuente es el REGISTRO DEL EVENTO de la secretaria
                                     (asistencia asamblea/faena, convenio/instalación firmados),
                                     NO el artefacto derivado DATA_boletas.
(v)   inputs + qué si falta   RESUELTO: validación al inicio → FileNotFoundError/ValueError
                                     descriptivo, nunca siembra parcial. audit_ref = fila cruda.
(vi)  idempotencia            RESUELTO: CARGO_ID = sha256[:8](JASS_ID, MZ, LT, CONCEPTO,
                                     SUB_CONCEPTO, MES_CARGO). Re-correr backfill no duplica.
```

**Mapeo de fuentes crudas → CARGO** (en `obligaciones/README.md`, tabla completa con columnas):

```
CONCEPTO  SUB          ARCHICO · HOJA                                        MONTO
MULTA     REUNION      FAENAS REUNIONES JASS PUEBLO.xlsx · Hoja1            tarifa fija 20 (vacía=faltó=1 cargo)
MULTA     FAENA        FAENAS REUNIONES JAS.xlsx · Hoja1 (col6=typo, ignorar) tarifa fija 30 (2da faena 31/5 sin cargar)
ACUERDOS  TECHADO      DEUDORES...TECHADO Y CAMPO.xlsx · hoja "Corregido"   monto directo celda
ACUERDOS  CAMPO        idem · col CAMPO                                      monto directo celda
CONVENIO  MEDIDOR      mayo-planilla...xlsx · "Cobro medidores" · Deuda     50 ó 100 (Deuda−ΣPago=Saldo)
CONVENIO  INSTALACION  SEGUMIENTO INSTALACIONES...xlsx · NUEVAS INSTALAC.   TOTAL directo
```

**Decisiones de dominio nuevas del 07-15:**
- **Reasignación COFOPRI obligatoria al leer:** las fuentes crudas usan predios viejos
  (pre-COFOPRI). Cada `(MZ,LT)` pasa por la tabla canónica `0_padron/reasignaciones_candidata.xlsx`
  (32 filas). Aplicar como **lookup RAW→SYS de golpe, nunca secuencial** (hay swaps/cadenas →
  double-shift). El remap SOLO aplica al sembrar; NO reescribe `padron_reconciliado` ni planilla.
- **MULTA se siembra BRUTA, el motor reconcilia (no manual):** los archivos de asistencia dan
  deuda bruta (≈S/12,540 = 234 ausentes reunión×20 + 262 faena×30) pero el residual real es
  ínfimo. **No se reconcilia a mano.** Se siembran todos los cargos de ausencia + todos los
  abonos a `caja`; el **motor** aplica la cascada y el residual cae solo → se valida contra
  `Deuda faena` (9 filas). No se puede parquear MULTA: la cascada acopla los conceptos.
- **`obligaciones` emite TODO el convenio, incl. instalación** → el hack
  `PREDIOS_INSTALACION_EXCLUIDOS` de la siembra vieja DESAPARECE.

Memoria ya guardada: [[project_obligaciones_fuentes_crudas]].

### Cascada cambió: P1-P5 → **P1-P6** (corte_reconexion insertado en P2)

El contrato de `estado_cuenta/README.md` (sección "Taxonomía ⑪") y el HTML nuevo coinciden:

```
P1  AGUA · MANTENIMIENTO · arrastre    FIFO por mes, sin sub
P2  CORTE_RECONEXION                    sin sub          ← NUEVO slot (antes multa era P2)
P3  MULTA          reunión → faena
P4  ACUERDOS       techado → campo
P5  CONVENIO       medidor → instalación
P6  OTROS          sin sub · slot residual, sin emisor hoy (especulativo)
```

> ⚠️ Deuda menor: quedan 2 rótulos stale "P1-P5" en `estado_cuenta/README.md` (línea ~220 y
> el resumen ⑪ ~línea 480) que no se actualizaron al insertar P2. El bloque canónico de
> taxonomía (líneas 360-367) sí dice P1-P6. Corregir esos 2 rótulos en la próxima pasada al README.

### Artefacto nuevo

```
docs/arquitectura_pipeline_futuro.html   visual completo del estado destino post-ledger:
                                         2 capas · pipeline 0-7 con emisor/abono/disolución ·
                                         obligaciones · substrato libro_mayor · principio
                                         HECHO/INTERPRETACIÓN · cascada P1-P6 · extracto · escala
```

---

## Decisión 6b (07-16) — SE DISUELVE en 6_corte

Bajo el ledger, `6_corte` y `6b` hacen las **mismas 4 operaciones** (riesgo=query · política
decide · emite CARGO `corte_reconexion` · `registro_cortes`+seguimiento@T2). Las diferencias
son 3 **parámetros** (elegibilidad, monto, condición de salvado), no otra máquina → por el
lente de escala (política ≠ arquitectura) es **un motor de corte único, config-driven**, no
dos módulos. `registro_cortes` ya es shared = un solo corte físico = un solo agregado.

```
DESTINO: corte = 1 capability policy-driven (no un módulo por tipo de deuda)

  dominio/politica_corte(jass, saldos_por_concepto) → [ (motivo, penalidad, salvado), ... ]
     triggers enumerados en el manifiesto del tenant:
        agua  : saldo_agua ≥ umbral_min(jass)     → corte_reconexion · monto_agua  · salvado: pago≥pen
        multa : saldo_multa+acuerdos > 0          → corte_reconexion · monto_multa · salvado: saldo==0
  6_corte:  riesgo_corte → politica_corte → registrar_cargo(corte_reconexion)
            → registro_cortes[MOTIVO=agua|multa] → seguimiento@T2 (query aplicaciones)
```

**Qué se resuelve solo bajo el ledger (no migra código de 6b):**
- El `DEUDA_MULTA = (MULTA+ACU) − max(0, pagado−agua)` de 6b **desaparece** — ese "excedente
  de agua se abona a multa" ES la cascada P1→P3→P4; el motor lo deriva. Query directa del saldo.
- CONVENIO excluido = el trigger multa mira solo conceptos MULTA+ACUERDOS (config), no CONVENIO.

**Sub-decisión (Q2, resuelta):** la penalidad del corte-por-multa **reutiliza el concepto
`corte_reconexion`** (un corte físico = una tarifa de reconexión, gatíllela agua o multa). NO
se agrega concepto/sub nuevo → la cascada P1-P6 no crece; `corte_reconexion` sigue sin
`SUB_CONCEPTO`. El `MOTIVO` (agua|multa) es metadato operativo en `registro_cortes`, no en el
cargo. Contrato ⑧ intacto: `corte_reconexion` lo sigue emitiendo `6_corte`.

**Costo aceptado:** se pierde la lista física separada agua/multa → se resuelve con un filtro
`MOTIVO` sobre la lista única (no justifica dos módulos).

---

## Estado del roadmap de diseño → implementación

```
✓ dominio/       spec cerrado (6 firmas) · céntimos int · sin TOL · tenant-agnóstico · cero I/O
✓ 5_cobranza     diseño cerrado (se disuelve → motor + queries + obligaciones)
✓ 5b_validacion  diseño cerrado (se disuelve → arqueo_caja + conciliar_caja)
✓ 6_corte        diseño cerrado (sobrevive reshaped; regla amount-based, umbral=deuda mínima)
✓ contrato ⑧     corregido y byte-idéntico (verificado 07-16: caja ↔ estado_cuenta OK)
✓ obligaciones/  diseño DETALLADO cerrado (07-15) — las 6 preguntas resueltas
✓ extracto ⑫     cerrado (5 decisiones + render stateless)
✓ 6b_corte_multas  SE DISUELVE en 6_corte (07-16) — spec destino ESCRITO y compatible en:
                     6_corte/README (autoridad: tabla de triggers agua|multa + MOTIVO + salvado)
                     · 6b/README (banner disuelve) · root README · los 2 HTML de arquitectura.
                     politica_corte.evaluar NO cambia (multa = caso degenerado umbral=0, pago_parcial=False).
✗ CÓDIGO         nada codificado. Orden (roadmap B) cuando cada pieza cierre:
                   1. dominio/ (lógica pura, bajo riesgo, testeable con dicts)
                      taxonomia → entidades → cascada → politica_corte → saldo → identidad
                   2. repos: caja_repo.py (writer único abonos) · cuenta_repo.py (writer cargos)
                   3. motor_aplicacion.py (aplicar(cargos, abonos) → aplicaciones)
                   4. importers de caja (efectivo/yape/egresos) + obligaciones/ + backfill (B4)
```

---

## Recordatorios operativos

- **6b NO bloquea el build de dominio/repos/motor.** Es un módulo del pipeline (corte de
  multas), no toca el substrato. Se puede cerrar en paralelo o arrancar `dominio/` ya.
- **Modelo Opus** para 6b (diseño) y el arranque de `dominio/` (lógica de negocio pura).
  Sonnet para código mecánico una vez el spec de cada pieza esté firme.
- **`docs/pendientes_plan.md` está descartado** por el usuario — NO usarlo. Este doc +
  RETOMAR_dominio §10/§11 + los README de cada pieza son la verdad.
- **Migración de `seguimiento_pueblo`:** NO se migra el histórico, se RE-DERIVA (Hueco 3).
  `obligaciones` es el emisor PERMANENTE que puebla `estado_cuenta`; el backfill es una de sus
  corridas. `seguimiento_pueblo` es el store viejo que DESAPARECE.
- **Al retomar, verificar** que la sección "## CONTRATO DE INTERFAZ" siga byte-idéntica entre
  `libro_mayor/caja/README.md` y `estado_cuenta/README.md`
  (`awk '/## CONTRATO DE INTERFAZ/,0'` + `diff`). Verificado OK el 07-16.
- **Todo `obligaciones/` + `libro_mayor/` son archivos UNTRACKED en git** (nunca commiteados).
  El último commit real es `51d8d26` (07-10). Todo el diseño del ledger vive sin commit —
  considerar un commit de diseño al cerrar 6b, antes de codificar.
