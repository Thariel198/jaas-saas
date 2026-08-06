# RETOMAR — blancos_efectivo (5º precursor) + Q-6/Garcilazo cerrado · Sesión 2026-07-21

Continúa `docs/RETOMAR_backfill_diseno_flujo_2026-07-20.md` — esa sesión dejó **una
sola pregunta de diseño abierta para hoy**: el flujo de backfill para G-12 (S/45,
blanco de efectivo que el bug B6 descartó en silencio en junio). Hoy se cerró esa
pregunta, se descubrió que el patrón era más grande (5 blancos, no 1), se codificó
el 5º precursor, y aparte se resolvió una alerta de validación (Q-6/Garcilazo,
S/410) que resultó ser un caso de segregación pendiente de correr, no un bug.

---

## 1. G-12 (Opus, diseño) → `blancos_efectivo.xlsx` (Sonnet, código) — CERRADO

**Auditoría de junio (Q3 del RETOMAR anterior) reveló 5 blancos, no 1** — B6
(efectivo no ruteaba blancos) descartó S/282 en junio, todos en la caja pero sin
dueño:

```
mesa_1 f23  BLANCO/''      S/69   Wilder Trujillo  05/06  (sin comentario)
mesa_1 f41  BLANCO/''      S/37   Wilder Trujillo  06/06  (sin comentario)
mesa_2 f7   BLANCO/BLANCO  S/24   Yerald Romero    05/06  "Exoneracion"
mesa_6 f7   BLANCO/''      S/45   Wagner Trujillo  05/06  "Hernestina Valladares?" ← G-12
mesa_6 f62  BLANCO/BLANCO  S/107  Wagner Trujillo  06/06  (sin comentario)
```

**Decisión (Opus):** los 5 en un solo archivo (Regla del Tres cumplida 5 veces,
costo marginal ≈0, cierra el gap de caja de una). Nombre: `blancos_efectivo.xlsx`
(análogo por canal de `blancos_acumulados.xlsx`, que es de-facto el de yape).

**Implementado (Sonnet):**
- `shared/blancos_efectivo.xlsx` — 5 filas, MZ/LT lleno solo en G-12; los otros 4
  quedan con MZ/LT vacío (pendientes de identificar).
- `5_cobranza/main.py` — `BLANCOS_EFECTIVO_PATH` + `_cargar_blancos_efectivo()` +
  overlay en `_cargar_planilla()`, mismo patrón/cascada que `abonos_rezagados`.
- **Verificado matemáticamente:** G-12 (Ernestina Valladares) — base
  `ACUERDOS_ASAMBLEA=50` en shared → overlay resta 45 → `5` en `planilla_cobrado`,
  exacto. Los 4 sin MZ/LT: sin efecto (gate funciona), quedan visibles en el
  archivo esperando identificación.
- `backfill_ledger/docs/cuaderno_backfill.html` actualizado: Lámina 2 ahora lista
  **5 precursores** (no 4); Lámina 4 nueva ("el otro caso difícil") explica cómo
  se identifica un blanco (escribir MZ/LT en su fila) y qué pasa con los que aún
  no tienen dueño (quedan como movimiento de caja real, sin imputar, hasta que
  alguien los nombre — no es "historial muerto").

**Pendiente menor, no bloqueante:** O-6 (S/107) y R-7 (S/24) — verificar que no se
hayan aplicado ya por otra vía en julio antes de identificarlos (venían de D2 como
"reclamarán en julio"). Nadie los tocó hoy.

---

## 2. Bloque mecánico — CERRADO

| Ítem | Resultado |
|---|---|
| C1-9 → `ESTADO=RESUELTO` en `jass_system - junio/.../arrastre_devolucion_2026-06.xlsx` | Ya estaba hecho (sesión 07-20, 2ª parte — el RETOMAR de las 11:11 quedó desactualizado en ese punto) |
| V-6 → reclamo julio | Marcado `RESUELTO` en `4b_reclamos/outputs/reclamos_2026-07.xlsx` (backup tomado antes). Verificado con números: base `MULTA=30 + MES_ANTERIOR=14 = 44` = exacto el crédito de `reidentificacion.xlsx` → `planilla_cobrado` V-6 = SALDO 0, CANCELADO. El reclamo "Ya pagué multa 30" era cierto. |
| T-7 → reclamo julio | Sin cambio (correcto — sigue `EN_REVISION`, tema distinto de convenio). Verificado: `ACUERDOS_ASAMBLEA` 50→5 por el crédito de reidentificación, coincide exacto. |
| `6_corte/generar_lista.py` | Corrido — **49 elegibles** (bajó de 63 por los overlays acumulados + V-6 resuelto bloqueando por reclamo). Esta lista sigue siendo la vigente/publicada al cierre de la sesión. |

---

## 3. Q-6/Garcilazo (S/410) — diagnosticado, corregido, CERRADO

**Cómo empezó:** `5b_validacion` (Nivel 1a, TE PAGÓ) dio alerta `dif=-410.00`.

**Primer diagnóstico (parcial/equivocado):** se enmarcó como un gap de
reprocesamiento del motor entre corridas (la planilla cambió, un pago quedó
"sin deuda en planilla" y no se re-clasificó). El usuario corrigió el enfoque:
**Q-6 (Dedicación Garcilazo Romero) presta su Yape a otros vecinos** — el motor
lo auto-matchea a su propia cuenta (es un vecino real, matchea por maestro) en
vez de segregar por lote. Ese es el mecanismo `CONCEPTO=comunitario` /
`forzar_comunitario.xlsx` que ya describe el README de `motor_matching`.

**Lo que ya estaba hecho (el usuario, antes de esta sesión):**
- `correcciones/forzar_comunitario.xlsx` — 2 filas forzando "Dedicacion Gar*"
  (04/07 S/361 y 05/07 S/49) a `comunitario`.
- `correcciones/pendientes.xlsx` hoja `Segregacion` — desglose ya completo por
  lote y submonto, `OK=SI` en las 7 filas:

```
361 (04/07) → Q-3(tanque,200) + D-5(20) + F-14(68) + K-7(23) + V-1(33) + H-4(17) = 361 ✓
49  (05/07) → G-23(49)                                                          = 49  ✓
```

**Lo que faltaba (ejecutado hoy):**
1. `4_pagos/yape/motor_matching/main.py` — corrido 2 veces (1ª falló por
   `PermissionError`, `pendientes.xlsx` abierto en Excel; 2ª OK tras cerrarlo).
   Los 7 lotes quedaron identificados en `pagos_yape_tepago.xlsx`
   (`FUENTE=comunitario`), ciclo cerrado sin pendientes, banco/procesado
   archivados (`2026-07_banco.xlsx` / `2026-07_procesado.xlsx`).
2. **Chequeo previo (Regla de lista publicada):** ninguno de los 6 lotes de agua
   (G-23, D-5, F-14, K-7, V-1, H-4) estaba en `lista_corte.xlsx` (49 elegibles,
   publicada) — y como la corrida solo *reduce* deuda, no podía meter a nadie de
   nuevo en la lista. Sin riesgo, sin exoneración necesaria.
3. `5_cobranza --force` (ciclo 15) — aplicó los 6 créditos de agua:
   `G-23 PARCIAL(22) · D-5 PARCIAL(30) · F-14 EXCESO(-48) · K-7/V-1/H-4 CANCELADO`.
4. `5b_validacion` — bajó de dif=-410 a **dif=-200** (faltaba el tanque de Q-3).
5. `4_pagos/consolidar_tanque.py` — recogió el aporte de Q-3 (S/200,
   `aportes_tanque.xlsx` → 12 filas, S/1800).
6. `5b_validacion` (2ª corrida) — **VALIDACION OK**, todos los niveles cuadran.

**Cierre:** el "problema" nunca fue un gap de diseño — era trabajo de segregación
ya preparado por el usuario, esperando que el motor lo corriera. Nada de código
nuevo se tocó en `motor_matching`/`5_cobranza`/`consolidar_tanque` para esto (solo
se ejecutaron, ya estaban listos).

---

## 4. Aprendizaje de método (guardado en memoria)

Los README de módulos con sección "⚠ DISEÑO POST-LEDGER" (`5_cobranza`,
`5b_validacion`, `6_corte`) pueden describir intencionalmente el **modelo de
agosto** (post-ledger) en el resto del documento — no necesariamente el código
actual. Para entender comportamiento vigente, **el código manda, no el README**.
Esto salió a la luz cuando se marcó `6_corte/README.md` como "desactualizado"
(describía `aplicar_penalidad.py` escribiendo a `shared/planilla_mes`) y el
usuario aclaró que es documentación adelantada a propósito — el código real ya
tiene el fix Modelo A (verificado: solo escribe audit, no shared).
Memoria: `feedback_readme_describe_modelo_agosto_no_codigo_actual.md`.

---

## Pendientes para la próxima sesión

- **Regla 7** — agregar `backfill_ledger/` al pipeline de `README.md` raíz (se
  creó ayer, 07-20, nunca se sumó al README — ofrecido, no confirmado).
- **Regla 8** — verificar sincronía completa de `README.md` raíz contra la
  estructura real de módulos antes de la próxima sesión larga.
- **6_corte Día 0→2** — `aplicar_penalidad.py` todavía no corrió sobre la
  `lista_corte.xlsx` de 49 elegibles (7 con `EJECUTAR_CORTE=SI`). Sigue
  pendiente el ciclo normal del módulo, sin urgencia nueva de esta sesión.
- **O-6 / R-7** — verificar que no se hayan aplicado ya por otra vía en julio
  antes de identificarlos en `blancos_efectivo.xlsx`.
- Nada bloqueante. `5b_validacion` cerró en OK — el ciclo julio está sano al
  momento de escribir esto.

---

## Estado git — nada commiteado

Todo lo de esta sesión sigue sin commitear, sobre la base ya sin commitear de
sesiones anteriores (ver RETOMAR del 07-20). Nuevo/modificado hoy:

```
shared/blancos_efectivo.xlsx                    NUEVO
5_cobranza/main.py                              overlay blancos_efectivo
backfill_ledger/docs/cuaderno_backfill.html     5º precursor + lámina G-12
4b_reclamos/outputs/reclamos_2026-07.xlsx       V-6 → RESUELTO (+ backup en
                                                 4b_reclamos/backup/reclamos/)
4_pagos/yape/motor_matching/outputs/*           regenerados (ciclo cerrado)
4_pagos/yape/motor_matching/correcciones/
  pendientes.xlsx                               vacío de nuevo (backup tomado)
shared/reporte_acumulado_procesado/
  2026-07_banco.xlsx · 2026-07_procesado.xlsx   archivados por el motor
4_pagos/outputs/aportes_tanque.xlsx             regenerado (Q-3 incluido)
5_cobranza/outputs/*                            regenerados (ciclo 15)
5b_validacion/outputs/validacion_diferencias.xlsx  regenerado (OK)
6_corte/outputs/lista_corte.xlsx                49 elegibles (sin cambios post-Q6)
```
