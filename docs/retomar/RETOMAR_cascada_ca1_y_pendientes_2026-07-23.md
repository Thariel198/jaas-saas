# RETOMAR — Reorden de cascada (CA1), casos cerrados y limpieza pendiente · Sesión 2026-07-23

## ⚡ Punto de entrada

Esta sesión (23/7) fue casi toda sobre un solo hilo: el caso **C-1** destapó un bug real en
`5_cobranza/main.py::_descomponer_saldo` y llevó a reabrir la decisión de dominio **CA1** del
ledger. Eso se cerró bien. Lo que quedó abierto es una limpieza de documentación (sección 3) y
unos pendientes operativos sueltos (sección 4).

---

## 1. Cascada reordenada — CA1 (CERRADO, con nota)

```
ANTES (hasta 22/7)              AHORA (23/7, CA1 reabierta)
P3 MULTA · P4 ACUERDOS          P3 CONVENIO · P4 ACUERDOS
P5 CONVENIO                     P5 MULTA
```

**Por qué:** el dinero cubre primero lo que SOLO el dinero puede saldar. La multa es la deuda
más "no-monetaria" (faena se paga con trabajo, reunión se exonera por asistencia); convenio y
acuerdos solo se saldan con plata.

**Origen:** caso C-1 — el orden viejo marcaba una multa real (S/50, reunión+faena, sin ningún
pago) como "pagada" solo porque el waterfall llegaba a ella antes que a convenio/acuerdos, que
era lo que el vecino sí había pagado.

**Estado del código: NO se tocó, a propósito.** `_descomponer_saldo` sigue con el orden viejo.
Solo se corrigió el **dominio** (`libro_mayor/dominio/README.md` + 8 archivos más sincronizados)
para que el **backfill de agosto** reparta bien desde el día 1, sin parches.

**Caso C-1 puntual:** su multa real (reunión 20 + faena 30) se exoneró en
`shared/exoneraciones_multa.xlsx` — el re-etiquetado (multa↔convenio/acuerdos) NO necesitó
ningún precursor, se autocorrige solo con el backfill.

Memoria: `project_cascada_reorden_ca1.md`.

---

## 2. Otros casos cerrados hoy

| Caso | Qué se hizo | Dónde |
|---|---|---|
| R-7 / M-7 (blanco junio S/24) | Reidentificado a M-7 (monto exacto, M-7 sin ningún pago en junio). `blancos_efectivo.xlsx` marcado identificado con `MES_ANO_APLICA` vacío para no duplicar el crédito | `shared/reidentificacion.xlsx` + `shared/blancos_efectivo.xlsx` — memoria: `project_r7_m7_reidentificacion.md` |
| G-23, D1-3, N-6 | Verificados con evidencia real (G-23 exonerado con fecha en `registro_cortes.xlsx`; D1-3 y N-6 ya estaban bien) | `4b_reclamos/outputs/reclamos_2026-07.xlsx` |
| B1-12 "13 consumo, compromiso" | Verificado: el 13 es exacto el arrastre de agua de junio, sin error | — |
| pendientes_convenio_multas.xlsx | Extendido a 4 conceptos (CONVENIO/MULTA/ACUERDOS_ASAMBLEA/MES_ANTERIOR), 23 lotes cargados con "Verificando" | `3_boletas/inputs/` + gancho en `3_boletas/main.py` |
| Penalidad de corte | Aplicada a 7 usuarios (+S/20 c/u) vía `6_corte/aplicar_penalidad.py` — **falta re-correr `5_cobranza --force`** para que se refleje en `planilla_cobrado.xlsx` | `6_corte/outputs/audit_penalidad.xlsx` |
| G-9 (apellidos invertidos) | Cargado en `overrides_padron.xlsx`, dormido hasta agosto | `0_padron/overrides_padron.xlsx` |

---

## 3. PRIMER PASO al retomar — limpieza de READMEs sin resolver

**El usuario planteó (correcto):** los 4 READMEs que documentan CA1 (`dominio`, `caja`,
`estado_cuenta`, `libro_mayor` raíz) mezclan **diseño puro** con **notas de estado/fecha**
("Reabierta 2026-07-23", "código vivo sigue con el orden viejo", "caso C-1") — eso va a estorbar
cuando alguien implemente el backfill mañana, porque para entonces esas notas quedan obsoletas
o confusas.

**Se intentó extraer ese contenido a un doc aparte (candidato: `libro_mayor/TRANSICION.md`) y
se revirtió** — no porque la idea esté mal, sino por errores de alcance en cómo lo ejecuté (me
metí a tocar READMEs de módulos no relacionados sin preguntar primero, dos veces seguidas).

**Los 4 READMEs quedaron tal cual estaban** (con la mezcla diseño+estado). Si se retoma:
1. Separar en cada uno: qué es diseño permanente (el orden P1-P6, el principio del "dinero
   cubre lo no-monetario primero", la consecuencia sobre corte-por-multa) vs qué es narración
   transicional (fechas, "caso C-1", "código vivo sigue con...").
2. Lo permanente se queda en el README. Lo transicional se mueve a un doc nuevo.
3. **Tocar SOLO esos 4 archivos** — no abrir la puerta a "sincronizar" READMEs de otros módulos
   de sesiones anteriores sin preguntar (ver lección en `feedback_cierre_readme_scope.md`).

Memoria: `feedback_cierre_readme_scope.md`.

---

## 4. Pendientes operativos sueltos

- **M-19** — reclamo "Cambio nombre" marcado `RECHAZADO`, pero el guardado quedó **sin confirmar**
  (archivo bloqueado varias veces). Revisar si se guardó.
- **B-19, F1-4** — multa/monto sin ningún respaldo en ningún archivo/mes. Preguntar directo a
  Yanet/secretaria. Ver `project_b19_f1-4_multas_sin_respaldo.md`.
- **Z-14** — está en riesgo **activo** de corte (`EJECUTAR_CORTE=SI` en la lista vigente), sin
  ninguna acción tomada todavía, a diferencia de G-23.
- **`5_cobranza --force`** — falta re-correr para que la penalidad aplicada hoy (7 usuarios)
  se refleje en `planilla_cobrado.xlsx`.
- **`resolucion_reclamos_2026-07.xlsx`** — sigue sin existir. 29 de 76 reclamos ya tienen
  `TIPO_RECLAMO` clasificado, falta correr `4b_reclamos/resolucion.py`.
- **O-6 (S/107)** — blanco de efectivo sin identificar (a diferencia de R-7, que sí se cerró hoy).
- **`0_padron/override.xlsx`** (el viejo, distinto de `overrides_padron.xlsx`) sigue roto —
  apunta a `03_llenado/`, que no existe. No bloqueante, deuda anotada desde el 22/7.

---

## 5. Dónde está todo

- Diario completo del día (con diagramas): `docs/diario/2026-07.html`, día 23/7.
- Decisión de dominio: `libro_mayor/dominio/README.md` (buscar "CA1").
- `0_padron/`, `4b_reclamos/outputs/`, `6_corte/outputs/` están en `.gitignore` — sus cambios
  de hoy son solo locales, nunca se van a commitear con el resto.
- `libro_mayor/` y `docs/cuaderno/` **nunca tuvieron un commit** — cualquier commit ahí mezcla
  varias sesiones, no solo hoy. Decidir aparte cuándo cerrarlo.
