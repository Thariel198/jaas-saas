# RETOMAR — 7b_historial_pagos: diseño cerrado, código parcial · Handoff sesión 2026-07-10

Dónde nos quedamos y qué sigue. Leer de arriba a abajo antes de tocar nada.
Esta sesión fue larga y mezcló 3 bloques distintos — ver "Bloque 1/2/3" abajo.
**Próxima sesión: Opus** para retomar el diseño de staging (es arquitectura, no
código mecánico). Cambiar a Sonnet recién para implementar lo ya decidido.

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **Bloque 1 (dedup reclamos + rechazos convenio) → ✅ COMMITEADO.** No re-hacer.
2. **Bloque 2 (módulo 7b_historial_pagos) → ⚠ DISEÑO CERRADO, CÓDIGO A MEDIO CAMINO
   Y PARCIALMENTE OBSOLETO.** La arquitectura cambió a mitad de sesión (de "parsear
   el libro directo" a "staging validado por el motor real") — el código ya escrito
   (`importar_libros.py`) queda **deprecado** por esa decisión, no se debe seguir
   usando tal cual. Ver sección "Qué cambia con la decisión de staging" abajo.
3. **Dato de prueba sucio en el store real:** `shared/reporte_acumulado_procesado/
   2026-05_historial.xlsx` tiene datos de una corrida de prueba con un bug conocido
   (127 filas efectivo en vez de 117 — investigar o descartar, ver abajo). **Borrar
   este archivo antes de sembrar mayo de verdad** con el flujo de staging.
4. Nada del Bloque 2 está commiteado — ver "Estado de archivos" al final.

---

## Bloque 1 — 4b_reclamos: dedup + rechazos convenio (✅ COMMITEADO, no tocar)

Continuación de la sesión anterior (ver `docs/retomar/RETOMAR_reclamos_duplicados_2026-07-10.md`,
ya puede archivarse/borrarse, este documento la reemplaza como handoff activo).

- **Commit `59810e7`**: `_reconciliar_duplicados` en `4b_reclamos/main.py` — dedup
  automático de filas con `FECHA_COBRO` corregida a mano en `mesa_N`, log de
  diagnóstico en `logs/duplicados.log` (no Excel de auditoría — ver doc de
  aprendizaje `auditoria_negocio_vs_log_diagnostico_20260710.html`, la distinción
  quedó documentada: auditoría de negocio = Excel permanente, diagnóstico técnico
  = log). Validado: 22 duplicados depurados en julio, idempotente, regresión
  6_corte OK.
- Además, sin commit propio (trabajo manual sobre el output, no código): se
  clasificaron 32 reclamos como `TIPO_RECLAMO=convenio` (keyword convenio/medidor),
  se cruzaron contra `CONVENIO_HISTORIAL` de `seguimiento_pueblo` y se marcaron
  `RECHAZADO` 13 con `SALDO ACTUAL=0` + 11 reclamos vagos ("Modificar"/"Coreccion"
  sin detalle) que también dieron `SALDO=0` — total 24 `RECHAZADO` en
  `reclamos_2026-07.xlsx` (gitignored, no aparece en git status).
- **Caso G-1/M-15 (pareja, multa)**: diagnosticado con `planilla_cobrado.xlsx` —
  G-1 debía 20 de multa, M-15 tenía 20 de exceso, mismo bolsillo (pareja). Se
  diseñó (no codeó) `correcciones_pago` en `5_cobranza` para reatribuir el pago —
  **este diseño quedó absorbido por 7b_historial_pagos** (`registrar_correccion`
  hace exactamente esto, de forma más general). No hace falta un módulo aparte en
  `5_cobranza` — pendiente de decidir si `5_cobranza` debe consultar
  `7b_historial_pagos` para esto o si se resuelve directo en el ledger.

**Pendiente suelto sin bloquear nada:** mesa_4 A-1/D-6/S-8 (reclasificados a
compromiso/exoneración en el input pero siguen como reclamo activo) — heredado de
la sesión anterior, no se tocó esta sesión tampoco.

---

## Bloque 2 — 7b_historial_pagos: diseño Fase 1 cerrado

### Por qué existe (resumen — detalle completo en `7b_historial_pagos/README.md`)

Los libros de Excel viejos tenían una hoja de pagos completa por mes — para
efectivo, esa capacidad se perdió al modularizar el sistema (yape sí la tiene en
`shared/reporte_acumulado_procesado/YYYY-MM_procesado.xlsx`). Se necesitó a mano
4 veces en esta sesión (convenio SALDO=0, reclamos vagos, pareja G-1/M-15,
"¿te pagué efectivo en mayo?") — motivo real, no "muchos módulos lo piden".

### Nombre y posición en el pipeline

Se llamó primero `8_historial_pagos`, se renombró a **`7b_historial_pagos`** — corre
**después de `7_cierre`**, nunca en caliente, porque durante el mes los reclamos
generan reatribuciones/rechazos que todavía están asentando el pago "verdadero".
Cargar antes fotografiaría un mes a medio corregir. `7_cierre` todavía no está
implementado (**"Diseñado — pendiente de implementación"** en el README raíz) —
por eso el import de meses HISTÓRICOS (Oct-2025 → hoy) no depende de que exista:
son meses ya cerrados de hecho, aunque el módulo `7_cierre` como código no exista
todavía.

README raíz (`README.md`) actualizado: pipeline muestra `[7_cierre + 7b_historial_pagos]`,
tabla con fila nueva, convención `b` documentada.

### Arquitectura (sin cambios desde que se cerró — ver README del módulo)

```
FUENTES → importar_*.py (productores thin) → historial_repo.py (WRITER ÚNICO)
   → STORE shared/reporte_acumulado_procesado/YYYY-MM_historial.xlsx (append-only)
   → consultar.py (LECTORES read-only) → analizar_reclamo(mz,lt) / reporte_pagos(mz,lt)
```

Invariantes: writer único, append-only, sin concepto (el reparto por concepto es
de `seguimiento_pueblo`), correcciones = eventos de primera clase (`FUENTE=correccion`),
conservación (`registrar_correccion` escribe 2 filas: −monto origen / +monto destino).

### El evento — columnas finales (ya corregidas, contrato cerrado)

```
¿Quién es?     MZ · LT · NOMBRE                    ← corregido: era LOTE, ahora LT
                                                       (consistencia con el resto del sistema)
¿Qué pagó?     CANAL · MONTO · FECHA · MES_CICLO
Referencia     REFERENCIA (1 sola columna, nunca vacía)  ← corregido: eran 5 columnas
                 efectivo: "MESA-COBRADOR"                  dispersas (MESA·COBRADOR·
                 yape:     ORIGEN                            ORIGEN_YAPE·DESTINO_YAPE·MENSAJE)
Auditoría      ESTADO · FUENTE · MOTIVO · AUDIT_REF · CICLO_CORRECCION ·
               ORIGEN_ARCHIVO (renombrado, antes "ORIGEN" — chocaba con el ORIGEN
               del yape que ahora vive dentro de REFERENCIA) · TIMESTAMP
```

`FECHA`: día solo para efectivo (`DD/MM/YYYY`), **día+hora:min:seg para yape**
(`DD/MM/YYYY HH:MM:SS`) — la precisión de segundos, junto con `REFERENCIA` (el
`ORIGEN`), es lo que identifica el pago yape de forma única (mismo criterio que
ya usa `motor_matching`).

Docs actualizados y consistentes entre sí:
- `7b_historial_pagos/README.md`
- `7b_historial_pagos/docs/arquitectura_7b_historial_pagos.html`
- `7b_historial_pagos/docs/diagrama_flujo_7b_historial_pagos.html`
- `7b_historial_pagos/docs/formato_evento_pago.html`

---

## Qué cambia con la decisión de staging (LO ÚLTIMO QUE SE DECIDIÓ — leer con cuidado)

### El problema que la motivó

Se codeó `importar_libros.py` para parsear el libro Excel legacy directo (hojas
"Efectivo" y "Reporte"). Al correrlo contra `libro_2026-05.xlsx`:

- **Bug real encontrado**: 117 filas de efectivo esperadas (log del primer run),
  pero al leer el store después aparecían **127** — no se investigó la causa raíz
  antes de que la conversación pivotara a la decisión de staging. **Queda sin
  resolver.**
- Un crash por un dato sucio real: una fila de "Reporte" con `mz="Rosa coronado"`
  (un nombre, no una manzana — una devolución mal codificada en el libro viejo).
  Se corrigió tratándolo como `blanco` en vez de crashear (fix ya aplicado en
  `importar_libros.py`, pero el archivo completo va a deprecarse igual).
- Suma efectivo importada (3429, con el bug) vs. ancla conocida del propio libro
  (`SALDO`=3188 en la hoja "INGRESOS Y EGRESOS") — no cuadraba, gap grande.

### La decisión: no parsear el libro crudo directo — validar con el motor real

Insight del usuario: el libro tiene **dos bloques espejo** en la hoja "Efectivo"
(cobrados por fórmula de Excel) que ya están reconciliados a mano en esa época.
En vez de reinventar reglas de validación ad-hoc, mejor **reusar el motor de
discrepancias que ya existe** (`4_pagos/efectivo/main.py`) tratando los dos
bloques como si fueran 2 cobradores de una mesa sintética — si el libro está bien,
el motor real debería dar 0 discrepancias.

Para yape: si hay yape crudo real de ese mes, correrlo por `motor_matching`
aislado (auto-match), y los `pendientes` que salgan se resuelven RÁPIDO copiando
el `mz/lt` que ya está resuelto a mano en la hoja "Reporte" del libro (matching de
esa época, reusado — no se rehace desde cero).

**Por qué NO se puede correr `4_pagos/efectivo/main.py` ni `motor_matching`
directo contra datos históricos:** ambos escriben en sus paths de producción
reales (`4_pagos/efectivo/outputs/pagos_efectivo.xlsx`, etc.) — el mismo archivo
que usa el ciclo EN CURSO. Hay que correrlos **aislados** (paths de config
monkey-parcheados a una carpeta temporal — mismo patrón que ya usa el proyecto
para tests de integración, ver `docs/metodologia_desarrollo.md` sección "Patrón:
test de integración con datos sintéticos").

### Arquitectura final acordada

```
PROCESO PUNTUAL POR MES (fuera de 7b, aislado, NO toca producción — trabajo manual
del usuario, no vive como código permanente en el módulo)

  EFECTIVO: libro → mesa_1.xlsx sintética (bloque1=Wilder, bloque2=Janet como
            2 "cobradores") → 4_pagos/efectivo/main.py corrido AISLADO (paths
            temporales) → 0 discrepancias esperadas → resultado con la MISMA
            forma que pagos_efectivo.xlsx real, para ese mes

  YAPE:     yape crudo del mes → motor_matching AISLADO → auto-match + pendientes
            → pendientes se llenan con el mz/lt ya resuelto en la hoja "Reporte"
            del libro (no se rematchea de cero) → resultado con la MISMA forma
            que YYYY-MM_procesado.xlsx real

  Ambos resultados se guardan en STAGING:
    7b_historial_pagos/inputs/validado/pagos_efectivo_YYYY-MM.xlsx
    7b_historial_pagos/inputs/validado/YYYY-MM_procesado.xlsx  (o el nombre que
        ya usa motor_matching para su output — a definir/confirmar)

PERMANENTE — 7b_historial_pagos (independiente, no cambia según cómo se validó
cada mes)

  importar_efectivo.py  ← lee pagos_efectivo_YYYY-MM.xlsx de STAGING (histórico)
                           o del path real (mes en curso, post 7_cierre) — MISMO
                           importador para los dos casos, solo cambia la fuente
  importar_yape.py      ← simétrico, lee de staging o del store real de yape
  historial_repo.py     ← sin cambios, writer único (YA CODEADO, ver abajo)
  consultar.py           ← sin cambios, lectores read-only (YA CODEADO, ver abajo)

  importar_libros.py    ← DEPRECADO por esta decisión. No se debe usar más tal
                           cual. Se reemplaza por importar_efectivo.py +
                           importar_yape.py leyendo de staging.
```

**Por qué separarlo así (y no que 7b invoque `4_pagos`/`motor_matching`
directamente):** si 7b importara y ejecutara esos módulos internamente, quedaría
acoplado a su versión actual (un cambio ahí podría romper la carga histórica sin
que nadie lo note) y con riesgo de pisar outputs reales si algún path no se aisló
bien. Separado, 7b no sabe ni le importa CÓMO se validó cada mes — solo lee un
archivo con forma conocida.

### Decisión de secuencia: sembrar mes a mes, NO esperar los 9 meses

Confirmado con el usuario: sembrar cada mes apenas su staging está listo, no
acumular los 9 meses de Drive antes de sembrar. Razón: `historial_repo` es
append-only e idempotente por mes — no hay costo técnico en sembrar incremental,
y sí hay costo en esperar (el bug de 127 vs 117 se encontró en el mes 1; si se
hubiera esperado a los 9, se habría propagado a los 9 antes de notarlo).

---

## Código ya escrito — estado exacto

Todo en `7b_historial_pagos/`, **nada commiteado**:

| Archivo | Estado | Qué hacer la próxima sesión |
|---|---|---|
| `historial_repo.py` | ✅ Completo, funcional, probado (writer único, `registrar_pago`/`registrar_correccion`/`identificar_blanco`, idempotencia por `ORIGEN_ARCHIVO` verificada con 2 corridas) | Revisar si el schema sigue sirviendo tal cual con el flujo de staging — probablemente sí, no cambia |
| `consultar.py` | ✅ Completo, `analizar_reclamo`/`reporte_pagos`, expone `historial_repo.COLUMNAS` (público, no `_COLS`) | Sin cambios necesarios |
| `importar_libros.py` | ⚠ Funcional pero **DEPRECADO** por la decisión de staging — parsea el libro crudo directo, tiene el bug 127 vs 117 sin resolver | Reemplazar por `importar_efectivo.py` + `importar_yape.py` (leen de staging). Decidir si se borra o se deja como referencia histórica del enfoque descartado |
| `README.md` + 3 HTMLs en `docs/` | ✅ Diseño Fase 1 cerrado y consistente entre sí | Agregar la sección de staging (este documento la tiene en prosa, falta pasarla al README + quizás un diagrama nuevo) |

**Bug sin resolver:** 127 filas de efectivo en el store vs 117 esperadas según el
log del primer run. No se investigó — quedó pendiente cuando la conversación
pivotó a la decisión de staging. Con el nuevo enfoque (staging validado por el
motor real) este bug específico de `importar_libros.py` deja de importar, pero
**si algo similar aparece en `importar_efectivo.py`, investigar la causa raíz
antes de confiar en cualquier número.**

---

## Dato sucio a limpiar antes de la próxima siembra real

`shared/reporte_acumulado_procesado/2026-05_historial.xlsx` — creado por las
corridas de prueba de `importar_libros.py` esta sesión. Tiene el bug de conteo
sin resolver. **Borrar este archivo antes de sembrar mayo de verdad** con el
flujo de staging (para no arrastrar el bug ni confundir "ya sembrado" con "sembrado
mal").

---

## SIGUIENTE_ACCIÓN — orden sugerido

1. **[Opus]** Retomar el diseño de staging: confirmar naming exacto de
   `inputs/validado/`, confirmar cómo se aísla `4_pagos/efectivo/main.py` y
   `motor_matching/main.py` (monkey-patch de paths — decidir si es un script
   reutilizable `_correr_aislado.py` o proceso manual cada vez).
2. **[Opus/Sonnet]** Codear `importar_efectivo.py` + `importar_yape.py` (leen de
   staging o de la fuente real, mismo importador para ambos casos).
3. **[Sonnet]** Borrar `2026-05_historial.xlsx` (dato de prueba sucio).
4. **[usuario]** Armar el staging de mayo (mesa sintética Wilder/Janet + corrida
   aislada de `4_pagos/efectivo`; yape si hay crudo disponible).
5. **[Sonnet]** Sembrar mayo con `importar_efectivo.py`/`importar_yape.py` contra
   el staging real, validar con `consultar.reporte_pagos` contra algún predio
   conocido.
6. Repetir 4-5 mes a mes hasta cubrir Oct-2025 → hoy. Un commit natural por mes.
7. Decidir destino de `importar_libros.py` (borrar vs. dejar documentado como
   enfoque descartado — ver `docs/decisiones/` si existe convención para esto).
8. Pendiente suelto de Bloque 1: mesa_4 A-1/D-6/S-8.

---

## Estado de archivos al cerrar

```
Bloque 1 — COMMITEADO (59810e7, 2904d64), nada pendiente de commit.

Bloque 2 — 7b_historial_pagos/           TODO sin commitear:
  README.md                               M (diseño Fase 1 + sección "cuándo corre")
  docs/arquitectura_7b_historial_pagos.html   nuevo
  docs/diagrama_flujo_7b_historial_pagos.html nuevo
  docs/formato_evento_pago.html               nuevo
  historial_repo.py                           nuevo, funcional
  importar_libros.py                          nuevo, DEPRECADO (ver arriba)
  consultar.py                                nuevo, funcional
  inputs/libros/libro_2026-05.xlsx            movido desde docs/exploracion/ (antes
                                                sin trackear, ahora vive acá)

README.md (raíz)                          M (pipeline + tabla con 7b_historial_pagos)
8_historial_pagos/README.md               D (carpeta vieja, ya no existe — el
                                             contenido se movió a 7b_historial_pagos/)

shared/reporte_acumulado_procesado/
  2026-05_historial.xlsx                  nuevo, DATO DE PRUEBA SUCIO — borrar
  estado_ciclo.json                       M — preexistente de sesión anterior (09/07),
                                             NO tocado esta sesión, no relacionado
  2026-07_banco.xlsx, 2026-07_procesado.xlsx  sin trackear, preexistentes (yape motor
                                                 en vivo), no relacionados a esta sesión

docs/retomar/RETOMAR_reclamos_duplicados_2026-07-10.md   handoff de la sesión anterior,
                                                    ya cubierto por Bloque 1 de este
                                                    documento — puede archivarse
docs/retomar/RETOMAR_historial_pagos_2026-07-10.md       este archivo
```
