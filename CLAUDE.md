# jass_system — Referencia de sesión

Reglas operativas que aplican en toda sesión. Detalle y contexto en `docs/metodologia_desarrollo.md`.

---

## Pipeline del proyecto

El orden y estado de los módulos está documentado en `README.md` —
consultarlo ahí, no asumir un orden fijo de memoria.

---

## Las 4 fases

| Fase | Qué se hace | Qué NO se hace | Modelo dominante |
|---|---|---|---|
| **0 · Investigación** | Entender el negocio JASS + explorar datos reales con scripts desechables | Escribir código de producción | Opus (negocio) · Haiku (scripts de exploración) |
| **1 · Diseño / Spec** | Iterar el spec hasta que todos los módulos, archivos y columnas sean compatibles entre sí | Escribir código | Opus |
| **2 · Implementación** | Traducir el spec cerrado a código | Iterar el diseño — si aparece una duda de diseño, volver a Fase 1 | Sonnet |
| **3 · Producción** | Correr contra datos reales · resolver pendientes por triage de causa raíz | — | Haiku (mecánico) · Sonnet (código medio) · Opus (causa raíz no obvia o decisión de negocio) |

**Dos modos de trabajar Fase 1:**
- **Modo A — por dolor:** cada módulo recorre su ciclo completo empezando por el más urgente. Entrega valor rápido, genera deuda de integración.
- **Modo B — prompt-first:** spec completo de todos los módulos antes de implementar. Sin deuda de integración.

---

## Asignación de modelo por tipo de tarea

| Tarea | Modelo |
|---|---|
| Correr main.py / test.py | Haiku |
| Limpieza, eliminar código muerto | Haiku |
| Leer/escribir MD, actualizar memoria | Haiku |
| Crear HTML | Haiku |
| Codificar features normales | Sonnet |
| Corregir errores de código (mayoría de casos) | Sonnet |
| Responder preguntas de aprendizaje | Sonnet (Opus si es muy conceptual/profundo) |
| Diseñar modelo de negocio / la solución | Opus |
| Bugs con causa raíz no obvia / errores raros | Opus |

**Regla:** no cambiar de modelo a mitad de sesión. Cada bloque de tareas del mismo modelo va en su propia sesión. Al cerrar, actualizar memoria/docs/pendientes con el estado actual.

---

## Convenciones de nombres

| Qué | Formato | Ejemplo |
|---|---|---|
| Módulos del sistema | `NN_nombre_modulo` | `1_lecturas`, `2_planilla` |
| Carpetas internas | `lower_snake_case` | `inputs/`, `outputs/`, `docs/` |
| Archivos Python | `lower_snake_case.py` | `main.py`, `config.py` |
| Archivos Excel con periodo | `nombre_YYYY-MM.xlsx` | `planilla_2026-06.xlsx` |
| Archivos Excel sin periodo | `lower_snake_case.xlsx` | `convenios.xlsx`, `multas.xlsx` |
| Variables Python | `lower_snake_case` | `df_lecturas`, `total_mes` |
| Constantes Python (`config.py`) | `UPPER_SNAKE_CASE` | `TARIFA_M3`, `INPUTS_DIR` |
| Hojas Excel (sheets) | `Title_Case` | `Sin_identificar`, `Pagos_multiples`, `Pagos_comunitarios` |
| Columnas Excel (todas) | `UPPER_SNAKE_CASE` | `MZ`, `LT`, `MES_ACTUAL`, `TOTAL_A_PAGAR` |

Todas las columnas Excel — sin importar si las llenó un humano o el sistema — van en `UPPER_SNAKE_CASE`. La distinción se comunica con color, no con capitalización.

---

## Reglas: siempre X antes de Y

1. **README antes de arquitectura visual.** Si se contradicen → el README manda.
2. **Proponer en consola → esperar aprobación → crear HTML.** Tres triggers que obligan a proponer primero:
   - (A) Decisión con dimensión visual (flechas, cajas, tabla de columnas).
   - (B) Comparar 3+ opciones de diseño.
   - (C) Cambios a un módulo existente (nueva columna, nueva lógica, nuevo archivo).
3. **Auditar `formato_*.html` de todos los outputs ANTES de escribir una línea de código** (paso 3.0). Si falta alguno, crearlo primero.
4. **Actualizar `diagrama_flujo_MODULO.html` antes de codificar cambios en un módulo existente.** Esperar confirmación del 80%, recién entonces codificar.
5. **Completar CRAD antes de escribir el prompt al LLM.** (Contexto · Realidad · Ambigüedades · Decisiones.)
6. **Fase 1 cerrada antes de empezar Fase 2.** Si aparece una duda de diseño durante la implementación, parar y volver a Fase 1.
7. **Actualizar README al mismo tiempo que la estructura.** Si un cambio agrega, elimina, reordena o cambia el estado de un módulo, actualizar el pipeline en README.md como parte de ese mismo trabajo — no como tarea aparte ni pendiente para después.
8. **Verificar sincronía del README al cerrar sesión.** Antes de cerrar, confirmar que README.md refleje la estructura real de módulos actual. Si hay desajuste, corregirlo antes de actualizar memoria/pendientes.

---

## Patrones de arquitectura — aplican a todos los módulos

| Patrón | Regla en una línea |
|---|---|
| **Phase gate** | Ciclo `BORRADOR → PUBLICADA → COMPROMETIDA`. El commit point bloquea re-generación de la lista. Auto-release por nuevo `MES_ANO`. |
| **Thin layer** | Primitivos puros en `shared/utils_*.py`. Orquestación (orden de pasos, qué columnas, cuándo aplicar) siempre en `main.py`. Regla del Tres: no abstraer antes de 3 usos reales idénticos. |
| **Reconciliación bidireccional** | `SET_DEBE − SET_TIENE` → aplicar · `SET_TIENE − SET_DEBE` → revertir. Columna `ACCION` en audit log. |
| **Preservación de trabajo manual** | Backup + leer decisiones humanas + set de ya-procesados. Las 3 capas juntas en cada re-corrida. |
| **Writer único** | Un solo módulo escribe el archivo compartido. Los demás solo leen. |
| **Resiliencia** | Guard + Journal + Idempotencia desde día 1. Sidecar / Incremental cuando el volumen lo justifica. |

---

## Validación y regresión (ampliación de la regla 5)

Después de cualquier fix de código:
1. Corré el main.py/test.py del módulo afectado y confirmá que el
   problema puntual quedó resuelto (evidencia concreta, no "debería
   funcionar").
2. Identificá TODOS los módulos que consumen el output del módulo
   tocado (no solo el siguiente en el orden del pipeline — cualquier
   módulo que lea un archivo que el módulo tocado escribe). Corré el
   main.py de CADA uno de esos módulos y confirmá que siguen
   funcionando igual que antes del cambio — esto es la verificación de
   regresión, no solo del fix puntual.
3. Si el módulo tocado es un primitivo compartido (ej. shared/), corré
   los main.py de TODOS los módulos que lo importan o consumen,
   sin excepción.
4. Recién después de 1-3, marcá la tarea como resuelta en
   docs/pendientes_plan.md.

---

## Manejo de tareas bloqueadas dentro de un bloque

Cuando una tarea individual esté bloqueada (falta ambiente, depende de
una decisión no tomada, causa raíz no clara):
1. NO te detengas a pedir /inventario por una sola tarea bloqueada.
2. Marcala como bloqueada en docs/pendientes_plan.md con el motivo
   (una línea), y CONTINUÁ con la siguiente tarea del mismo bloque que
   sí se pueda resolver con el modelo activo.
3. Solo cuando TODAS las tareas restantes del bloque activo estén
   bloqueadas o completadas, evaluá si corresponde pasar al siguiente
   bloque/modelo (según el campo SIGUIENTE_ACCION) o si hace falta
   correr /inventario de nuevo.
4. NO corras /inventario por iniciativa propia. Solo sugerilo cuando se
   cumpla la condición del punto 3, o cuando se cumpla lo que dice la
   sección "Cuándo correr /inventario otra vez" — y esperá confirmación
   antes de correrlo.

---

## Repos de experimentación (jass_system2 o similares)

Antes de borrar un repo de experimentación, documentar en `docs/pendientes_plan.md` qué quedó hecho y qué no. Un repo de copia puede contener trabajo valioso no mergeado — revisarlo con `diff` antes de eliminar.

---

## Artefactos de módulo — dónde viven

```
MODULO/
├── docs/
│   ├── diagrama_flujo_MODULO.html   ← flujo en 5 segundos (LEE + GENERA por script)
│   ├── diagrama_MODULO.html         ← detalle de reglas y columnas
│   ├── arquitectura_MODULO.html     ← resumen visual del README
│   └── formato_ARCHIVO.html         ← contrato de formato por cada output Excel
├── inputs/
├── outputs/
└── README.md                        ← fuente de verdad
```
