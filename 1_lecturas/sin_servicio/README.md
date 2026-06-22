# sin_servicio

Submódulo de `1_lecturas` que justifica por qué un usuario no tiene lectura de medidor este mes
y mantiene el catálogo de usuarios sin servicio activo (`lista_sin_servicio.xlsx`).

Vive dentro de `1_lecturas/` porque su output (`validacion_ausencias_YYYY-MM.xlsx`) solo lo consume
ese módulo. Si en el futuro otros módulos lo necesitan (ej. `6_corte` para no cortar a estos usuarios),
se extrae como módulo independiente.

---

## El problema que resuelve

El operario llena `MARC_ACT` para cada usuario que recorre con el medidor. Algunos usuarios están en
el template como filas pero con `MARC_ACT` **vacío** — no porque estén ausentes del archivo, sino
porque no tienen medidor instalado o el operario los pasó por alto.

El código actual detecta eso como `SIN_LECTURA` (anomalía bloqueante), pero no sabe distinguir si es:

- **Sin medidor confirmado** — el usuario está catalogado en `lista_sin_servicio.xlsx`. No hay nada que corregir.
- **Posible error de captura** — el usuario no está en la lista. El operario pudo haberse saltado esa fila.

Sin este submódulo, todos los `SIN_LECTURA` sin `obs=P` son bloqueantes, incluyendo los 74 usuarios
que ya se sabe que no tienen medidor — lo cual genera ruido innecesario y ralentiza el cierre del ciclo.

---

## Los dos escenarios

### Escenario A — principal

Usuario **presente en el template** con `MARC_ACT` vacío (sin obs que lo justifique):

```
MARC_ACT vacío + obs vacío
    └─ ¿está en lista_sin_servicio con TIPO=SIN_MEDIDOR?
           SÍ → INFORMATIVA "sin_medidor_confirmado" (no bloquea)          ← NUEVO
           NO → BLOQUEANTE SIN_LECTURA (supervisor decide en reporte)       ← existente mejorado
```

### Escenario B — secundario

Usuario **completamente ausente del template** (fila no existe en `registro_operario_mes.xlsx`),
pero sí está en el historial del acumulado. Ocurre cuando el padrón creció y el template viejo
no tiene esa fila:

```
Usuario en acumulado · NO aparece como fila en el template
    └─ ¿está en lista_sin_servicio?
           SÍ → INFORMATIVA (ausencia de fila justificada)                  ← NUEVO
           NO → BLOQUEANTE USUARIO_FANTASMA (supervisor decide)             ← existente mejorado
```

En ambos casos, las anomalías que no están justificadas por la lista generan filas en
`validacion_ausencias_YYYY-MM.xlsx` que el supervisor debe resolver antes de cerrar el ciclo.

---

## Qué hace cada tool

### `migrar_lista.py` — one-shot

Convierte el archivo plano original (`Lista de usuarios sin servicio.xlsx` con solo MZ, LT, NOMBRE)
al schema nuevo con columnas TIPO, FECHA_INICIO, ÚLTIMA_REVISIÓN, MESES_SIN_LECTURA, NOTAS.
Todas las 74 filas migran con `TIPO = SIN_MEDIDOR`. Se corre **una sola vez** antes del primer ciclo.

### `validar_ausencias.py` — read-only, core del submódulo

Compara `registro_operario_acumulado.xlsx` vs `registro_operario_mes.xlsx` para identificar usuarios
con historial que no tienen lectura este mes. Cruza cada caso contra `lista_sin_servicio.xlsx` y clasifica.

**No muta nada.** Devuelve dos listas al orquestador (`main.py`):
- `justificadas[]` → informativas (no bloquean el ciclo)
- `revisar[]` → se suman a los bloqueantes del ciclo

Si existe `validacion_ausencias_YYYY-MM.xlsx` de una ejecución anterior, **preserva las decisiones
ya llenas por el supervisor**. Nunca sobreescribe trabajo hecho.

### `actualizar_lista.py` — writer único

Lee `validacion_ausencias_YYYY-MM.xlsx` con las decisiones del supervisor ya completas y aplica
los cambios a `lista_sin_servicio.xlsx`. Es el **único script que puede escribir la lista**.

Hace backup timestamped antes de cada escritura y registra cada cambio en `outputs/audit_lista.xlsx`
(append-only). Solo se puede correr cuando **todas las filas tienen DECISIÓN llena**.

`main.py` nunca llama a este script. Lo avisa al cerrar el ciclo, pero el supervisor lo corre manualmente.

### `clasificar_tipo.py` — read-only, fase 2

No implementado todavía. Disponible cuando el acumulado tenga ≥ 6 ciclos de historial.
Ver sección **Fase 2** al final.

---

## Cuándo se corre cada script

| Momento | Script | Quién | Condición |
|---|---|---|---|
| Una sola vez (migración inicial) | `migrar_lista.py` | Supervisor | Antes del primer ciclo con el nuevo schema |
| Automático — dentro de `main.py` paso 6 | `validar_ausencias.py` | Sistema | Cada vez que se corre `main.py` |
| Manual — después de llenar el reporte | `actualizar_lista.py` | Supervisor | Cuando todas las filas de `validacion_ausencias_YYYY-MM.xlsx` tienen DECISIÓN |

---

## Integración con `main.py`

Este submódulo se integra en **dos puntos exactos** dentro de `_detectar_anomalias()` en `main.py`:

**Punto 1 — al detectar `SIN_LECTURA` (escenario A):**
```python
if marc_act_val is None:
    if obs == "P":
        → informativa (predio cerrado)                        # ya existe
    elif (mz, lt) in lista_sin_servicio:                      # NUEVO
        → informativa "sin_medidor_confirmado"                # NUEVO
    else:
        → bloqueante SIN_LECTURA                              # ya existe
```

**Punto 2 — al detectar `USUARIO_FANTASMA` (escenario B):**
```python
for key in historial:
    if key not in keys_en_mes:
        if key in lista_sin_servicio:                         # NUEVO
            → informativa (ausencia justificada)              # NUEVO
        else:
            → bloqueante USUARIO_FANTASMA                     # ya existe
```

El gate de validación **bloquea el cierre del ciclo** si quedan filas con `DECISIÓN` vacía en
`validacion_ausencias_YYYY-MM.xlsx`. Estas filas suman a los bloqueantes del ciclo con la misma
lógica que `correcciones_YYYY-MM.xlsx`.

---

## Flujo de operación mensual

```
# PASO 1 — main.py corre normalmente
python main.py
   ← inputs/registro_operario_mes.xlsx
   ← inputs/registro_operario_acumulado.xlsx
   ← sin_servicio/inputs/lista_sin_servicio.xlsx         (nuevo — lee pero NO escribe)
   → sin_servicio/outputs/validacion_ausencias_YYYY-MM.xlsx   (si hay casos a revisar)
   → outputs/correcciones_YYYY-MM.xlsx                        (bloqueantes de lecturas)

# Si validacion_ausencias tiene filas con DECISIÓN vacía → ciclo no cierra.

# PASO 2 — Supervisor abre el reporte y llena DECISIÓN por cada fila
sin_servicio/outputs/validacion_ausencias_YYYY-MM.xlsx
   DECISIÓN = sin_servicio    → el usuario se agregará a la lista como SIN_MEDIDOR
   DECISIÓN = error_captura   → hay que agregar la lectura al template y re-correr
   DECISIÓN = investigar      → se agrega a la lista como EN_INVESTIGACIÓN
   DECISIÓN = ignorar         → no entra a la lista, solo queda en audit

# PASO 3 — Volver a correr main.py para re-verificar
python main.py
   → preserva las DECISIONES ya llenas en validacion_ausencias
   → si todas tienen DECISIÓN y no quedan bloqueantes → ciclo cierra

# PASO 4 — Aplicar decisiones a la lista (manual, después del cierre)
python sin_servicio/actualizar_lista.py
   ← sin_servicio/outputs/validacion_ausencias_YYYY-MM.xlsx   (con DECISIÓN completa)
   ← sin_servicio/inputs/lista_sin_servicio.xlsx
   → sin_servicio/inputs/lista_sin_servicio.xlsx               (actualizada)
   → sin_servicio/outputs/audit_lista.xlsx                     (append-only)

# main.py avisa al cerrar: "Hay N decisiones listas → python sin_servicio/actualizar_lista.py"
```

---

## Schema: `lista_sin_servicio.xlsx`

Writer único: `actualizar_lista.py`. Ubicación: `sin_servicio/inputs/lista_sin_servicio.xlsx`.
Nunca se edita manualmente. Backup automático antes de cada escritura.

| Columna | Tipo | Descripción |
|---|---|---|
| `MZ` | texto | Manzana del usuario |
| `LT` | texto | Lote del usuario |
| `NOMBRE` | texto | Nombre completo |
| `TIPO` | dropdown | `SIN_MEDIDOR` · `SIN_AGUA` · `EN_INVESTIGACIÓN` |
| `FECHA_INICIO` | YYYY-MM-DD | Fecha en que se agregó a la lista |
| `ÚLTIMA_REVISIÓN` | YYYY-MM-DD | Última vez que el supervisor confirmó o reclasificó |
| `MESES_SIN_LECTURA` | entero | Calculado del acumulado al leer la lista |
| `NOTAS` | texto libre | Observaciones del supervisor |

### TIPO — los 3 estados y sus reglas

| TIPO | Significado | Comportamiento en ciclos futuros |
|---|---|---|
| `SIN_MEDIDOR` | Sin medidor instalado — caso cerrado | **No aparece en el reporte** de validación. El sistema lo acepta automáticamente como informativa. |
| `SIN_AGUA` | Tiene medidor pero sin consumo este mes — temporal | **Aparece en el reporte** si falta su lectura, con DECISIÓN default `sin_servicio`. Se re-evalúa cada ciclo. |
| `EN_INVESTIGACIÓN` | Caso sin clasificar — bajo revisión activa | **Siempre aparece en el reporte** hasta que el supervisor tome una decisión. Auto-promoción a `SIN_MEDIDOR` tras 3 ciclos consecutivos sin lectura (registrado en NOTAS). |

**Migración inicial:** las 74 filas del archivo plano original arrancan todas con `TIPO = SIN_MEDIDOR`.
No aparecerán en el reporte en ciclos futuros. La reclasificación ocurrirá cuando `clasificar_tipo.py`
(fase 2) analice el historial del acumulado.

---

## Schema: `validacion_ausencias_YYYY-MM.xlsx`

Generado por `validar_ausencias.py`, llenado por el supervisor. Una fila por usuario sin lectura
que no está justificado en la lista. Las filas con `DECISIÓN` llena se preservan entre re-ejecuciones.

| Columna | Quién llena | Descripción |
|---|---|---|
| `MZ` | Sistema | Manzana |
| `LT` | Sistema | Lote |
| `NOMBRE` | Sistema | Nombre completo |
| `ESCENARIO` | Sistema | `SIN_LECTURA` (A) o `USUARIO_FANTASMA` (B) |
| `ÚLTIMO_MES_CON_LECTURA` | Sistema | Último ciclo YYYY-MM con MARC_ACT registrado en acumulado |
| `MARCACION_ÚLTIMA` | Sistema | Valor de esa última marcación (vacío si nunca tuvo) |
| `MESES_SIN_LECTURA` | Sistema | Ciclos consecutivos sin lectura hasta el mes actual |
| `ESTADO_LISTA` | Sistema | `NO_ESTÁ` o `EN_INVESTIGACIÓN` (los `SIN_MEDIDOR` nunca llegan aquí) |
| `DECISIÓN` | **Supervisor** | `sin_servicio` · `error_captura` · `investigar` · `ignorar` · vacío = bloqueante |
| `NOTAS` | Supervisor (opcional) | Observación libre |
| `CICLO` | Sistema | Número de ciclo del mes en que se detectó |
| `FECHA_DETECCIÓN` | Sistema | Timestamp de la ejecución |

### Opciones de `DECISIÓN` y su efecto

| DECISIÓN | Qué ocurre al correr `actualizar_lista.py` |
|---|---|
| `sin_servicio` | Agrega a la lista con `TIPO = SIN_MEDIDOR` y `FECHA_INICIO = hoy` |
| `error_captura` | No modifica la lista. Supervisor agrega la lectura al template y re-corre `main.py` |
| `investigar` | Agrega a la lista con `TIPO = EN_INVESTIGACIÓN` y `FECHA_INICIO = hoy` |
| `ignorar` | No modifica la lista. Queda registrado en `audit_lista.xlsx` con motivo `ignorado` |
| (vacío) | **Bloqueante** — impide el cierre del ciclo |

---

## Estructura de carpetas

```
sin_servicio/
├── README.md                              ← este archivo — fuente de verdad
├── config.py                              ← paths y constantes del submódulo
├── inputs/
│   └── lista_sin_servicio.xlsx           ← catálogo mutable — writer único: actualizar_lista.py
├── outputs/
│   ├── validacion_ausencias_YYYY-MM.xlsx ← reporte mensual — supervisor llena DECISIÓN
│   └── audit_lista.xlsx                  ← log append-only de cambios a la lista
├── docs/
│   ├── arquitectura_sin_servicio.html    ← visual de este README
│   ├── formato_lista_sin_servicio.html   ← contrato de formato del catálogo
│   └── formato_validacion_ausencias.html ← contrato de formato del reporte
├── tests/
├── validar_ausencias.py                  ← tool 1: read-only · core del submódulo
├── actualizar_lista.py                   ← tool 2: writer único de lista_sin_servicio.xlsx
├── migrar_lista.py                       ← tool 3: one-shot · 74 filas → schema nuevo
└── clasificar_tipo.py                    ← tool 4: futuro · fase 2
```

---

## Lifecycle de outputs

| Archivo | Lifecycle |
|---|---|
| `inputs/lista_sin_servicio.xlsx` | **PERMANECE** — mutable solo vía `actualizar_lista.py` · backup antes de cada escritura |
| `outputs/validacion_ausencias_YYYY-MM.xlsx` | Mensual — uno por ciclo · se regenera conservando decisiones llenas |
| `outputs/audit_lista.xlsx` | **PERMANECE** — append-only · historial de todas las mutaciones a la lista |

---

## Reglas de operación — las 10 decisiones de diseño

| # | Regla | Razón |
|---|---|---|
| 1 | El submódulo vive dentro de `1_lecturas/` | Su output solo lo consume `1_lecturas`. Se extrae si lo usan 2+ módulos. |
| 2 | Migración inicial: 74 filas → todas `SIN_MEDIDOR` | Arrancar con TIPO conocido. Reclasificación posterior con datos reales. |
| 3 | Gate en paso 6 de `main.py`, después de detectar anomalías | Necesita el estado post-correcciones para no duplicar trabajo ni generar falsos positivos. |
| 4 | Archivo separado `validacion_ausencias_YYYY-MM.xlsx` | Correcciones = valores anómalos. Ausencias sin lectura = usuarios sin medidor. Mezclarlos confunde. |
| 5 | TIPO con 3 estados: `SIN_MEDIDOR` · `SIN_AGUA` · `EN_INVESTIGACIÓN` | Cada estado define una conducta distinta del sistema en ciclos futuros. |
| 6 | `DECISIÓN` vacía bloquea el cierre del ciclo | Sin disposición explícita, los casos se acumulan. El estado "pendiente" no es válido para cierre contable. |
| 7 | `EN_INVESTIGACIÓN` se auto-promueve a `SIN_MEDIDOR` tras 3 ciclos sin lectura | Sin escape, la categoría se vuelve cementerio. La promoción se registra en NOTAS, no es silenciosa. |
| 8 | `clasificar_tipo.py` queda para fase 2 | Solo hay 2 meses de acumulado. Con poca data, la clasificación automática genera ruido. Mínimo 6 ciclos. |
| 9 | Decisiones llenas se preservan en re-ejecuciones | El supervisor puede correr `main.py` varias veces. Nunca perder trabajo ya hecho. |
| 10 | `actualizar_lista.py` es paso manual explícito — `main.py` no lo llama | Las mutaciones de estado son siempre explícitas. `main.py` es read-only respecto a la lista. |

---

## Lo que este submódulo NO hace

- **No decide si un usuario tiene o no medidor** — esa decisión siempre es del supervisor via `DECISIÓN`.
- **No modifica `registro_operario_acumulado.xlsx`** — ese archivo lo gobierna `1_lecturas/main.py` y `aplicar_sincronizacion.py`.
- **No modifica `registro_operario_mes.xlsx`** — si `DECISIÓN = error_captura`, el supervisor agrega la lectura manualmente.
- **No elimina usuarios de la lista** — `actualizar_lista.py` solo agrega y actualiza TIPO. Las bajas son manuales con NOTAS.
- **No calcula cobros para usuarios sin medidor** — eso lo maneja `2_planilla` con la columna `SIN_SERVICIO` del acumulado.
- **No detecta patrones automáticamente** — esa es la responsabilidad de `clasificar_tipo.py` (fase 2).

---

## Errores comunes

**"lista_sin_servicio.xlsx no existe"**
→ Correr `migrar_lista.py` primero. Si el archivo original `Lista de usuarios sin servicio.xlsx` tampoco existe,
crearlo con las 3 columnas mínimas (MZ, LT, NOMBRE) y correr la migración.

**"Todas las filas del reporte son DECISIÓN vacía y el ciclo no cierra"**
→ El supervisor tiene que abrir `validacion_ausencias_YYYY-MM.xlsx` y llenar la columna DECISIÓN por cada fila.
Ver opciones válidas en la tabla de DECISIÓN arriba.

**"Hay usuarios con `error_captura` pero el ciclo sigue sin cerrar"**
→ `error_captura` significa que hay que agregar la lectura al template del operario y re-correr `main.py`.
Los usuarios con esa decisión no desaparecen del reporte hasta que el template tenga su `MARC_ACT`.

**"`actualizar_lista.py` falla con 'decisiones incompletas'"**
→ Hay filas con `DECISIÓN` vacía en el reporte. Completar todas las filas antes de correr el writer.
Solo se puede aplicar cuando el reporte está 100% resuelto.

**"Un usuario `SIN_MEDIDOR` apareció en el reporte de validación"**
→ Inconsistencia entre la lista y el acumulado. Verificar que `MESES_SIN_LECTURA` en la lista sea correcto.
Si el usuario tuvo lectura este mes, `clasificar_tipo.py` (fase 2) lo detectaría y propondría reclasificación.
Por ahora, revisar manualmente y actualizar `TIPO` a `SIN_AGUA` si corresponde.

**"El reporte tiene el mismo usuario en múltiples ciclos sin resolverse"**
→ Si `TIPO = EN_INVESTIGACIÓN`, aparece en el reporte cada ciclo hasta que el supervisor decida.
Después de 3 ciclos sin lectura, se auto-promueve a `SIN_MEDIDOR` (ya no aparece más).
Si lleva menos de 3 ciclos, es esperado — el supervisor debe cerrar el caso.

---

## Señales de alerta

| Señal | Diagnóstico |
|---|---|
| Más de 10 filas nuevas en `validacion_ausencias` en un ciclo | El padrón puede haber tenido una actualización masiva. Revisar si `aplicar_sincronizacion.py` corrió este mes. |
| `EN_INVESTIGACIÓN` con más de 2 ciclos sin decisión | El supervisor está ignorando el reporte. En el siguiente ciclo se auto-promueven a `SIN_MEDIDOR`. |
| `audit_lista.xlsx` no crece después de correr `actualizar_lista.py` | El script no encontró decisiones nuevas que aplicar. Verificar que el reporte tiene DECISIÓN llena y que es del ciclo correcto. |
| Usuario con `error_captura` que nunca desaparece del reporte | El operario no agregó la lectura al template. Recordar que `error_captura` requiere acción manual del operario. |
| Lista con más de 150 usuarios `SIN_MEDIDOR` | El padrón creció significativamente. Considerar activar `clasificar_tipo.py` (fase 2) para auditar si algunos ya tienen medidor. |

---

## Fase 2 — `clasificar_tipo.py`

Disponible cuando el acumulado tenga ≥ 6 ciclos de historial. Analiza patrones de consumo y
propone reclasificaciones:

- `SIN_MEDIDOR` con lectura > 0 detectada → propone `SIN_AGUA`
- `SIN_AGUA` con consumo > 0 consistente en últimos 3 ciclos → propone salida de la lista
- `EN_INVESTIGACIÓN` con 3+ ciclos sin lectura → auto-promoción a `SIN_MEDIDOR` (ya activo desde fase 1)

El tool propone — nunca aplica. Las sugerencias van al reporte como columna `RECLASIFICACIÓN_SUGERIDA`.
El supervisor acepta o rechaza cada una antes de correr `actualizar_lista.py`.

**Condición para implementar:** cuando el supervisor haya procesado al menos 3 ciclos de
`validacion_ausencias` y se tenga evidencia real de qué patrones son confiables en esta JASS.

---

## Referencias

- **Arquitectura visual:** `docs/arquitectura_sin_servicio.html`
- **Contrato de formato del catálogo:** `docs/formato_lista_sin_servicio.html`
- **Contrato de formato del reporte:** `docs/formato_validacion_ausencias.html`
- **Módulo padre:** `1_lecturas/README.md`
- **Integración en el código:** `1_lecturas/main.py` función `_detectar_anomalias()` — puntos SIN_LECTURA y USUARIO_FANTASMA
