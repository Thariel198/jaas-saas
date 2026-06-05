# Metodología de Desarrollo — jass_system
## Cómo construir sistemas de IA sin improvisar

Versión 1.2 — Mayo 2026
Basada en la experiencia construyendo motor_matching y jass_system

---

## Filosofía base

> "Conozco el inicio y el fin del héroe. Los personajes y sus historias
> se descubren en el camino."

Esto es **diseño iterativo** — la forma profesional de construir sistemas.
No significa improvisar. Significa tener un mapa claro antes de caminar,
y ajustar el mapa cuando la realidad lo pide.

Lo que sí debes saber desde el inicio: input, output y orden del pipeline.
Lo que se descubre en el camino: las reglas internas de cada módulo.

### Documentation-Driven Design

El README se escribe **antes** que la arquitectura visual. No al revés.

Si no puedes explicar el módulo en texto claro, no puedes diseñarlo bien.
El visual es un resumen del README — no la fuente de verdad.

**Por qué esto importa:**
- El README obliga a pensar con precisión antes de dibujar
- Si README y arquitectura se contradicen → el README manda
- Es más rápido corregir texto que corregir HTML
- El README es lo que el código implementa — el visual solo ayuda a visualizar

```
README (fuente de verdad)
    ↓
Arquitectura visual (resumen del README)
    ↓
Código (implementa el README)
```

### HTML como Contrato de Diseño

Un archivo HTML aprobado es una especificación visual que el código debe replicar exactamente.

**Regla:** Antes de tocar código, pedir el HTML actual. Nunca generarlo desde cero si ya existe uno aprobado. El HTML es el contrato — el código lo implementa.

---

## Los 3 activos que construyes en cada proyecto

### Activo 1 — Skills
Carpeta de mejores prácticas reutilizables. Lo que aprendiste con errores
en proyectos anteriores, documentado para no repetirlos.

```
skills/
├── metodologia_desarrollo.md  ← este archivo
├── motor_matching.md          ← cómo construir motores de matching
├── trazabilidad.md            ← cómo diseñar sistemas de trazabilidad
└── override.md                ← cómo diseñar correcciones post-validación
```

**Regla:** cada vez que resuelves algo que te costó tiempo, lo conviertes en skill.

### Activo 2 — Metodología
El proceso que sigues en cada proyecto. Lo que estás leyendo ahora.
Se optimiza con el tiempo — cada proyecto la mejora.

### Activo 3 — Arquitectura del sistema
El mapa del proyecto actual. Vive en el repositorio del proyecto.
No en tu cabeza.

---

## Cómo usar las herramientas IA — división profesional

No todas las tareas son para la misma herramienta. Usar la herramienta equivocada
cuesta tiempo, dinero y contexto acumulado.

### La regla fundamental

**Chat = pensar. Code = ejecutar. Scheduled = cosechar.**

El costo oculto de usar Code para explorar: no son solo los tokens. Es el costo
de dar instrucciones imprecisas, recibir algo incorrecto, corregir, reiterar.
Ese ciclo sale mucho más caro que haber explorado primero en Chat.

### Claude Chat (claude.ai)

Úsalo cuando **no necesitas ver el código real** del proyecto.

| Cuándo usarlo | Por qué |
|---|---|
| Ordenar ideas, explorar el problema | Iterás muchas veces sin costo por token |
| Evaluar enfoques antes de diseñar | Brainstorm barato antes de comprometerte |
| Diseñar outputs HTML desde cero | No necesita leer el repo, solo el contrato de diseño |
| Escribir el README de un módulo nuevo | No hay código que leer aún |
| Preparar el prompt que le darás a Claude Code | Un prompt bien escrito vale 10x en Code |
| Discutir si un enfoque es correcto | Exploración barata antes de codificar |

**Regla económica:** Chat tiene costo fijo (suscripción). Iterá gratis hasta que el diseño es claro.

### Claude Code

Úsalo cuando **necesitas el contexto real del proyecto**.

| Cuándo usarlo | Por qué |
|---|---|
| Implementar un diseño ya aprobado | Llegás con prompt listo, ejecuta con precisión |
| Debuggear errores reales | Lee el traceback, el archivo, el contexto exacto |
| Diseñar cuando el diseño depende del código existente | Lee los archivos sin copy/paste |
| Git, tests, validaciones | Son operaciones técnicas sobre el repo real |
| Refactorizar código existente | Necesita leer el código actual para proponer cambios |

**Regla económica:** Code cobra por token. Llegá con instrucciones precisas. Ejecuta una vez, bien.

**Problema de la compresión:** En sesiones largas, el contexto anterior se comprime y Code
pierde el hilo. Para tareas de diseño que no requieren ver el código, usá Chat.

### Claude Scheduled (fin de día)

Úsalo para **actualizar tus activos** sin ocupar la sesión de trabajo.

| Cuándo usarlo | Por qué |
|---|---|
| Actualizar skill_tracker | Lee git log del día, sabe exactamente qué practicaste |
| Escribir reporte_aprendizaje del día | Contexto limpio, no acumulado de horas de trabajo |
| Actualizar CHANGELOG | Resume el día sin interrumpir el flujo de código |

**Regla:** el Claude de trabajo construye. El Claude de cierre cosecha.

### El flujo completo

```
CHAT (exploración barata)        CODE (ejecución sobre repo)     SCHEDULED (cierre)
         │                                   │                          │
1. Ordenar ideas                    4. Leer archivos reales      7. git log del día
2. Evaluar enfoques                 5. Implementar diseño        8. Actualizar skills
3. Escribir README nuevo            6. Tests + git commit        9. Reporte aprendizaje
4. Preparar prompt para Code
         │
         └──── prompt claro ─────────────────►
```

---

## La metodología — fases de desarrollo

La metodología tiene 4 fases obligatorias (1 a 4) más una fase 0 condicional.

---

### FASE 0 — Exploración *(solo cuando no conocés el problema o los datos)*

**Objetivo:** entender el problema real antes de diseñarlo. Reducir la incertidumbre
antes de comprometerte con un README y una arquitectura.

Si llegás a Fase 1 sin entender bien los datos o las reglas de negocio, escribís
un README incorrecto, que produce un diseño incorrecto, que produce código incorrecto.
Lo descubrís en Fase 4 cuando corrés contra datos reales. Eso es exactamente el costo
que esta metodología busca evitar.

#### Cuándo aplicar Fase 0

Aplicar cuando al menos una de estas condiciones es verdadera:
- Primera vez que trabajás con este tipo de datos
- Las reglas de negocio no están claras todavía
- No sabés qué casos borde existen en los datos reales
- Tenés dudas sobre si tu enfoque va a funcionar con los datos reales

No aplicar cuando:
- Ya implementaste módulos similares antes y conocés bien los datos
- Es una extensión de un módulo existente con reglas conocidas

#### Duración máxima

2 a 4 horas. Si necesitás más, dividir en bloques de 2-4 horas con un output concreto
al final de cada bloque. Si después de 4 horas seguís sin entender el problema,
el problema es más complejo de lo que pensabas — documentar eso como hallazgo
y redefinir el alcance antes de continuar.

#### Regla fundamental de Fase 0

**Este código no es producción.** No importa si está sucio, sin manejo de errores,
sin logging, sin estructura. Su único objetivo es responder preguntas sobre el problema.
No te autoengañes pensando que después lo vas a "limpiar" para usarlo en producción:
ese camino produce el código sin diseño que la metodología busca evitar.

#### Pasos

**0.1 — Escribir las preguntas que no podés responder sin ver los datos**

Antes de tocar el código, listar las dudas concretas:
- ¿Qué columnas existen realmente y qué contienen?
- ¿Qué casos borde pueden existir? (registros sin MZ, nombres duplicados, etc.)
- ¿Las reglas de negocio que asumo son correctas?
- ¿El volumen de datos es manejable con el enfoque que pienso usar?
- ¿Hay datos sucios o inesperados que cambiarían el diseño?

Escribirlas antes de explorar — no improvises la exploración,
porque sin preguntas claras terminas mirando datos sin saber qué buscás.

**0.2 — Script de exploración rápida**

Escribir un script desechable que responda esas preguntas una por una.
Puede ser sucio. No necesita estructura. No necesita manejo de errores.

```python
# exploracion_MODULO.py — DESECHABLE, no es producción
import pandas as pd

df = pd.read_excel("inputs/padron_secundario.xlsx")

# Pregunta 1: ¿qué columnas existen realmente?
print(df.columns.tolist())
print(df.dtypes)

# Pregunta 2: ¿cuántos registros sin MZ?
print(df[df["MZ"].isna()].shape[0], "registros sin MZ")

# Pregunta 3: ¿hay MZ+LT duplicados?
dupes = df[df.duplicated(subset=["MZ", "LT"], keep=False)]
print(dupes.shape[0], "registros duplicados por MZ+LT")
print(dupes.head(10))
```

**0.3 — Documentar los hallazgos**

Al terminar la exploración, escribir una lista de lo que encontraste:
- Qué encontraste que no esperabas
- Qué reglas de negocio se confirmaron o descartaron
- Qué casos borde existen en los datos reales
- Cualquier cosa que cambie el diseño que tenías en mente

Esta lista no necesita formato — puede ser texto plano o bullet points.
Es el input de Fase 1.

**0.4 — Decidir si el diseño original sigue en pie**

Con los hallazgos del paso anterior, preguntarse:
- ¿El problema que creía tener es el problema real?
- ¿El enfoque que tenía en mente sigue siendo válido con estos datos?
- ¿Hay algo que cambia el diseño que iba a hacer en Fase 1?

Si los hallazgos cambian el diseño → ajustar el enfoque antes de avanzar.
Si el diseño sigue siendo válido → avanzar a Fase 1 con certeza, no con suposiciones.

#### Destino del código de exploración

El script de exploración **no se mantiene en el proyecto**.
Moverlo a `backup/exploracion/exploracion_MODULO.py` cuando termines.
Si no existe esa carpeta, crearla.

#### Entregables de Fase 0
- [ ] Lista de preguntas respondidas (hallazgos)
- [ ] Script de exploración movido a `backup/exploracion/`
- [ ] Decisión documentada: ¿el diseño original sigue en pie o se ajusta?

**Señal de que podés pasar a Fase 1:**
- Podés describir el problema con precisión (no con suposiciones)
- Los casos borde están identificados
- Sabés qué reglas de negocio son reales y cuáles eran suposiciones

---

### FASE 1 — Mapa general
**Objetivo:** conocer el inicio y el fin. No los detalles internos.

#### Pasos

**1.1 README del sistema** ← primero siempre
Crear `README.md` en la raíz con:
- Qué hace el sistema completo
- El orden del pipeline
- Qué vive en shared/
- Cómo instalar y correr

**1.2 Arquitectura general** ← resumen visual del README
Crear `arquitectura_NOMBRE_SISTEMA.html` con:
- Todos los módulos en orden de pipeline
- Input y output de cada módulo
- Estado de cada módulo (hecho / siguiente / futuro / usa IA)
- Recursos shared entre módulos

> Si README y arquitectura se contradicen → el README manda.
> Corregir la arquitectura para que cuente la misma historia que el README.

**1.3 Crear rutas de carpetas**
Crear físicamente la estructura de carpetas del sistema.

**1.4 Validar el mapa**
Leer README junto con arquitectura y preguntarse:
- ¿El orden del pipeline tiene sentido?
- ¿Los inputs y outputs de cada módulo conectan correctamente?
- ¿Hay módulos que faltan?
- ¿Hay responsabilidades en el lugar equivocado?
- ¿README y arquitectura cuentan la misma historia?

Corregir antes de avanzar.

#### Entregables de Fase 1
- [ ] `README.md` raíz
- [ ] `arquitectura_SISTEMA.html`
- [ ] Carpetas físicas creadas
- [ ] README y arquitectura cuentan la misma historia

---

### FASE 2 — Diseño de módulo
**Objetivo:** conocer los personajes y su historia antes de codificar.
Se hace módulo por módulo, empezando por el que más dolor causa.

#### Criterio para elegir qué módulo atacar primero
No el primero del pipeline. El que más valor entrega hoy o más dolor causa.

#### Pasos — en este orden exacto

**2.0 Elegir el enfoque correcto** ← antes de cualquier diseño

Codificar la primera idea que se te ocurre es el error más caro de este proceso.
El costo de cambiar un enfoque antes del README es cero.
El costo de cambiar un enfoque cuando el código ya está escrito es empezar desde cero.

> Caso real: 02_matching_backup nació porque se codificó la primera idea (enfoque de reportes).
> La versión correcta nació de preguntarse "¿cómo lo haría manual?". Se perdieron semanas.

Este paso se hace en **Claude Chat** — no en Claude Code. No necesitás ver el código aún.

**Paso 2.0.1 — Describir el problema, no la solución**

Escribir en Chat:
> "Tengo este problema: [qué pasa]. Input: [X]. Output esperado: [Y].
> Restricciones: [tiempo disponible, si necesita validación manual, volumen de datos, etc.]"

Clave: describir el *problema*, no la solución que ya se te ocurrió.

**Paso 2.0.2 — Definir criterios de éxito antes de ver soluciones**

Escribir en Chat:
> "Antes de diseñar nada, ayudame a definir: ¿cómo sabré que la solución es correcta?
> ¿Qué casos borde existen? ¿Qué es lo que más temo que falle?"

Escribir las respuestas. Estos criterios son el filtro que vas a usar en el paso 2.0.5.

Ejemplo de criterio bien escrito:
> "La solución debe ser auditable fila por fila por una persona no técnica."

Este solo criterio ya descarta automáticamente cualquier enfoque complejo de reportes
o transformaciones encadenadas. Es el trabajo de 5 minutos que ahorra semanas.

**Paso 2.0.3 — Solución manual (el ancla)**

Escribir en Chat:
> "¿Cómo resolvería esto una persona con papel y lápiz, paso a paso?"

La solución manual siempre existe, siempre es auditable, y revela la lógica natural
del problema. Si el código imita el proceso manual, es más fácil de entender,
verificar y corregir cuando algo falla. Empezar por aquí no es limitante — es la base.

**Paso 2.0.4 — Al menos 3 enfoques distintos**

Escribir en Chat:
> "Dame 3 enfoques distintos para resolver este problema. Para cada uno:
> cómo funciona, cuándo es la mejor opción, cuándo falla, complejidad de implementar."

No preguntes cuál es mejor todavía. Primero que existan las opciones sin juzgarlas.
Siempre incluir la solución manual del paso anterior como uno de los enfoques.

**Paso 2.0.5 — Evaluar contra criterios**

Escribir en Chat:
> "Evalúa estos 3 enfoques contra los criterios que definimos en el paso anterior.
> ¿Cuál recomendás? ¿Qué señal temprana me diría que el enfoque elegido está fallando?"

La señal de alerta es tan importante como la evaluación. Si esa señal aparece en la
primera semana de implementación, es el momento de cambiar — no cuando ya está todo hecho.

**Paso 2.0.6 — Registrar la decisión**

Antes de escribir el README, documentar en 5 líneas en `docs/decisiones/MODULO.md`:

```
Problema: [qué problema resuelve este módulo]
Criterios: [cómo se ve el éxito]
Enfoque elegido: [cuál y por qué]
Alternativas descartadas: [cuáles y por qué cada una]
Señal de alerta: [qué indicaría que el enfoque elegido está fallando]
```

Cuando en 6 meses te preguntes "¿por qué lo hice así?", este archivo tiene la respuesta.
Cuando alguien proponga cambiar el enfoque, este archivo explica por qué se descartó.

**Señal de que podés pasar al 2.1:**
- [ ] Hay un enfoque elegido con criterios claros documentados
- [ ] Las alternativas están descartadas con razón escrita
- [ ] Sabés cuál es la señal de alerta temprana

---

**2.1 README del módulo** ← primero siempre
Crear `README_MODULO.md` con:
- Qué hace
- Cuándo se corre
- Estructura de carpetas
- Reglas de negocio detalladas
- Flujo paso a paso con comandos
- Tabla de lifecycle (qué se borra, qué permanece)
- Lo que NO hace este módulo
- Errores comunes

**2.2 Arquitectura del módulo** ← resumen visual del README
Crear `arquitectura_MODULO.html` con:
- Estructura de carpetas del módulo — exactamente como dice el README
- Inputs con su origen
- Outputs con su destino
- Prioridades o reglas de negocio principales
- Flujo mensual o de uso
- Qué NO hace este módulo

> Si arquitectura y README se contradicen → el README manda.

**2.3 Validar que cuentan la misma historia**
- ¿Las carpetas que muestra la arquitectura son las del README?
- ¿Los outputs mencionados en arquitectura coinciden con el README?
- ¿Hay algo en la arquitectura que el README no menciona? → eliminar o agregar al README primero

**2.4 Crear rutas de carpetas del módulo**
Crear físicamente las carpetas internas según el README.

**2.5 Diseño de outputs** ← README de cada archivo complejo
Por cada archivo que produce el módulo, diseñar:
- Columnas y su significado
- Quién escribe cada campo (sistema / usuario / banco)
- Reglas de negocio del archivo
- README propio si el archivo es complejo

Orden sugerido:
1. Inputs que el usuario llena (ej: pendientes.xlsx)
2. Archivos de trazabilidad
3. Output principal (ej: pagos_yape.xlsx)

**2.6 Definir el esquema de inputs esperado**
Por cada input del módulo, documentar:
- Columnas requeridas y sus nombres exactos (incluyendo alias aceptados)
- Tipos de dato esperados
- Qué pasa si falta un archivo o una columna

Esto se convierte en la validación que corre al inicio del módulo (ver Fase 3).

**2.7 Validar el diseño completo**
- ¿Las columnas diseñadas son suficientes para el código?
- ¿Hay reglas de negocio que no están documentadas?
- ¿Los archivos temporales y permanentes están bien identificados?
- ¿El módulo hace algo que no le corresponde?

Corregir antes de avanzar.

#### Entregables de Fase 2 — por módulo
- [ ] `README_MODULO.md` — fuente de verdad
- [ ] `arquitectura_MODULO.html` — resumen visual del README
- [ ] README y arquitectura cuentan la misma historia
- [ ] Carpetas físicas del módulo creadas
- [ ] Diseño de cada output (columnas + reglas)
- [ ] README por archivo complejo
- [ ] Esquema de inputs documentado

---

### FASE 3 — Código
**Objetivo:** traducir el diseño a código limpio en una sola pasada.
Solo se llega aquí cuando Fase 2 está completa y validada.

#### Regla de oro
Si durante el código descubres algo que cambia el diseño:
1. Para el código
2. Actualiza el README primero
3. Actualiza la arquitectura para que coincida con el README
4. Continúa el código

Nunca dejes que el código contradiga la documentación.

#### Pasos

**3.1 Revisar skills relevantes**
Antes de escribir una línea, leer las skills que aplican al módulo.

**3.2 Validación de inputs al inicio del módulo**
El primer bloque de código verifica que los archivos de entrada existen
y tienen las columnas requeridas según el esquema de Fase 2:

```python
def validar_inputs():
    if not MAESTRO_PATH.exists():
        raise FileNotFoundError(f"Falta: {MAESTRO_PATH}")
    df = pd.read_excel(MAESTRO_PATH)
    requeridas = {"MZ", "LOTE", "NOMBRE"}
    faltantes = requeridas - set(df.columns)
    if faltantes:
        raise ValueError(f"Columnas faltantes en maestro_yape.xlsx: {faltantes}")
```

Esto evita que el módulo falle a la mitad con un error confuso.

**3.3 Logging a archivo**
Cada módulo escribe un log con timestamp, no solo print a consola:

```python
import logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("outputs/run.log", encoding="utf-8"),
    ]
)
log = logging.getLogger(__name__)

log.info("Iniciando motor_matching")
log.warning("3 registros sin MZ/LOTE — van a blancos")
log.error("maestro_yape.xlsx no encontrado")
```

El archivo `run.log` queda en outputs/ junto al Excel. Si algo falla, el log dice exactamente qué y cuándo.

**3.4 Documentar el por qué en el código**
Cada función importante lleva un comentario del por qué existe,
no solo del qué hace:

```python
# Prioridad máxima al ciclo 2+ porque el usuario ya validó manualmente
# este origen — no tiene sentido volver a procesar el matching automático
if ciclo > 1 and corr_mult:
```

**3.5 Manejo de errores descriptivo**
Cada error dice qué archivo falló, qué columna, qué valor:

```python
raise ValueError(f"MZ-LOTE '{mz}-{lote}' no existe en planilla del mes")
```

**3.6 Tests**
Por cada función crítica, tests con casos conocidos:

```python
# test_matching.py
assert extraer_mz_lote("Mz A Lt 7") == ("A", "7")
assert extraer_mz_lote("MzE Lt7, MzP Lt11A") == [("E","7"), ("P","11A")]
```

Niveles de test según la criticidad del módulo:
- **Unitario** — función por función con datos inventados
- **Integración** — flujo completo con datos reales conocidos
- **Regresión** — comparar output actual con output de referencia guardado

**3.7 Commits con Conventional Commits**
Formato estándar que hace el git log legible como historia del proyecto:

```
tipo: descripción corta en presente

feat:     nueva funcionalidad
fix:      corrección de bug
docs:     cambio solo en documentación
refactor: cambio de código sin cambiar comportamiento
test:     agregar o corregir tests
chore:    mantenimiento (dependencias, configuración)
```

Ejemplos:
```bash
git commit -m "feat: validación de inputs antes de correr motor"
git commit -m "fix: ciclo 2 no duplicaba registros en trazabilidad"
git commit -m "docs: README motor_matching con tabla lifecycle actualizada"
```

Al terminar un ciclo mensual completo, taggear:
```bash
git tag v2026-05 -m "Ciclo Mayo 2026 — completado"
```

#### Entregables de Fase 3
- [ ] Validación de inputs al inicio
- [ ] Logging a archivo
- [ ] Código funcional
- [ ] Comentarios del por qué en funciones importantes
- [ ] Manejo de errores descriptivo
- [ ] Tests básicos
- [ ] Commits en formato Conventional Commits

---

### FASE 4 — Validación y cierre
**Objetivo:** verificar que el sistema hace lo que el diseño dice.

#### Pasos

**4.1 Correr con datos reales**
Usar datos del mes actual y verificar que los outputs coinciden
con lo diseñado en Fase 2.

**4.2 Verificar el log**
Leer `outputs/run.log` y confirmar que no hay warnings ni errores inesperados.

**4.3 Tag del ciclo**
Al confirmar que el mes cerró correctamente:
```bash
git tag v2026-05 -m "Ciclo Mayo 2026 — completado"
```

**4.4 Actualizar skills**
Si encontraste algo nuevo que costó tiempo resolverlo, convertirlo en skill.

**4.5 Actualizar CHANGELOG**
Anotar en `docs/CHANGELOG.md` (no en el README):
- Qué funcionó bien
- Qué cambiaría la próxima vez
- Qué decisiones de diseño resultaron correctas

#### Entregables de Fase 4
- [ ] Módulo validado con datos reales
- [ ] Log revisado sin errores inesperados
- [ ] Tag de ciclo en Git
- [ ] Skills actualizadas
- [ ] CHANGELOG actualizado

---

## Reglas generales — siempre

**README primero — Documentation-Driven Design**
Siempre escribir el README antes de la arquitectura visual.
El README es la fuente de verdad. La arquitectura es su resumen visual.
Si se contradicen → el README manda, se corrige la arquitectura.

**Diseño antes de código**
Nunca escribir código sin tener el README y la arquitectura completos.
El costo de cambiar un diseño es una línea. El costo de cambiar código es horas.

**Una sola fuente de verdad**
Si un archivo es usado por varios módulos → vive en `shared/`.
Si una regla está documentada en dos lugares → solo debe estar en uno.

**Deuda técnica visible**
Si algo está mal ubicado o es código viejo que no pertenece → va a `backup/`
con un README que explica por qué existe y cuándo se elimina.

**Git es la memoria del proyecto**
Todo cambio va a Git con mensaje en formato Conventional Commits.
Un archivo en una carpeta sin Git no existe para el equipo.
Al cerrar cada ciclo mensual → tag con la versión del mes.

**El mapa se actualiza cuando cambia el territorio**
Si descubres que un módulo no va donde pensabas → actualiza el README primero,
luego la arquitectura. Un mapa desactualizado es peor que no tener mapa.

**Los inputs se validan antes de correr**
Ningún módulo asume que sus archivos de entrada están bien.
La validación al inicio del módulo es la primera línea de defensa.

**El log es la memoria de cada ejecución**
Cada corrida escribe un `run.log`. Si algo falla, el log dice qué y cuándo.
El log va en `outputs/` junto al Excel que generó.

**Lo manual es el ancla**
Ante cualquier problema de diseño, la primera pregunta es siempre:
"¿Cómo lo resolvería una persona con papel y lápiz?"
La respuesta revela la lógica natural del problema. El código que imita ese proceso
es más fácil de entender, auditar y corregir que el código que inventa una solución
desde cero usando lo que la IA puede hacer.

**Nunca codifiques la primera idea**
Explorar al menos 3 enfoques antes de comprometerte con uno es parte del proceso,
no una pérdida de tiempo. Ver la sección 2.0 para el flujo completo.
El costo de explorar en Chat es cero. El costo de cambiar de enfoque cuando el
código ya está terminado puede ser empezar desde cero.

---

## Protocolo de rollback

Cuando un módulo falla a mitad del ciclo:

1. **No entrar en pánico** — el backup está en Git
2. Identificar en qué paso falló leyendo el `run.log`
3. Si el output fue sobreescrito parcialmente → restaurar desde Git:
   ```bash
   git checkout -- outputs/pagos_yape_tepago.xlsx
   ```
4. Si el fallo fue en un input que tú llenaste → corregir el input
5. Volver a correr desde el módulo fallido — no desde el inicio
6. Verificar que el output coincide con el diseño antes de continuar

**Regla:** nunca borrear archivos de outputs manualmente. Si algo está mal,
restaurar desde Git o desde la carpeta `backup/` del módulo.

---

## Gestión de dependencias

Mantener un `requirements.txt` en la raíz del proyecto:

```
pandas>=2.0
openpyxl>=3.1
python-docx>=1.0
docx2pdf>=0.1
```

Cada vez que instalas una librería nueva:
```bash
pip freeze > requirements.txt
git commit -m "chore: actualizar requirements.txt"
```

Esto garantiza que el sistema funcione igual en otra máquina o después de meses.

---

## Señales de que vas bien

- Puedes explicar el sistema completo en 2 minutos sin abrir código
- README y arquitectura cuentan exactamente la misma historia
- Cuando algo cambia, sabes exactamente qué documentos actualizar — primero el README
- El código no tiene sorpresas — hace lo que el README dice
- Los errores tienen mensajes claros que dicen exactamente qué falló
- Puedes retomar el proyecto después de semanas sin perder contexto
- El `run.log` de la última ejecución no tiene errores inesperados
- El `git log` se lee como la historia del proyecto

## Señales de que algo está mal

- README y arquitectura se contradicen
- La arquitectura tiene información que el README no menciona
- Hay código que no sabes para qué sirve
- Un archivo vive en una carpeta que no le corresponde
- Descubres funcionalidad nueva mientras codificas (debería estar en diseño)
- No puedes explicar por qué tomaste una decisión de arquitectura
- Un módulo falla sin decir qué archivo o columna causó el problema

---

## Orden de prioridad cuando hay conflicto

1. **README del módulo** — fuente de verdad · las reglas de negocio mandan
2. **Arquitectura visual** — resumen del README · debe coincidir siempre
3. **Diseño de outputs** — columnas y reglas de cada archivo
4. **Código** — implementa lo que dicen los 3 anteriores

Si el código descubre algo que contradice el diseño → para y corrige el README primero.

---

## Historial de versiones

| Versión | Fecha | Qué cambió |
|---------|-------|------------|
| 1.0 | Mayo 2026 | Primera versión — basada en motor_matching y jass_system |
| 1.1 | Mayo 2026 | Documentation-Driven Design — README primero · arquitectura es resumen del README · orden de prioridad actualizado |
| 1.2 | Mayo 2026 | Eliminado contenido duplicado (v1.0 que contradecía v1.1) · Fase 2.6: esquema de inputs · Fase 3.2: validación de inputs · Fase 3.3: logging a archivo · Fase 3.7: Conventional Commits + tags de ciclo · Fase 4: CHANGELOG separado del README · Protocolo de rollback · Gestión de dependencias |
| 1.3 | Mayo 2026 | División profesional de herramientas IA (Chat/Code/Scheduled) · Fase 2.0: flujo de evaluación de enfoques antes de diseñar · Reglas: "Lo manual es el ancla" y "Nunca codifiques la primera idea" |
| 1.4 | Mayo 2026 | Fase 0 (Exploración condicional) — para problemas nuevos o datos desconocidos · Renombrado "4 fases" a "fases de desarrollo" |
