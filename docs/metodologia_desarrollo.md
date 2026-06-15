# Metodología de Desarrollo — jass_system
## Cómo construir sistemas de IA sin improvisar

Versión 3.0 — Junio 2026
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

#### División de modelos dentro de Code — Opus codifica, Sonnet acompaña

Dentro de Code hay dos modelos disponibles. El error caro no es elegir cuál usar:
es codificar lógica compleja con el modelo más débil y descubrirlo después en debugging.

**Principio:** las tareas que requieren razonamiento sobre el código — debatir su estructura, planearlo, escribirlo, debuggearlo — van con Opus. Las mecánicas (leer, transcribir, git) van con Sonnet.

| Tarea dentro de Code | Modelo | Por qué |
|---|---|---|
| **Debatir el problema a codificar** (¿qué hojas tiene cada output? ¿una clase o dos? ¿dónde vive esta responsabilidad?) | **Opus** | El debate de diseño tiene el mismo costo de error que el código mismo: si la conversación sale floja, el código bien escrito implementa una mala decisión. |
| **Planear la implementación** antes de escribir código | **Opus** | Mismo razonamiento que el debate. El plan dicta el código. |
| **Escribir o refactorizar código Python no trivial** | **Opus** | Un bug en lógica compleja cuesta más que el token premium. Opus minimiza fallas en cross-checks, ciclos, transformaciones de datos. |
| **Debugging** de errores reales en código | **Opus** | El razonamiento sobre causa-raíz es donde Opus saca más ventaja. |
| **Escribir HTMLs de contrato de formato** (cualquiera: formato_*.html, arquitectura, diagramas) | Sonnet | Cuando se escribe el HTML el diseño YA está cerrado por Opus. Sonnet solo lee la decisión y la transcribe a HTML — no diseña, no decide, no agrega. |
| Leer archivos, hacer Glob/Grep, explorar el repo | Sonnet | Tareas mecánicas — el modelo no es el cuello de botella. |
| **Correr el módulo, leer el run.log, validar outputs** | Sonnet | Ejecutar un script, leer el log, verificar que el Excel tiene las columnas correctas — mecánico. |
| Actualizar README, arquitectura, metodología, memorias | Sonnet | Escritura conversacional, no algorítmica. |
| Git, tests, validaciones de output | Sonnet | Operaciones mecánicas sobre el repo. |

**Cómo aplicarlo en práctica:**
1. Apenas el usuario empieza a **debatir cómo se va a estructurar el código** → cambiar a Opus.
2. Mantener Opus durante el debate, el plan y la implementación.
3. Al terminar el bloque de código → volver a Sonnet para transcribir HTMLs aprobados, commitear, documentar.

**Regla:** si la próxima conversación es "¿cómo deberíamos hacer X en el código?", el modelo se cambia
*antes* de empezar a debatir. No solo antes de escribir.

**Regla de recordatorio activo — la IA avisa antes de empezar:**

La IA no espera que el usuario recuerde cambiar el modelo. Lo dice ella primero, antes de empezar la tarea:

| Situación | Qué dice la IA antes de empezar |
|---|---|
| Va a escribir o actualizar un HTML | *"Voy a escribir el HTML — estás en Sonnet, es el modelo correcto para esto."* |
| Va a empezar a debatir diseño de código | *"Esto es debate de diseño — cambia a Opus antes de empezar."* |
| Va a escribir código Python no trivial | *"Voy a codificar — cambia a Opus si todavía no lo hiciste."* |
| Va a hacer debugging de un error real | *"Para debuggear esto bien — cambia a Opus."* |
| Va a correr el módulo, leer el log o validar outputs | *"Voy a correr y validar — cambia a Sonnet."* |
| Va a hacer git, tests, actualizar docs | *"Esto es mecánico — cambia a Sonnet si no lo hiciste."* |

El recordatorio es una sola línea, antes del trabajo, no después. Si ya está en el modelo correcto, no hace falta decir nada.

**Error frecuente a evitar:** terminar de codificar en Opus y seguir directamente a correr el código sin avisar el cambio. El run, el log y la validación de outputs van con Sonnet — avisar antes de ejecutar el primer comando.

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

**Regla: el orden del pipeline refleja la realidad del negocio**

El orden de los módulos debe coincidir con el flujo real del negocio, no con la conveniencia técnica. Si en la vida real el cliente recibe la boleta antes de pagar, `3_boletas` va antes que `4_pagos`. Un pipeline cuyo orden no coincide con el flujo real genera confusión cuando alguien lo lee o lo audita.

Señal de que el orden está mal: el nombre del módulo M describe una acción que lógicamente ocurre después de la acción del módulo M+1.

**Regla: sub-módulo vs módulo independiente**

Cuando un paso intermedio produce un archivo que solo consume el siguiente paso dentro del mismo módulo, pertenece como sub-carpeta del módulo padre — no como módulo independiente de nivel superior.

- Sub-módulo: el output es específico del módulo y no lo usan otros módulos
- Módulo independiente: el output es consumido por 2+ módulos del sistema

Ejemplo: `3_boletas/enriquecimiento/` (3.1) produce `data_boletas.xlsx` que solo usa `3_boletas/render/` (3.2). Pertenece adentro de `3_boletas`, no como módulo propio.

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

**2.1 Diagrama visual del módulo** ← primero siempre, antes del README

Un diagrama de una sola página que muestra de un vistazo: qué entra, qué hace el módulo,
qué sale y quién consume el output. Si el diagrama está mal, se corrige antes de escribir
el README. Es más barato corregir una imagen que tres documentos.

Formato: HTML en `MODULO/docs/diagrama_MODULO.html` — dentro de la carpeta `docs/`
del propio módulo, no en la carpeta `docs/` global del sistema.

Ejemplo: `4_pagos/efectivo/docs/diagrama_efectivo.html`

Contenido mínimo:
- Bloque de inputs con su origen (¿viene de otro módulo, es manual, es shared?)
- Bloque central con las reglas clave del módulo (3-5 líneas máximo)
- Bloque de output con las columnas que genera
- Flecha al módulo siguiente

**Señal de aprobación:** el usuario confirma que el diagrama refleja bien la idea.
Solo entonces se escribe el README.

---

**2.2 README del módulo** ← después del diagrama aprobado
Formato: **Markdown (`.md`)** — nunca Word ni PDF. MD es el estándar profesional:
vive en Git, se lee en GitHub, cualquier editor lo renderiza, y no requiere instalar nada.

Crear `README.md` en la raíz del módulo con:
- Qué hace
- Cuándo se corre
- Estructura de carpetas
- Reglas de negocio detalladas
- Flujo paso a paso con comandos
- Tabla de lifecycle (qué se borra, qué permanece)
- Lo que NO hace este módulo
- Errores comunes

**2.3 Arquitectura del módulo** ← resumen visual del README
Crear `MODULO/docs/arquitectura_MODULO.html` (dentro de la carpeta `docs/` del módulo) con:
- Estructura de carpetas del módulo — exactamente como dice el README
- Inputs con su origen
- Outputs con su destino
- Prioridades o reglas de negocio principales
- Flujo mensual o de uso
- Qué NO hace este módulo

> Si arquitectura y README se contradicen → el README manda.

**2.4 Validar que cuentan la misma historia**
- ¿Las carpetas que muestra la arquitectura son las del README?
- ¿Los outputs mencionados en arquitectura coinciden con el README?
- ¿Hay algo en la arquitectura que el README no menciona? → eliminar o agregar al README primero

**Regla de actualización atómica**

Cuando un diseño se acuerda en conversación y se plasma en un HTML (diagrama, formato, arquitectura),
actualizar **todos los documentos relacionados en la misma pasada**: README + arquitectura + diagrama + formato.

No esperar a que el usuario abra y apruebe el HTML para tocar los demás archivos.
La aprobación en conversación es suficiente para actualizar todos los artefactos.

La única excepción es el paso 2.1 (diagrama visual inicial): ahí sí esperar aprobación
antes de escribir el README, porque es la primera vez que se plasma el diseño en papel
y puede haber malentendidos de fondo que cambiarían el README entero.

**2.5 Crear rutas de carpetas del módulo**
Crear físicamente las carpetas internas según el README.

**2.6 Diseño de outputs** ← contrato visual + reglas de formato

Por cada archivo Excel que produce el módulo, crear `MODULO/docs/formato_ARCHIVO.html`
(dentro de la carpeta `docs/` del propio módulo) con:
- **Historia visual**: agrupación de columnas en secciones con nombre (¿Quién es? / ¿Qué hizo? / ¿Cuánto debe? / etc.)
- **Paleta de colores**: un color por sección (header bg, data bg, border) — reutilizar la paleta del proyecto
- **Tabla HTML de muestra**: 2–3 filas con datos reales para validar que la historia tiene sentido
- **Reglas de formato para openpyxl**: tipo de dato, formato de número, alineación, negrita, ancho de columna
- **Leyenda**: qué significa cada color
- **Columnas vacías**: identificar cuáles salen vacías y qué módulo las completa después

Este archivo es el contrato que el código debe replicar exactamente con openpyxl.
Sin este contrato, cada módulo queda con el formato que "parece bien" al momento de codificar.

**Columnas completadas por módulos posteriores — escribir como fórmula Excel**

Si una columna es la suma de otras columnas en la misma fila, y algunas de esas columnas serán llenadas por un módulo posterior, escribirla como fórmula Excel — no como valor calculado en Python:

```python
# Ejemplo: TOTAL_A_PAGAR depende de BLANCO y DEVOLUCION que llena 4_pagos
ws[f"Q{row}"] = f"=H{row}+I{row}+J{row}+K{row}+L{row}+M{row}+N{row}+O{row}+P{row}"
```

Si Python calcula el valor al momento de escribir (BLANCO=0, DEVOLUCION=0), ese valor queda congelado. Cuando 4_pagos llena esas columnas, el total no se actualiza. Con fórmula, Excel recalcula automáticamente al abrir o al editar.

Orden de columnas: debe contar una historia de izquierda a derecha.
El lector que abre el Excel debe entender el flujo sin leer ningún README.

Paleta base del proyecto (reutilizar, no inventar colores):

| Sección | Header bg | Header text | Data bg | Border |
|---------|-----------|-------------|---------|--------|
| Identidad | `#EBF5FB` | `#1A5276` | `#F4FAFF` | `#AED6F1` |
| Datos operario / lectura | `#E6F1FB` | `#0C447C` | `#F0F8FF` | `#AED6F1` |
| Calculado por sistema | `#E9F7EF` | `#1E5C3A` | `#F4FBF7` | `#A9DFBF` |
| Total / columna clave | `#1E8449` (bg) | `#FFFFFF` | `#D5F5E3` | — |
| Pendiente / otro módulo llena | `#F3E8FF` | `#5B21B6` | `#FAF5FF` | `#C4B5FD` |

**Registro de auditoría vs vista operacional**

Al diseñar los outputs, clasificar cada archivo en una de dos categorías:

- **Registro de auditoría** — contiene todos los eventos sin filtrar. Solo crece. Nunca se filtra ni se borra. Es la fuente de verdad institucional. Pregunta que responde: *"¿qué ocurrió exactamente?"*
- **Vista operacional** — proyección filtrada del registro de auditoría para un consumidor específico. Se puede regenerar desde la fuente. Pregunta que responde: *"¿qué necesita ver el módulo X?"*

Regla: si un módulo downstream necesita un subconjunto de los datos → crear una vista operacional separada, no pasarle el archivo de auditoría completo ni modificarlo.

Ejemplo: `pagos_yape_pagaste.xlsx` (auditoría: todos los egresos JASS) produce `pagos_yape_devolucion.xlsx` (vista: solo devoluciones con MZ+LOTE → 5_cobranza). Si mañana la lógica de filtrado cambia, el registro de auditoría no cambia — solo se regenera la vista.

**2.7 Definir el esquema de inputs esperado**
Por cada input del módulo, documentar:
- Columnas requeridas y sus nombres exactos (incluyendo alias aceptados)
- Tipos de dato esperados
- Qué pasa si falta un archivo o una columna

Esto se convierte en la validación que corre al inicio del módulo (ver Fase 3).

**Retrocompatibilidad de lectura — cuando el schema va a cambiar**

Si este input va a renombrar una columna en el futuro, documentar el nombre anterior como alias aceptado.
El código lector maneja ambos mientras conviven archivos con el schema viejo y el nuevo.

```python
try:
    i_ciclo = idx("CICLO_CORRECCION")  # schema nuevo
except KeyError:
    i_ciclo = idx("CICLO")             # fallback: schema viejo
```

No es lo mismo que la migración. La migración actualiza el archivo. La retrocompatibilidad cubre el período de transición donde ambos schemas coexisten — en un sistema de uso mensual puede durar ciclos enteros.

Documentar en el esquema de inputs: nombre actual · alias aceptado · desde qué versión aplica.

**2.8 Validar el diseño completo**
- ¿Las columnas diseñadas son suficientes para el código?
- ¿Hay reglas de negocio que no están documentadas?
- ¿Los archivos temporales y permanentes están bien identificados?
- ¿El módulo hace algo que no le corresponde?

Corregir antes de avanzar.

---

**2.9 Diseño de la migración** ← solo cuando hay datos manuales en archivos del schema anterior

Cuando el rediseño cambia las columnas de un archivo que los usuarios ya llenaron a mano,
hay trabajo en riesgo. Antes de codificar el módulo nuevo, planear cómo migrar esos datos.

**Señal de que aplica:**
- Los usuarios ya entregaron archivos con el schema viejo
- Se agregan, renombran o reordenan columnas en esos archivos
- Si el archivo solo tiene datos del sistema (no manual) → no aplica, basta regenerar

Si aplica, seguir el **Protocolo de Migración** documentado al final de este archivo antes de avanzar a Fase 3.

#### Entregables de Fase 2 — por módulo
- [ ] `README_MODULO.md` — fuente de verdad
- [ ] `arquitectura_MODULO.html` — resumen visual del README
- [ ] README y arquitectura cuentan la misma historia
- [ ] Carpetas físicas del módulo creadas
- [ ] Diseño de cada output (columnas + reglas)
- [ ] README por archivo complejo
- [ ] Esquema de inputs documentado
- [ ] Migración de schema planificada y ejecutada (si hay datos manuales en el schema anterior)

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

**3.0 Auditar contratos de formato — ANTES de escribir una sola línea**

Este paso es obligatorio aunque hayas hecho bien la Fase 2.
Cuando retomás un módulo o lo extendés, puede que falten HTMLs de outputs nuevos.

```
1. Glob MODULO/docs/*.html  →  listar todos los HTMLs presentes
2. Listar todos los archivos Excel que produce este módulo (outputs/ + trazabilidad/)
3. Verificar que existe formato_ARCHIVO.html para CADA uno
   - Si falta alguno → crearlo ahora, NO empezar el código
4. Leer cada formato_*.html de arriba a abajo (columnas, anchos, colores, formatos)
   - No asumir que se recuerda el contenido — releerlo siempre
```

**Señal de que podés avanzar al 3.1:**
- [ ] Hay un `formato_ARCHIVO.html` aprobado por cada output Excel del módulo
- [ ] Leíste cada contrato en esta sesión, no en una sesión anterior

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

def main() -> None:
    config.OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(config.OUTPUTS_DIR / "run.log", encoding="utf-8"),
        ],
        force=True,
    )
    log = logging.getLogger(__name__)
    log.info("Iniciando modulo")
```

El archivo `run.log` queda en outputs/ junto al Excel. Si algo falla, el log dice exactamente qué y cuándo.

**Dos reglas obligatorias:**

1. **`force=True`** — sin este flag, si el root logger ya estaba configurado (por ejemplo, por un test runner o por un import previo), `basicConfig` no hace nada y el `FileHandler` nunca se agrega. El log a consola funciona, el log a archivo silenciosamente no.

2. **Dentro de `main()`**, no a nivel de módulo — la razón es que `config.OUTPUTS_DIR` puede haber sido monkey-patcheado por el test antes de llamar a `main()`. Si `basicConfig` corre al importar el módulo, apunta al path original de producción, no al path temporal del test.

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

**Patrón: test de integración con datos sintéticos**

Cuando los datos reales ya están limpios y no gatillan los casos borde, construir
fixtures sintéticos que fuercen cada regla deliberadamente.

```python
# tests/test_MODULO_integracion.py
import shutil, sys
from pathlib import Path
from collections import Counter

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))
import config  # noqa: E402

TEST_ROOT = THIS.parent / "_tmp_integracion"

def _setup_paths():
    """Redirige todos los paths de config a la carpeta temporal."""
    if TEST_ROOT.exists():
        shutil.rmtree(TEST_ROOT)
    (TEST_ROOT / "inputs").mkdir(parents=True)
    (TEST_ROOT / "outputs").mkdir(parents=True)
    # Monkey-patch: el módulo usa config.RUTA, no strings hardcodeados
    config.INPUTS_DIR  = TEST_ROOT / "inputs"
    config.OUTPUTS_DIR = TEST_ROOT / "outputs"
    config.MI_ARCHIVO_PATH = TEST_ROOT / "inputs" / "mi_archivo.xlsx"
    # ... etc por cada path del módulo

def _build_fixture(casos: list) -> None:
    """Construye el archivo de entrada con casos diseñados para forzar cada regla."""
    # Escribir el Excel con openpyxl directamente
    ...

def _verificar() -> tuple[bool, list[str]]:
    """Lee los outputs y compara con lo esperado."""
    msgs, ok = [], True
    filas = leer_output(config.OUTPUTS_DIR / "output.xlsx")
    tipos = Counter(f["tipo"] for f in filas)
    for tipo, esperado in ESPERADO.items():
        obtenido = tipos.get(tipo, 0)
        marca = "✓" if obtenido == esperado else "✗"
        msgs.append(f"  {marca} {tipo:<24} esperado={esperado}  obtenido={obtenido}")
        if obtenido != esperado:
            ok = False
    return ok, msgs

def main():
    _setup_paths()
    _build_fixture(CASOS_SINTETICOS)
    import main as mod_main
    mod_main.main()      # corre el flujo completo
    ok, msgs = _verificar()
    for m in msgs:
        print(m)
    return 0 if ok else 1
```

Reglas del patrón:
- Los fixtures van en `tests/_tmp_integracion/` (en `.gitignore`) — nunca en producción
- El monkey-patch se hace antes de importar `main` para que los paths ya estén redefinidos
- Diseñar UN caso por regla, nombrado explícitamente en el comentario
- La verificación cuenta instancias por tipo — no verifica el contenido exacto de cada celda
- `_tmp_integracion/` se elimina al inicio de cada corrida → idempotente

**3.6b Patrón: bypass por resoluciones en ciclos**

En módulos con ciclos de corrección (Ciclo 1 → correcciones → Ciclo 2+), las filas ya
resueltas en el ciclo anterior no deben re-evaluarse con las reglas de detección.
El patrón es un mapa `resoluciones` que la función de detección consulta antes de aplicar
cualquier regla.

```python
# Al aplicar correcciones del ciclo previo, construir el mapa
resoluciones: dict = {}   # (mz, lt) → resuelto_por

for fila in correcciones_anteriores:
    key = (fila["mz"], fila["lt"])
    if fila["resuelto"]:
        resoluciones[key] = fila["resuelto"]

# En la función de detección, consultar antes de evaluar reglas
def _detectar_anomalias(filas, historial, resoluciones=None):
    resoluciones = resoluciones or {}
    for fila in filas:
        key = (fila["mz"], fila["lt"])
        resuelto = resoluciones.get(key)
        if resuelto in ("acepta_original", "maquillaje", "campo", "corrige_dato"):
            # Ya resuelto — va directo a confirmados con origen='corregido'
            confirmados.append({**fila, "origen": "corregido"})
            continue
        # recién aquí se evalúan las reglas normales
        ...
```

Reglas del patrón:
- Las resoluciones que bypasean son las que el supervisor aplicó manualmente
  (`acepta_original`, `maquillaje`, `campo`, `corrige_dato`)
- Las resoluciones administrativas (`borra_duplicado`, `marca_baja`) se manejan
  antes, en la función que aplica correcciones — eliminan la fila del flujo
- El mapa `resoluciones` se pasa como argumento — la función de detección no
  importa ni lee el archivo de correcciones directamente
- El bypass debe producir el origen correcto en el output (`"corregido"`)
  para que el archivo final sea auditable

**3.6c Regla: verificar comportamientos nuevos con test sintético — no con datos reales**

Cuando se agrega un patrón de comportamiento nuevo a un módulo (preservación de ediciones manuales, bypass por resoluciones, backup automático, etc.), la verificación correcta es un `test_verificacion.py` que:

1. Construye el estado mínimo necesario para activar el comportamiento (archivo de entrada con valores pre-llenados)
2. Corre el módulo
3. Afirma que el invariante se cumple (los valores sobrevivieron, el backup existe, el contador es correcto)
4. Falla con mensaje claro si el invariante se rompe

**Lo que NO se hace:** editar manualmente archivos reales de producción, borrar un output, correr el módulo y verificar visualmente. Ese proceso:
- Toca datos reales de producción (riesgo de pérdida si algo falla)
- No es repetible — la próxima vez hay que volver a hacerlo manualmente
- No detecta regresiones futuras
- Requiere que el desarrollador recuerde exactamente qué verificar

```python
# tests/test_preservacion.py  — verifica que los valores manuales sobreviven al re-run
import shutil, sys
from pathlib import Path
import openpyxl

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))
import config

TEST_ROOT = THIS.parent / "_tmp_preservacion"

def _setup():
    if TEST_ROOT.exists():
        shutil.rmtree(TEST_ROOT)
    (TEST_ROOT / "outputs").mkdir(parents=True)
    config.OUTPUTS_DIR = TEST_ROOT / "outputs"

def _build_discrepancias_con_edicion():
    """Crea discrepancias.xlsx con una fila que ya tiene MZ_CORRECTO/LT_CORRECTO."""
    path = TEST_ROOT / "outputs" / "discrepancias.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "pago_multi_mesa"
    ws.append(["MZ", "LT", "MESA", "OK", "MZ_CORRECTO", "LT_CORRECTO"])
    ws.append(["R", "7", "mesa_2", None, "R", "7"])   # supervisor ya llenó MZ/LT
    wb.save(path)

def _verificar(expected_mz, expected_lt):
    path = TEST_ROOT / "outputs" / "discrepancias.xlsx"
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["pago_multi_mesa"]
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        if d.get("MZ") == "R" and d.get("LT") == "7":
            assert d.get("MZ_CORRECTO") == expected_mz, f"MZ_CORRECTO perdido: {d}"
            assert d.get("LT_CORRECTO") == expected_lt, f"LT_CORRECTO perdido: {d}"
            print("✓ Preservación verificada")
            return
    raise AssertionError("Fila R-7 no encontrada en el output")

def main():
    _setup()
    _build_discrepancias_con_edicion()
    import main as mod_main
    mod_main.main()
    _verificar("R", "7")

if __name__ == "__main__":
    main()
```

Reglas del patrón:
- El test construye el estado mínimo necesario — no depende de los archivos reales del mes en curso
- Un assert fallido dice exactamente qué se esperaba y qué se obtuvo
- El test es idempotente: borra y recrea su carpeta temp en cada corrida
- Vive en `tests/` como cualquier otro test del módulo

**3.6d Patrón: preservación de trabajo manual en re-corridas**

En sistemas donde humanos y automatización comparten el mismo archivo, cada re-corrida del módulo
debe distinguir entre "lo que el sistema generó" (regenerable) y "lo que el humano decidió" (sagrado).
El patrón tiene tres capas que trabajan juntas:

| Capa | Función | Sin esta capa |
|------|---------|---------------|
| 1. Backup | Copia el archivo antes de tocar nada | Si falla la escritura, no hay punto de restauración |
| 2. Leer decisiones humanas | Extrae lo que el humano escribió antes de regenerar | Las decisiones del humano se borran al regenerar |
| 3. Set de ya-procesados | Deduplicación por contenido entre corridas | Los registros confirmados reaparecen como pendientes |

```python
def main():
    if PENDIENTES_FILE.exists():
        _backup(PENDIENTES_FILE)                     # capa 1
        preservados = _leer_decisiones_humanas(...)  # capa 2
    ya_confirmados = _cargar_ya_confirmados()         # capa 3
    # regenerar solo lo que no está en ya_confirmados ni en preservados
```

Este patrón no es solo para migraciones — aplica en **cada ejecución** del módulo mientras el mes está abierto.
El módulo puede re-correrse 10 veces en el mismo ciclo: el trabajo manual siempre sobrevive.

**3.6e Patrón: thin layer of shared primitives**

Cuando varios módulos implementan el mismo patrón conceptual, hay tres caminos:

```
(A) Duplicar todo          (B) Módulo compartido grande    (C) Thin layer ← profesional
    main_limpieza.py           preservar.py (8 parámetros)     shared/utils_preservacion.py
    main_matching.py           callbacks, if modulo == X           solo primitivos puros
    main_pendientes.py         acoplamiento fuerte                 orquestación en cada main.py
```

El camino (B) parece DRY pero termina con 8 parámetros configurables y callbacks porque cada módulo tiene sus propias reglas. El punto medio correcto es (C).

**Test rápido para decidir qué va en `shared/`:**

| Pregunta | Respuesta → decisión |
|---|---|
| ¿La función necesita un `if modulo == "X"` adentro? | Sí → **no compartir**, vive en cada módulo |
| ¿Es genuinamente la misma cosa sin importar el llamador? | Sí → **compartir** en `shared/` |
| ¿Es orquestación (orden de pasos, qué columnas, cuándo aplicar)? | Siempre en `main.py`, nunca en `shared/` |

**Qué vive en `shared/utils_*.py`:** funciones puras sin lógica de negocio — `backup_con_timestamp()`, `normalize()`, `leer_xlsx_a_mapa()`. Estables: si cambian, el cambio aplica igual para todos los módulos.

**Regla del Tres:** no abstraer hasta tener 3+ usos REALES de lógica GENUINAMENTE IDÉNTICA donde el cambio futuro afectaría a todos por igual. Si dos módulos comparten el 80% pero el 20% difiere por razones de negocio, la abstracción completa es prematura.

Este patrón lo usan kernel Linux, React y pandas: primitivos compartidos estables + orquestación local en cada módulo.

**3.6f Patrón: columna REVISADO en archivos de autorización**

En archivos donde el operario escribe "Si" para aprobar un cambio, la columna AUTORIZAR sola es ambigua:

| Estado real | AUTORIZAR | El sistema no puede distinguir |
|---|---|---|
| Fila nueva, el operario no la vio | vacío | ¿No la vio o decidió no autorizar? |
| Vista, decidió conservar lo que había | vacío | Idéntico — imposible diferenciar |

**Solución: agregar columna REVISADO**

| REVISADO | AUTORIZAR | Significado | Acción del sistema |
|---|---|---|---|
| vacío | vacío | Nueva — el operario no la vio | Resaltar en rojo (requiere atención) |
| Si | vacío | Vista, decidió conservar | Sin cambios (fondo neutro) |
| Si | Si | Vista, autoriza el cambio | Aplicar automáticamente en próxima corrida |

Las filas rojas son las que requieren acción del operario. El sistema puede re-correrse sin confundir "nuevo" con "decidido, mantener". La preservación por clave recupera REVISADO y AUTORIZAR de la corrida anterior.

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

### Convenciones de nombres — una sola fuente para todo

Estas reglas aplican en carpetas, archivos, código Python y columnas Excel.
No hay excepciones: si el nombre ya existe en un input externo, se conserva tal como viene.

| Qué | Formato | Ejemplo |
|-----|---------|---------|
| Módulos del sistema | `NN_nombre_modulo` | `1_lecturas`, `2_planilla` |
| Carpetas internas | `lower_snake_case` | `inputs/`, `outputs/`, `docs/` |
| Archivos Python | `lower_snake_case.py` | `main.py`, `config.py` |
| Archivos Excel con periodo | `nombre_YYYY-MM.xlsx` | `planilla_2026-06.xlsx`, `arrastre_deuda_2026-06.xlsx` |
| Archivos Excel sin periodo | `lower_snake_case.xlsx` | `convenios.xlsx`, `multas.xlsx` |
| Variables Python | `lower_snake_case` | `df_lecturas`, `total_mes` |
| Constantes Python (`config.py`) | `UPPER_SNAKE_CASE` | `TARIFA_M3`, `MANT_FIJO`, `INPUTS_DIR` |
| Columnas Excel (todas) | `UPPER_SNAKE_CASE` | `MZ`, `LT`, `MES_ACTUAL`, `TOTAL_A_PAGAR`, `MONTO_YAPE` |

**Columnas Excel: una sola regla**
Todas las columnas Excel — sin importar si las llenó un humano o las calculó el sistema → `UPPER_SNAKE_CASE`.
La distinción campo/sistema se comunica con **color** en el Excel (verde = sistema, morado = 4_pagos), no con capitalización.

**Nombres cortos y descriptivos:** el nombre debe decir qué es, no de dónde viene ni cómo se calcula.
`MES_ACTUAL` (no `TOTAL_MES_ACTUAL_CALCULADO`) · `MES_ANTERIOR` (no `ARRASTRE_DEUDA_MONTO`)

---

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

## Protocolo de Migración

Un "schema change" ocurre cuando rediseñas las columnas de un archivo que los usuarios ya llenaron manualmente.
Es el momento de mayor riesgo del ciclo de desarrollo: hay que cambiar el formato SIN perder el trabajo existente.

### Cuándo aplica

Aplicar cuando el rediseño afecta archivos con datos manuales (inputs que llenan usuarios, no outputs del sistema).
Si el archivo solo tiene datos calculados por el sistema → no aplica, basta regenerar.

### El error más común — cómo ocurre

Un script de migración importa otro módulo que tiene código a nivel de módulo (sin guard).
Al importar, Python ejecuta ese código y sobreescribe los archivos que ibas a migrar.

```python
# PELIGROSO — si crear_templates.py tiene código fuera del guard, esto ejecuta la migración al importar
from crear_templates import GRUPOS, COLUMNAS

# CORRECTO — definir las constantes inline en el script de migración, o importar solo módulos puros
```

**Regla:** toda importación en un script de migración debe ser de módulos que NO tienen efectos al importarse.

### Pasos obligatorios

**M.1 — Backup ANTES que todo**
```
1. Identificar todos los archivos afectados (glob por patrón si son N archivos)
2. Crear backup/migracion_YYYY-MM/
3. Copiar TODOS los archivos afectados al backup
4. Verificar que el backup está completo (contar archivos)
5. Recién ahora continuar — no hay M.2 sin M.1 completo
```

**M.2 — Asegurar que existe un output consolidado reciente**

Antes de migrar, correr el módulo actual para que el output consolidado (ej. `pagos_efectivo.xlsx`) esté al día.
Este consolidado es el plan B si la migración falla y el backup se corrompe.

**M.3 — Guard obligatorio en todo script que toca archivos**
```python
# Al final del script — sin esto, cualquier `import migrar` ejecuta la migración automáticamente
if __name__ == "__main__":
    main()
```

**M.4 — Migración idempotente — detectar y saltar si ya migrado**

El script detecta si un archivo ya tiene el nuevo schema y lo salta.
Correr dos veces produce el mismo resultado que correr una vez.

```python
headers = [c.value for c in ws[2]]  # fila 2 = headers en el patrón del proyecto
if "NUEVA_COLUMNA" in headers:
    print(f"  {archivo.name} ya v2 → skip")
    continue
```

**M.5 — Leer por nombre de columna, no por posición**

Al leer el archivo viejo, acceder por nombre de columna.
El orden puede haber cambiado entre versiones.

```python
# FRÁGIL — rompe si el orden cambió
valor = fila[3]

# ROBUSTO — funciona independientemente del orden
headers = [c.value for c in ws[2]]
idx = {h: i for i, h in enumerate(headers) if h}
valor = fila[idx["COLUMNA"]]
```

**M.6 — Columnas nuevas vacías**

Columnas nuevas que el usuario debe llenar → dejar en `None`.
No inventar valores. No propagar otra columna. El usuario decide.

```python
ws.cell(row=r, column=col_nueva).value = None  # usuario llenará
```

**M.7 — Probar en un archivo antes de correr en N**

```
1. Correr la migración en un solo archivo
2. Abrir en Excel, verificar visualmente:
   - ¿Mismo número de filas?
   - ¿Los datos existentes se preservaron exactamente?
   - ¿Las columnas nuevas están vacías?
3. Confirmar con el usuario si hay datos de verdad en juego
4. Recién entonces correr en todos los archivos
```

### Plan de recuperación — escribirlo antes de ejecutar

Antes de correr la migración, escribir en 3 líneas qué hacer si algo sale mal:
- Si falla a la mitad → restaurar de `backup/migracion_YYYY-MM/`
- Si el backup se corrompe → reconstruir desde el output consolidado del mes
- Si el output consolidado tampoco existe → escalar a recuperación manual

Tener el plan escrito antes de necesitarlo elimina el pánico.

### Señales de que la migración fue exitosa

- Mismo número de filas antes y después
- Los datos existentes tienen exactamente los mismos valores que antes
- Las columnas nuevas están vacías
- El script es idempotente: correrlo de nuevo no cambia nada

### Caso real — módulo efectivo junio 2026

**Schema change:** `crear_templates.py` pasó de 7 a 9 columnas (agregó MONTO_EFECTIVO y MONTO_YAPE).
Los usuarios ya habían llenado 7 mesas con el schema viejo.

**Error cometido:** `migrar_formato_v2.py` importó `crear_templates.py` que tenía código a nivel de módulo.
El import ejecutó la creación de templates, sobreescribió los 7 archivos de mesa con plantillas vacías.

**Recuperación:** Los datos existían en `outputs/pagos_efectivo.xlsx` (output consolidado del mes).
Se escribió `recuperar_mesas.py` que leyó ese consolidado y reconstruyó los 7 archivos.

**Lección clave:** El output consolidado salvó la situación (paso M.2 aplicado sin saberlo).
Si no hubiera existido, los datos se habrían perdido permanentemente.

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
| 1.5 | Junio 2026 | Fase 3.6: patrón de test de integración con datos sintéticos · Fase 3.6b: patrón de bypass por resoluciones en ciclos de corrección |
| 1.6 | Junio 2026 | Convenciones de nombres — sección única para carpetas, archivos, Python y columnas Excel |
| 1.7 | Junio 2026 | Fase 2.6: diseño de outputs ampliado — formato visual Excel obligatorio (historia, paleta, openpyxl) antes de codificar |
| 1.8 | Junio 2026 | Regla de actualización atómica en 2.4 — cuando el diseño se acuerda en conversación, actualizar todos los documentos en la misma pasada sin esperar aprobación del HTML |
| 1.9 | Junio 2026 | Paso 3.3: logging dentro de `main()` con `force=True` — obligatorio para que FileHandler funcione en tests · Paso 2.6: columnas completadas por módulos posteriores deben ser fórmulas Excel, no valores Python · Fase 1.4: regla de orden de pipeline por realidad de negocio · Fase 1.4: regla de sub-módulo vs módulo independiente |
| 2.0 | Junio 2026 | Pasos 2.1, 2.3, 2.6: los artefactos de diseño de módulo (diagrama, arquitectura, formato outputs) viven en `MODULO/docs/` — no en la carpeta `docs/` global del sistema |
| 2.1 | Junio 2026 | Paso 3.0: auditoría obligatoria de contratos HTML antes de codificar — Glob *.html, verificar que existe formato para cada output, leer cada contrato en la sesión actual |
| 2.2 | Junio 2026 | División de modelos dentro de Code — Opus codifica lógica Python no trivial y hace debugging; Sonnet hace diseño, HTMLs, lectura, planning, docs y git |
| 2.3 | Junio 2026 | Corrección a 2.2 — Opus también para **debatir y planear** el código (no solo escribir). El razonamiento sobre el código y el código mismo son la misma actividad cognitiva |
| 2.4 | Junio 2026 | Aclaración a 2.3 — **escribir cualquier HTML** (formato_*, arquitectura, diagramas) va con Sonnet; el diseño YA lo cerró Opus, Sonnet solo transcribe |
| 2.5 | Junio 2026 | Regla de recordatorio activo — la IA avisa antes de empezar si el modelo no es el correcto para la tarea (HTML → Sonnet, código/debate/debug → Opus) |
| 2.6 | Junio 2026 | Explícito en tabla y recordatorio: correr módulo + leer log + validar outputs → Sonnet. Error frecuente documentado: no avisar el cambio al terminar de codificar en Opus antes de correr. |
| 2.7 | Junio 2026 | Paso 3.6c: verificar comportamientos nuevos con test sintético — no con datos reales. Verificación manual (editar archivos reales, borrar output, re-correr) no es repetible, toca producción, y no detecta regresiones. La forma correcta es un `test_verificacion.py` con estado mínimo sintético, assert con mensaje claro, y carpeta temp idempotente. |
| 2.8 | Junio 2026 | Paso 2.9: diseño de migración (obligatorio cuando el schema cambia y hay datos manuales). Protocolo de Migración completo con 7 pasos: backup primero, output consolidado como plan B, guard `if __name__ == "__main__":` obligatorio, idempotencia, lectura por nombre de columna, columnas nuevas vacías, probar en uno antes de N. Caso real documentado (módulo efectivo junio 2026). |
| 2.9 | Junio 2026 | Tres patrones nuevos. Paso 2.6: Registro de auditoría vs vista operacional — clasificar outputs en fuente de verdad (nunca se filtra) vs proyección para downstream (regenerable). Paso 2.7: Retrocompatibilidad de lectura — try/except por nombre de columna para coexistencia de schemas durante transición, diferente a migración. Paso 3.6d: Preservación en tres capas (backup + leer decisiones humanas + set ya-procesados) — aplica en cada re-corrida del módulo, no solo en migraciones. |
| 3.0 | Junio 2026 | Paso 3.6e: Thin layer of shared primitives — el punto medio entre duplicación pura y módulo compartido grande. Regla del Tres + test rápido (`if modulo == X` → no compartir). `shared/utils_*.py` solo para primitivos puros sin lógica de negocio; la orquestación siempre en el `main.py` del módulo. Paso 3.6f: Columna REVISADO en archivos de autorización — distingue "no vista aún" (rojo) de "vista, decidida" (neutro) de "autorizada para aplicar". Elimina ambigüedad cuando AUTORIZAR está vacío. |
