# Metodología de Desarrollo — jass_system
## Cómo construir sistemas de IA sin improvisar

Versión 1.1 — Mayo 2026
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

## La metodología — 4 fases

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
Crear físicamente la estructura de carpetas del sistema:
```
jass_system/
├── 01_lecturas/
├── 02_planilla/
├── 03_pagos/
├── 04_cobranza/
├── 04b_validacion/
├── 05_corte/
├── 05b_override/
├── 06_boletas/
├── 07_nueva_planilla/
└── shared/
```

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
> La arquitectura es un resumen — nunca tiene más detalle ni información diferente.

**2.3 Validar que cuentan la misma historia**
Leer README y arquitectura juntos y verificar:
- ¿Las carpetas que muestra la arquitectura son las del README?
- ¿Los outputs mencionados en arquitectura coinciden con el README?
- ¿El flujo mensual de la arquitectura sigue los mismos pasos que el README?
- ¿Hay algo en la arquitectura que el README no menciona? → eliminar o agregar al README primero

**2.4 Crear rutas de carpetas del módulo**
Crear físicamente las carpetas internas según el README:
```
motor_matching/
├── Inputs/
│   ├── Maestro/
│   ├── Planilla_anterior/
│   ├── Planilla_mes/
│   └── Reporte_mes/
├── Correcciones/
├── Outputs/
├── Trazabilidad/
└── Blancos/
```

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

**2.6 Validar el diseño completo**
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

**3.2 Documentar el por qué en el código**
Cada función importante lleva un comentario del por qué existe,
no solo del qué hace:

```python
# Prioridad máxima al ciclo 2+ porque el usuario ya validó manualmente
# este origen — no tiene sentido volver a procesar el matching automático
if ciclo > 1 and corr_mult:
```

**3.3 Manejo de errores descriptivo**
Cada error dice qué archivo falló, qué columna, qué valor:

```python
raise ValueError(f"MZ-LOTE '{mz}-{lote}' no existe en planilla del mes")
```

**3.4 Test simple**
Por cada función crítica, un test con casos conocidos:

```python
# test_matching.py
assert extraer_mz_lote("Mz A Lt 7") == ("A", "7")
assert extraer_mz_lote("MzE Lt7, MzP Lt11A") == [("E","7"), ("P","11A")]
```

**3.5 Commit a Git**
Al terminar cada función o bloque lógico:

```bash
git add .
git commit -m "feat: extraer_mz_lote con soporte para múltiples lotes"
```

#### Entregables de Fase 3
- [ ] Código funcional
- [ ] Comentarios del por qué en funciones importantes
- [ ] Manejo de errores descriptivo
- [ ] Tests básicos
- [ ] Commits en Git

---

### FASE 4 — Validación y cierre
**Objetivo:** verificar que el sistema hace lo que el diseño dice.

#### Pasos

**4.1 Correr con datos reales**
Usar datos del mes actual y verificar que los outputs coinciden
con lo diseñado en Fase 2.

**4.2 Actualizar skills**
Si encontraste algo nuevo que costó tiempo resolverlo, convertirlo en skill.

**4.3 Retrospectiva del módulo**
Anotar en el README:
- Qué funcionó bien
- Qué cambiaría la próxima vez
- Qué decisiones de diseño resultaron correctas

#### Entregables de Fase 4
- [ ] Módulo validado con datos reales
- [ ] Skills actualizadas
- [ ] README con retrospectiva

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
Todo cambio va a Git con un mensaje que explica qué y por qué.
Un archivo en una carpeta sin Git no existe para el equipo.

**El mapa se actualiza cuando cambia el territorio**
Si descubres que un módulo no va donde pensabas → actualiza el README primero,
luego la arquitectura. Un mapa desactualizado es peor que no tener mapa.

---

## Señales de que vas bien

- Puedes explicar el sistema completo en 2 minutos sin abrir código
- README y arquitectura cuentan exactamente la misma historia
- Cuando algo cambia, sabes exactamente qué documentos actualizar — primero el README
- El código no tiene sorpresas — hace lo que el README dice
- Los errores tienen mensajes claros que dicen exactamente qué falló
- Puedes retomar el proyecto después de semanas sin perder contexto

## Señales de que algo está mal

- README y arquitectura se contradicen
- La arquitectura tiene información que el README no menciona
- Hay código que no sabes para qué sirve
- Un archivo vive en una carpeta que no le corresponde
- Descubres funcionalidad nueva mientras codificas (debería estar en diseño)
- No puedes explicar por qué tomaste una decisión de arquitectura

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


---

## Los 3 activos que construyes en cada proyecto

### Activo 1 — Skills
Carpeta de mejores prácticas reutilizables. Lo que aprendiste con errores
en proyectos anteriores, documentado para no repetirlos.

```
skills/
├── diseño_arquitectura.md    ← esta metodología
├── motor_matching.md         ← cómo construir motores de matching
├── trazabilidad.md           ← cómo diseñar sistemas de trazabilidad
└── override.md               ← cómo diseñar correcciones post-validación
```

**Regla:** cada vez que resuelves algo que te costó tiempo, lo conviertes en skill.

### Activo 2 — Metodología
El proceso que sigues en cada proyecto. Lo que estás leyendo ahora.
Se optimiza con el tiempo — cada proyecto la mejora.

### Activo 3 — Arquitectura del sistema
El mapa del proyecto actual. Vive en el repositorio del proyecto.
No en tu cabeza.

---

## La metodología — 4 fases

---

### FASE 1 — Mapa general
**Objetivo:** conocer el inicio y el fin. No los detalles internos.

#### Pasos

**1.1 Arquitectura general**
Crear `arquitectura_NOMBRE_SISTEMA.html` con:
- Todos los módulos en orden de pipeline
- Input y output de cada módulo
- Estado de cada módulo (hecho / siguiente / futuro / usa IA)
- Recursos shared entre módulos

**1.2 README del sistema**
Crear `README.md` en la raíz con:
- Qué hace el sistema completo
- El orden del pipeline
- Qué vive en shared/
- Cómo instalar y correr

**1.3 Crear rutas de carpetas**
Crear físicamente la estructura de carpetas del sistema:
```
jass_system/
├── 01_lecturas/
├── 02_planilla/
├── 03_pagos/
├── 04_cobranza/
├── 04b_validacion/
├── 05_corte/
├── 05b_override/
├── 06_boletas/
├── 07_nueva_planilla/
└── shared/
```

**1.4 Validar el mapa**
Leer README junto con arquitectura y preguntarse:
- ¿El orden del pipeline tiene sentido?
- ¿Los inputs y outputs de cada módulo conectan correctamente?
- ¿Hay módulos que faltan?
- ¿Hay responsabilidades en el lugar equivocado?

Corregir antes de avanzar.

#### Entregables de Fase 1
- [ ] `arquitectura_SISTEMA.html`
- [ ] `README.md` raíz
- [ ] Carpetas físicas creadas
- [ ] Arquitectura validada y corregida

---

### FASE 2 — Diseño de módulo
**Objetivo:** conocer los personajes y su historia antes de codificar.
Se hace módulo por módulo, empezando por el que más dolor causa.

#### Criterio para elegir qué módulo atacar primero
No el primero del pipeline. El que más valor entrega hoy o más dolor causa.

#### Pasos

**2.1 Arquitectura del módulo**
Crear `arquitectura_MODULO.html` con:
- Estructura de carpetas del módulo
- Inputs con su origen
- Outputs con su destino
- Prioridades o reglas de negocio principales
- Flujo mensual o de uso
- Qué NO hace este módulo

**2.2 README del módulo**
Crear `README_MODULO.md` con:
- Qué hace
- Cuándo se corre
- Estructura de carpetas
- Reglas de negocio detalladas
- Flujo paso a paso con comandos
- Tabla de lifecycle (qué se borra, qué permanece)
- Errores comunes

**2.3 Crear rutas de carpetas del módulo**
Crear físicamente las carpetas internas:
```
motor_matching/
├── Inputs/
│   ├── Maestro/
│   ├── Planilla_anterior/
│   ├── Planilla_mes/
│   └── Reporte_mes/
├── Correcciones/
├── Outputs/
├── Trazabilidad/
└── Blancos/
```

**2.4 Diseño de outputs**
Por cada archivo que produce el módulo, diseñar:
- Columnas y su significado
- Quién escribe cada campo (sistema / usuario / banco)
- Reglas de negocio del archivo
- README propio si el archivo es complejo

Orden sugerido para diseñar outputs:
1. Inputs que el usuario llena (ej: overrides.xlsx, pendientes.xlsx)
2. Archivos de trazabilidad
3. Output principal (ej: pagos_yape.xlsx)

**2.5 Validar el diseño**
Leer README junto con arquitectura del módulo y preguntar:
- ¿Las columnas diseñadas son suficientes para el código?
- ¿Hay reglas de negocio que no están documentadas?
- ¿Los archivos temporales y permanentes están bien identificados?
- ¿El módulo hace algo que no le corresponde?

Corregir antes de avanzar.

#### Entregables de Fase 2 — por módulo
- [ ] `arquitectura_MODULO.html`
- [ ] `README_MODULO.md`
- [ ] Carpetas físicas del módulo creadas
- [ ] Diseño de cada output (columnas + reglas)
- [ ] README por archivo complejo
- [ ] Diseño validado y corregido

---

### FASE 3 — Código
**Objetivo:** traducir el diseño a código limpio en una sola pasada.
Solo se llega aquí cuando Fase 2 está completa y validada.

#### Regla de oro
Si durante el código descubres algo que cambia el diseño:
1. Para el código
2. Actualiza la arquitectura y el README
3. Continúa el código

Nunca dejes que el código contradiga la documentación.

#### Pasos

**3.1 Revisar skills relevantes**
Antes de escribir una línea, leer las skills que aplican al módulo.

**3.2 Documentar el por qué en el código**
Cada función importante lleva un comentario del por qué existe,
no solo del qué hace:

```python
# Prioridad máxima al ciclo 2+ porque el usuario ya validó manualmente
# este origen — no tiene sentido volver a procesar el matching automático
if ciclo > 1 and corr_mult:
```

**3.3 Manejo de errores descriptivo**
Cada error dice qué archivo falló, qué columna, qué valor:

```python
raise ValueError(f"MZ-LOTE '{mz}-{lote}' no existe en planilla del mes")
```

**3.4 Test simple**
Por cada función crítica, un test con casos conocidos:

```python
# test_matching.py
assert extraer_mz_lote("Mz A Lt 7") == ("A", "7")
assert extraer_mz_lote("MzE Lt7, MzP Lt11A") == [("E","7"), ("P","11A")]
```

**3.5 Commit a Git**
Al terminar cada función o bloque lógico:

```bash
git add .
git commit -m "feat: extraer_mz_lote con soporte para múltiples lotes"
```

#### Entregables de Fase 3
- [ ] Código funcional
- [ ] Comentarios del por qué en funciones importantes
- [ ] Manejo de errores descriptivo
- [ ] Tests básicos
- [ ] Commits en Git

---

### FASE 4 — Validación y cierre
**Objetivo:** verificar que el sistema hace lo que el diseño dice.

#### Pasos

**4.1 Correr con datos reales**
Usar datos del mes actual y verificar que los outputs coinciden
con lo diseñado en Fase 2.

**4.2 Actualizar skills**
Si encontraste algo nuevo que costó tiempo resolverlo, convertirlo en skill.

**4.3 Retrospectiva del módulo**
Anotar en el README:
- Qué funcionó bien
- Qué cambiaría la próxima vez
- Qué decisiones de diseño resultaron correctas

#### Entregables de Fase 4
- [ ] Módulo validado con datos reales
- [ ] Skills actualizadas
- [ ] README con retrospectiva

---

## Reglas generales — siempre

**Diseño antes de código**
Nunca escribir código sin tener la arquitectura y el README del módulo completos.
El costo de cambiar un diseño es una línea. El costo de cambiar código es horas.

**Una sola fuente de verdad**
Si un archivo es usado por varios módulos → vive en `shared/`.
Si una regla está documentada en dos lugares → solo debe estar en uno.

**Deuda técnica visible**
Si algo está mal ubicado o es código viejo que no pertenece → va a `backup/`
con un README que explica por qué existe y cuándo se elimina.

**Git es la memoria del proyecto**
Todo cambio va a Git con un mensaje que explica qué y por qué.
Un archivo en una carpeta sin Git no existe para el equipo.

**El mapa se actualiza cuando cambia el territorio**
Si descubres que un módulo no va donde pensabas → actualiza la arquitectura
antes de continuar. Un mapa desactualizado es peor que no tener mapa.

---

## Señales de que vas bien

- Puedes explicar el sistema completo en 2 minutos sin abrir código
- Cuando algo cambia, sabes exactamente qué documentos actualizar
- El código no tiene sorpresas — hace lo que el README dice
- Los errores tienen mensajes claros que dicen exactamente qué falló
- Puedes retomar el proyecto después de semanas sin perder contexto

## Señales de que algo está mal

- Hay código que no sabes para qué sirve
- El README dice algo diferente al código
- Un archivo vive en una carpeta que no le corresponde
- Descubres funcionalidad nueva mientras codificas (debería estar en diseño)
- No puedes explicar por qué tomaste una decisión de arquitectura

---

## Orden de prioridad cuando hay conflicto

1. Arquitectura del sistema — el mapa general manda
2. README del módulo — las reglas de negocio mandan
3. Diseño de outputs — las columnas y datos mandan
4. Código — implementa lo que dicen los 3 anteriores

Si el código descubre algo que contradice el diseño → para y corrige el diseño primero.

---

## Historial de versiones

| Versión | Fecha | Qué cambió |
|---------|-------|------------|
| 1.0 | Mayo 2026 | Primera versión — basada en motor_matching y jass_system |
