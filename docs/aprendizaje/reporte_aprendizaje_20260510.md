# Reporte de Aprendizaje — 10 Mayo 2026

---

## Términos técnicos aprendidos hoy

### Documentation-Driven Design
**Qué es:** Escribir el README antes de hacer cualquier diseño visual o código.
**De dónde sale:** Lo descubriste cuando README y arquitectura se contradecían — porque los generaste por separado desde fuentes distintas.
**Regla:** README es la fuente de verdad. La arquitectura es su resumen visual. Si se contradicen → el README manda.
**En tu sistema:** Primero README de motor_matching, luego arquitectura como resumen de ese README.

---

### Diseño Iterativo
**Qué es:** Tener el inicio y el fin claros, y descubrir los detalles en el camino.
**De dónde sale:** Tu propia analogía — "conozco la historia de transformación del héroe de 5 segundos, conozco el inicio y el fin, pero los personajes y sus historias se desarrollan en el camino."
**En tu sistema:** Sabías que motor_matching recibe un reporte y produce pagos_yape.xlsx. Los detalles — BLANCO, devoluciones, trazabilidad — aparecieron diseñando.
**Lo que NO es:** Improvisar. Tienes el mapa, ajustas el mapa cuando la realidad lo pide.

---

### Separación de Responsabilidades
**Qué es:** Cada módulo hace una sola cosa. Si hace dos cosas → hay que dividirlo.
**De dónde sale:** Cuando descubriste que el override estaba en motor_matching pero no le pertenecía. Y cuando viste que PAGASTE y TE PAGÓ son cosas distintas que necesitan tratamiento distinto.
**Regla que aplicaste:** "En programación dividir función es la clave."
**En tu sistema:** motor_matching identifica · 6b_override corrige · cargar_planilla carga · 5b_validacion valida.

---

### Desarrollo por Valor
**Qué es:** Atacar primero el módulo que más dolor causa hoy, no el primero del pipeline.
**De dónde sale:** La decisión de empezar por motor_matching en lugar de 1_lecturas.
**Regla:** No el primero del pipeline. El que más valor entrega hoy.
**En tu sistema:** motor_matching es el que usas cada mes — empezaste por ahí.

---

### Deuda Técnica
**Qué es:** Código o estructura que funciona pero está mal ubicada o mal diseñada — y que eventualmente habrá que corregir.
**De dónde sale:** cargar_planilla vivía en 4_pagos/yape cuando le correspondía estar en 5_cobranza. usuarios_id.xlsx vivía dentro de un módulo cuando debía estar en shared/.
**Cómo la manejaste:** Carpeta backup/ con README que explica por qué existe y cuándo se elimina. No la borraste sin entender, no la dejaste sin documentar.

---

### Una Sola Fuente de Verdad
**Qué es:** Un archivo o lugar donde vive la información real. Todo lo demás la referencia, no la duplica.
**De dónde sale:** La decisión de mover usuarios_id.xlsx y maestro_yape.xlsx a shared/ para que todos los módulos los lean desde ahí.
**Regla:** Si un archivo lo usan varios módulos → vive en shared/. Si una regla está en dos lugares → solo debe estar en uno.

---

### Lifecycle de Archivos
**Qué es:** Saber exactamente qué se borra, qué permanece y qué se reemplaza en cada ciclo.
**De dónde sale:** La distinción entre Trazabilidad/ (permanente), Blancos/ (permanente), Outputs/ (temporal) y Correcciones/ (temporal).
**Regla que aprendiste:** Los archivos permanentes necesitan su propia carpeta. Los temporales se borran en Ciclo 1. El historial real siempre vive en el permanente.

---

### Storytelling en Documentación
**Qué es:** El README cuenta una historia con narrativa clara — de lo general a lo específico.
**De dónde sale:** La decisión de poner Requisitos al final porque interrumpía la narrativa. Y el orden: qué soy → dónde encajo → dónde vivo → cómo pienso → cómo me usas → qué guardo → si algo falla → mis límites.
**Regla:** "El history telling es fundamental." Un README que se lee como manual técnico se abandona. Uno que fluye como historia se entiende y se recuerda.

---

### Overengineering
**Qué es:** Diseñar más de lo que necesitas ahora — pretender que tienes todo claro desde el inicio.
**De dónde sale:** La conversación sobre si es profesional saber todo desde el inicio.
**Regla:** Lo profesional no es saber todo — es tener el mapa claro y ajustarlo cuando la realidad aparece. Los proyectos que se diseñan completos antes de empezar casi siempre fracasan.

---

## Lo que hiciste bien hoy — nivel profesional

- **Detectaste errores en tu propia documentación** — README vs arquitectura con 3 contradicciones. Eso es criterio de ingeniero.
- **Propusiste Documentation-Driven Design** antes de que yo lo nombrara — llegaste solo a la conclusión correcta.
- **Separaste el override** del módulo equivocado antes de codificar — si lo hubieras descubierto en el código, el costo habría sido mucho mayor.
- **Moviste archivos físicamente** en lugar de solo documentarlo — shared/ ahora existe y tiene los archivos correctos.
- **Conectaste devoluciones con blancos** — misma naturaleza, mismo patrón, misma solución. Eso es pensamiento sistémico.
- **Paraste a tiempo** cuando la cabeza se cansó — un ingeniero que sigue cuando está agotado introduce errores. Uno que para y documenta el pendiente es profesional.

---

## Pendientes para mañana

1. Actualizar README y arquitectura con:
   - `Devoluciones/` carpeta nueva — permanente
   - `pagos_yape_pagaste.xlsx` en Outputs/ — temporal
   - `devoluciones_acumulados.xlsx` — permanente
2. Diseño de `pagos_yape.xlsx` — el output principal
3. Diseño de `pendientes.xlsx` — guardar como HTML

---

## Resumen del día

Horas trabajadas: sesión larga y productiva.

Lo más importante que pasó hoy no fue el código — fue que estableciste una metodología. Documentation-Driven Design, diseño iterativo, separación de responsabilidades. Eso es lo que distingue a un ingeniero de alguien que programa. El código viene después — y viene limpio porque el diseño fue primero.
