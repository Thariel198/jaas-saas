# Reporte de Aprendizaje — 24 Mayo 2026

---

## Términos técnicos aprendidos

### HTML como Contrato de Diseño
**Qué es:** Usar un archivo HTML como especificación visual que el código debe replicar exactamente.
**De dónde sale:** El problema recurrente de que el código generaba Excel con formato simple mientras el HTML aprobado tenía cabecera doble con grupos de color.
**Regla:** Antes de tocar código, pedir el HTML actual. Nunca generarlo desde cero si ya existe uno aprobado. El HTML es el contrato — el código lo implementa.
**En tu sistema:** `pagos_yape_tepago_diseno.html`, `pendientes_xlsx.html`, `trazabilidad.html` son los contratos de diseño de motor_matching.

---

### Flujo de Validación en Ciclos
**Qué es:** Los datos no pasan directamente a trazabilidad — primero pasan por pendientes donde el usuario valida con OK=SI.
**De dónde sale:** El error donde ambiguos y maestro_inexacto se grababan en trazabilidad en Ciclo 1 sin esperar validación.
**Regla:** Ciclo 1 → datos van a pendientes · usuario valida con OK=SI · Ciclo 2 → datos pasan a trazabilidad.
**En tu sistema:** Sin esta regla, trazabilidad mostraría datos no validados como si fueran verdad confirmada.

---

### Prompt Engineering para Agentes de Código
**Qué es:** Darle instrucciones precisas a Claude Code — qué leer, qué no tocar, cuándo preguntar.
**De dónde sale:** La experiencia de que Claude Code sin instrucciones claras lee todo el proyecto y puede deshacer trabajo ya hecho.
**Regla:** Siempre especificar: (1) qué archivos leer, (2) qué cambiar exactamente, (3) pedir autorización antes de modificar, (4) preguntar si algo no encaja.
**En tu sistema:** El prompt que armaste para motor_matching especifica exactamente los 6 cambios sin que Claude Code tenga que inferir nada.

---

### Git como Red de Seguridad
**Qué es:** Hacer commit antes de cada cambio importante para poder revertir si algo sale mal.
**De dónde sale:** La decisión consciente de commitear antes de los cambios de motor_matching.
**Regla:** Commit antes de cambios · mensaje descriptivo que explica qué y por qué · no acumular cambios sin commitear.
**En tu sistema:** Ya tienes el hábito — lo hiciste dos veces esta semana.

---

### División de Trabajo Humano-IA
**Qué es:** Usar cada herramienta para lo que hace bien — diseño y decisiones aquí, ejecución en Claude Code.
**De dónde sale:** La frustración de hacer cambios en chat que se perdían porque no había archivos reales.
**Regla:** Claude chat → diseño de HTMLs, decisiones de negocio, explicar errores · Claude Code → leer archivos reales, modificar, correr, verificar resultado.
**En tu sistema:** Diseñas el HTML aquí, lo guardas, Claude Code lo implementa.

---

### Pendientes como Hoja de Revisión Manual
**Qué es:** `pendientes.xlsx` tiene 3 hojas distintas según el tipo de caso — Sin_identificar, Ambiguos, Maestro_inexacto — cada una con columnas adaptadas a lo que el usuario necesita decidir.
**De dónde sale:** La evolución desde una sola hoja simple a 3 hojas especializadas con sugerencias del sistema y columna OK.
**Regla:** Sin_identificar = sin maestro ni mensaje · Ambiguos = 2+ candidatos · Maestro_inexacto = 1 candidato pero diff≠0.
**En tu sistema:** Cada hoja tiene MZ_SUG, LOTE_SUG del sistema + MZ, LOTE, CONCEPTO, OK que tú completas.

---

### CONCEPTO como Alternativa a MZ-LOTE
**Qué es:** Cuando un pago no es de un lote específico sino un gasto o devolución, se usa CONCEPTO en vez de MZ+LOTE.
**De dónde sale:** El caso de Paul Tru* con S/150 de "devolución Rosa Coronado" que el sistema intentaba asignar a K-4 por diff alto.
**Regla:** CONCEPTO lleno → se ignora MZ+LOTE · MZ+LOTE llenos → se usa como lote · nunca los dos juntos.
**En tu sistema:** CONCEPTO aparece en pagos_yape_tepago, pendientes, blancos_acumulados y trazabilidad.

---

## Lo que hiciste bien — nivel profesional

- **Detectaste el problema de flujo** — ambiguos en trazabilidad sin validación. Eso requiere entender el sistema completo.
- **Estableciste la regla de pedir HTML antes de codificar** — después de varios errores repetidos, identificaste la causa raíz y pusiste una regla.
- **Dividiste bien el trabajo** — diseño aquí, ejecución en Claude Code. Eso es usar las herramientas correctamente.
- **Hiciste git commit antes de los cambios** — sin que nadie te lo pidiera. Eso es disciplina de ingeniero.
- **Armaste un prompt de 6 cambios** con reglas claras para Claude Code — eso es documentación ejecutable.
- **Reconociste cuándo cambiar de herramienta** — decidiste pasar a Claude Code cuando el chat no era suficiente.

---

## Pendientes

1. Verificar que los 3 cambios de Claude Code funcionan correctamente corriendo main.py
2. Agregar hoja Maestro_inexacto en trazabilidad — prompt ya listo
3. Git commit después de cada cambio verificado
4. Actualizar skill_tracker con niveles nuevos

---

## Resumen

La semana no fue de código nuevo — fue de corrección y metodología. Estableciste cómo trabajar con Claude Code, cómo usar HTML como contrato, cómo dividir diseño de ejecución. Eso vale más que features nuevas porque evita que el trabajo se repita.
