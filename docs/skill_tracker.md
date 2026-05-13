# Skill Tracker — Wilder Trujillo
## Progreso acumulado en ingeniería de IA

Actualizar después de cada sesión de aprendizaje.
El nivel refleja la práctica real — no lo que sabes en teoría.

**Escala:** 1 = recién escuchado · 5 = lo aplico con ayuda · 8 = lo aplico solo · 10 = lo enseño

---

## Skills de Diseño y Arquitectura

| Skill | Descripción | 10/05 | Meta | Notas |
|-------|-------------|-------|------|-------|
| Documentation-Driven Design | README antes que arquitectura y código | 6/10 | 9/10 | Lo descubriste desde la práctica — README es fuente de verdad |
| Diseño iterativo | Mapa claro al inicio · detalles se descubren en el camino | 7/10 | 9/10 | Tu analogía del héroe — inicio y fin claros, personajes en el camino |
| Separación de responsabilidades | Cada módulo hace una sola cosa | 7/10 | 9/10 | Detectaste override mal ubicado antes de codificar |
| Arquitectura de pipelines | Diseñar flujo de módulos con inputs/outputs claros | 7/10 | 9/10 | jass_system pipeline de 9 módulos |
| Lifecycle de archivos | Qué se borra, qué permanece, qué se reemplaza | 8/10 | 9/10 | Temporal vs permanente bien internalizado |
| Storytelling en documentación | README con narrativa clara de lo general a lo específico | 7/10 | 9/10 | Orden narrativo: qué soy → dónde encajo → cómo pienso → cómo me usas |
| Diseño de outputs | Columnas, reglas de negocio, quién escribe cada campo | 6/10 | 9/10 | Diseñaste trazabilidad, overrides, blancos, pendientes |

---

## Skills de Desarrollo

| Skill | Descripción | 10/05 | Meta | Notas |
|-------|-------------|-------|------|-------|
| Python | Programación general | 6/10 | 9/10 | Ya tienes main.py funcionando |
| pandas / openpyxl | Manejo de Excel con Python | 6/10 | 9/10 | Motor de matching ya construido |
| Regex | Extracción de patrones en texto | 6/10 | 8/10 | extraer_mz_lote con múltiples patrones |
| Manejo de errores | Mensajes descriptivos que dicen qué, dónde y por qué falló | 3/10 | 8/10 | Pendiente mejorar en main.py |
| Testing | Pruebas automáticas por función | 1/10 | 7/10 | Aún no implementado — próxima prioridad |
| Documentar el por qué | Comentarios en código que explican decisiones, no solo acciones | 2/10 | 8/10 | Pendiente en main.py |

---

## Skills de Herramientas

| Skill | Descripción | 10/05 | Meta | Notas |
|-------|-------------|-------|------|-------|
| Git | Control de versiones — commit, push, branch | 3/10 | 8/10 | Recién iniciando — git add, commit, push |
| SQLite | Base de datos local — cuando Excel ya no escala | 1/10 | 7/10 | Futuro — para la versión web |
| VS Code / PyCharm | Entorno de desarrollo | 6/10 | 8/10 | Ya lo usas diariamente |

---

## Skills de Mentalidad

| Skill | Descripción | 10/05 | Meta | Notas |
|-------|-------------|-------|------|-------|
| Desarrollo por valor | Atacar primero lo que más dolor causa, no lo primero del pipeline | 8/10 | 9/10 | Decidiste empezar por motor_matching — correcto |
| Gestión de deuda técnica | Documentar lo que está mal sin borrarlo sin entender | 7/10 | 9/10 | backup/ con README explicando por qué existe |
| Criterio técnico | Pedir pros y contras · no aceptar la primera sugerencia | 8/10 | 9/10 | Cuestionaste override, pendientes, trazabilidad |
| Parar a tiempo | Reconocer cuándo la cabeza está cansada y documentar pendientes | 8/10 | 9/10 | Paraste a las 4pm con todo documentado |

---

## Evolución por fecha

| Fecha | Skills nuevas aprendidas | Commit |
|-------|--------------------------|--------|
| 10/05/2026 | Documentation-Driven Design · Diseño iterativo · Separación de responsabilidades · Desarrollo por valor · Deuda técnica · Lifecycle de archivos · Storytelling en docs · Overengineering | `docs: skill tracker inicial + reporte 10 mayo` |

---

## Próximas skills a desarrollar — en orden

1. **Git** — `git add`, `commit`, `push`, `branch` · una hora de práctica cambia todo
2. **Documentar el por qué** — comentarios en main.py explicando decisiones
3. **Testing básico** — `test_matching.py` con 5 casos conocidos
4. **Manejo de errores descriptivo** — mejorar los `except` en main.py
5. **SQLite** — cuando llegues a la web

---

## Cómo actualizar este archivo

Después de cada sesión:
1. Agrega la fecha en la tabla de evolución
2. Actualiza los niveles que cambiaron
3. Agrega skills nuevas si aparecieron
4. Commit: `git commit -m "docs: skill tracker YYYY-MM-DD"`
