# Skill Tracker — Wilder Trujillo
## Progreso acumulado en ingeniería de IA

Actualizar después de cada sesión de aprendizaje.
El nivel refleja la práctica real — no lo que sabes en teoría.

**Escala:** 1 = recién escuchado · 5 = lo aplico con ayuda · 8 = lo aplico solo · 10 = lo enseño

---

## Skills de Diseño y Arquitectura

| Skill | Descripción | 10/05 | 24/05 | 31/05 | 05/06 | 07/06 | 09/06 | 11/06 | 13/06 | 16/06 | Meta | Notas |
|-------|-------------|-------|-------|-------|-------|-------|-------|-------|-------|-------|------|-------|
| Documentation-Driven Design | README antes que arquitectura y código | 6/10 | 7/10 | — | — | — | 8/10 | — | — | — | 9/10 | Fase 2 completa aplicada de cero para módulo efectivo |
| Diseño iterativo | Mapa claro al inicio · detalles se descubren en el camino | 7/10 | 7/10 | — | — | — | — | — | — | — | 9/10 | Estable |
| Separación de responsabilidades | Cada módulo hace una sola cosa | 7/10 | 8/10 | — | — | — | — | — | — | — | 9/10 | Estable |
| Arquitectura de pipelines | Diseñar flujo de módulos con inputs/outputs claros | 7/10 | 8/10 | — | — | 9/10 | — | — | — | — | 9/10 | Estable |
| Lifecycle de archivos | Qué se borra, qué permanece, qué se reemplaza | 8/10 | 8/10 | — | — | — | 9/10 | — | — | — | 9/10 | cleanup_temporales() patrón · sagrado para inputs manuales |
| Storytelling en documentación | README con narrativa clara | 7/10 | 7/10 | — | — | — | — | — | — | — | 9/10 | Estable |
| Diseño de outputs | Columnas, reglas de negocio, quién escribe cada campo | 6/10 | 8/10 | — | — | 8/10 | 9/10 | — | — | — | 9/10 | 5 contratos HTML para un solo módulo en una sesión |
| HTML como contrato de diseño | Usar HTML aprobado como especificación que el código replica | — | 7/10 | — | 8/10 | — | 9/10 | — | — | — | 9/10 | Feedback clave: leer siempre los HTML ANTES de escribir código |
| Evaluación de enfoques | Explorar mínimo 3 soluciones con pros/cons/cuándo-falla antes de elegir | — | — | 2/10 | 5/10 | — | 6/10 | — | — | — | 9/10 | Aplicado para rediseño efectivo (mia/amiga vs mesa vs columna MESA) |
| Criterios de éxito | Definir cómo se ve el éxito antes de ver soluciones | — | — | 2/10 | 5/10 | — | 6/10 | — | — | — | 9/10 | 5 criterios definidos para módulo efectivo antes de diseñar |
| Registro de decisiones | Documentar enfoque elegido y alternativas descartadas con razón | — | — | 1/10 | 6/10 | — | 7/10 | — | — | — | 8/10 | ADR para efectivo: mesa=archivo con hasta 3 hojas |
| Idempotencia | Diseñar operaciones que producen el mismo resultado al re-correr | — | — | — | 5/10 | — | 7/10 | — | — | — | 8/10 | cleanup al FINAL del pipeline (7_cierre), no al inicio del módulo |
| DRY en presentación | Centralizar escritura de Excel para no repetir formato entre módulos | — | — | — | 6/10 | — | — | — | — | — | 8/10 | Estable |
| Diseño por acción del operario | El código responde a lo que el operario hizo en campo (obs M/F/P) | — | — | — | 6/10 | — | — | — | — | — | 8/10 | Estable |
| Cleanup en cierre de pipeline | Módulos exponen cleanup_temporales() · 7_cierre la llama post git-tag | — | — | — | — | — | 7/10 | — | — | — | 9/10 | Patrón profesional: no borrar al inicio del módulo sino al cierre del mes |
| Protección de trabajo manual | Inputs manuales son sagrados — backup antes de cualquier migración | — | — | — | — | — | 8/10 | 9/10 | — | — | 9/10 | Vivido en producción: sin backup → datos perdidos; output consolidado → recuperados |
| Docs de módulo en carpeta propia | Artefactos de diseño viven en MODULO/docs/, no en docs/ global | — | — | — | — | — | 7/10 | — | — | — | 8/10 | Metodología v2.0 — pasos 2.1, 2.3, 2.6 actualizados |
| Migración segura de schema | Cambiar columnas sin perder datos manuales — backup, guard, idempotencia, leer por nombre | — | — | — | — | — | — | 7/10 | — | — | 9/10 | Protocolo M.1–M.7 · caso real efectivo junio 2026 · output consolidado como plan B |
| Registro de auditoría vs vista operacional | Separar la fuente de verdad completa (crece, nunca se filtra) de la proyección filtrada para consumidores downstream | — | — | — | — | — | — | 6/10 | — | — | 9/10 | pagaste = auditoría · devolucion = vista para 5_cobranza · cambiar filtrado no toca la fuente |
| Preservación en tres capas | Backup + leer decisiones humanas antes de regenerar + set ya-procesados — aplica en cada re-corrida, no solo en migraciones | — | — | — | — | — | — | 6/10 | — | — | 9/10 | Patrón para cualquier sistema donde humanos y automatización comparten archivos |
| Thin layer of shared primitives | shared/utils_*.py solo para primitivos puros — la orquestación vive en cada main.py | — | — | — | — | — | — | — | 6/10 | — | 9/10 | Regla del Tres + test rápido: si necesita `if modulo == X` → no compartir |
| Columna REVISADO para autorización | Distingue "no vista aún" (rojo) de "vista, decidida" (neutro) de "autorizada" | — | — | — | — | — | — | — | 5/10 | — | 8/10 | Elimina ambigüedad cuando AUTORIZAR vacío no es suficiente para el operario |
| Reconciliación bidireccional en writers | SET_DEBE vs SET_TIENE + ACCION en audit — re-correr aplica nuevos, revierte sobrantes, skipea correctos | — | — | — | — | — | — | — | — | 6/10 | 9/10 | aplicar_penalidad v2 · 10 cargos erróneos revertidos automáticamente al excluir CORTADOS+EXONERADOS |
| Estado persistente con máquina de exclusión | CORTADO / EXONERADO / REACTIVADO en registro persistente — cada estado con semántica clara y efecto distinto en generar_lista | — | — | — | — | — | — | — | — | 6/10 | 8/10 | EXONERADO para locales comunales y casos especiales — excluye igual que CORTADO |
| Trazabilidad de origen del pago | Enriquecer lista de decisión (lista_corte) con metadata de origen: MESA + COBRADOR desde pagos_efectivo | — | — | — | — | — | — | — | — | 6/10 | 8/10 | Cross por (MZ,LT) · columna ámbar nueva · permite auditar "me dijeron que pagó, en qué mesa?" |
| Verificación de fuente antes de diseñar | Escribir script exploración para confirmar qué contiene realmente un archivo antes de asumir | — | — | — | — | — | — | — | — | 7/10 | 9/10 | MONTO en pagos_efectivo = MONTO_EFECTIVO (no total) — confirmado con 4 casos mixtos reales |

---

## Skills de Desarrollo

| Skill | Descripción | 10/05 | 24/05 | 31/05 | 05/06 | 07/06 | 09/06 | 11/06 | 13/06 | 16/06 | Meta | Notas |
|-------|-------------|-------|-------|-------|-------|-------|-------|-------|-------|-------|------|-------|
| Python | Programación general | 6/10 | 6/10 | — | — | — | — | — | — | — | 9/10 | Estable |
| Retrocompatibilidad de lectura | Código que maneja dos versiones del mismo schema durante la transición — try nuevo nombre, fallback a viejo | — | — | — | — | — | — | 6/10 | — | — | 8/10 | CICLO_CORRECCION → fallback CICLO · diferente a migrar: la migración actualiza el archivo |
| pandas / openpyxl | Manejo de Excel con Python | 6/10 | 7/10 | — | 8/10 | 9/10 | — | — | — | — | 9/10 | Estable |
| Regex | Extracción de patrones en texto | 6/10 | 6/10 | — | — | — | — | — | — | — | 8/10 | Estable |
| Manejo de errores | Mensajes descriptivos | 3/10 | 3/10 | — | — | — | — | — | — | — | 8/10 | Pendiente |
| Testing | Pruebas automáticas por función | 1/10 | 1/10 | — | 6/10 | — | — | — | — | — | 7/10 | Estable |
| Test de integración | Fixtures sintéticos · monkey-patch · verificar end-to-end sin tocar producción | — | — | — | 6/10 | 7/10 | — | — | — | — | 8/10 | Estable |
| Documentar el por qué | Comentarios en código que explican decisiones | 2/10 | 2/10 | — | — | — | — | — | — | — | 8/10 | Pendiente |

---

## Skills de Herramientas

| Skill | Descripción | 10/05 | 24/05 | 31/05 | 05/06 | 07/06 | 09/06 | 11/06 | 13/06 | 16/06 | Meta | Notas |
|-------|-------------|-------|-------|-------|-------|-------|-------|-------|-------|-------|------|-------|
| Git | Control de versiones | 3/10 | 5/10 | — | — | — | — | — | 7/10 | — | 8/10 | Primer push a GitHub — gitignore avanzado para excluir PDFs grandes (160MB+) |
| Claude Code | Agente de código — leer, modificar, correr, verificar | — | 6/10 | — | — | — | — | — | — | — | 9/10 | Estable |
| Prompt Engineering | Instrucciones precisas para agentes IA | — | 6/10 | — | — | — | — | — | — | — | 9/10 | Estable |
| PowerShell bulk rename | Reemplazar referencias en masa con ordered hashtable · UTF-8 NoBOM | — | — | — | — | 6/10 | — | — | — | — | 8/10 | Estable |
| SQLite | Base de datos local | 1/10 | 1/10 | — | — | — | — | — | — | — | 7/10 | Futuro |
| VS Code / PyCharm | Entorno de desarrollo | 6/10 | 6/10 | — | — | — | — | — | — | — | 8/10 | Estable |

---

## Skills de Mentalidad

| Skill | Descripción | 10/05 | 24/05 | 31/05 | 05/06 | 07/06 | 09/06 | 11/06 | 13/06 | Meta | Notas |
|-------|-------------|-------|-------|-------|-------|-------|-------|-------|-------|------|-------|
| Desarrollo por valor | Atacar primero lo que más dolor causa | 8/10 | 8/10 | — | — | — | — | — | — | 9/10 | Estable |
| Gestión de deuda técnica | Documentar lo que está mal sin borrarlo sin entender | 7/10 | 7/10 | — | — | — | — | — | — | 9/10 | Estable |
| Criterio técnico | Pedir pros y contras · no aceptar la primera sugerencia | 8/10 | 8/10 | — | — | 9/10 | — | — | — | 9/10 | Estable |
| Parar a tiempo | Reconocer cuándo la cabeza está cansada | 8/10 | 8/10 | — | — | — | — | — | — | 9/10 | Estable |
| División de trabajo Humano-IA | Saber qué herramienta usar para cada tarea | — | 7/10 | 7/10 | — | — | — | — | — | 9/10 | Estable |

---

## Evolución por fecha

| Fecha | Skills nuevas aprendidas | Commit |
|-------|--------------------------|--------|
| 10/05/2026 | Documentation-Driven Design · Diseño iterativo · Separación de responsabilidades · Desarrollo por valor · Deuda técnica · Lifecycle de archivos · Storytelling en docs · Overengineering | `docs: skill tracker inicial + reporte 10 mayo` |
| 24/05/2026 | HTML como contrato de diseño · Flujo de validación en ciclos · Prompt Engineering para Claude Code · Git como red de seguridad · División trabajo Humano-IA · CONCEPTO como alternativa a MZ-LOTE · Pendientes 3 hojas | `docs: skill tracker 24 mayo` |
| 31/05/2026 | Evaluación de enfoques antes de codificar · Criterios de éxito · Registro de decisiones · División Chat/Code/Scheduled | `docs: skill tracker 31 mayo` |
| 05/06/2026 | Test de integración con datos sintéticos · Idempotencia · DRY en presentación (formato_excel.py) · Diseño por acción del operario (obs M/F/P) | `docs: skill tracker 05 junio` |
| 07/06/2026 | Orden del pipeline por realidad de negocio · sub-módulo vs módulo independiente · 7_cierre como coordinación pura · fórmulas Excel para columnas completadas por módulos posteriores · logging dentro de main() con force=True · monkey-patch de config antes de importar main · PowerShell bulk rename con ordered hashtable | `docs: skill tracker 07 junio` |
| 09/06/2026 | cleanup_temporales() patrón · docs de módulo en carpeta propia MODULO/docs/ (metodología v2.0) · mesa como unidad de cross-check (1 mesa = 1 archivo, hasta 3 hojas) · regla de mayoría 2/3 · protección de trabajo manual (sagrado + backup antes de migrar) · HTML como contrato — leer SIEMPRE antes de codificar | `docs: skill tracker 09 junio` |
| 11/06/2026 | Migración segura de schema (M.1–M.7) · Registro de auditoría vs vista operacional · Preservación en tres capas (backup + leer decisiones humanas + set ya-procesados) · Retrocompatibilidad de lectura (try/except por nombre de columna) · if __name__ == "__main__": como protección contra import con side effects | `docs: metodología v2.9 + skill tracker 11 junio` |
| 13/06/2026 | Thin layer of shared primitives (Regla del Tres + test rápido) · Columna REVISADO para archivos de autorización · Git push a GitHub con gitignore avanzado (excluir PDFs 160MB+) · Análisis de colisión en padrón catastral multi-lote | `docs: metodología v3.0 + skill tracker 13 junio` |

---

## Próximas skills a desarrollar — en orden

1. **Documentar el por qué** — comentarios en código que explican decisiones, no lo que hacen
2. **Manejo de errores descriptivo** — mejorar los `except` en main.py con mensajes de acción
3. **Implementar thin layer en 00_padron** — crear `shared/utils_preservacion.py` + REVISADO en reporte_conflictos
4. **3_boletas enriquecimiento** — primer sub-módulo con estructura 3.1/3.2 real
5. **SQLite** — cuando llegues a la web

---

## Cómo actualizar este archivo

Después de cada sesión:
1. Agrega la fecha en la tabla de evolución
2. Actualiza los niveles que cambiaron
3. Agrega skills nuevas si aparecieron
4. Commit: `git commit -m "docs: skill tracker YYYY-MM-DD"`
