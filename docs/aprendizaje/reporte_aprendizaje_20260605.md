# Reporte de Aprendizaje — 05 Junio 2026

---

## Términos técnicos aprendidos

### Test de Integración con Datos Sintéticos
**Qué es:** Un test que construye fixtures completos desde cero, conecta todos los módulos, corre el flujo end-to-end y verifica los resultados automáticamente — sin tocar datos de producción.
**De dónde sale:** Los tests unitarios de `test_lecturas.py` solo gatillaban 2 tipos de anomalía con datos reales (ya limpios). Para verificar las 14 reglas se necesitaban datos "sucios" diseñados.
**Regla:** Fixtures sintéticos en carpeta temporal · monkey-patch de `config` para apuntar a esa carpeta · llamar `main.main()` · verificar conteos con `Counter`. Nunca mezclar fixtures con datos reales.
**En tu sistema:** `tests/test_anomalias_integracion.py` fuerza las 14 anomalías (8 bloqueantes + 6 informativas) y verifica que todos los outputs sean correctos en Ciclo 1.

---

### Idempotencia
**Qué es:** Una operación es idempotente si correrla dos veces produce el mismo resultado que correrla una sola vez. Sin efectos dobles, sin acumulación de basura.
**De dónde sale:** El diseño del ciclo en `1_lecturas`: si `correcciones_YYYY-MM.xlsx` no existe → Ciclo 1 (estado limpio). Si existe → Ciclo 2+ (lee lo que hay). Re-correr Ciclo 1 siempre produce el mismo resultado.
**Regla:** Diseñar primero la señal de estado (¿qué archivo indica en qué ciclo estamos?) · la operación limpia su propio rastro antes de escribir · no acumular en Ciclo 1.
**En tu sistema:** `trazabilidad.xlsx` evita duplicados por `(MZ, LT, tipo)` antes de agregar. `correcciones.xlsx` se elimina cuando no quedan bloqueantes — eso es la señal de cierre.

---

### DRY en Presentación (formato_excel.py)
**Qué es:** "Don't Repeat Yourself" aplicado al formato de Excel — centralizar la lógica de escritura en un solo módulo para que los exportadores solo declaren qué datos escribir, no cómo formatearlos.
**De dónde sale:** Los 3 exportadores (`_exportar_correcciones`, `_exportar_trazabilidad`, `_exportar_lecturas_planilla`) necesitaban el mismo estilo de grupos de columnas con cabecera de color. Repetir ese código en cada función garantiza inconsistencias.
**Regla:** Si dos exportadores comparten estructura visual → va a `formato_excel.py`. Si es específico de uno → queda inline.
**En tu sistema:** `formato_excel.py` define `GRUPOS_CORRECCIONES`, `GRUPOS_TRAZABILIDAD`, `GRUPOS_LECTURAS_PLANILLA` como configuración declarativa. Los exportadores llaman `fe.escribir_con_grupos(ws, grupos, filas)` y no saben nada del formato.

---

### Diseño por Acción del Operario (obs M/F/P)
**Qué es:** El código no solo valida datos — responde a lo que el operario hizo en campo. Una misma anomalía toma caminos distintos según el código que el operario dejó en `obs_operario`.
**De dónde sale:** El problema de que anomalías legítimas (medidor cambiado, predio cerrado) llegaban como bloqueantes porque el sistema no tenía forma de saber que el operario ya las había visto y confirmado.
**Regla:** Diseñar primero los códigos de acción (mínimo, máximo 3 caracteres) · imprimirlos en el mismo Excel donde el operario trabaja · el código cambia el tipo de anomalía y el flujo completo.
**En tu sistema:** `obs=M` convierte POSIBLE_CAMBIO_MEDIDOR en MEDIDOR_CAMBIADO informativo · `obs=F` agrega FUGA_REPORTADA sin bloquear · `obs=P` convierte SIN_LECTURA bloqueante en SIN_LECTURA informativa.

---

### Separar una Anomalía en Dos (por magnitud)
**Qué es:** Cuando una sola categoría cubre dos situaciones que requieren acciones distintas en campo, separarlas por un umbral medible es más útil que un campo "motivo".
**De dónde sale:** `RETROCESO` podía ser medidor instalado al revés (diff chica, −5 m³) o medidor cambiado sin reportar (diff enorme, −1588 m³). El operario recibía "RETROCESO" y no sabía qué buscar.
**Regla:** Si la acción del operario en campo es diferente → son anomalías distintas. El umbral debe tener sentido de negocio (en este caso: = M3_EXCESIVO = 50 m³, simétrico con el techo de consumo razonable).
**En tu sistema:** `MEDIDOR_INVERTIDO` cuando `ant − act ≤ 50` · `POSIBLE_CAMBIO_MEDIDOR` cuando `ant − act > 50`.

---

## Lo que hiciste bien — nivel profesional

- **Diseñaste 8 decisiones documentadas** con alternativas descartadas y razón — el archivo `docs/decisiones/1_lecturas.md` es auditable meses después.
- **Detectaste que RETROCESO necesitaba dividirse** — sin que nadie te lo dijera. La señal fue que el PDF de campo no le decía al operario qué buscar.
- **Definiste el umbral con lógica de negocio** — UMBRAL_INVERSION = M3_EXCESIVO = 50 m³ no es arbitrario, es simétrico con el techo de consumo.
- **Construiste el test de integración antes de declarar terminado** — sin el test, los 14 casos eran solo promesas. Con el test, son hechos verificables.
- **El test pasó 100% al primer intento** — los fixtures sintéticos estaban bien diseñados porque el diseño del módulo era sólido.
- **Monkey-patcheaste `config` en el test** — en vez de modificar el código de producción para testear, modificaste el entorno del test. Eso es la dirección correcta.

---

## Pendientes

1. Actualizar `metodologia_desarrollo.md` con el patrón de test de integración y el patrón de bypass por resoluciones
2. Hacer commit del módulo `1_lecturas` completo
3. Iniciar `2_planilla` con la metodología: criterios → diseño → HTML → código

---

## Resumen

El módulo `1_lecturas` fue el primero construido con la metodología completa: decisiones documentadas, HTML como contrato, test de integración verificable. El resultado es un módulo que detecta 14 tipos de anomalía, maneja ciclos de corrección y cierra el mes cuando no quedan bloqueantes. La inversión en metodología se vio en el resultado: test 100% al primer intento.
