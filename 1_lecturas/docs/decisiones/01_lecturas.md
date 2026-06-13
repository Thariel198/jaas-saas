# Registro de decisiones — 1_lecturas

Sigue la Fase 2.0.6 de la metodología: documentar enfoque elegido + alternativas descartadas con razón.

---

## Decisión 1 — Doble header en registro_operario_acumulado

**Problema:** el archivo de historial guardaba solo MARC_ACT por mes. El M3 que declaró el operario se perdía (se recalculaba como `act - ant`), lo que impedía detectar `SALTO_HISTORICO` y auditar correcciones del operario.

**Criterios:**
- Conservar lo que el operario declaró (auditable)
- Habilitar anomalías históricas (saltos, patrones)
- Que módulos posteriores puedan migrar sin romper masivamente

**Enfoque elegido:** doble header `(YYYY-MM.MARCACION, YYYY-MM.M3)` por ciclo.

**Alternativas descartadas:**
- *Una columna por mes con MARCACION solamente* — descarta el M3 declarado, impide SALTO_HISTORICO. Era el enfoque original.
- *Dos archivos separados (marcaciones.xlsx + consumos.xlsx)* — duplica la fuente de verdad, complica el join. Descartado.

**Señal de alerta:** si en 3 meses ningún módulo descendente está leyendo la sub-columna M3, sobra y se elimina.

---

## Decisión 2 — registro_operario_acumulado vive en inputs/, no en outputs/

**Problema:** el archivo es leído al inicio (historial) y actualizado al final (nueva columna del ciclo). ¿Es input o output?

**Criterios:**
- Coherente con el principio "una sola fuente de verdad"
- Que `crear_template.py` lo lea sin ambigüedad
- Que el supervisor sepa dónde encontrarlo

**Enfoque elegido:** vive en `inputs/`. Es semánticamente un input/output simultáneo — el módulo lo lee y lo actualiza in-place.

**Alternativas descartadas:**
- *En `outputs/`* — la arquitectura original lo ponía ahí, pero entonces `crear_template.py` tendría que ir a otra carpeta a leerlo. Inconsistente.
- *En `shared/`* — error histórico de la versión anterior. Era ambiguo "compartido con quién" (en realidad solo lo usa 1_lecturas).

---

## Decisión 3 — Una sola hoja con columna `categoria` en trazabilidad

**Problema:** mezclar bloqueantes resueltas con informativas registradas en el mismo archivo, o separarlas en hojas.

**Criterios:**
- Auditoría: poder filtrar por usuario y ver TODO lo que le pasó
- Consistencia con `correcciones.xlsx` (una hoja)

**Enfoque elegido:** una hoja `Trazabilidad` con columna `categoria` ∈ {bloqueante, informativa}.

**Alternativas descartadas:**
- *Dos hojas separadas (Bloqueantes / Informativas)* — obliga a saltar de hoja para auditar al mismo usuario. Descartado.
- *Hojas por tipo (RETROCESO, EXCESIVO...)* — 11 hojas. Excesivo. Descartado.

---

## Decisión 4 — Columna `obs_operario` en el template con códigos M/F/P

**Problema:** muchas anomalías son legítimas en el campo (medidor cambiado físicamente, predio cerrado, fuga visible) pero el sistema las marca como bloqueantes y obliga a maquillar.

**Criterios:**
- Que el operario pueda justificar lo que vio en campo de una vez (sin ida-y-vuelta)
- Lenguaje mínimo (un solo carácter)
- No romper compatibilidad con templates que vienen vacíos

**Enfoque elegido:** columna `obs_operario` opcional con códigos:
- `M` = medidor cambiado → legitima RETROCESO como MEDIDOR_CAMBIADO informativo
- `F` = fuga visible → dispara informativo siempre
- `P` = predio cerrado → legitima SIN_LECTURA informativo

La leyenda va impresa en el mismo Excel encima de la tabla.

**Alternativas descartadas:**
- *Columna libre de texto* — invita a errores ortográficos y normalizaciones costosas. Descartado.
- *Hoja separada para anotaciones* — obliga al operario a buscar otra hoja. Descartado.

---

## Decisión 5 — Ciclos detectados por existencia de `lecturas_planilla_YYYY-MM.xlsx`

**Problema:** distinguir Ciclo 1 (primer pass del mes, limpiar lo temporal anterior) de Ciclo 2+ (leer correcciones, regenerar).

**Criterios:**
- Sin archivo de estado adicional
- Idempotente con re-corridas del mismo ciclo

**Enfoque elegido:** mismo patrón que `motor_matching`: si `lecturas_planilla_YYYY-MM.xlsx` NO existe → Ciclo 1. Si SÍ existe → Ciclo 2+.

**Alternativas descartadas:**
- *Archivo `resumen_ciclo.txt` con número de ciclo* — agrega un archivo extra a mantener. Descartado.
- *Columna `ciclo` en el log* — el log puede borrarse, no es señal confiable. Descartado.

---

## Decisión 6 — PDF orden_verificacion solo lleva 3 anomalías de campo

**Problema:** no todas las anomalías requieren ir al predio. Si todas van al PDF, el operario carga 12 páginas pero solo 4 son útiles en campo.

**Criterios:**
- Minimizar peso del documento físico
- Que el operario no se distraiga con casos administrativos

**Enfoque elegido:** PDF solo incluye `SIN_LECTURA`, `MEDIDOR_INVERTIDO`, `POSIBLE_CAMBIO_MEDIDOR`, `EXCESIVO`. Las demás (DIFERENCIA_M3, MARC_ACT_NO_NUMERICO, DUPLICADO, USUARIO_FANTASMA) se resuelven en escritorio y van solo a `correcciones.xlsx`.

**Alternativas descartadas:**
- *Todas las anomalías en el PDF* — sobrecarga al operario con casos no aplicables. Descartado.
- *Un PDF para campo + un PDF para escritorio* — duplica trabajo y confusión. Descartado.

---

## Decisión 7 — Separar RETROCESO en MEDIDOR_INVERTIDO + POSIBLE_CAMBIO_MEDIDOR

**Problema:** una sola anomalía `RETROCESO` cubría dos situaciones físicamente distintas con acciones diferentes en campo:
- Medidor instalado al revés → calibración (diferencia chica, ej: −5 m³)
- Medidor cambiado sin reportar → confirmar con el usuario (diferencia grande, ej: −1588 m³)

El operario que recibía "RETROCESO" no sabía a priori qué buscar; las causas en el PDF eran genéricas.

**Criterios:**
- Que el PDF dirija al operario a la acción correcta sin ambigüedad
- Que la trazabilidad refleje qué pasó realmente en el predio
- Umbral simple y configurable

**Enfoque elegido:** dos anomalías separadas:
- `MEDIDOR_INVERTIDO` cuando `MARC_ANT − MARC_ACT ≤ UMBRAL_INVERSION`
- `POSIBLE_CAMBIO_MEDIDOR` cuando `MARC_ANT − MARC_ACT > UMBRAL_INVERSION`

`UMBRAL_INVERSION = M3_EXCESIVO = 50.0` m³ — simétrico con el techo de consumo razonable. Si la diferencia negativa supera el máximo de consumo posible, claramente no es invertido sino otro medidor.

**Alternativas descartadas:**
- *Una sola anomalía con motivo distinto según magnitud* — el operario sigue viendo "RETROCESO" en la tabla, no diferencia. Descartado.
- *Tres anomalías (RETROCESO genérico + las dos específicas)* — la diferencia siempre se puede clasificar con el umbral. Mantener RETROCESO genérico solo agrega ruido. Descartado.

**Señal de alerta:** si aparecen falsos POSIBLE_CAMBIO_MEDIDOR en predios donde nadie cambió el medidor (lecturas tomadas con error grosero por el operario), subir el umbral o agregar verificación del segundo dígito.

---

## Decisión 8 — Test de integración con datos sintéticos

**Problema:** la prueba con datos reales (`DATA_boletas.xlsx`) solo gatillaba 2 tipos de bloqueante (SIN_LECTURA + EXCESIVO) porque los datos ya estaban limpios. Confirmar que las 11+ reglas funcionan requería cargar datos "sucios" reales (improbable) o sintéticos.

**Criterios:**
- Forzar cada anomalía con un caso diseñado
- No tocar archivos de producción
- Resultados verificables automáticamente

**Enfoque elegido:** `tests/test_anomalias_integracion.py` que:
- Construye fixtures en `tests/_tmp_integracion/` (descartable)
- Monkey-patcha `config` para apuntar paths ahí
- Llama `main.main()` end-to-end
- Verifica conteo por tipo en correcciones, trazabilidad, lecturas_planilla, PDF

**Alternativas descartadas:**
- *Tests unitarios por función* — útiles pero no validan el flujo completo. Quedan en `test_lecturas.py`. Descartado como reemplazo.
- *Pytest con fixtures* — más estándar pero agrega dependencia. Por ahora script standalone. Reconsiderar si se suma pytest al proyecto.

**Señal de alerta:** si el test pasa pero un caso real falla, agregar el caso real al test como regresión.

---

## Historial de cambios

| Fecha | Cambio |
|---|---|
| 2026-06-04 | Versión inicial — registro de las 6 decisiones de diseño del módulo 01 |
| 2026-06-05 | Decisiones 7 y 8 — separación de RETROCESO + test de integración |
