# Reporte de Aprendizaje — 11 Junio 2026

---

## Lo que se construyó esta sesión

### Parte 1 — Feature pagos_yape_devolucion (motor_matching)

Se completó la Fase 3 de la feature `pagos_yape_devolucion`. Motor matching ahora produce
tres outputs en lugar de dos: `pagos_yape_tepago.xlsx` (pagos entrantes), `pagos_yape_pagaste.xlsx`
(todos los egresos JASS) y el nuevo `pagos_yape_devolucion.xlsx` (subset de pagaste donde MZ+LOTE
están llenos → devoluciones a usuarios → consumido por 5_cobranza).

Cambios aplicados en la sesión anterior (resumidos aquí por continuidad):
- `exportar_motor.py`: nueva función `exportar_pagos_devolucion()` con paleta morada
- `main.py`: nueva constante `DEVOLUCION_FILE`, lectura de MZ+LOTE desde pendientes.xlsx, nueva función `exportar_devolucion_xlsx()`, cleanup_temporales() actualizado
- `tests/test_fixes.py`: fixtures actualizados con nuevo schema de 13 columnas, nuevo test `test_mz_lote_se_leen_cuando_estan`
- HTMLs: `pagos_yape_devolucion_diseno.html`, `pagos_yape_pagaste_diseno.html`, `pendientes_xlsx.html` (tab Pagaste), `arquitectura_motor_matching.html` — todos actualizados atómicamente

### Parte 2 — Protocolo de Migración de schema

Se formalizó el protocolo de migración a partir del incidente de junio 2026 (módulo efectivo).

**Qué se documentó:**
- Paso 2.9 en Fase 2 de la metodología — señal de que aplica, referencia al protocolo
- Protocolo de Migración completo (7 pasos: M.1–M.7)
- Entregable nuevo en Fase 2: "Migración de schema completada (si aplica)"
- Versión metodología: 2.7 → 2.8
- Nueva skill en skill_tracker: "Migración segura de schema" → 7/10

---

## Términos técnicos aprendidos

### Protocolo de Migración — 7 pasos obligatorios

**Qué es:** Proceso para cambiar el schema de un archivo que los usuarios ya llenaron manualmente,
sin perder su trabajo.

**De dónde sale:** El incidente de junio 2026. `migrar_formato_v2.py` importó `crear_templates.py`
que tenía código a nivel de módulo (sin guard). Al importar, Python ejecutó ese código y sobreescribió
los 7 archivos de mesa con plantillas vacías. El trabajo manual de los cobradores desapareció.

**Los 7 pasos:**
1. **M.1 Backup ANTES que todo** — copiar todos los archivos afectados a `backup/migracion_YYYY-MM/`
2. **M.2 Output consolidado reciente** — correr el módulo actual para tener un plan B si el backup falla
3. **M.3 Guard obligatorio** — `if __name__ == "__main__": main()` en todo script que toca archivos
4. **M.4 Idempotencia** — detectar si el archivo ya tiene el nuevo schema y saltarlo
5. **M.5 Leer por nombre de columna** — no por posición (el orden puede haber cambiado)
6. **M.6 Columnas nuevas vacías** — `None`, el usuario decide; no inventar valores
7. **M.7 Probar en uno antes de N** — abrir en Excel, verificar visualmente, confirmar, recién ejecutar en todos

**Regla central:** No hay M.2 sin M.1 completo. El backup no es opcional.

---

### import-time side effects — el error silencioso

**Qué es:** Código que se ejecuta al nivel de módulo en Python (fuera de funciones y del guard)
se ejecuta automáticamente cuando otro archivo hace `import ese_modulo`.

**Por qué es peligroso en scripts de migración:**
```python
# crear_templates.py  — ANTES (código a nivel de módulo)
wb = Workbook()
# ... código que crea archivos ...
wb.save("mesa_1.xlsx")   # ← esto corre al hacer import

# migrar_formato_v2.py
from crear_templates import GRUPOS, COLUMNAS   # ← dispara la creación de archivos
```

Al importar para obtener solo `GRUPOS` y `COLUMNAS`, se ejecutó todo el código que creaba las mesas.
Los 7 archivos con datos reales de los cobradores fueron sobreescritos con plantillas vacías.

**La corrección:**
```python
# crear_templates.py — DESPUÉS (con guard)
if __name__ == "__main__":
    wb = Workbook()
    # ... código que crea archivos ...
    wb.save("mesa_1.xlsx")   # solo corre cuando se ejecuta directamente
```

**Regla:** Todo script que crea, modifica o borra archivos debe tener el guard. Sin excepción.

---

### Output consolidado como plan B

**Qué es:** El archivo de output del módulo (ej. `pagos_efectivo.xlsx`) que consolida todos los
inputs manuales en un solo lugar. Si los inputs se corrompen, el consolidado permite reconstruirlos.

**Por qué importa en migraciones:** Antes de migrar, correr el módulo garantiza que el consolidado
está al día. Si el backup y la migración fallan, el consolidado es el último recurso.

**En el incidente real:** `pagos_efectivo.xlsx` tenía 340 filas con todos los registros de las 7 mesas.
Se escribió `recuperar_mesas.py` que leyó ese archivo y reconstruyó los 7 archivos de mesa.

**Lección:** El output consolidado salvó la situación. Si no hubiera existido (si el módulo nunca
se hubiera corrido antes de migrar), los datos habrían sido irrecuperables.

---

## Errores cometidos — y lo que revelan

### No tener el guard — incidente de junio 2026

`crear_templates.py` y `migrar_formato_v2.py` fueron escritos sin el guard `if __name__ == "__main__":`.
La migración importó el primero para reutilizar sus constantes y disparó la creación de plantillas.

**Lo que revela:** Cualquier código que tiene efectos secundarios (crear archivos, conectar a BD, enviar emails)
debe estar protegido con el guard. No es un detalle de estilo — es una garantía de que el código
solo corre cuando el desarrollador lo decide explícitamente, no como efecto de un import.

**Corrección:** Regla M.3 del Protocolo de Migración. Se agrega como pregunta de auditoría en Fase 3:
"¿Todos los scripts que tocan archivos tienen el guard?"

---

## Lo que hiciste bien — nivel profesional

- **Identificaste el patrón de protección.** Después del incidente, no solo pediste corregir el bug
  sino convertirlo en metodología: "quiero que en metodología pongas los pasos adecuados para hacer
  una migración de forma profesional". El error se convirtió en activo.

- **Conectaste la teoría con la experiencia.** El Protocolo de Migración no es abstracto — cada paso
  tiene una razón que viene de algo que ocurrió en este proyecto. Ese es el tipo de documentación
  que se lee y se aplica.

- **Pediste actualizar los tres activos al mismo tiempo.** Metodología + aprendizaje + skill_tracker
  en una sola pasada — no como tres pedidos separados. Eso es la regla de actualización atómica
  aplicada a los activos de IA.

---

## Tres conceptos que aparecieron sin ser nombrados

### Registro de auditoría vs vista operacional

**Qué es:** Los outputs se clasifican en dos tipos con propósitos distintos.

- **Registro de auditoría** — contiene todos los eventos sin filtrar. Solo crece. Nunca se modifica. Responde: *"¿qué ocurrió exactamente?"*
- **Vista operacional** — proyección filtrada para un consumidor específico. Se puede regenerar. Responde: *"¿qué necesita ver el módulo X?"*

**Dónde apareció:** `pagos_yape_pagaste.xlsx` es el registro de auditoría de todos los egresos JASS. `pagos_yape_devolucion.xlsx` es la vista para 5_cobranza — solo las filas con MZ+LOTE. Si mañana cambia la lógica de filtrado (por ejemplo, excluir devoluciones de ciclos anteriores), el archivo de auditoría no se toca — solo se regenera la vista.

**Por qué importa:** Cuando el registro de auditoría y la vista operacional son el mismo archivo, cualquier cambio en la lógica de negocio corre el riesgo de corromper el historial. Separarlos desde el diseño protege la fuente de verdad institucional.

---

### Preservación en tres capas

**Qué es:** Patrón para sistemas donde humanos y automatización comparten el mismo archivo. No es solo para migraciones — aplica en cada re-corrida del módulo durante el mes.

| Capa | Función | Sin esta capa |
|------|---------|---------------|
| 1. Backup | Copia antes de tocar | Sin punto de restauración si algo falla |
| 2. Leer decisiones humanas | Extrae lo que el humano decidió antes de regenerar | Las decisiones del humano desaparecen |
| 3. Set de ya-procesados | Deduplicación entre corridas | Los registros confirmados reaparecen como pendientes |

**Dónde apareció:** `main.py` de motor_matching: `_backup_pendientes()` (capa 1) + `_leer_pendientes_preservados()` (capa 2) + `ya_confirmados` (capa 3). El módulo puede re-correrse 10 veces en el mismo ciclo y el trabajo manual siempre sobrevive.

**Por qué importa:** La protección de trabajo manual no es un backup antes de migrar — es una garantía que el sistema ofrece en cada ejecución normal. Un sistema sin este patrón obliga al usuario a no re-correr para no perder su trabajo, lo cual es inaceptable.

---

### Retrocompatibilidad de lectura

**Qué es:** Cuando un schema cambia (columna renombrada), el código lector maneja ambas versiones durante la transición. Es diferente a migrar: la migración actualiza el archivo; la retrocompatibilidad permite leer archivos que todavía no fueron migrados.

```python
try:
    i_ciclo = idx("CICLO_CORRECCION")  # schema nuevo
except KeyError:
    i_ciclo = idx("CICLO")             # fallback: schema viejo
```

**Dónde apareció:** `_cargar_pagaste_existentes()` en motor_matching. El archivo `pagos_yape_pagaste.xlsx` del mes anterior tenía columna "CICLO"; el nuevo schema la llama "CICLO_CORRECCION". El fallback permite leer ambas versiones sin forzar al usuario a migrar el archivo manualmente.

**Por qué importa:** En sistemas de uso mensual, un schema change en junio significa que julio arranca con archivos de junio en el schema viejo. Si el lector no maneja ambas versiones, el primer run del mes siguiente falla con un KeyError en producción.

---

## Pendientes

1. Correr `main.py` de motor_matching con los datos del mes para generar pagos_yape_devolucion.xlsx
2. Verificar que 89 tests siguen en verde después de los cambios
3. Segunda pregunta pendiente del usuario (mencionó "son 2" — nunca se reveló la segunda)
4. Actualizar 5_cobranza para consumir `pagos_yape_devolucion.xlsx` además de `pagos_yape_tepago.xlsx`

---

## Resumen

Sesión en dos partes. La primera completó la feature `pagos_yape_devolucion` en motor_matching:
tres outputs bien definidos con contratos HTML aprobados, MZ+LOTE viajan desde pendientes hasta
el archivo de devoluciones, backwards-compatible con schema viejo. La segunda convirtió el incidente
de migración de junio en metodología formal: Protocolo de Migración con 7 pasos, Fase 2.9 añadida,
nueva skill en el tracker. El aprendizaje central: un script de migración que importa módulos con
side effects es un riesgo silencioso — el guard `if __name__ == "__main__":` no es opcional.
