# Reporte de Aprendizaje — 07 Junio 2026

---

## Lo que se construyó esta sesión

### Parte 1 — Módulo 2_planilla completo

Se construyeron los tres archivos del módulo siguiendo la metodología:

- **`config.py`** — paths, constantes, paleta de colores por sección, lista ordenada de 21 columnas. Centraliza todo lo que el módulo necesita saber sobre su entorno.
- **`main.py`** — `build_planilla()` calcula MES_ACTUAL, normaliza MZ/LT, hace left join con arrastres opcionales. `write_excel()` escribe 2 filas de headers, aplica colores por sección, escribe TOTAL_A_PAGAR como fórmula Excel.
- **`test.py`** — test de integración con 4 fixtures sintéticos, usa `tests/_tmp_integracion/`, monkey-patchea config, verifica resultados automáticamente.

El test pasó sin problemas después de corregir dos violaciones a la metodología detectadas en el camino (ver abajo).

### Parte 2 — Rediseño completo de la arquitectura del sistema

Se rediseñó el orden de los módulos, se renombraron las 9 carpetas y se actualizaron las referencias en 59 archivos a través de todo el proyecto.

**Nuevo orden:** `1_lecturas → 2_planilla → 3_boletas → 4_pagos → 5_cobranza → 5b_validacion → 6_corte → 6b_override → 7_cierre`

Cambios clave:
- `3_boletas` antes que `4_pagos` (el orden correcto por flujo de negocio)
- `enriquecimiento` como sub-módulo `3_boletas/enriquecimiento/` (3.1), no módulo independiente
- `7_cierre` nuevo — verifica que los arrastres existen y los transfiere al siguiente mes; no calcula nada

---

## Términos técnicos aprendidos

### Fórmula Excel para columnas completadas por módulos posteriores

**Qué es:** En vez de calcular el valor en Python (que quedaría congelado), escribir la celda como fórmula Excel. Cuando el módulo posterior llene las columnas de origen, Excel recalcula automáticamente.

**De dónde sale:** `TOTAL_A_PAGAR` es la suma de 9 columnas. Las columnas `BLANCO` y `DEVOLUCION` las completa `4_pagos` más adelante. Si Python calcula el total al escribir (cuando BLANCO=0 y DEVOLUCION=0), ese valor queda congelado para siempre.

**Regla:** Identificar en el diseño de outputs qué columnas serán completadas por módulos posteriores. Esas columnas → fórmula. Las demás → valor Python.

```python
ws[f"Q{row}"] = f"=H{row}+I{row}+J{row}+K{row}+L{row}+M{row}+N{row}+O{row}+P{row}"
```

**En tu sistema:** `2_planilla/main.py` — columna `TOTAL_A_PAGAR` en `write_excel()`.

---

### Logging dentro de main() con force=True

**Qué es:** La llamada a `logging.basicConfig()` debe ir dentro de `main()`, no a nivel de módulo, y debe incluir `force=True`.

**De dónde sale:** Dos bugs distintos que se descubrieron implementando 2_planilla:

1. Sin `force=True`: si el root logger ya estaba configurado (por el test runner, por otro import), `basicConfig` no hace nada. El log a consola funciona, pero el `FileHandler` para el archivo `run.log` nunca se agrega. El bug no es visible — el código corre, pero no hay log en disco.

2. A nivel de módulo: el `FileHandler` apunta a `config.OUTPUTS_DIR / "run.log"` al momento de importar. Si el test monkey-patcha `config.OUTPUTS_DIR` después del import, el handler ya apunta al path original de producción, no al path temporal del test.

**Regla:** Siempre dentro de `main()`, siempre `force=True`, siempre crear `OUTPUTS_DIR` antes de configurar el logging.

**En tu sistema:** `2_planilla/main.py` — `def main()` primero crea el directorio, luego configura logging.

---

### El orden del pipeline refleja la realidad del negocio

**Qué es:** El orden de los módulos en el sistema debe coincidir con el orden en que ocurren las cosas en el mundo real, no con la conveniencia técnica.

**De dónde sale:** En el diseño original los módulos estaban en orden `01_lecturas → 02_planilla → 03_pagos → 04_boletas...`. El problema: en la vida real el usuario primero recibe la boleta (aviso de cobro) y después paga. El sistema tenía el orden invertido — `pagos` antes que `boletas`.

**Regla:** Antes de numerar los módulos, describir el flujo real: "el usuario hace X, luego Y, luego Z". El orden de los módulos debe contar esa misma historia.

**Señal de que está mal:** El módulo M describe una acción que en la vida real ocurre después de la acción del módulo M+1.

**En tu sistema:** `3_boletas` (aviso de cobro) → `4_pagos` (el usuario paga) — ahora el orden es correcto.

---

### Sub-módulo vs módulo independiente

**Qué es:** Cuando un paso intermedio produce un archivo que solo usa el siguiente paso dentro del mismo módulo, pertenece como sub-carpeta — no como módulo de nivel superior.

**De dónde sale:** El debate sobre dónde poner `enriquecimiento` (preparar datos para la boleta). Opciones evaluadas:
- Módulo independiente `4_enriquecimiento` — incorrecto porque su único propósito es preparar datos para `render`
- Sufijo `b` (`3b_enriquecimiento`) — incorrecto, los sufijos `b` en este sistema son para módulos de corrección/validación
- Sub-módulo `3_boletas/enriquecimiento/` (3.1) — correcto

**Criterio de decisión:**
- Sub-módulo: el output es específico de este módulo y no lo consume nadie más
- Módulo independiente: el output es consumido por 2+ módulos del sistema

**Argumento de trazabilidad:** `data_boletas.xlsx` (output de enriquecimiento) es la prueba auditable de qué datos se imprimieron en cada boleta. Si vive dentro de `3_boletas`, la relación es explícita.

**En tu sistema:** `3_boletas/enriquecimiento/` (paso 3.1) → produce `data_boletas.xlsx` → `3_boletas/render/` (paso 3.2) → produce los PDFs.

---

### Bulk rename con ordered hashtable en PowerShell

**Qué es:** Para reemplazar strings en muchos archivos, usar un hashtable ordenado y poner los strings más específicos/largos primero.

**De dónde sale:** Al renombrar `06_boletas → 3_boletas`, existía el riesgo de que `06a_enriquecimiento` se reemplazara a `3a_enriquecimiento` si la regla general `06 → 3` corría primero. Con un hashtable ordenado, `06a_enriquecimiento → 3_boletas/enriquecimiento` corre antes que `06_boletas → 3_boletas`.

**Herramienta:** `[ordered]@{}` en PowerShell + `[System.IO.File]::WriteAllText` con `New-Object System.Text.UTF8Encoding($false)` para evitar BOM en archivos Python.

---

## Errores cometidos — y lo que revelan

### test.py versión 1 escribía a inputs/ reales

El primer test.py construía los fixtures en `inputs/lecturas/`, `inputs/deuda_anterior/`, etc. — las carpetas de producción reales. Si se corría por error con datos reales, los sobreescribía.

**Violación de metodología:** paso 3.6 dice que los fixtures van en `tests/_tmp_integracion/` — nunca en producción.

**Corrección:** reescribir completamente usando `tests/_tmp_integracion/` con `shutil.rmtree` al inicio y monkey-patch de config.

**Lo que revela:** La metodología ya tiene la regla correcta. El error fue no releerla antes de escribir el test. Paso 3.1 dice: "leer las skills relevantes antes de escribir una línea".

### logging.basicConfig a nivel de módulo

La primera versión de `main.py` tenía `logging.basicConfig` fuera de `main()`. Funcionaba en producción pero fallaría en tests con monkey-patch porque el FileHandler apuntaría al path equivocado.

**Lo que revela:** Los errores silenciosos son los más peligrosos. Este habría pasado desapercibido hasta el primer test que fallara misteriosamente sin log en disco.

---

## Lo que hiciste bien — nivel profesional

- **Detectaste el orden incorrecto del pipeline sin que nadie te lo dijera.** La señal fue interna: "en la vida real primero recibes la boleta, después pagas". Ese criterio de negocio venció cualquier argumento técnico.
- **Propusiste 7_cierre espontáneamente.** Viste que los módulos 5b y 6 ya calculaban los arrastres, y que faltaba un módulo que coordinara el cierre del ciclo. La distinción clave: 7_cierre no calcula nada, solo verifica y coordina.
- **Ejecutaste el rename de 59 archivos con criterio ordenado.** Aplicaste el principio de strings largos primero para evitar matches parciales — sin que nadie lo dijera explícitamente.
- **Aceptaste la corrección del test sin resistencia.** Cuando se detectó la violación de metodología, pediste reescribirlo directamente, sin defender la primera versión.

---

## Pendientes

1. Hacer commit del módulo 2_planilla completo (config.py + main.py + test.py)
2. Diseñar `3_boletas/enriquecimiento/` — criterios → README → HTML → código
3. Re-integrar `4_pagos/yape/motor_matching` con la nueva estructura de 2_planilla
4. Implementar `7_cierre` cuando todos los módulos upstream estén funcionando

---

## Resumen

Esta sesión tuvo dos mitades. La primera construyó el módulo 2_planilla completo con dos errores de metodología detectados y corregidos en el camino — lo que demuestra que la metodología funciona como red de seguridad. La segunda rediseñó la arquitectura del sistema completo: orden correcto por flujo de negocio, sub-módulos bien definidos, y un nuevo módulo de coordinación (7_cierre) que no existía. Los 9 módulos ahora cuentan una historia que cualquiera puede leer de izquierda a derecha y entender.
