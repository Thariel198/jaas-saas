# crear_maestro

Consolida los orígenes brutos de `crear_origenes` y descarta intrusos.
Produce `maestro_yape.xlsx` — el archivo que consume `motor_matching` para identificar pagos.
Los MZ-LOTE sin orígenes válidos van a `sin_confirmar.xlsx` para revisión manual por ciclos.

---

## Posición en el pipeline

```
1_lecturas
2_planilla
4_pagos/yape/construir_maestro/crear_origenes   → produce origenes_yape.xlsx
        ↓
4_pagos/yape/construir_maestro/crear_maestro    ← estás aquí
        ↓
4_pagos/yape/motor_matching                     → consume maestro_yape.xlsx · identifica pagos
        ↓
5_cobranza
        ↓
5b_validacion
        ↓
6_corte
        ↓
6b_override
        ↓
3_boletas
        ↓
7_cierre
```

---

## Estructura de carpetas

```
crear_maestro/
├── outputs/
│   ├── maestro_yape.xlsx         → orígenes consolidados · lo consume motor_matching
│   └── sin_confirmar.xlsx        → sin orígenes válidos · revisión manual por ciclos
├── main.py
└── README.md                     → este archivo
```

**Lee desde crear_origenes/ — no tiene inputs propios en shared/:**

```
crear_origenes/
└── outputs/
    └── origenes_yape.xlsx        → input de este módulo
```

---

## Cómo piensa el sistema

Para cada fila de `origenes_yape.xlsx`:

```
1. Toma todos los orígenes brutos del MZ-LOTE (ORIGEN_1, ORIGEN_2, ORIGEN_3...)
2. Para cada origen evalúa 3 criterios:
        Criterio 1 — comparte token exacto con otro origen
                     → detecta familia por apellido compartido
                     → ej: Adelina Chi* + Marylie Chi* comparten CHI
        Criterio 2 — comparte token exacto con el nombre registrado
                     → ej: Linda Paj* comparte LINDA con LINDA PAJUELO COCHACHIN
        Criterio 3 — similitud fuzzy >= 50% vs nombre registrado
                     → detecta nombres truncados con *
                     → ej: Camile Hid* ≈ CAMILA HIDALGO
3. Origen válido → cumple al menos 1 criterio
   Intruso       → no cumple ninguno → se descarta
4. Si hay orígenes válidos → va al maestro_yape
   Si no hay ninguno       → va a sin_confirmar para revisión manual
```

> El fuzzy solo se usa para comparar origen vs nombre — nunca entre orígenes.
> La detección de familia usa tokens exactos — más precisa que fuzzy para apellidos.
> Umbral fuzzy configurable en `utils/consolidar.py` → `UMBRAL_NOMBRE = 50`

---

## Cuándo correrlo

Se corre cada vez que `crear_origenes` regenera `origenes_yape.xlsx`.
Siempre después de `crear_origenes` — nunca antes.
Se puede correr múltiples veces en ciclos para procesar los autorizados de `sin_confirmar.xlsx`.

---

## Flujo de uso

### Ciclo 1 — primera corrida del mes

**1. Verificar que el input está listo:**

```
crear_origenes/outputs/origenes_yape.xlsx   → existe y fue generado por crear_origenes
```

**2. Correr:**

```bash
python main.py
```

El sistema genera:
- `outputs/maestro_yape.xlsx` — orígenes válidos consolidados
- `outputs/sin_confirmar.xlsx` — MZ-LOTE sin orígenes válidos identificados

**3. Revisar sin_confirmar.xlsx:**

| Caso | Qué hacer |
|---|---|
| Reconoces quién paga | escribe el origen correcto en `ORIGEN_CORRECTO` · pon `AUTORIZADO = si` |
| No reconoces quién paga | deja `AUTORIZADO = no` · volverá el próximo mes con más historial |

### Ciclo 2+ — procesar autorizados

**4. Correr nuevamente:**

```bash
python main.py
```

Los autorizados de `sin_confirmar.xlsx` suben al maestro con `VALIDADO_MANUAL = si`
y no vuelven a aparecer en `sin_confirmar`.

---

## Lifecycle — qué se regenera y qué permanece

| Archivo | ¿Qué pasa cada vez que corre? | Motivo |
|---|---|---|
| `outputs/maestro_yape.xlsx` | Se regenera completo | se recalcula desde origenes_yape cada vez |
| `outputs/sin_confirmar.xlsx` | Se regenera completo | refleja el estado actual · los autorizados suben al maestro |
| `crear_origenes/outputs/origenes_yape.xlsx` | **NO SE TOCA** — solo lectura | lo gestiona crear_origenes |

> Los autorizados se leen **antes** de borrar outputs/ — nunca se pierden.
> `VALIDADO_MANUAL = si` también se preserva entre regeneraciones.

---

## maestro_yape.xlsx

Output principal. Lo consume `motor_matching`. No edites orígenes directamente.

| Grupo | Columnas | Origen |
|---|---|---|
| ¿Quién es? | USER_ID · NOMBRE | desde origenes_yape.xlsx |
| ¿Dónde vive? | MZ · LOTE | desde origenes_yape.xlsx |
| ¿Cuáles son sus orígenes válidos? | ORIGEN_1 · ORIGEN_2... | filtrados por token y fuzzy vs nombre |
| ¿Validado manual? | VALIDADO_MANUAL | `si` si vino de sin_confirmar · `no` si fue automático |
| ¿Cuándo se consolidó? | FECHA_REGISTRO | fecha en que corrió crear_maestro |

**Reglas:**
- Una fila por MZ-LOTE — nunca duplicados
- `ORIGEN_N` — solo orígenes válidos · intrusos ya descartados
- `VALIDADO_MANUAL = si` → origen confirmado manualmente · nunca vuelve a sin_confirmar

---

## sin_confirmar.xlsx

Output de auditoría. Se procesa por ciclos — no tiene fecha límite.

| Grupo | Columnas | Quién la llena |
|---|---|---|
| ¿Quién es? | USER_ID · NOMBRE | sistema |
| ¿Dónde vive? | MZ · LOTE | sistema |
| ¿Qué orígenes vimos? | ORIGEN_1 · ORIGEN_2... | sistema — todos los brutos |
| ¿Cuánto lo hemos visto? | TOTAL_APARICIONES · MESES_EN_REGISTROS | sistema |
| ¿Qué decides? | ORIGEN_CORRECTO · AUTORIZADO | tú |
| ¿Cuándo se registró? | FECHA_REGISTRO | sistema |

**2 casos únicamente:**

| AUTORIZADO | ORIGEN_CORRECTO | Qué pasa |
|---|---|---|
| `no` | vacío | sigue apareciendo cada mes · aún sin decidir |
| `si` | lleno | sube al maestro con VALIDADO_MANUAL = si · nunca más aparece |

---

## Errores comunes

**"origenes_yape.xlsx no encontrado"**
→ Verificar que corriste `crear_origenes` primero y que el archivo existe en `crear_origenes/outputs/`.

**"Columna ORIGEN_1 no encontrada"**
→ El archivo `origenes_yape.xlsx` no tiene orígenes — puede estar vacío o con formato distinto.

**"NOMBRE vacío para USER_ID X"**
→ El usuario existe en usuarios_id.xlsx pero sin nombre — completar en `shared/usuarios_id.xlsx`.

---

## Lo que NO hace este módulo

- **No acumula orígenes brutos** — eso lo hace `crear_origenes`
- **No identifica pagos del banco** — eso lo hace `motor_matching`
- **No agrega usuarios nuevos** — eso lo hace `1_lecturas`
- **No modifica usuarios_id.xlsx** — es solo lectura aquí
- **No sobreescribe validaciones manuales** — `VALIDADO_MANUAL = si` se preserva siempre
- **No asigna nivel de confianza** — los orígenes son válidos o no · sin gradación subjetiva

---

## Antes de codificar — verificar inputs

### ¿De dónde viene cada archivo?

| Archivo | Lo produce | Vive en |
|---|---|---|
| `origenes_yape.xlsx` | `crear_origenes` | `crear_origenes/outputs/` |

### Columnas requeridas — verifica antes de codificar

**`crear_origenes/outputs/origenes_yape.xlsx`**
```
✓ USER_ID                → identificador único del usuario
✓ NOMBRE                 → nombre completo · se usa para fuzzy y token vs orígenes
✓ MZ                     → manzana
✓ LOTE                   → número de lote
✓ ORIGEN_1               → al menos un origen · ORIGEN_2 · ORIGEN_3 son opcionales
✓ TOTAL_APARICIONES      → número entero
✓ MESES_EN_REGISTROS     → número entero
✓ FECHA_REGISTRO         → formato DD/MM/YYYY
```

### Checklist antes de codificar

```
[ ] crear_origenes/outputs/origenes_yape.xlsx existe
[ ] origenes_yape.xlsx tiene USER_ID · NOMBRE · MZ · LOTE · ORIGEN_1
[ ] origenes_yape.xlsx tiene TOTAL_APARICIONES · MESES_EN_REGISTROS
[ ] outputs/ existe o el código la crea automáticamente
[ ] utils/consolidar.py existe con filtrar_origenes_validos
[ ] UMBRAL_NOMBRE definido en utils/consolidar.py — por defecto 50
```

---

## Requisitos

- Python 3.10 o superior
- Librerías: `openpyxl` · `rapidfuzz`

```bash
C:\Users\wilde\AppData\Local\Programs\Python\Python313\python.exe -m pip install openpyxl rapidfuzz
```

---

## Cómo leer este README

Este README sigue siempre el mismo orden narrativo:

```
1.  Qué hace el módulo
2.  Dónde encaja en el pipeline
3.  Dónde viven los archivos
4.  Cómo piensa el sistema
5.  Cuándo correrlo
6.  Cómo se usa — flujo paso a paso · ciclos
7.  Qué se regenera y qué permanece — lifecycle
8.  Los archivos que produce — uno por uno con columnas
9.  Si algo falla — errores comunes
10. Sus límites — lo que NO hace
11. Antes de codificar — verificar inputs · columnas · checklist
12. Requisitos
```

Todos los README del sistema siguen este mismo orden.
Si encuentras algo fuera de orden — corrígelo para mantener el estándar.
