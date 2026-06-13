# utils

Funciones de lógica compartida entre módulos de `construir_maestro`.
No produce archivos — solo expone funciones que otros módulos importan.

---

## Posición en el sistema

```
construir_maestro/utils/   ← estás aquí
        ↑
        lo importan:
        construir_maestro/crear_maestro/main.py
        construir_maestro/test/test.py
```

---

## Estructura de carpetas

```
utils/
├── consolidar.py    → lógica de similitud · filtrado · nivel de confianza
└── README.md        → este archivo
```

---

## consolidar.py

### `similitud(a, b) → float`

Calcula similitud entre dos strings ignorando orden de palabras y asteriscos.
Retorna float entre 0 y 100.

```python
from utils.consolidar import similitud

similitud("Linda Paj*", "LINDA PAJUELO COCHACHIN")  # → ~72.0
similitud("Janet Vil*", "LINDA PAJUELO COCHACHIN")  # → ~21.0
```

---

### `filtrar_origenes_validos(origenes, umbral, nombre) → list`

Conserva solo los orígenes que tienen al menos:
- un par similar con otro origen del mismo MZ-LOTE (similitud >= umbral)
- O similitud >= umbral con el nombre registrado del usuario

Los intrusos — sin par ni similitud con nombre — se descartan.

```python
from utils.consolidar import filtrar_origenes_validos

origenes = ["Linda Paj*", "Linda F. Pajuelo C.", "Janet Villanueva A."]
validos  = filtrar_origenes_validos(origenes, umbral=60, nombre="LINDA PAJUELO COCHACHIN")
# → ["Linda Paj*", "Linda F. Pajuelo C."]
# Janet descartada — sin par · sin similitud con nombre
```

---

### `calcular_nivel(origenes, nombre, meses) → (str, list)`

Calcula el nivel de confianza y los orígenes válidos para un MZ-LOTE.

```python
from utils.consolidar import calcular_nivel

nivel, validos = calcular_nivel(
    origenes = ["Linda Paj*", "Linda F. Pajuelo C.", "Janet Villanueva A."],
    nombre   = "LINDA PAJUELO COCHACHIN",
    meses    = 6
)
# → ("alto", ["Linda Paj*", "Linda F. Pajuelo C."])
```

**Reglas de nivel:**

| Nivel | Condición |
|---|---|
| `alto` | 2+ orígenes con similitud >= 70% entre sí · O 1 origen con similitud >= 70% vs nombre y meses >= 2 |
| `medio` | 2+ orígenes con similitud >= 40% entre sí · O 1 origen con similitud >= 40% vs nombre · O 1 origen con meses >= 3 |
| `bajo` | ninguna condición anterior · no va al maestro |

**Regla de intrusos:**
- Un origen es válido si tiene par similar O similitud con nombre
- Un origen sin par y sin similitud con nombre se descarta siempre

---

## Cómo importar

Desde `crear_maestro/main.py`:
```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from utils.consolidar import calcular_nivel, similitud
```

Desde `test/test.py`:
```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from utils.consolidar import calcular_nivel, similitud
```

---

## Umbrales configurables

Los umbrales viven en `consolidar.py` — ajústalos según tus datos reales:

```python
UMBRAL_ALTO  = 70   # similitud >= 70% → alto
UMBRAL_MEDIO = 40   # similitud >= 40% → medio
```

> Nota: bajar `UMBRAL_ALTO` a 60 mejora la detección de nombres truncados con *

---

## Lo que NO hace este módulo

- **No lee archivos** — solo funciones puras
- **No escribe archivos** — solo retorna valores
- **No tiene estado** — cada llamada es independiente

---

## Requisitos

- Python 3.10 o superior
- Librerías: `rapidfuzz`

```bash
C:\Users\wilde\AppData\Local\Programs\Python\Python313\python.exe -m pip install rapidfuzz
```
