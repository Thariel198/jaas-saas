# crear_origenes

Acumula todos los orígenes de pago vistos por MZ-LOTE en el historial del banco.
No filtra ni decide — solo observa y registra.
Su output es el insumo de `crear_maestro`, que consolida y calcula nivel de confianza.

---

## Posición en el pipeline

```
01_lecturas
02_planilla
03_pagos/yape/construir_maestro/crear_origenes   ← estás aquí
        ↓
03_pagos/yape/construir_maestro/crear_maestro    → consolida · calcula nivel_confianza · produce maestro_yape.xlsx
        ↓
03_pagos/yape/motor_matching                     → consume maestro_yape.xlsx · identifica pagos
        ↓
04_cobranza
        ↓
04b_validacion
        ↓
05_corte
        ↓
05b_override
        ↓
06_boletas
        ↓
07_nueva_planilla
```

---

## Estructura de carpetas

```
crear_origenes/
├── outputs/
│   ├── origenes_yape.xlsx        → orígenes brutos acumulados por MZ-LOTE → lo lee crear_maestro
│   └── mz_no_encontradas.xlsx    → alerta · MZ-LOTE sin match en usuarios_id
├── main.py
└── README.md                     → este archivo
```

**Lee desde shared/ — no tiene inputs propios:**

```
shared/
├── reporte_acumulado_procesado/            → reportes procesados del banco · uno por mes
│   └── reporte_tepago_YYYY_MM.xlsx      → nombre estándar · snake_case
└── usuarios_id.xlsx              → identidad de cada MZ-LOTE
```

---

## Cómo piensa el sistema

Para cada archivo en `shared/reporte_acumulado_procesado/`:

```
1. Lee el archivo del banco
2. Filtra solo filas "TE PAGÓ"
3. Extrae: origen · mz · lote
4. Cruza mz-lote con usuarios_id.xlsx
        → existe    → acumula el origen para ese mz-lote
        → no existe → va a mz_no_encontradas.xlsx · no se incluye en origenes_yape
5. Cuenta total_apariciones y meses_en_registros por mz-lote
6. Exporta origenes_yape.xlsx y mz_no_encontradas.xlsx
```

> El sistema no juzga si un origen es válido — eso lo hace `crear_maestro`.
> Un MZ-LOTE sin match en usuarios_id es casi siempre error de tipeo — no un lote nuevo.
> Si ves algo en mz_no_encontradas con 2+ meses → revísalo en `01_lecturas`.

---

## Cuándo correrlo

No es mensual — se corre cuando hay nuevos reportes disponibles en `shared/reporte_acumulado_procesado/`.
Cada vez regenera `origenes_yape.xlsx` completo desde cero leyendo todo el historial.

---

## Flujo de uso

**1. Verificar que los inputs están listos:**

```
shared/reporte_acumulado_procesado/     → al menos un archivo .xlsx del banco procesado
shared/usuarios_id.xlsx       → existe y tiene USER_ID · NOMBRE · MZ · LOTE
```

**2. Correr:**

```bash
python main.py
```

El sistema corre y genera:
- `outputs/origenes_yape.xlsx` — todos los MZ-LOTE con sus orígenes acumulados
- `outputs/mz_no_encontradas.xlsx` — MZ-LOTE que no existen en usuarios_id

**3. Revisar mz_no_encontradas.xlsx:**

| MESES_EN_REGISTROS | Qué hacer |
|---|---|
| `1` | ignorar · probable error de tipeo · se corrige solo |
| `2+` | revisar · puede ser lote genuinamente nuevo → agrégalo en `01_lecturas` |

**4. Pasar el output a crear_maestro:**

```
outputs/origenes_yape.xlsx   → crear_maestro lo lee directo desde esta carpeta
```

---

## Lifecycle — qué se regenera y qué permanece

| Archivo | ¿Qué pasa cada vez que corre? | Motivo |
|---|---|---|
| `outputs/origenes_yape.xlsx` | Se regenera completo | lee todo el historial cada vez |
| `outputs/mz_no_encontradas.xlsx` | Se regenera completo | refleja el estado actual |
| `shared/reporte_acumulado_procesado/` | **PERMANECE** — nunca se toca | fuente de verdad · solo crece |
| `shared/usuarios_id.xlsx` | **PERMANECE** — nunca se toca | lo gestiona otro módulo |

---

## origenes_yape.xlsx

Output principal. Lo consume `crear_maestro`. No lo edites manualmente.

| Grupo | Columnas | Origen |
|---|---|---|
| ¿Quién es? | USER_ID · NOMBRE | cruzado desde usuarios_id.xlsx |
| ¿Dónde vive? | MZ · LOTE | extraído del reporte del banco |
| ¿Quién ha pagado? | ORIGEN_1 · ORIGEN_2 · ORIGEN_3... | todos los orígenes vistos · sin filtrar |
| ¿Cuánto lo hemos visto? | TOTAL_APARICIONES · MESES_EN_REGISTROS | calculado del historial |
| ¿Cuándo se registró? | FECHA_REGISTRO | fecha en que corrió crear_origenes |

**Reglas:**
- Una fila por MZ-LOTE — nunca duplicados
- `TOTAL_APARICIONES` — cuántas filas TE PAGÓ se registraron en todo el historial para ese MZ-LOTE
- `MESES_EN_REGISTROS` — en cuántos archivos distintos de `reporte_acumulado_procesado/` apareció
- `ORIGEN_N` — columnas dinámicas · tantas como el máximo de orígenes distintos vistos
- MZ-LOTE sin match en usuarios_id no aparecen aquí — van a `mz_no_encontradas.xlsx`

---

## mz_no_encontradas.xlsx

Output de alerta. Solo lectura — no lo edites.

| Grupo | Columnas | Origen |
|---|---|---|
| ¿Qué vio el banco? | ORIGEN | tal como viene del reporte |
| ¿Dónde dijo que vivía? | MZ · LOTE | lo que escribió el usuario |
| ¿Cuántas veces lo vimos? | TOTAL_APARICIONES · MESES_EN_REGISTROS | calculado del historial |
| ¿Cuándo se registró? | FECHA_REGISTRO | fecha en que corrió crear_origenes |

---

## Errores comunes

**"No hay archivos .xlsx en reporte_acumulado_procesado/"**
→ Verificar que `shared/reporte_acumulado_procesado/` tiene al menos un reporte del banco procesado.

**"usuarios_id.xlsx no encontrado"**
→ Verificar que `shared/usuarios_id.xlsx` existe y tiene las columnas USER_ID · NOMBRE · MZ · LOTE.

**"Hoja 'reporte' no encontrada en archivo X"**
→ El archivo del banco no tiene hoja llamada `reporte` — verificar que el nombre de la hoja está en minúsculas.

**"Sin filas TE PAGÓ en archivo X"**
→ El archivo no tiene filas con tipo `TE PAGÓ` — puede estar vacío o con formato distinto.

---

## Lo que NO hace este módulo

- **No decide qué orígenes son válidos** — eso lo hace `crear_maestro`
- **No calcula nivel de confianza** — eso lo hace `crear_maestro`
- **No agrega usuarios nuevos** — eso lo hace `01_lecturas`
- **No modifica usuarios_id.xlsx** — es solo lectura aquí
- **No es mensual** — corre cuando hay nuevo historial disponible

---

## Antes de codificar — verificar inputs

Checklist completo antes de escribir una línea de código.
Si algún archivo falta o tiene columnas distintas → el código fallará.

### ¿De dónde viene cada archivo?

| Archivo | Lo produce | Vive en |
|---|---|---|
| Reportes del banco procesados | `motor_matching` · tú los mueves aquí cada mes | `shared/reporte_acumulado_procesado/` |
| `usuarios_id.xlsx` | se gestiona en `01_lecturas` | `shared/` |

### Columnas requeridas — verifica antes de codificar

**Reportes del banco — `shared/reporte_acumulado_procesado/*.xlsx`**
```
✓ Tipo de Transacción   → valor exacto: "TE PAGÓ"
✓ Origen                → nombre del pagador · puede venir truncado con *
✓ MZ                    → manzana · puede llamarse "Manzana" o "MZ"
✓ LT                    → lote · puede llamarse "Lote" o "LT"
✓ Fecha de operación    → formato: DD/MM/YYYY HH:MM
```

**`shared/usuarios_id.xlsx`**
```
✓ USER_ID               → identificador único del usuario
✓ NOMBRE                → nombre completo
✓ MZ                    → manzana del primer lote
✓ LT                    → número del primer lote
✓ MZ_2                  → manzana del segundo lote · si tiene un segundo lote
✓ LT_2                  → número del segundo lote · si tiene un segundo lote
```

> Una fila por usuario — los lotes se expanden horizontalmente en columnas.
> Un usuario con 2 lotes tiene 4 columnas de ubicación: MZ · LT · MZ_2 · LT_2
> Un usuario con 3 lotes tendría 6 columnas: MZ · LT · MZ_2 · LT_2 · MZ_3 · LT_3
> La clave de cruce es siempre el par (MZ, LT) — el código debe buscar en todos los pares disponibles.

### Checklist antes de codificar

```
[ ] shared/reporte_acumulado_procesado/ existe y tiene al menos un .xlsx
[ ] Cada .xlsx tiene hoja llamada "reporte"
[ ] Columnas verificadas — especialmente "Tipo de Transacción" · "MZ" · "LT"
[ ] shared/usuarios_id.xlsx existe
[ ] usuarios_id.xlsx tiene USER_ID · NOMBRE · MZ · LOTE
[ ] outputs/ existe o el código la crea automáticamente
```

---

## Requisitos

- Python 3.10 o superior
- Librerías: `pandas` · `openpyxl`

```bash
pip install pandas openpyxl
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
6.  Cómo se usa — flujo paso a paso
7.  Qué se regenera y qué permanece — lifecycle
8.  Los archivos que produce — uno por uno con columnas
9.  Si algo falla — errores comunes
10. Sus límites — lo que NO hace
11. Antes de codificar — verificar inputs · columnas · checklist
12. Requisitos
```

Todos los README del sistema siguen este mismo orden.
Si encuentras algo fuera de orden — corrígelo para mantener el estándar.
