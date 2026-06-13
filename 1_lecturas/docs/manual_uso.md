# 1_lecturas

Registra, valida y depura las lecturas mensuales del operario. Genera el archivo limpio
que alimenta a `2_planilla` para calcular el cobro.
Corre en ciclos hasta que no queden anomalías bloqueantes.

---

## Posición en el pipeline

```
1_lecturas    ← estás aquí
        ↓
2_planilla
4_pagos
5_cobranza
5b_validacion
6_corte
6b_override
3_boletas
3_boletas/enriquecimiento
7_cierre
```

---

## Estructura de carpetas

```
1_lecturas/
├── inputs/
│   ├── registro_operario_mes.xlsx        → operario llena MARC_ACT, M3, obs_operario
│   └── registro_operario_acumulado.xlsx  → historial permanente · NUNCA borrar · input/output
│
├── outputs/                              → temporal del ciclo + permanentes
│   ├── correcciones_YYYY-MM.xlsx         → cola de maquillaje · se borra al cerrar
│   ├── trazabilidad_YYYY-MM.xlsx         → auditoría permanente · uno por mes
│   ├── lecturas_planilla_YYYY-MM.xlsx    → input limpio para 2_planilla · PERMANECE
│   ├── orden_verificacion_YYYY-MM.pdf    → para el operario en campo · se regenera
│   └── run.log
│
├── docs/                                 → contratos y diseño · nunca se borran
│   ├── arquitectura_1_lecturas.html
│   ├── contrato_registro_operario_mes.html
│   ├── contrato_registro_operario_acumulado.html
│   ├── contrato_correcciones.html
│   ├── contrato_trazabilidad.html
│   ├── contrato_lecturas_planilla.html
│   ├── contrato_orden_verificacion.html
│   ├── flujo_modulo_1_lecturas.svg
│   ├── manual_uso.md                     → este archivo
│   └── decisiones/
│       └── 1_lecturas.md
│
├── tests/
├── main.py
├── crear_template.py
└── config.py                             → umbrales, paths, códigos obs_operario
```

---

## Las 13 anomalías

El sistema valida cada lectura contra el historial y aplica las reglas en orden.

### Bloqueantes (8) — impiden generar lecturas_planilla hasta resolverse

| Tipo | Cuándo | Acción |
|---|---|---|
| `MEDIDOR_INVERTIDO` | MARC_ACT < MARC_ANT y la diferencia ≤ `UMBRAL_INVERSION` (50 m³), sin obs=M | verificar si el medidor está físicamente al revés |
| `POSIBLE_CAMBIO_MEDIDOR` | MARC_ACT < MARC_ANT y la diferencia > `UMBRAL_INVERSION`, sin obs=M | confirmar si cambiaron el medidor sin reportar |
| `DIFERENCIA_M3` | `\|M3_operario − (MARC_ACT − MARC_ANT)\| > 0.05` (sin mínimo) | corregir cálculo con el operario |
| `EXCESIVO` | consumo calculado > `M3_EXCESIVO` (50 m³) | verificar fuga o lectura |
| `SIN_LECTURA` | MARC_ACT vacío (sin obs=P) | el operario va al predio a leer |
| `DUPLICADO` | misma (MZ, LT) aparece 2+ veces en el template | borrar la fila duplicada |
| `USUARIO_FANTASMA` | usuario en acumulado falta en el template | confirmar si sigue activo o marcar baja |
| `MARC_ACT_NO_NUMERICO` | MARC_ACT contiene caracteres no numéricos | corregir el dato (suelen ser O/0, l/1) |

### Informativas (4) — se aceptan automáticamente, quedan en trazabilidad

| Tipo | Cuándo |
|---|---|
| `SIN_HISTORIAL` | usuario nuevo sin ciclos anteriores en el acumulado |
| `CONSUMO_CERO` | MARC_ACT = MARC_ANT y M3 = 0 (sin mínimo aplicado) |
| `SALTO_HISTORICO` | M3 actual > `SALTO_FACTOR` × promedio de últimos `SALTO_MESES` ciclos |
| `MEDIDOR_CAMBIADO` | MEDIDOR_INVERTIDO o POSIBLE_CAMBIO_MEDIDOR con obs=M (legitimado por el operario) |

### Códigos de obs_operario (legitiman situaciones de campo)

| Código | Efecto |
|---|---|
| `M` | MARC_ACT < MARC_ANT → no es bloqueante (MEDIDOR_INVERTIDO o POSIBLE_CAMBIO_MEDIDOR), es MEDIDOR_CAMBIADO informativo |
| `F` | dispara informativo siempre (fuga visible) — sirve de alerta al supervisor |
| `P` | MARC_ACT vacío → no es SIN_LECTURA bloqueante, es SIN_LECTURA informativo |
| vacío | lectura normal sin observación |

### Constantes (en `config.py`)

```python
M3_MINIMO        = 5.0    # mínimo de cobro aplicado por el operario
M3_EXCESIVO      = 50.0   # umbral para EXCESIVO
UMBRAL_INVERSION = 50.0   # separa MEDIDOR_INVERTIDO (≤) de POSIBLE_CAMBIO_MEDIDOR (>)
SALTO_MESES      = 4      # ciclos previos para promedio
SALTO_FACTOR     = 3.0    # multiplicador para SALTO_HISTORICO
```

---

## Ciclos — Ciclo 1 vs Ciclo 2+

El sistema detecta automáticamente en qué ciclo está:

| Ciclo | Señal | Qué hace |
|---|---|---|
| **Ciclo 1** | `outputs/lecturas_planilla_YYYY-MM.xlsx` NO existe | Primer pass del mes · limpia `correcciones_YYYY-MM.xlsx` anterior si quedó · no toca acumulado/trazabilidad/lecturas_planilla |
| **Ciclo 2+** | ese archivo SÍ existe | Lee `correcciones_YYYY-MM.xlsx` con maquillajes aplicados · mueve filas resueltas a trazabilidad · regenera correcciones con las que faltan · al llegar a 0 → genera `lecturas_planilla` final |

---

## Flujo mensual

### Paso 1 — Generar el template del operario

Al inicio del mes, una sola vez:

```bash
python crear_template.py
```

Genera `inputs/registro_operario_mes.xlsx` con todos los usuarios pre-cargados desde
el acumulado: MZ, LT, NOMBRE, MES_ANO, MARC_ANT. Columnas `MARC_ACT`, `M3` y `obs_operario`
en amarillo para que el operario complete. Incluye la leyenda M/F/P impresa encima.

### Paso 2 — El operario completa las lecturas

Recorre las manzanas anotando:
- `MARC_ACT` — número del display del medidor
- `M3` — metros cúbicos calculados a mano
- `obs_operario` — solo si aplica (M/F/P)

### Paso 3 — Correr el módulo

```bash
python main.py
```

El sistema valida, separa confirmados de anomalías, y según el ciclo:

**Si todo pasa sin anomalías bloqueantes:**
- Actualiza `registro_operario_acumulado.xlsx` con el par de columnas del ciclo
- Genera `lecturas_planilla_YYYY-MM.xlsx` (input para 2_planilla)
- Registra informativas en `trazabilidad_YYYY-MM.xlsx`
- Listo

**Si hay anomalías bloqueantes:**
- Genera `correcciones_YYYY-MM.xlsx` con la cola de maquillaje
- Genera `orden_verificacion_YYYY-MM.pdf` con los casos de campo
- Espera ciclo de corrección

### Paso 4 — Ciclo de corrección (repetir hasta bloqueantes = 0)

1. Abrir `outputs/orden_verificacion_YYYY-MM.pdf` — enviar al operario para casos de campo
2. Abrir `outputs/correcciones_YYYY-MM.xlsx` — llenar el bloque verde para cada fila pendiente:
   - `MARC_ACT_corregido` y `M3_corregido`
   - `motivo_correccion` (texto libre)
   - `resuelto_por` (campo / maquillaje / acepta_original / corrige_dato / borra_duplicado / marca_baja)
3. Volver a correr `python main.py`. Las filas con `resuelto_por` lleno se mueven a trazabilidad. Las que faltan regeneran el archivo.
4. Repetir hasta que `correcciones_YYYY-MM.xlsx` quede en 0 filas y desaparezca.

### Paso 5 — Cierre

Cuando bloqueantes = 0:
- `lecturas_planilla_YYYY-MM.xlsx` queda generado y listo
- `correcciones_YYYY-MM.xlsx` se elimina
- Las informativas del mes ya están en `trazabilidad_YYYY-MM.xlsx`

### Paso 6 — Listo → 2_planilla

```bash
python ../2_planilla/main.py
```

---

## Lifecycle — qué se borra y qué permanece

| Archivo | ¿Qué pasa? | Motivo |
|---|---|---|
| `inputs/registro_operario_acumulado.xlsx` | **PERMANECE** · se actualiza | historial puro del operario — nunca borrar |
| `inputs/registro_operario_mes.xlsx` | Tú reemplazas cada mes | correr `crear_template.py` al inicio |
| `outputs/trazabilidad_YYYY-MM.xlsx` | **PERMANECE** · uno por mes | auditoría permanente |
| `outputs/lecturas_planilla_YYYY-MM.xlsx` | **PERMANECE** · uno por mes | fuente de verdad para 2_planilla |
| `outputs/orden_verificacion_YYYY-MM.pdf` | Se regenera cada ciclo | refleja anomalías del ciclo actual |
| `outputs/correcciones_YYYY-MM.xlsx` | Se elimina al cerrar | datos volcados a trazabilidad antes de eliminar |
| `outputs/run.log` | Se sobreescribe cada run | log del último run |

---

## Archivos producidos — uno por uno

Cada archivo tiene su contrato HTML detallado en `docs/`. Esto es solo el resumen.

### registro_operario_acumulado.xlsx

Historial permanente con doble header. Una fila por usuario, dos sub-columnas por ciclo.

| Grupo | Columnas |
|---|---|
| Identificación | MZ · LT · NOMBRE |
| Por ciclo (par de columnas) | YYYY-MM.MARCACION · YYYY-MM.M3 |

Reglas resumidas: solo se escriben confirmados · vacío = sin dato · mes nuevo al final · último mes resaltado en azul · re-correr el mismo mes sobrescribe el par. Detalle completo en `docs/contrato_registro_operario_acumulado.html`.

### correcciones_YYYY-MM.xlsx

Cola de maquillaje. Una hoja `Correcciones`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ · LT · NOMBRE |
| ¿Qué entregó el operario? (no editar) | MES_ANO · MARC_ANT · MARC_ACT_original · M3_original · obs_operario_original |
| ¿Cuál es la anomalía? | tipo_anomalia · motivo_detectado |
| ¿Cómo se corrige? (supervisor llena) | MARC_ACT_corregido · M3_corregido · motivo_correccion · resuelto_por |
| ¿Estado y cuándo? | estado · ciclo · fecha_correccion |

Detalle en `docs/contrato_correcciones.html`.

### trazabilidad_YYYY-MM.xlsx

Registro permanente de auditoría. Una hoja `Trazabilidad`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ · LT · NOMBRE |
| ¿Qué entregó el operario? (inmutable) | MES_ANO · MARC_ANT · MARC_ACT_original · M3_original · obs_operario_original |
| ¿Cuál fue la anomalía? | categoria · tipo_anomalia · motivo_detectado |
| ¿Cómo se resolvió? | MARC_ACT_final · M3_final · motivo_correccion · resuelto_por |
| ¿Cuándo? | ciclo · fecha_correccion |

Detalle en `docs/contrato_trazabilidad.html`.

### lecturas_planilla_YYYY-MM.xlsx

Lecturas facturables. Una hoja `LecturasPlanilla`.

| Grupo | Columnas |
|---|---|
| ¿Quién es? | MZ · LT · NOMBRE · MES_ANO |
| ¿Qué se factura? | MARC_ANT · MARC_ACT · M3 |
| ¿Cómo se llegó? | origen · ciclo |

`origen` ∈ {directo, minimo_aplicado, corregido, informativa_legitimada}. Detalle en `docs/contrato_lecturas_planilla.html`.

### orden_verificacion_YYYY-MM.pdf

Documento para el operario en campo. Solo incluye `SIN_LECTURA`, `MEDIDOR_INVERTIDO`, `POSIBLE_CAMBIO_MEDIDOR`, `EXCESIVO`. Una tarjeta por caso con espacios en blanco para escribir a mano. Detalle en `docs/contrato_orden_verificacion.html`.

---

## Errores comunes

**"No se encontró registro_operario_mes.xlsx"**
→ Correr `python crear_template.py` primero. Si es el primer mes del sistema, el template queda vacío y el operario lo completa manualmente.

**"Columnas faltantes en registro_operario_acumulado.xlsx"**
→ El archivo no tiene el doble header. Posiblemente se generó con una versión vieja del módulo. Reconstruir migrando manualmente o desde otra fuente.

**"obs_operario tiene valor inválido: X"**
→ Solo se aceptan M, F, P, o celda vacía. Revisar el template y corregir.

**"Anomalía registrada pero el operario dice que está bien"**
→ El operario puede legitimar con obs (M para retroceso, F para fuga, P para predio cerrado). Si no aplica obs y el valor es correcto, marcar `resuelto_por = acepta_original` en correcciones.

---

## Lo que NO hace este módulo

- **No calcula cobros** — eso lo hace `2_planilla` a partir de `lecturas_planilla_YYYY-MM.xlsx`
- **No genera boletas** — eso es `3_boletas`
- **No corrige lecturas ya facturadas** — para eso existe `6b_override`
- **No mantiene tarifas** — vive en `shared/tarifas.xlsx` y lo lee `2_planilla`

---

## Antes de codificar — verificar inputs

### Checklist antes de correr

```
[ ] inputs/registro_operario_mes.xlsx existe (generado por crear_template.py)
[ ] inputs/registro_operario_acumulado.xlsx existe o es primer mes del sistema
[ ] El operario completó las columnas MARC_ACT y M3 del template
[ ] config.py tiene los umbrales correctos para el ciclo actual
```

### Columnas requeridas

**registro_operario_mes.xlsx** (verifica antes de correr `main.py`):
```
✓ MZ                 → texto (mayúsculas)
✓ LT                 → texto o número (acepta 8A)
✓ NOMBRE             → texto
✓ MES_ANO            → YYYY-MM (ej: 2026-06)
✓ MARC_ANT           → número (lectura del ciclo anterior)
✓ MARC_ACT           → número (operario llena)
✓ M3                 → número (operario llena)
✓ obs_operario       → M / F / P / vacío
```

**registro_operario_acumulado.xlsx** (verifica antes de correr `crear_template.py`):
```
✓ MZ · LT · NOMBRE   → identificación
✓ YYYY-MM.MARCACION  → última columna del último ciclo (la sub-columna marcación)
✓ YYYY-MM.M3         → última columna del último ciclo (la sub-columna M3)
```

---

## Requisitos

- Python 3.10+
- Librerías: `pandas`, `openpyxl`, `python-docx`, `docx2pdf`

```bash
pip install pandas openpyxl python-docx docx2pdf
```

`docx2pdf` requiere Microsoft Word instalado (se usa también en `3_boletas`).

---

## Cómo leer este README

Sigue el orden narrativo estándar del proyecto:

```
1.  Qué hace el módulo
2.  Dónde encaja en el pipeline
3.  Dónde viven los archivos
4.  Cómo piensa el sistema — las 11 anomalías
5.  Cómo detecta el contexto — ciclos
6.  Cómo se usa mes a mes — flujo
7.  Qué se borra y qué permanece — lifecycle
8.  Los archivos que produce — uno por uno con sus columnas
9.  Si algo falla — errores comunes
10. Sus límites — lo que NO hace
11. Antes de codificar — verificar inputs
12. Cómo instalarlo — requisitos
```
