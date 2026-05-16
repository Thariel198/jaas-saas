# motor_matching

Identifica a qué MZ-LOTE corresponde cada pago Yape descargado del banco.
Corre en ciclos hasta que todos los pagos estén identificados (pendientes = 0).
Solo `Sin_identificar` requiere intervención manual — Ambiguos y Pagos_multiples
se resuelven automáticamente.

---

## Posición en el pipeline

```
01_lecturas
02_planilla
03_pagos/yape/motor_matching   ← estás aquí
        ↓
04_cobranza                    → consume pagos_yape_tepago.xlsx · blancos_acumulados.xlsx · devoluciones_acumulados.xlsx
        ↓
04b_validacion                 → verifica que el dinero cuadra · usa devoluciones para conciliación
        ↓
05_corte
        ↓
05b_override                   → correcciones de usuarios si las hay
        ↓
06_boletas
        ↓
07_nueva_planilla
```

---

## Estructura de carpetas

```
motor_matching/
├── correcciones/                → temporal · se borra en Ciclo 1
│   └── pendientes.xlsx          → 1 hoja: Sin_identificar · tú completas aquí
│
├── outputs/                     → temporal · se borra en Ciclo 1
│   ├── pagos_yape_tepago.xlsx   → resultado final → 04_cobranza
│   └── resumen_validacion.txt   → log del ciclo · su ausencia indica Ciclo 1
│
├── trazabilidad/                → PERMANENTE · nunca se borra · un archivo por mes
│   └── trazabilidad_YYYY_MM.xlsx
│
├── blancos/                     → PERMANENTE · nunca se borra · un solo archivo
│   └── blancos_acumulados.xlsx  → lo lee 04_cobranza para DESCUENTO_YAPE_BLANCO
│
├── devoluciones/                → PERMANENTE · nunca se borra
│   ├── devoluciones_acumulados.xlsx → historial permanente de todos los PAGASTE
│   └── pagos_yape_pagaste.xlsx  → PAGASTE del mes · tú completas · se borra en Ciclo 1
│
├── docs/                        → documentos de diseño · nunca se borran
├── main.py                      → ciclo de matching
└── README.md                    → este archivo
```

**Lee desde shared/ — no tiene inputs propios:**

```
shared/
├── planilla_mes/                → planilla del mes actual con deuda por MZ-LOTE
├── planilla_acumulado/          → planillas históricas · ancla de corte
├── reporte_mes_crudo/           → ReporteTransacciones.xlsx descargado del banco
├── reporte_acumulado_procesado/ → TE PAGÓ históricos · lo genera motor_matching
└── usuarios_id.xlsx
```

---

## Prioridad de identificación

El sistema intenta identificar cada pago TE PAGÓ en este orden:

```
1. Corrección acumulada (ciclo 2+)   → prioridad máxima · ya fue validado manualmente
2. Múltiples en mensaje              → "MzE Lt7, MzP Lt11A" · automático · va a trazabilidad
3. Mensaje simple                    → regex extrae 1 MZ-LOTE · automático
4. Maestro ambiguo                   → elige candidato con menor diff · automático · va a trazabilidad
5. Maestro único                     → 1 solo lote en maestro · automático
P. Sin identificar                   → va a pendientes.xlsx · único caso manual
```

> Ambiguos y Pagos_multiples nunca van a pendientes — el sistema los resuelve solo.

Los PAGASTE se separan automáticamente en `Devoluciones/pagos_yape_pagaste.xlsx`
para que tú los revises y confirmes.

---

## Qué pasa después de identificar

Una vez que el sistema asigna MZ-LOTE a un pago TE PAGÓ, lo valida contra
la planilla del mes y calcula la diferencia:

| Resultado | Condición | Significado |
|-----------|-----------|-------------|
| `exacto` | monto = deuda | el pago cubre exactamente la deuda |
| `exceso` | monto > deuda | pagó de más — diferencia positiva |
| `parcial` | monto < deuda | pagó de menos — diferencia negativa |
| `pendiente` | MZ-LOTE no existe en planilla | va a Sin_identificar para corrección manual |

> La planilla del mes vive en `shared/planilla_mes.xlsx` — la leen tanto
> motor_matching como 04_cobranza desde el mismo lugar.

---

## Ciclos — Ciclo 1 vs Ciclo 2+

El sistema detecta automáticamente en qué ciclo está:

| Ciclo | Señal | Qué hace |
|-------|-------|----------|
| **Ciclo 1** | `Outputs/pagos_yape_tepago.xlsx` NO existe | mes nuevo · limpia Correcciones/ · Outputs/ · pagos_yape_pagaste.xlsx · Trazabilidad/ · Blancos/ · devoluciones_acumulados.xlsx intactos |
| **Ciclo 2+** | `Outputs/pagos_yape_tepago.xlsx` SÍ existe | lee correcciones · actualiza trazabilidad · regenera pendientes con los que faltan |

---

## Flujo mensual

### Inicio de mes

**1. Verificar que los inputs están listos en shared/:**

```
shared/reporte_mes_crudo/           → ReporteTransacciones.xlsx del banco · tú lo copias
shared/planilla_mes/                → planilla del mes con deudas actualizadas
shared/planilla_acumulado/          → planillas históricas · para ancla de corte
shared/usuarios_id.xlsx             → identidad de cada MZ-LOTE
```

**El código lee directo desde shared/ y crear_maestro/ — no copies nada a inputs:**

```
crear_maestro/outputs/maestro_yape.xlsx   → el código lo lee directo desde aquí
```

### Ciclo de matching

**2. Correr el matching:**

```bash
python main.py
```

El sistema corre y genera:
- `Outputs/pagos_yape_tepago.xlsx` — todos los TE PAGÓ con su estado y diferencia
- `Devoluciones/pagos_yape_pagaste.xlsx` — todos los PAGASTE del mes · tú completas
- `Correcciones/pendientes.xlsx` — TE PAGÓ y PAGASTE que no se pudieron identificar

**3. Revisar pendientes:**

Abrir `Correcciones/pendientes.xlsx` — 1 sola hoja `Sin_identificar`:

| TIPO | Qué hacer |
|------|-----------|
| `TE PAGÓ` | Escribir MZ y LOTE · si no sabes → BLANCO en MZ |
| `PAGASTE` | MZ-LOTE si es devolución · CONCEPTO si es gasto JASS · BLANCO si no sabes |

> BLANCO en TE PAGÓ → va directo a `Blancos/blancos_acumulados.xlsx` · se identificará en el futuro para aplicar descuento
> BLANCO en PAGASTE → queda en `devoluciones_acumulados.xlsx` con ESTADO `no aplica` · solo para auditoría · no afecta planilla

**4. Revisar pagos_yape_pagaste.xlsx:**

Abrir `Devoluciones/pagos_yape_pagaste.xlsx` y confirmar cada PAGASTE:
- Sistema acierta → pon 1 en CONFIRMADO
- Sistema falla → corrige MZ-LOTE manual → pon 1
- Gasto JASS → escribe CONCEPTO → pon 1

**5. Repetir hasta pendientes = 0:**

```bash
python main.py   # Ciclo 2
python main.py   # Ciclo 3
# ...hasta que el sistema reporte: Pendientes: 0
```

### Fin de mes

**6. Pasar los outputs a 04_cobranza:**

```
outputs/pagos_yape_tepago.xlsx               → 04_cobranza
blancos/blancos_acumulados.xlsx              → 04_cobranza (vía shared/)
devoluciones/devoluciones_acumulados.xlsx    → 04_cobranza y 04b_validacion
```

**7. Archivar el reporte procesado:**

```
outputs/pagos_yape_tepago.xlsx   → copiar a shared/reporte_acumulado_procesado/
                                    como reporte_tepago_YYYY_MM.xlsx
                                    lo lee crear_origenes para construir maestro
```

---

## Lifecycle — qué se borra y qué permanece

| Archivo / Carpeta | ¿Qué pasa? | Motivo |
|-------------------|------------|--------|
| `Trazabilidad/*.xlsx` | **PERMANECE** | historial mensual permanente |
| `Blancos/blancos_acumulados.xlsx` | **PERMANECE** | acumulado de todos los BLANCO |
| `Devoluciones/devoluciones_acumulados.xlsx` | **PERMANECE** | historial permanente de PAGASTE |
| `Outputs/pagos_yape_tepago.xlsx` | Se borra | ya fue consumido por 04_cobranza |
| `Outputs/resumen_validacion.txt` | Se borra | su ausencia indica Ciclo 1 |
| `Devoluciones/pagos_yape_pagaste.xlsx` | Se borra | historial ya en devoluciones_acumulados |
| `Correcciones/pendientes.xlsx` | Se borra | correcciones ya en trazabilidad |
| `Inputs/` (todas) | Tú reemplazas | copias los archivos del mes nuevo |

---

## pagos_yape_tepago.xlsx

Resultado final del matching. Solo lectura — generado automáticamente. Lo consume 04_cobranza.
Pagos múltiples aparecen como una fila por lote.

| Grupo | Columnas | Origen |
|-------|----------|--------|
| ¿Quién es? | USER_ID · NOMBRE · FUENTE | sistema |
| ¿Qué hizo el banco? | TIPO · ORIGEN · DESTINO · MONTO_PAGO · MONTO_ASIGNADO · MENSAJE · FECHA | banco |
| ¿Dónde vive? | MZ · LOTE | matching |
| ¿Cuánto debe? | DEUDA_TOTAL · DIFERENCIA · ESTADO_PAGO | calculado |
| ¿Cuándo? | CICLO | sistema |

**Reglas:**
- `FUENTE` — cómo se identificó: `mensaje` · `maestro` · `ambiguo` · `múltiple` · `manual`
- `MONTO_ASIGNADO` — solo en pagos múltiples · vacío en pagos normales
- `CICLO` — ciclo 1 = automático · ciclo 2+ = necesitó corrección manual
- Pagos múltiples — USER_ID asignado por el sistema cruzando MZ-LOTE con `usuarios_id.xlsx`

---

## pagos_yape_pagaste.xlsx

Vive en `Devoluciones/`. Se borra en Ciclo 1. Historial permanente en `devoluciones_acumulados.xlsx`.
Tú lo completas y confirmas cada registro.

| Grupo | Columnas | Origen |
|-------|----------|--------|
| ¿Qué hizo el banco? | TIPO · ORIGEN · DESTINO · MONTO_PAGO · MONTO_ASIGNADO · MENSAJE · FECHA | banco — no editar |
| ¿A qué lote? (sistema propone) | MZ_AUTO · LOTE_AUTO | sistema — desde regex del mensaje |
| ¿A qué lote? (tú decides) | MZ · LOTE · CONCEPTO | tú — corriges si falla o pones concepto si es gasto JASS |
| ¿Confirmado? | CONFIRMADO | tú — pon 1 cuando todo está correcto |

**5 casos:**

| Caso | Qué haces |
|------|-----------|
| Devolución simple · sistema acierta | solo pon 1 en CONFIRMADO |
| Devolución simple · sistema falla | corrige MZ-LOTE manual → pon 1 |
| Devolución múltiple | una fila por lote · confirma cada una |
| Gasto JASS | escribe CONCEPTO → pon 1 |
| BLANCO · regresa de pendientes | MZ-LOTE o CONCEPTO ya llenos → pon 1 |

**Regla del código:** si MZ manual tiene valor → usa ese · si está vacío → usa MZ_AUTO

---

## trazabilidad_YYYY_MM.xlsx

Un archivo por mes. Nunca se borra. Se acumula cada ciclo sin borrar filas anteriores.

### Hoja Sin_identificar

| Grupo | Columnas |
|-------|----------|
| ¿Quién es? | USER_ID · NOMBRE |
| ¿Qué hizo el banco? | TIPO · ORIGEN · DESTINO · MONTO · MENSAJE · FECHA |
| ¿Cómo se resolvió? | MZ · LOTE · CONCEPTO · MOTIVO |
| ¿Cuándo? | CICLO · FECHA_CORRECCION |

### Hoja Ambiguos

Una fila por candidato. MZ_FINAL y LOTE_FINAL se repiten en todas las filas del mismo origen.

| Grupo | Columnas |
|-------|----------|
| ¿Quién es? | USER_ID · NOMBRE |
| ¿Qué hizo el banco? | ORIGEN · MONTO · MENSAJE · FECHA |
| Candidatos evaluados | MZ_CAND · LOTE_CAND · NOMBRE_CAND · DEUDA_CAND · DIFF_CAND |
| Resultado final | ELEGIDO · MZ_FINAL · LOTE_FINAL |
| ¿Cuándo? | CICLO · FECHA_CORRECCION |

### Hoja Pagos_multiples

Una fila por lote. ORIGEN y MONTO_TOTAL se repiten por cada lote del mismo pago.

| Grupo | Columnas |
|-------|----------|
| ¿Quién es? | USER_ID · NOMBRE |
| ¿Qué hizo el banco? | ORIGEN · MONTO_TOTAL · MENSAJE · FECHA |
| Asignación por lote | MZ · LOTE · DEUDA · MONTO_ASIGNADO · DIFF |
| ¿Cuándo? | CICLO · FECHA_CORRECCION |

---

## blancos_acumulados.xlsx

Archivo permanente. Acumula todos los TE PAGÓ sin dueño identificado.
Va directo aquí cuando el usuario escribe BLANCO en pendientes — sin archivo intermedio.

| Grupo | Columnas |
|-------|----------|
| ¿De qué mes? | MES |
| ¿Quién es? | USER_ID · NOMBRE — vacíos hasta identificar |
| ¿Qué hizo el banco? | TIPO · ORIGEN · DESTINO · MONTO · MENSAJE · FECHA |
| ¿Cómo se identificó? | MZ · LOTE · MOTIVO — vacíos hasta identificar |
| ¿Ya se aplicó? | ESTADO |

| ESTADO | Significado |
|--------|-------------|
| `pendiente` | sin dueño · USER_ID y MZ-LOTE vacíos |
| `identificado` | dueño encontrado · 04_cobranza aplica descuento |
| `aplicado` | descuento ya aplicado en planilla |

---

## devoluciones_acumulados.xlsx

Archivo permanente. Acumula todos los PAGASTE del sistema.
Lo leen 04_cobranza para descuentos y 04b_validacion para conciliación.

| Grupo | Columnas |
|-------|----------|
| ¿De qué mes? | MES |
| ¿Quién es? | USER_ID · NOMBRE |
| ¿Qué hizo el banco? | TIPO · ORIGEN · DESTINO · MONTO_PAGO · MONTO_ASIGNADO · MENSAJE · FECHA |
| ¿A qué lote? | MZ · LOTE · CONCEPTO |
| ¿Ya se aplicó? | ESTADO |

| ESTADO | Significado | Quién lo usa |
|--------|-------------|--------------|
| `pendiente` | devolución a usuario · aún no descontada | 04_cobranza aplica descuento |
| `aplicado` | descuento ya aplicado | historial |
| `no aplica` | gasto JASS · no genera descuento | 04b_validacion lo suma como egreso en conciliación |

---

## Errores comunes

**"No se encontró maestro_yape.xlsx"**
→ Verificar que corriste `crear_maestro` y que el archivo existe en `crear_maestro/outputs/`.

**"No hay archivo en Planilla_anterior"**
→ El sistema corre sin ancla de corte — incluirá todos los pagos sin filtro de fecha.

**"MZ-LOTE no existe en planilla"**
→ El MZ-LOTE que escribiste en pendientes.xlsx no existe en la planilla del mes.

**"Ningún archivo válido en Reporte_mes"**
→ Verificar que `shared/reporte_mes_crudo/` tiene el archivo del banco con columna "Tipo de Transacción".

---

## Lo que NO hace este módulo

- **No valida que el dinero cuadre** — eso lo hace `04b_validacion/`
- **No carga la planilla** — eso lo hace `04_cobranza/cargar_planilla/`
- **No resuelve Ambiguos ni Pagos_multiples manualmente** — se resuelven solos
- **No corrige pagos ya identificados** — eso lo hace `05b_override/`

---

## Antes de codificar — verificar inputs

Checklist completo antes de correr `python main.py` por primera vez.
Si algún archivo falta o tiene columnas distintas → el código fallará.

### ¿De dónde viene cada archivo y dónde va?

**Tú copias manualmente cada mes:**

| Archivo | Lo produces / descargas | Lo copias a |
|---------|------------------------|-------------|
| `ReporteTransacciones.xlsx` | banco — tú lo descargas | `shared/reporte_mes_crudo/` |

**El código lee directo — no copias nada más:**

| Archivo | Vive en | El código lo lee desde |
|---------|---------|----------------------|
| `maestro_yape.xlsx` | `crear_maestro/outputs/` | directo desde ahí |
| `planilla_mes/` | `shared/` | `shared/planilla_mes/` |
| `planilla_acumulado/` | `shared/` | `shared/planilla_acumulado/` |
| `usuarios_id.xlsx` | `shared/` | `shared/usuarios_id.xlsx` |

---

### Columnas requeridas — verifica antes de correr

**ReporteTransacciones.xlsx — banco**
```
✓ Tipo de Transacción   → valores exactos: "TE PAGÓ" · "PAGASTE"
✓ Origen                → nombre del pagador
✓ Destino               → nombre del receptor
✓ Monto                 → número · sin símbolo S/
✓ Mensaje               → texto libre · puede estar vacío
✓ Fecha de operación    → formato: DD/MM/YYYY HH:MM:SS
```

**maestro_yape.xlsx — desde crear_maestro/outputs/**
```
✓ ORIGEN_1 · ORIGEN_2...  → orígenes válidos del pagador · puede haber varios
✓ MZ                      → manzana
✓ LOTE                    → número de lote
✓ VALIDADO_MANUAL         → si · no
```

**planilla_mes.xlsx — desde shared/**
```
✓ mz                    → manzana
✓ lote o lt             → número de lote
✓ nombre o nombres      → nombre del usuario
✓ total                 → deuda total del mes
```

**planilla_anterior.xlsx — desde shared/ del mes pasado**
```
✓ hoja llamada "Reporte" o que contenga "reporte"
✓ columna fecha         → para calcular el ancla de corte
```

**usuarios_id.xlsx — desde shared/**
```
✓ USER_ID               → identificador único del usuario
✓ NOMBRE                → nombre completo
✓ MZ                    → manzana
✓ LOTE                  → número de lote
```

---

### Checklist antes de correr

```
[ ] shared/reporte_mes_crudo/ tiene el reporte del banco del mes actual
[ ] shared/planilla_mes/ tiene la planilla del mes con deudas actualizadas
[ ] shared/planilla_acumulado/ tiene las planillas históricas
[ ] shared/usuarios_id.xlsx existe y tiene USER_ID · NOMBRE · MZ · LT · LT2
[ ] crear_maestro/outputs/maestro_yape.xlsx existe y tiene ORIGEN_1 · MZ · LOTE
[ ] Columnas del banco verificadas — especialmente "Tipo de Transacción"
```

---

- Python 3.10 o superior
- Librerías: `pandas`, `openpyxl`

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
4.  Cómo piensa el sistema — prioridades
5.  Qué calcula después de identificar
6.  Cómo detecta el contexto — ciclos
7.  Cómo se usa mes a mes — flujo
8.  Qué se borra y qué permanece — lifecycle
9.  Los archivos que produce — uno por uno con sus columnas
10. Si algo falla — errores comunes
11. Sus límites — lo que NO hace
12. Antes de codificar — verificar inputs y paths
13. Cómo instalarlo — requisitos
```

Todos los README del sistema siguen este mismo orden.
Si encuentras algo fuera de orden — corrígelo para mantener el estándar.
