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
├── Inputs/                      → se copian manualmente cada mes
│   ├── Maestro/                 → maestro_yape.xlsx (desde shared/)
│   ├── Planilla_anterior/       → planilla del mes pasado (ancla de corte)
│   ├── Planilla_mes/            → planilla del mes actual con deuda por MZ-LOTE
│   └── Reporte_mes/             → ReporteTransacciones.xlsx descargado del banco
│
├── Correcciones/                → temporal · se borra en Ciclo 1
│   └── pendientes.xlsx          → 1 hoja: Sin_identificar · tú completas aquí
│
├── Outputs/                     → temporal · se borra en Ciclo 1
│   ├── pagos_yape_tepago.xlsx   → resultado final → 04_cobranza
│   └── resumen_validacion.txt   → log del ciclo · su ausencia indica Ciclo 1
│
├── Trazabilidad/                → PERMANENTE · nunca se borra · un archivo por mes
│   └── trazabilidad_YYYY_MM.xlsx
│
├── Blancos/                     → PERMANENTE · nunca se borra · un solo archivo
│   └── blancos_acumulados.xlsx  → lo lee 04_cobranza para DESCUENTO_YAPE_BLANCO
│
├── Devoluciones/                → PERMANENTE · nunca se borra
│   ├── devoluciones_acumulados.xlsx → historial permanente de todos los PAGASTE
│   └── pagos_yape_pagaste.xlsx  → PAGASTE del mes · tú completas · se borra en Ciclo 1
│
├── main.py                      → ciclo de matching
└── README.md                    → este archivo
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

**1. Copiar los inputs del mes nuevo:**

```
Inputs/Maestro/           → maestro_yape.xlsx desde shared/
Inputs/Planilla_anterior/ → planilla del mes que acaba de cerrar
Inputs/Planilla_mes/      → planilla del mes nuevo con deudas actualizadas
Inputs/Reporte_mes/       → ReporteTransacciones.xlsx descargado del banco
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
Outputs/pagos_yape_tepago.xlsx           → 04_cobranza
Blancos/blancos_acumulados.xlsx          → 04_cobranza (vía shared/)
Devoluciones/devoluciones_acumulados.xlsx → 04_cobranza y 04b_validacion
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
→ Verificar que copiaste el archivo a `Inputs/Maestro/` desde `shared/`.

**"No hay archivo en Planilla_anterior"**
→ El sistema corre sin ancla de corte — incluirá todos los pagos sin filtro de fecha.

**"MZ-LOTE no existe en planilla"**
→ El MZ-LOTE que escribiste en pendientes.xlsx no existe en la planilla del mes.

**"Ningún archivo válido en Reporte_mes"**
→ El archivo del banco no tiene la columna "Tipo de Transacción" o no hay filas "TE PAGÓ".

---

## Lo que NO hace este módulo

- **No valida que el dinero cuadre** — eso lo hace `04b_validacion/`
- **No carga la planilla** — eso lo hace `04_cobranza/cargar_planilla/`
- **No resuelve Ambiguos ni Pagos_multiples manualmente** — se resuelven solos
- **No corrige pagos ya identificados** — eso lo hace `05b_override/`

---

## Requisitos

- Python 3.10 o superior
- Librerías: `pandas`, `openpyxl`

```bash
pip install pandas openpyxl
```

---

## Cómo leer este README

Este README sigue siempre el mismo orden narrativo:

```
1. Qué hace el módulo
2. Dónde encaja en el pipeline
3. Dónde viven los archivos
4. Cómo piensa el sistema — prioridades
5. Qué calcula después de identificar
6. Cómo detecta el contexto — ciclos
7. Cómo se usa mes a mes — flujo
8. Qué se borra y qué permanece — lifecycle
9. Los archivos que produce — uno por uno con sus columnas
10. Si algo falla — errores comunes
11. Sus límites — lo que NO hace
12. Cómo instalarlo — requisitos
```

Todos los README del sistema siguen este mismo orden.
Si encuentras algo fuera de orden — corrígelo para mantener el estándar.
