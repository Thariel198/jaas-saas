# blancos_acumulados.xlsx — reglas de negocio

Archivo permanente. Registra todos los pagos BLANCO del sistema — pagos que no pudieron
identificarse durante el matching y quedaron sin dueño. Nunca se borra. Crece mes a mes.

Lo produce `main.py`. Lo consume `cargar_planilla` para aplicar descuentos cuando
se identifica al dueño.

---

## Ubicación

```
motor_matching/
    └── Blancos/
        └── blancos_acumulados.xlsx  ← único archivo · permanente
```

---

## Ciclo de vida

- **Se crea** la primera vez que aparece un pago BLANCO en el sistema
- **Se actualiza** en cada ciclo de cada mes cuando hay nuevos BLANCO en `pendientes.xlsx`
- **Se actualiza también** cuando `apply_overrides.py` identifica el dueño de un BLANCO histórico
- **Nunca se borra** — el Ciclo 1 del mes siguiente NO lo toca
- **No tiene un archivo por mes** — es un único archivo acumulado de todo el historial

---

## Columnas

| Columna | Tipo | Quién lo pone | Descripción |
|---------|------|---------------|-------------|
| MES | texto | sistema | mes de origen del pago — formato YYYY-MM (ej: 2026-01) |
| FECHA_PAGO | datetime | banco | fecha y hora exacta del pago del banco |
| ORIGEN | texto | banco | nombre del pagador tal como viene del Reporte_mes |
| MONTO | decimal | banco | monto del pago |
| MZ | texto | sistema/usuario | vacío cuando es BLANCO · se llena cuando se identifica el dueño |
| LOTE | texto | sistema/usuario | vacío cuando es BLANCO · se llena cuando se identifica el dueño |
| MOTIVO_IDENTIFICACION | texto | usuario | texto libre — cómo se identificó el dueño (ej: "usuario reclamo con comprobante") |
| FECHA_IDENTIFICACION | datetime | sistema | `datetime.now()` cuando apply_overrides.py marcó el pago como identificado |
| ESTADO | texto | sistema | `pendiente` o `identificado` |

---

## Estados

| ESTADO | Significado |
|--------|-------------|
| `pendiente` | el dueño del pago aún no se conoce |
| `identificado` | el dueño fue encontrado · MZ y LOTE están llenos · cargar_planilla puede aplicar el descuento |

---

## Reglas

- Se agrega una fila por cada pago que el usuario marcó como BLANCO en `pendientes.xlsx`
- La clave de unicidad es ORIGEN + FECHA_PAGO — no solo ORIGEN
- Cuando se identifica el dueño el registro NO se elimina — solo cambia ESTADO a `identificado` y se llenan MZ, LOTE, MOTIVO_IDENTIFICACION y FECHA_IDENTIFICACION
- Un pago BLANCO de cualquier mes puede identificarse en cualquier mes futuro — sin límite de tiempo
- Si el mismo pago BLANCO se intenta identificar dos veces, la segunda sobreescribe la primera (no duplica)

---

## Flujo completo de un pago BLANCO

```
Ciclo N del mes actual
    ↓
Usuario escribe BLANCO en pendientes.xlsx
    ↓
main.py agrega fila en blancos_acumulados.xlsx con ESTADO = pendiente
main.py también lo registra en trazabilidad_YYYY_MM.xlsx hoja Sin_identificar
    ↓
Meses después — el dueño aparece y reclama su pago
    ↓
Usuario llena overrides.xlsx con FECHA_PAGO · MZ_NUEVO · LOTE_NUEVO · MOTIVO
    ↓
apply_overrides.py:
    1. Actualiza blancos_acumulados.xlsx → ESTADO = identificado · llena MZ · LOTE · FECHA_IDENTIFICACION
    2. Registra en trazabilidad del mes actual hoja Overrides
    ↓
cargar_planilla lee blancos_acumulados.xlsx
    ↓
Aplica DESCUENTO_YAPE_BLANCO en la planilla del mes actual
```

---

## Relación con cargar_planilla

`cargar_planilla` lee este archivo buscando filas con `ESTADO = identificado` que aún
no hayan sido descontadas. Aplica el monto como `DESCUENTO_YAPE_BLANCO` en la planilla
del mes, reduciendo la deuda del usuario identificado.

Esto compensa el pago que se hizo en un mes anterior pero no se reconoció en su momento.

---

## Lo que NO hace este archivo

- No reemplaza a `trazabilidad_YYYY_MM.xlsx` — la trazabilidad mensual sigue registrando el BLANCO en hoja Sin_identificar
- No modifica pagos_yape.xlsx de meses anteriores — el descuento se aplica en el mes actual via cargar_planilla
- No elimina registros — solo cambia el ESTADO de pendiente a identificado
