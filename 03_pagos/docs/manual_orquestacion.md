# Manual de Orquestación — 03_pagos

## ¿Qué hace este módulo?

`03_pagos` procesa dos tipos de pago de manera independiente:
- **Yape** (`TE PAGÓ` y `PAGASTE`): pipeline de 4 pasos con matching automático.
- **Efectivo**: reconciliación entre dos registros de cobro (mía vs. amiga).

El archivo `03_pagos/main.py` ejecuta ambos pipelines en el orden correcto.

---

## Antes de ejecutar: llenar los inputs

### Inputs Yape

| Archivo | Ubicación | Descripción |
|---|---|---|
| Reportes mensuales | `shared/reporte_mes_crudo/` | Descarga del banco/Yape del mes actual |
| Reportes acumulados | `shared/reporte_acumulado_procesado/` | Histórico procesado de meses anteriores |
| Planilla del mes | `shared/planilla_mes/` | Planilla de deudas generada en `02_planilla` |
| Usuarios ID | `shared/usuarios_id.xlsx` | Tabla maestra de MZ/LT → ID de usuario |

> Si hay correcciones manuales del ciclo anterior, verificar que
> `yape/motor_matching/correcciones/` esté actualizado antes de correr.

### Inputs Efectivo

| Archivo | Ubicación | Descripción |
|---|---|---|
| `registro_mia.xlsx` | `efectivo/inputs/` | Cobros registrados por la tesorera |
| `registro_amiga.xlsx` | `efectivo/inputs/` | Cobros registrados por la cobradora |

> Si los templates no existen, ejecutar una sola vez:
> `python efectivo/crear_templates.py`

---

## Flujo de ejecución

```
shared/reporte_mes_crudo/          shared/planilla_mes/
shared/reporte_acumulado_procesado/      ↓
shared/usuarios_id.xlsx                  │
        │                                │
        ▼                                │
[1] crear_origenes ──────────────────────┤
        │ origenes_yape.xlsx             │
        ▼                                │
[2] crear_maestro                        │
        │ maestro_yape.xlsx              │
        ▼                                ▼
[3] motor_matching ─────────────────────►
        │ pagos_yape_tepago.xlsx
        │ pendientes.xlsx (si hay)
        ▼
[4] validacion
        │ saldos_pendientes.xlsx
        │ resumen_validacion.txt
        │
efectivo/inputs/
        │
        ▼
[5] efectivo
        │ pagos_efectivo.xlsx
        │ discrepancias.xlsx (si hay)
        ▼
   04_cobranza/inputs/
```

---

## Cómo ejecutar

```bash
# Desde la raíz del proyecto
python 03_pagos/main.py

# O desde dentro de 03_pagos/
python main.py
```

---

## Outputs generados

### Yape
| Archivo | Ubicación | Destino |
|---|---|---|
| `pagos_yape_tepago.xlsx` | `yape/motor_matching/outputs/` | `04_cobranza/inputs/pagos_yape/` |
| `saldos_pendientes.xlsx` | `yape/validacion/` | Referencia interna |
| `resumen_validacion.txt` | `yape/validacion/` | Revisión manual |

### Efectivo
| Archivo | Ubicación | Destino |
|---|---|---|
| `pagos_efectivo.xlsx` | `efectivo/outputs/` | `04_cobranza/inputs/pagos_efectivo/` |

---

## Ciclos de corrección

El sistema maneja múltiples ciclos. No es de una sola pasada.

### Ciclo Yape (motor_matching)

```
Ejecutar main.py
    ↓
¿pendientes.xlsx generado?
    ├── NO  → proceso completo, pasar a 04_cobranza
    └── SÍ  → abrir correcciones/, registrar resoluciones manuales
                    ↓
              Volver a ejecutar main.py (ciclo 2, 3, ...)
```

Los ciclos anteriores quedan guardados en `yape/motor_matching/trazabilidad/`.
El motor no duplica registros ya confirmados.

### Ciclo Efectivo

```
Ejecutar main.py
    ↓
¿discrepancias.xlsx generado?
    ├── NO  → proceso completo
    └── SÍ  → abrir discrepancias.xlsx
              escribir "Sí" en columna OK para cada fila resuelta
              completar MONTO_FINAL si corresponde
                    ↓
              Volver a ejecutar main.py (ciclo 2, 3, ...)
```

---

## Módulos que se pueden ejecutar solos

Si solo necesitas re-correr un paso específico sin ejecutar todo el pipeline:

```bash
python yape/construir_maestro/crear_origenes/main.py   # solo origenes
python yape/construir_maestro/crear_maestro/main.py    # solo maestro
python yape/motor_matching/main.py                     # solo matching
python yape/validacion/main.py                         # solo validacion
python efectivo/main.py                                # solo efectivo
```

---

## Qué hacer cuando termina

1. Copiar `pagos_yape_tepago.xlsx` → `04_cobranza/inputs/pagos_yape/`
2. Copiar `pagos_efectivo.xlsx` → `04_cobranza/inputs/pagos_efectivo/`
3. Ejecutar `04_cobranza/main.py`
