# Manual de uso — 3_boletas/enriquecimiento

**Para quién es este manual:** la persona que prepara las boletas cada mes
— toma las lecturas del operario y genera el archivo listo para imprimir.

---

## Qué hace este módulo

Antes de que llegue el día de cobro, los usuarios deben recibir su boleta
para saber cuánto pagar. Este módulo toma la planilla del mes (con las
lecturas del operario) y genera automáticamente el archivo `data_boletas.xlsx`
que el módulo de boletas convierte en PDFs.

```
planilla_base.xlsx  ←  operario llena lecturas
config_mes.xlsx     ←  tú llenas una vez por mes
                            ↓
                  3_boletas/enriquecimiento
                            ↓
              3_boletas/inputs/data_boletas.xlsx
                            ↓
                       3_boletas
                            ↓
                    RECIBO_XXXXX.pdf × N
```

---

## Antes de correr este módulo

- [ ] El operario ya tomó las lecturas y llenó `planilla_base.xlsx`
- [ ] `5_cobranza` ya corrió el mes anterior → los valores de ARRASTRE,
      CORTE_RECONEXION, CONVENIO, REUNION_FAENA, TECHADO en la planilla
      reflejan lo que quedó pendiente del mes anterior
- [ ] `config_mes.xlsx` está completado con los datos de este mes

---

## Paso 1 — Completar config_mes.xlsx

El archivo está en `inputs/config_mes.xlsx`. Si no existe, créalo corriendo:

```
python crear_config.py
```

Luego completar los campos con los datos del mes:

| Campo | Descripción | Ejemplo |
|---|---|---|
| **PERIODO** | Período que cubre el recibo | `11/03/2026 al 10/04/2026` |
| **FECHA_VENCIMIENTO** | Fecha límite de pago | `2026-05-02` |
| **FECHA_EMISION** | Fecha en que se emite el recibo | `2026-04-26` |
| **LECTURA_ANT_FECHA** | Fecha en que se tomó la lectura anterior | `2026-03-10` |
| **LECTURA_ACT_FECHA** | Fecha en que se tomó la lectura de este mes | `2026-04-10` |
| **FECHA_PAGO** | Día y mes del cobro presencial | `02/05` |
| **HORA_PAGO** | Hora del cobro | `4-6 pm` |
| **LUGAR_PAGO** | Lugar del cobro | `LOCAL DEL PUEBLO` |
| **TELEFONO** | Número Yape | `948 227 636` |
| **NUMERO_RECIBO_INICIO** | Número del primer recibo de este mes | `16900` |

> El número de recibo continúa desde donde quedó el mes anterior.
> Si el mes pasado el último fue 16896, este mes empieza en 16897.

---

## Paso 2 — Correr el módulo

Abrir una terminal en la carpeta `3_boletas/enriquecimiento/` y ejecutar:

```
python main.py
```

Salida esperada:

```
═════════════════════════════════════════════════
  3_boletas/enriquecimiento — Preparación DATA_boletas
═════════════════════════════════════════════════

[1/4] Validando inputs...
[2/4] Cargando config y planilla base...
[3/4] Enriqueciendo 165 registros...
[4/4] Exportando data_boletas.xlsx...

═════════════════════════════════════════════════
  165 recibos preparados — recibos 16900 al 17064
  → data_boletas.xlsx
  Siguiente paso: correr 3_boletas/main.py
═════════════════════════════════════════════════
```

---

## Cómo se calcula cada campo

### Marcación anterior
Viene de la columna `MARC_ANT` en la planilla. El operario la llena con la
lectura del mes pasado.

### M3 y Total mes actual
```
M3             = MARC_ACT − MARC_ANT
Total mes actual = M3 × S/1.00 por m³
```

### Estado del recibo
- **USTED ESTÁ AL DÍA** → no tiene arrastre ni deudas pendientes
  (solo paga el consumo de este mes + mantenimiento)
- **NO ESTÁ AL DÍA** → tiene arrastre, corte y reconexión, convenio,
  multa u otro cargo de meses anteriores

> El Estado se calcula a partir de la planilla (lo que el usuario debía
> antes de este mes), NO de si pagó o no este mes.

### Total e Importe a pagar
```
Total = Total mes actual + MES ANTERIOR + Corte y reconexion
      + Convenio + Mantenimiento + Multa + Cuota directa − DEVOLUCION + AJUSTE
```
Nunca puede ser negativo (si DEVOLUCION > cargos, Total = 0).

---

## Revisar el output

El archivo generado es `3_boletas/inputs/data_boletas.xlsx`, hoja `Data`.

Antes de generar las boletas, verificar al menos una fila para confirmar
que los valores son razonables:
- NUMERO DE RECIBO correlativo
- M3 no negativo
- Total coherente con lo que el usuario acostumbra pagar

---

## Qué hacer después

Una vez generado `data_boletas.xlsx`, correr el módulo de boletas:

```
cd ../3_boletas
python main.py
```

Los PDFs quedarán en `3_boletas/Outputs/`.

---

## Referencia rápida

```
COMPLETAR UNA VEZ POR MES
  inputs/config_mes.xlsx   ← fechas, teléfono, número de recibo inicial

INPUTS (no copiar — el módulo los lee directo)
  5_cobranza/inputs/planilla_base/planilla_base.xlsx
  inputs/config_mes.xlsx

OUTPUT
  3_boletas/inputs/data_boletas.xlsx   ← listo para 3_boletas

SIGUIENTE PASO
  3_boletas/main.py   → genera RECIBO_XXXXX.pdf por usuario
```
