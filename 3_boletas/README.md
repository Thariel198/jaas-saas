# 3_boletas — Generación de boletas y constancias en papel

Convierte los datos del ciclo en **documentos imprimibles** para la mesa de cobro:
la boleta mensual de cada predio y las constancias especiales (deuda sin servicio,
medidor cancelado). Es el último paso antes de que el socio reciba papel.

> **Orden real del negocio:** el socio recibe la boleta **antes** de pagar, por eso
> `3_boletas` va antes que `4_pagos` en el pipeline.

---

## Qué hace

Toma `DATA_boletas.xlsx` (una fila por predio, ya enriquecida con consumo, deuda y
datos del ciclo) y una plantilla Word, y produce un `.docx` + `.pdf` por predio más
un consolidado para imprimir en lote. Cubre además dos universos que no entran en el
ciclo normal: predios con deuda pero sin lectura, y constancias de medidor cancelado.

## Cuándo se corre

Después de `2_planilla` (que produce la planilla del mes) y antes de `4_pagos`.
El enriquecimiento lee la planilla del mes; el render corre cuando `DATA_boletas.xlsx`
está listo y validado.

---

## Estructura

```
3_boletas/
├── enriquecimiento/            ← 3.1 — sub-módulo: arma DATA_boletas.xlsx
│   ├── main.py                    lee planilla_YYYY-MM + config_mes.xlsx → DATA_boletas.xlsx
│   ├── crear_config.py            genera la plantilla config_mes.xlsx
│   ├── validar_enriquecimiento.py chequeos post-enriquecimiento
│   └── inputs/config_mes.xlsx     datos del ciclo (periodo, fechas, serie de recibos)
├── main.py                     ← 3.2 — render de la boleta mensual (DOCX→PDF)
├── boletas_sin_servicio.py     ← constancias de DEUDA para predios sin lectura
├── recibos_medidor_pagado.py   ← constancias de medidor CANCELADO (serie MP-)
├── correcciones.py             ← reimpresión puntual por (MZ, LT)
├── validar_boletas.py          ← validación del lote generado
├── Inputs/                     ← DATA_boletas.xlsx, PLANTILLA_boletas.docx, imágenes (logo, QR, caritas)
├── Outputs/                    ← RECIBO_*.docx/.pdf + CONSOLIDADO*.pdf + Imagenes/
│   ├── Correcciones/              reimpresiones acumuladas del mes
│   ├── Medidor_pagado/            constancias MP-
│   └── Sin_servicio/              boletas de deuda sin servicio
└── docs/
    └── formato_recibo_medidor_pagado.html   ← contrato visual de la constancia MP-
```

> Nota: `main.py` y `correcciones.py` usan `Inputs/`/`Outputs/` (mayúscula, Windows es
> case-insensitive); `enriquecimiento/` usa `inputs/`/`outputs/`.

---

## Los cuatro productos de este módulo

| Script | Produce | Universo | Numeración |
|---|---|---|---|
| `enriquecimiento/main.py` + `main.py` | **Boleta mensual** (DOCX+PDF) | todos los predios con lectura en `DATA_boletas.xlsx` | serie del ciclo (correlativa, arranca en `NUMERO_RECIBO_INICIO`) |
| `boletas_sin_servicio.py` | **Boleta de deuda sin servicio** | saldo > 0 en el ledger pero sin lectura del ciclo (deuda de pueblo + agua/corte de arrastre) | continúa la serie del ciclo (`max + 1`) |
| `recibos_medidor_pagado.py` | **Constancia CANCELADO** | `SALDO ACTUAL == 0 AND DEUDA > 0` de convenio de **medidor** | serie propia `MP-001…` (no consume la del ciclo) |
| `correcciones.py` | **Reimpresión** de boletas puntuales | pares `(MZ, LT)` editados a mano en el script | reusa el número original |

---

## Reglas de negocio

- **`DATA_boletas.xlsx` es el input canónico del render** — una fila por predio.
  Lo arma `enriquecimiento/main.py` desde `2_planilla/outputs/planilla_YYYY-MM.xlsx`
  + `config_mes.xlsx` (periodo, fechas de lectura/emisión/vencimiento, lugar y hora de
  pago, teléfono, número de recibo inicial).
- **La boleta refleja "esto debes ahora" de UN mes** — es documento de cobro, no
  historial. El historial multi-mes es el **Extracto de cuenta** (⑫), que vive en
  `libro_mayor/estado_cuenta`, no acá.
- **`boletas_sin_servicio.py` factura deuda real sin lectura** — predios con
  MULTA/ACUERDOS/CONVENIO en el ledger que no reciben boleta del ciclo (sin servicio)
  pero cuya deuda la mesa necesita cobrar. Suma agua vieja/corte del
  `arrastre_consolidado` si existe.
  **Excluye** lotes viejos de reasignaciones pendientes de génesis (ej. B-29→B-20,
  C-45→C-43): su deuda se factura cuando los cargos se muevan al lote nuevo.
- **`recibos_medidor_pagado.py`** lee `shared/vista_seguimiento_pueblo.xlsx` (hoja
  `CONVENIO_HISTORIAL`, en vivo); solo medidor cancelado (inscripción/multa/acuerdos
  no entran ni bloquean); PDF directo con PyMuPDF (sin Word/COM), regenerable siempre.
- **Reimpresión no re-numera** — `correcciones.py` reusa el número de recibo original,
  borra los `.docx/.pdf/.jpg` viejos del predio y acumula la copia en `Outputs/Correcciones/`.

---

## Flujo paso a paso

```bash
# 3.1 — enriquecer: planilla del mes → DATA_boletas.xlsx
cd 3_boletas/enriquecimiento
python crear_config.py         # (una vez por mes) genera config_mes.xlsx para llenar
python main.py                 # produce ../Inputs/DATA_boletas.xlsx
python validar_enriquecimiento.py

# 3.2 — render del lote mensual
cd 3_boletas
python main.py                 # RECIBO_*.docx/.pdf + Outputs/CONSOLIDADO.pdf
python validar_boletas.py

# universos especiales (cuando aplican)
python boletas_sin_servicio.py
python recibos_medidor_pagado.py

# reimpresión puntual (editar RECIBOS_A_REIMPRIMIR dentro del script)
python correcciones.py
```

---

## Acoplamiento con otros módulos

```
2_planilla/outputs/planilla_YYYY-MM.xlsx   ──►  enriquecimiento/main.py
shared/seguimiento_pueblo.xlsx (repo)       ──►  boletas_sin_servicio.py
shared/vista_seguimiento_pueblo.xlsx        ──►  recibos_medidor_pagado.py
5_cobranza/outputs/arrastre_consolidado_*   ──►  boletas_sin_servicio.py (agua/corte vieja)
```

---

## Pendiente Fase 2 — integración con el ledger

Cuando `libro_mayor/estado_cuenta` esté implementado, la boleta deja de leer el
descuento de la planilla y pasa a:

1. **Leer el `SALDO` derivado** de `estado_cuenta` (no una columna de descuento manual).
2. Mostrar un bloque **"pago reconocido este mes"** (las aplicaciones del ledger),
   con el mismo diseño que `estado_cuenta/docs/formato_vista_estado_cuenta.html`.
3. La plantilla `PLANTILLA_boletas.docx` **no tiene hoy** línea de descuento/reconocimiento
   — hay que agregarla.

**Arquitectura de render (decidida 2026-07-13):** `3_boletas` **no se renombra a
`3_impresor`** y **sigue donde está**. Ningún módulo de negocio "imprime": cada dueño
de datos arma sus filas (`3_boletas` la boleta, `estado_cuenta` el extracto) y un
servicio **stateless** `render(plantilla, filas) → PDF` (hoy candidato a
`shared/utils_render.py`, mañana su propio contenedor) hace la conversión. El extracto
de cuenta vive con su dueño de datos (`estado_cuenta`), no acá.

---

## Lo que NO hace

- **No calcula la deuda** — la recibe ya calculada en `DATA_boletas.xlsx` (hoy de la
  planilla; Fase 2, del ledger).
- **No registra pagos** — eso es `4_pagos`.
- **No es el historial del socio** — el extracto multi-mes es de `estado_cuenta` (⑫).

## Errores comunes

- Correr `main.py` (render) sin haber corrido `enriquecimiento/main.py` → `DATA_boletas.xlsx`
  viejo o inexistente.
- `docx2pdf` depende de Word/COM (Windows) — `recibos_medidor_pagado.py` evita esa
  dependencia usando PyMuPDF directo.
- Constantes de fecha/periodo hardcodeadas en `main.py` (LECTURA_ANTERIOR, PERIODO,
  etc.): el flujo correcto es que vengan de `config_mes.xlsx` vía enriquecimiento.
