# 4_pagos/efectivo — Módulo de pagos en efectivo

## Qué hace

Consolida los registros de cobranza en efectivo levantados en campo por los cobradores
de cada mesa. Compara los registros dentro de una misma mesa para detectar coincidencias
o discrepancias, y entrega una lista limpia de pagos confirmados lista para `5_cobranza`.

## Cuándo se corre

Una vez por mes, después de que todos los cobradores han entregado sus registros y antes
de correr `5_cobranza`. Si quedan discrepancias sin resolver, el módulo puede correrse
de nuevo hasta que `discrepancias.xlsx` desaparezca.

> **Alimenta el ledger `libro_mayor/caja` (Fase 2):** `pagos_efectivo.xlsx` es una **fuente de
> ABONOS**. `libro_mayor/caja/importar_efectivo.py` lo lee post-cierre y genera un abono por
> pago. La clave natural del `ABONO_ID` es `(jass, MESA, COBRADOR, FECHA, MONTO, MZ, LT)`
> — esas columnas deben existir en el output y no cambiar de nombre. Ver el contrato en
> `libro_mayor/caja/README.md`.

---

## Estructura de carpetas

```
4_pagos/efectivo/
├── inputs/
│   ├── mesa_1.xlsx          ← hojas: registro_1, [registro_2], [registro_3]
│   ├── mesa_2.xlsx
│   ├── ...
│   ├── mesa_7.xlsx
│   ├── entregas_hoja.xlsx   ← la tesorera la llena a mano (staging físico)
│   └── entregas.xlsx        ← log append-only (lo escribe entregas_repo, NO se edita a mano)
├── outputs/
│   ├── pagos_efectivo.xlsx       ← resultado limpio → 5_cobranza
│   ├── discrepancias.xlsx        ← temporal, desaparece al resolverse todo
│   ├── reclamos_YYYY-MM.xlsx     ← vista operacional de reclamos del mes
│   ├── arqueo_YYYY-MM.xlsx       ← cuadre papel vs tesorera, por día y cobrador
│   └── verificacion_lotes_YYYY-MM.xlsx  ← temporal, MZ-LT escrito vs boleta emitida
├── trazabilidad/
│   ├── consolidado_YYYY-MM.xlsx        ← todo lo procesado (permanente)
│   ├── incidencias_YYYY-MM.xlsx        ← anomalías del mes (permanente)
│   └── trazabilidad_reclamos.xlsx      ← auditoría histórica de reclamos cerrados
├── backup/
│   └── migracion_YYYY-MM/           ← archivos anteriores al rediseño
├── docs/
│   ├── diagrama_efectivo.html
│   ├── diagrama_reclamos.html
│   ├── diagrama_flujo_arqueo.html
│   ├── diagrama_flujo_verificacion_lotes.html
│   ├── diagrama_verificacion_lotes.html
│   ├── arquitectura_efectivo.html
│   ├── formato_pagos_efectivo.html
│   ├── formato_reclamos.html
│   ├── formato_trazabilidad_reclamos.html
│   ├── formato_entregas.html
│   ├── formato_arqueo.html
│   └── formato_verificacion_lotes.html
├── tests/
├── verificar_lotes.py       ← cruza mesas vs DATA_boletas → verificacion_lotes_YYYY-MM.xlsx (corre ANTES de main.py)
├── main.py
├── reclamos.py
├── entregas_repo.py         ← escritor único de entregas.xlsx (append-only)
├── importar_entregas.py     ← crea entregas_hoja.xlsx y la vuelca al ledger (flujo B)
├── registrar_entrega.py     ← CLI alterno: registrar una entrega directo por consola
├── arqueo.py                ← cruza mesas vs entregas → arqueo_YYYY-MM.xlsx
└── crear_templates.py
```

---

## Formato de inputs — mesa_N.xlsx

Cada archivo representa una mesa física de cobranza. Puede tener 1, 2 o 3 hojas
nombradas `registro_1`, `registro_2`, `registro_3`.

Columnas de cada hoja — schema v3 (todas requeridas salvo COMENTARIO, CONCEPTO y CATEGORIA):

| Columna          | Tipo    | Descripción                              |
|------------------|---------|------------------------------------------|
| COBRADOR         | texto   | Nombre de quien cobra                    |
| FECHA_REGISTRO   | fecha   | Día en que se llenó el registro          |
| MZ               | texto   | Manzana del predio                       |
| LT               | texto   | Lote del predio                          |
| MONTO            | decimal | Monto total cobrado en soles             |
| MONTO_EFECTIVO   | decimal | Parte en billetes/monedas (0 si todo Yape) |
| MONTO_YAPE       | decimal | Parte por Yape (0 si todo efectivo)      |
| FECHA            | fecha   | Fecha del pago (puede diferir del cobro) |
| COMENTARIO       | texto   | Nota libre, opcional — no clasifica nada |
| CONCEPTO         | texto   | Qué es la plata: vacío=agua · tanque · honorario · gasto · comunitario |
| CATEGORIA        | texto   | Qué pasó en mesa (dropdown): vacío=pago normal · reclamo · compromiso · exoneracion · otros |

`CONCEPTO` rutea el dinero (5_cobranza excluye del cálculo de agua todo pago con
CONCEPTO no vacío). `CATEGORIA` marca eventos de mesa y no afecta montos:
`reclamo` lo detecta `4b_reclamos`; `compromiso`/`exoneracion`/`otros` quedan registrados y
filtrables en `pagos_efectivo.xlsx`. Contrato visual: `docs/formato_registro.html`.
Migración v2→v3: `migrar_formato_v3.py` (backup en `backup/migracion_2026-07/`).

---

## Reglas de negocio

### Cross-check dentro de la mesa

El cross-check ocurre **dentro del mismo archivo** (misma mesa). Mesas distintas tienen
registros distintos por diseño: cada mesa atiende una zona diferente.

| Situación | Estado resultante | Acción |
|---|---|---|
| 1 sola hoja en el archivo | `solo_un_cobrador` | Se acepta como verdad sin comparar |
| 2-3 hojas y todas coinciden en MZ+LT+MONTO | `confirmado` | Pasa a `pagos_efectivo.xlsx` |
| 2-3 hojas y hay diferencias | `discrepancia` | Va a `discrepancias.xlsx` para revisión |
| 2 de 3 hojas coinciden (mayoría) | `mayoria_aplicada` | Mayoría pasa, minoría se traza |

### Regla de mayoría (2 de 3 cobradores)

Si una mesa tiene 3 hojas y 2 coinciden pero 1 difiere:
- La fila de la mayoría pasa a `pagos_efectivo.xlsx` con estado `mayoria_aplicada`
- La fila de la minoría va a `trazabilidad/incidencias_YYYY-MM.xlsx`
- No bloquea el proceso

### Pago en múltiples mesas

Si el mismo `MZ+LT` aparece en más de una mesa (usuario que pagó en dos lugares):
- Ambas filas se marcan como `pago_multi_mesa`
- Se registra en `trazabilidad/incidencias_YYYY-MM.xlsx`
- **No** pasan automáticamente a `pagos_efectivo.xlsx` — requieren revisión manual

### Discrepancias sin resolver

Si al terminar quedan filas en `discrepancias.xlsx`:
- El archivo permanece en `outputs/`
- El módulo termina con advertencia, no con error
- Correr de nuevo después de editar `discrepancias.xlsx` (columna RESOLUCION)
- Cuando todas las discrepancias están resueltas, `discrepancias.xlsx` se elimina automáticamente

---

## Verificación de lotes

El cross-check dentro de la mesa (arriba) y el arqueo de caja verifican **cobrador vs
cobrador** y **papel vs plata entregada** — ninguno de los dos cuestiona si el `MZ-LT` que
el cobrador escribió es el predio correcto. Un error de una letra o un dígito acredita el
pago a otro vecino sin que nada lo detecte: el monto entregado cuadra, y si la mesa tiene
un solo cobrador no hay con qué comparar.

`verificar_lotes.py` corre **antes** de `main.py` y cruza cada fila de las mesas contra
`3_boletas/inputs/DATA_boletas.xlsx` (no la planilla — su `TOTAL_A_PAGAR` es fórmula Excel
y pandas la lee `NaN`). Cuatro capas, en orden:

1. **Cuadre** — ¿`MONTO` es alguna combinación de los cargos de ESE lote? (`CUADRA` /
   `NO CUADRA`)
2. **Fuerza de evidencia** — ¿cuántos lotes del pueblo deben exactamente ese importe?
   (`ALTA` / `MEDIA` / `BAJA`) — un `S/8` que comparten 101 vecinos no prueba nada, y el
   reporte lo dice en vez de dar un OK vacío.
3. **Vecindad de confusión** (solo si `NO CUADRA`) — de los lotes que sí explican el
   monto, ¿cuáles están a un error de tipeo/OCR del `MZ-LT` escrito? (`U/V/W`, `G/O/Q/C`,
   dígitos que se confunden, `X↔X1`, transposición, etc.)
4. **Filtro de realidad** — el candidato solo se propone cuando queda exactamente uno y no
   está ya pagado en esta misma corrida.

Filas con `CONCEPTO` no vacío (tanque, honorario, gasto, comunitario) o `MONTO = 0` se
omiten (`EVIDENCIA = OMITIDO`) — esa plata ya no es deuda de agua y nunca va a cuadrar
contra `DATA_boletas`.

Solo reporta — **nunca escribe en `mesa_N.xlsx`** ("manual — sagrado"). Avisa, no bloquea,
igual que `discrepancias.xlsx`. Guard previo: exige `shared/ciclo_activo.json` declarado y
que la `FECHA` del pago caiga en la ventana de emisión→vencimiento de la boleta — si no,
para con el mensaje de qué módulo anterior falta correr.

Contratos: `docs/diagrama_flujo_verificacion_lotes.html`, `docs/diagrama_verificacion_lotes.html`,
`docs/formato_verificacion_lotes.html`. Decisión de diseño completa, con la evidencia medida
contra datos reales: `docs/decisiones/verificacion_lotes_efectivo.md`.

```bash
python verificar_lotes.py --mes 2026-08
```

### Señal de alerta

Si más del 50% de las filas cae en evidencia `BAJA`, el monto dejó de discriminar en esta
JASS — la respuesta no es afinar el algoritmo, es que falta la columna `NOMBRE` en la hoja
de papel del cobrador.

---

## Flujo paso a paso

```bash
# 1. Asegurarse de que los archivos mesa_N.xlsx están en inputs/
#    (cada archivo puede tener 1, 2 o 3 hojas)

# 2. Verificar que el MZ-LT escrito es el correcto (antes de consolidar)
python verificar_lotes.py --mes 2026-08

# 3. Revisar outputs/verificacion_lotes_YYYY-MM.xlsx si existe
#    Corregir mesa_N.xlsx a mano donde EVIDENCIA=NO CUADRA lo indique

# 4. Correr el módulo
python main.py

# 5. Revisar outputs/discrepancias.xlsx si existe
#    Llenar columna RESOLUCION en cada fila (acepta / corrige)

# 6. Volver a correr si había discrepancias
python main.py

# 7. Cuando no hay discrepancias, outputs/pagos_efectivo.xlsx está listo
#    → pasar a 5_cobranza
```

---

## Tabla de lifecycle

| Archivo | Tipo | Cuándo se crea | Cuándo desaparece |
|---|---|---|---|
| `outputs/pagos_efectivo.xlsx` | permanente | cada corrida | nunca (se sobreescribe) |
| `outputs/verificacion_lotes_YYYY-MM.xlsx` | temporal | cada corrida de `verificar_lotes.py` | cuando todas las filas `NO CUADRA` tienen `RESOLUCION` |
| `outputs/discrepancias.xlsx` | temporal | si hay discrepancias | cuando todas se resuelven |
| `outputs/reclamos_YYYY-MM.xlsx` | mensual | cada corrida de `reclamos.py` | nunca (se sobreescribe) |
| `trazabilidad/consolidado_YYYY-MM.xlsx` | permanente | cada corrida | nunca |
| `trazabilidad/incidencias_YYYY-MM.xlsx` | permanente | si hay anomalías | nunca |
| `trazabilidad/trazabilidad_reclamos.xlsx` | permanente | primera vez que hay reclamos cerrados | nunca — solo crece |
| `inputs/mesa_N.xlsx` | manual — sagrado | el cobrador lo llena | nunca se borra sin backup |
| `inputs/entregas_hoja.xlsx` | manual — staging | 1a corrida de `importar_entregas.py` | nunca se borra sin backup |
| `inputs/entregas.xlsx` | permanente — append-only | primera entrega importada | nunca — solo crece, nunca se edita |
| `outputs/arqueo_YYYY-MM.xlsx` | mensual | cada corrida de `arqueo.py` | nunca (se regenera) |

---

---

## Reclamos

Los cobradores marcan `CATEGORIA = reclamo` (dropdown, schema v3) cuando un usuario
cuestiona su cobro; el detalle va en `COMENTARIO` (texto libre). Fallback de transición
jul-2026: filas con CATEGORIA vacía cuyo COMENTARIO contiene "reclamo" también se
detectan (con warning en el log) — retirar en agosto. El módulo `4b_reclamos` convierte
esas marcas en un sistema de seguimiento auditable.

### Cómo funciona

```bash
# Después de correr main.py (pagos_efectivo.xlsx actualizado):
python reclamos.py --mes 2026-06
```

1. **Detecta** todas las filas de `pagos_efectivo.xlsx` con `CATEGORIA = reclamo`
   (fallback jul-2026: CATEGORIA vacía + `COMENTARIO` contiene "reclamo", sin distinción
   de mayúsculas/minúsculas).
2. **Preserva** el trabajo manual del supervisor: si ya existe `reclamos_YYYY-MM.xlsx`, copia
   las columnas `RECLAMO`, `ESTADO` y `FECHA_RESOLUCION` usando la clave `(MESA, MZ, LT, FECHA_COBRO)`.
3. **Arrastra** reclamos PENDIENTE o EN_REVISION del mes anterior que no tienen match en el
   mes actual (el usuario sigue sin resolver su reclamo).
4. **Cierra** filas con `ESTADO = RESUELTO` o `RECHAZADO`: las mueve a
   `trazabilidad/trazabilidad_reclamos.xlsx` y las elimina de la vista del mes.

### Estados de reclamo

| Estado | Significado | Siguiente corrida |
|--------|-------------|-------------------|
| `PENDIENTE` | Detectado, sin gestionar | Se mantiene en vista; se arrastra si no hay match |
| `EN_REVISION` | El supervisor está gestionando | Se mantiene en vista; se arrastra si no hay match |
| `RESUELTO` | Cerrado favorablemente | Mueve a trazabilidad, sale de vista |
| `RECHAZADO` | Evaluado y descartado | Mueve a trazabilidad, sale de vista |

El supervisor llena `RECLAMO` (texto libre), `ESTADO` (dropdown) y `FECHA_RESOLUCION`
directamente en `outputs/reclamos_YYYY-MM.xlsx`. Ese trabajo se preserva en cada re-corrida.

### Señales de alerta

- Vista del mes con >50 filas PENDIENTE → revisar si el filtro captura demasiados falsos positivos.
- Trazabilidad crece pero vista del mes no baja → ningún reclamo se está cerrando; problema de proceso.

---

## Arqueo de caja

Valida que el dinero que cada cobrador **anotó en su mesa** (papel) sea igual al que
**la tesorera recibió** físicamente — efectivo y yape, por día y por cobrador. Responde
"¿cuadra la caja?" y, cuando no, "¿a quién le reclamo cuánto?".

### Las dos fuentes

```
mesa_1..7.xlsx (papel del cobrador)      entregas.xlsx (declaración de la tesorera)
  Σ MONTO_EFECTIVO / MONTO_YAPE            Σ EFECTIVO / YAPE
  por (FECHA, COBRADOR)                    por (FECHA, COBRADOR)
            └──────────► arqueo.py ◄───────────┘
                    arqueo_YYYY-MM.xlsx
```

- **FECHA** = el día que pagó el usuario (no el día que se registró). Misma clave en las dos fuentes.
- **Efectivo esperado** = Σ `MONTO_EFECTIVO` bruto de las mesas. No se netea nada — la plata llega
  primero a la tesorera; nadie gasta antes de entregar (`CONCEPTO=gasto` no interviene en el cuadre).
- **Yape**: el usuario yapea a la cuenta del cobrador, el cobrador reenvía a la cuenta de la tesorera,
  la tesorera registra el monto. Se cuadra por cobrador → `DIF_YAPE ≠ 0` dice a quién reclamar.

### Registrar entregas — flujo B (Excel físico + import)

Hay dos archivos, con roles distintos:

| Archivo | Quién escribe | Rol |
|---|---|---|
| `inputs/entregas_hoja.xlsx` | la tesorera, a mano | staging físico — llena una fila por entrega |
| `inputs/entregas.xlsx` | solo `entregas_repo.py` | ledger append-only — nunca se edita a mano |

La tesorera llena `entregas_hoja.xlsx` como si fuera una mesa; `importar_entregas.py` vuelca
cada fila al ledger append-only. El Excel es la superficie de captura; el ledger es la fuente
de verdad event-sourced (el arqueo pliega sus eventos → `RECIBIDO = Σ filas`). Mismo patrón que
`shared/seguimiento_repo.py`.

```bash
python importar_entregas.py    # 1a vez: crea entregas_hoja.xlsx. Luego: importa lo pendiente.
```

- **Fila 3** = ejemplo guía (se ignora). Datos desde la fila 4.
- **Idempotente** por `(SOURCE, AUDIT_REF = HOJA-fecha-cobrador-rN)` — re-importar no duplica.
  La columna `IMPORTADO` marca lo ya volcado; re-correr es seguro aunque no se guarde la marca.
- **Segunda entrega el mismo día** = otra fila; `arqueo.py` SUMA las filas de cada `(FECHA, COBRADOR)`.
- **Corrección** = fila con `MOTIVO` lleno; ahí `EFECTIVO`/`YAPE` son el **total correcto** y el
  import appendea el **delta** (nunca edita el ledger). La tesorera no hace la resta.

Alterno rápido por consola (una entrega suelta, sin abrir Excel): `python registrar_entrega.py --fecha … --cobrador … --efectivo … --yape …`.

Este diseño es el que escala a SaaS: mañana un endpoint POST reemplaza al import y llama a la
misma función `registrar_entrega()` — cambia quién la invoca, no el ledger ni la lógica.

### Correr el arqueo

```bash
python arqueo.py --mes 2026-07
```

Lee las mesas crudas (no `pagos_efectivo.xlsx`, cuyo dedupe cross-cobrador distorsionaría el
total físico por cobrador) y `entregas.xlsx`, agrupa por `(FECHA, COBRADOR)` y compara.

### Estados del cuadre

| ESTADO | Significado | Acción |
|--------|-------------|--------|
| `CUADRA` | `DIF_EFECTIVO = 0` y `DIF_YAPE = 0` | ninguna |
| `DESCUADRE` | hay papel y entrega, alguna DIF ≠ 0 | reclamar al cobrador el monto de la DIF |
| `SIN_ENTREGA` | hay papel, la tesorera no declaró ese día/cobrador | que la tesorera registre la entrega |
| `SIN_MESA` | la tesorera declaró, no hay mesa ese día | el cobrador no registró lo que entregó |

`DIF = RECIBIDO − PAPEL`: negativo = falta (entregó menos de lo anotado), positivo = sobra.

Contratos: `docs/formato_entregas.html`, `docs/formato_arqueo.html`, `docs/diagrama_flujo_arqueo.html`.

---

## Lo que este módulo NO hace

- No calcula la deuda ni valida su cálculo (eso es `2_planilla` + `5_cobranza`) —
  `verificar_lotes.py` solo usa el monto de la boleta ya emitida como evidencia de qué
  lote pagó, no recalcula cuánto debía
- No cruza datos con Yape (eso lo hace `4_pagos/yape/`)
- No decide si un usuario está al día (eso lo hace `5_cobranza`)
- No borra los archivos de inputs — son trabajo manual sagrado
- No fusiona mesas distintas automáticamente (cada mesa es independiente por diseño)

---

## Señal de alerta

Si más del 30% de las filas salen como `solo_un_cobrador` durante 2 meses seguidos,
la metodología de doble registro por mesa no se está aplicando en campo.
Revisar el procedimiento con los cobradores antes del siguiente ciclo.

---

## Errores comunes

| Error | Causa | Solución |
|---|---|---|
| `FileNotFoundError: inputs/mesa_N.xlsx` | No se creó el archivo antes de correr | Correr `crear_templates.py` y llenar el archivo |
| `ValueError: hoja 'registro_1' no encontrada` | El archivo existe pero está vacío o mal nombrado | Verificar nombre de hojas en Excel |
| `pagos_efectivo.xlsx vacío` | Todos los registros quedaron en discrepancias | Resolver `discrepancias.xlsx` y volver a correr |
| `discrepancias.xlsx` no desaparece | Hay filas sin columna RESOLUCION llenada | Llenar RESOLUCION en todas las filas y volver a correr |

---

## Migración junio 2026

Los archivos `registro_01.xlsx … registro_07.xlsx` (diseño anterior) fueron movidos a
`backup/migracion_2026_06/`. Los datos se re-llenaron manualmente en `mesa_1.xlsx … mesa_7.xlsx`.
Este mes solo hay 1 cobrador por mesa → todas las filas saldrán como `solo_un_cobrador`.
