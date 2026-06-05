# Manual de uso — 04_cobranza

**Para quién es este manual:** la persona que registra los pagos y cierra la
cobranza cada mes (tesorero o administrador de la JASS).

---

## Antes de empezar

Verificar que estos dos pasos ya estén completos:

- [ ] El módulo `03_pagos` corrió sin errores y los pagos Yape están identificados
- [ ] Los pagos en efectivo están registrados en `pagos_efectivo.xlsx`

Si alguno falta, completarlo primero. Este módulo necesita ambos archivos.

---

## Paso 1 — Preparar la planilla base

La planilla base es el archivo que tú llenas con las lecturas del mes y los
conceptos de deuda de cada usuario.

### 1.1 Crear el archivo si no existe

La primera vez del mes, abrir una terminal en la carpeta `04_cobranza/` y ejecutar:

```
python crear_template.py
```

Esto crea el archivo vacío en `inputs/planilla_base/planilla_base.xlsx` con todas
las columnas correctas y una fila de ejemplo.

> Si el archivo ya existe de un mes anterior, **no** correr `crear_template.py` —
> abrirlo directamente y actualizar los datos.

### 1.2 Llenar la planilla

Abrir `inputs/planilla_base/planilla_base.xlsx`. Llenar desde la fila 4 en
adelante — las filas 2 y 3 son guía, eliminarlas antes de correr `main.py`.

**Columna por columna:**

| Columna | Qué poner | Ejemplo |
|---|---|---|
| MZ | Letra o código de la manzana | `A`, `B`, `G1` |
| LT | Número o código del lote | `1`, `7`, `11A` |
| NOMBRE | Nombre completo del usuario | `ROSA AURORA ROCA` |
| MARC_ANT | Lectura del medidor del mes pasado | `2678` |
| MARC_ACT | Lectura del medidor de este mes | `2703` |
| ARRASTRE | Lo que debía del mes anterior (0 si pagó todo) | `33` o `0` |
| CONVENIO | Cuota de convenio de pago si tiene acuerdo activo | `50` o `0` |
| MANT | Mantenimiento mensual fijo | `3` |
| REUNION_FAENA | Multa por no asistir a reunión o faena | `30` o `0` |
| TECHADO | Cuota techado si está en el programa | `0` |
| DEVOLUCION | Si se le reconoció un exceso del mes pasado, va aquí | `5` o `0` |
| AJUSTE | Corrección manual (+/-). Usarlo solo si hay un error confirmado | `0` |

> El sistema calcula solo el consumo en m3 y el costo. No hace falta escribir
> m3 ni el total — eso lo hace `main.py`.

### 1.3 Casos especiales al llenar

**Usuario que no tiene medidor activo (lectura cero):**
Poner `0` en MARC_ANT y MARC_ACT. Solo pagará los conceptos fijos (MANT, REUNION_FAENA, etc.).

**Usuario nuevo sin arrastre:**
Poner `0` en ARRASTRE.

**Usuario con convenio de pago:**
El monto de la cuota mensual del convenio va en CONVENIO. Si pagó la cuota, igual
la dejas ahí — el cruce con los pagos lo resuelve el sistema.

**Usuario con blanco (descuento por pago bancario no identificado):**
No escribir nada en la planilla. Los blancos los aplica automáticamente el sistema
desde `shared/blancos_acumulados.xlsx` — siempre que el registro tenga MZ y LOTE
identificados y estado `pendiente`.

**Usuario con devolución reconocida:**
Si en el mes anterior pagó de más y ya se le reconoció, poner el monto en DEVOLUCION.
El sistema lo resta del total. No escribirlo como negativo — siempre positivo.

---

## Paso 2 — Copiar los archivos de pagos

Copiar estos dos archivos a sus carpetas dentro de `04_cobranza/inputs/`:

| Archivo | Origen | Destino |
|---|---|---|
| `pagos_yape_tepago.xlsx` | `03_pagos/yape/motor_matching/outputs/` | `inputs/pagos_yape/` |
| `pagos_efectivo.xlsx` | `03_pagos/efectivo/outputs/` | `inputs/pagos_efectivo/` |

> No renombrar los archivos — el sistema los busca por ese nombre exacto.

---

## Paso 3 — Ejecutar el módulo

Abrir una terminal en la carpeta `04_cobranza/` y ejecutar:

```
python main.py
```

El sistema mostrará el progreso:

```
═══════════════════════════════════════════════════════
  04_cobranza — Procesamiento de cobranza
═══════════════════════════════════════════════════════

[1/5] Validando inputs...
[2/5] Cargando datos...
[3/5] Calculando cobranza...
[4/5] Actualizando blancos...
[5/5] Exportando outputs...

═══════════════════════════════════════════════════════
  Cobranza completada → revisar outputs/
  Pasar lista_corte.xlsx → 05_corte
═══════════════════════════════════════════════════════
```

Si hay un error, el mensaje indica exactamente qué falta. Ver sección
[Problemas frecuentes](#problemas-frecuentes).

---

## Paso 4 — Revisar cobranza_final.xlsx

Abrir `outputs/cobranza_final.xlsx`. Este archivo tiene todos los usuarios del mes
con el resultado de su cobranza.

### Cómo leer el archivo

El archivo está organizado en cinco grupos de columnas:

| Grupo | Qué muestra |
|---|---|
| ¿Quién es? | MZ, Lote, Nombre |
| ¿Cuánto consumió? | Lecturas anteriores/actuales, m3, costo, total del mes |
| ¿Cuánto debía? | Todos los conceptos de deuda desglosados + total |
| ¿Cómo pagó? | Yape, Efectivo, Total pagado |
| ¿Qué queda? | Saldo y estado final |

### Cómo interpretar el ESTADO

| Estado | Color | Qué significa | Qué hacer |
|---|---|---|---|
| **CANCELADO** | Verde | Pagó exacto — saldo 0 | Nada |
| **EXCESO** | Azul | Pagó de más — saldo negativo | Registrar el exceso. Se le reconoce el próximo mes en DEVOLUCION cuando lo reclame |
| **PARCIAL** | Ámbar | Pagó algo pero quedó debiendo | Registrar el saldo como ARRASTRE en la planilla del próximo mes |
| **PENDIENTE** | Rojo | No pagó nada | Registrar el saldo como ARRASTRE. Si persiste el mes siguiente, irá a corte |

### Qué revisar antes de cerrar

1. ¿Todos los usuarios aparecen? El total de filas debe coincidir con la planilla base.
2. ¿Hay algún CANCELADO con saldo ≠ 0? Indicaría un error de carga.
3. ¿Los EXCESO son razonables? Montos muy grandes pueden ser errores de identificación en `03_pagos`.
4. ¿Los blancos aplicados son correctos? La columna BLANCOS mostrará el descuento en verde donde se aplicó.

---

## Paso 5 — Revisar lista_corte.xlsx

Abrir `outputs/lista_corte.xlsx`. Contiene solo los usuarios que cumplen la
condición de corte:
- Tienen saldo pendiente **este mes**
- Tenían arrastre de **≥ S/8 el mes anterior** (confirma que no pagaron nada el mes pasado)

### Filtrar usuarios ya cortados

Antes de pasar esta lista a `05_corte`, revisar manualmente si algún usuario
ya fue cortado en un mes anterior y aún no se ha reconectado. Esos usuarios
**no deben aparecer de nuevo** en la lista — eliminar su fila.

> Este filtro es manual. El sistema no lleva registro de qué usuarios están
> cortados actualmente.

### Qué datos trae la lista

| Columna | Descripción |
|---|---|
| MZ / LT / NOMBRE | Identificación del usuario |
| DEUDA_ARRASTRE | Lo que debía del mes pasado (≥ S/8) |
| SALDO | Deuda total pendiente este mes |
| PENALIDAD | S/20 fijo por reconexión |
| TOTAL_A_PAGAR | SALDO + S/20 — lo que debe pagar para reconectarse |

Una vez revisada, pasar a `05_corte`.

---

## Paso 6 — Revisar resumen_recaudacion.xlsx

Abrir `outputs/resumen_recaudacion.xlsx`. Muestra los totales del mes:

- Cuánto se debía en total según la planilla
- Cuánto se recaudó por Yape y por efectivo
- Cuánto saldo queda pendiente
- Cuántos usuarios en cada estado
- Cuántos van a corte

Usar este resumen para el reporte mensual a la directiva.

---

## Problemas frecuentes

**"Falta: inputs/planilla_base/planilla_base.xlsx"**
No existe el archivo. Ejecutar `python crear_template.py` para crearlo y luego llenarlo.

**"Columnas faltantes: {'MARC_ANT', 'MARC_ACT', ...}"**
La planilla tiene columnas del formato anterior (DEUDA_ARRASTRE, DEUDA_MES, DEUDA_TOTAL).
Regenerar el template con `crear_template.py` y trasladar los datos al nuevo formato.

**"Falta: inputs/pagos_yape/pagos_yape_tepago.xlsx"**
No se copió el archivo desde `03_pagos`. Copiar antes de ejecutar.

**"PermissionError" al actualizar blancos**
`blancos_acumulados.xlsx` está abierto en Excel. Cerrarlo y volver a ejecutar.

**La lista_corte salió vacía**
Normal si es el primer mes con deudas nuevas — nadie tiene arrastre ≥ 8 todavía.
También puede pasar si todos los usuarios con arrastre del mes pasado pagaron este mes.

**Un usuario no aparece en cobranza_final**
La fila en planilla_base tiene MZ o LT vacío, o el nombre de columna no es exacto.
Revisar que las columnas se llamen exactamente como indica el template.

---

## Referencia rápida — archivos del mes

```
INPUTS (llenar/copiar antes de ejecutar)
  inputs/planilla_base/planilla_base.xlsx   ← tú lo llenas
  inputs/pagos_yape/pagos_yape_tepago.xlsx  ← copia de 03_pagos
  inputs/pagos_efectivo/pagos_efectivo.xlsx ← copia de 03_pagos
  shared/blancos_acumulados.xlsx            ← automático (no tocar)

OUTPUTS (revisar después de ejecutar)
  outputs/cobranza_final.xlsx               → fuente de verdad del mes
  outputs/lista_corte.xlsx                  → filtrar y pasar a 05_corte
  outputs/resumen_recaudacion.xlsx          → reporte mensual
  outputs/run.log                           → log de ejecución
```
