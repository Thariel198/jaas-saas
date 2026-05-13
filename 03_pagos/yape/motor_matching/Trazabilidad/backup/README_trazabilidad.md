# trazabilidad_YYYY_MM.xlsx — reglas de negocio

Archivo mensual permanente. Registra todo lo que pasó en el motor matching del mes.
Nunca se borra. Crece mes a mes en la carpeta Trazabilidad/.

---

## Ciclo de vida

- **Se crea** en Ciclo 1 de cada mes con nombre `trazabilidad_YYYY_MM.xlsx` (ej: `trazabilidad_2026_05.xlsx`)
- **Se actualiza** en cada ciclo del mes — agrega filas nuevas, nunca borra las anteriores
- **Se actualiza también** cuando corre `apply_overrides.py` — agrega filas en hoja Overrides
- **Nunca se borra** — el Ciclo 1 del mes siguiente crea un archivo nuevo, este queda intacto
- **Quién lo escribe** — `main.py` escribe hojas 1, 2 y 3. `apply_overrides.py` escribe hoja 4

---

## Hoja 1 — Sin_identificar

Registra los pagos que no tenían mensaje ni maestro y fueron resueltos manualmente por el usuario en `pendientes.xlsx`.

### Columnas

| Columna | Tipo | Quién lo pone | Descripción |
|---------|------|---------------|-------------|
| FECHA | datetime | banco | fecha y hora del pago — copia de Reporte_mes |
| ORIGEN | texto | banco | nombre del pagador — copia de Reporte_mes |
| MONTO | decimal | banco | monto del pago |
| MENSAJE | texto | banco | mensaje del pago aunque esté vacío |
| MZ | texto | usuario | lo que el usuario escribió en pendientes.xlsx — puede ser BLANCO |
| LOTE | texto | usuario | lo que el usuario escribió — vacío si MZ = BLANCO |
| MOTIVO | texto | usuario | texto libre — explicación de cómo se identificó o por qué es BLANCO |
| CICLO | entero | sistema | número de ciclo en que se resolvió (2, 3, 4...) |
| FECHA_CORRECCION | datetime | sistema | `datetime.now()` del momento en que main.py procesó la corrección |

### Reglas
- Se agrega una fila por cada pago resuelto desde `pendientes.xlsx` hoja `Sin_identificar`
- Si MZ = BLANCO, LOTE queda vacío
- MOTIVO es texto libre — puede estar vacío si el usuario no escribió nada
- La clave de unicidad es ORIGEN + FECHA — no solo ORIGEN. Un mismo pagador puede tener varios pagos en el mes para distintos lotes y todos se registran como filas independientes
- Cuando MZ = BLANCO el registro también se agrega automáticamente a `Blancos/blancos_acumulados.xlsx` con ESTADO = pendiente

---

## Hoja 2 — Ambiguos

Registra los pagos donde el maestro tenía 2 o más candidatos. Una fila por candidato.

### Columnas

| Columna | Tipo | Quién lo pone | Descripción |
|---------|------|---------------|-------------|
| FECHA | datetime | banco | fecha y hora del pago |
| ORIGEN | texto | banco | nombre del pagador — se repite por cada candidato |
| MONTO | decimal | banco | monto del pago — se repite por cada candidato |
| MZ_CANDIDATO | texto | maestro | MZ de este candidato |
| LOTE_CANDIDATO | texto | maestro | LOTE de este candidato |
| NOMBRE_CANDIDATO | texto | maestro | nombre del dueño de este candidato |
| DEUDA_CANDIDATO | decimal | planilla | deuda del candidato en la planilla del mes |
| DIFF_CANDIDATO | decimal | sistema | `MONTO - DEUDA_CANDIDATO` — positivo = exceso, negativo = parcial, 0 = exacto |
| ELEGIDO | texto | sistema | `SI` en la fila del candidato seleccionado, `NO` en el resto |
| MZ_FINAL | texto | sistema | MZ que quedó en pagos_yape.xlsx — se repite en todas las filas del mismo origen |
| LOTE_FINAL | texto | sistema | LOTE que quedó en pagos_yape.xlsx — se repite en todas las filas del mismo origen |
| CICLO | entero | sistema | ciclo en que se procesó |
| FECHA_CORRECCION | datetime | sistema | `datetime.now()` del momento en que se procesó |

### Reglas
- Una fila por candidato — si hay 3 candidatos hay 3 filas para el mismo ORIGEN
- ELEGIDO = SI solo en una fila por ORIGEN + FECHA
- El sistema elige automáticamente el candidato con menor diferencia absoluta (prioridad diff=0) — nunca llega a pendientes.xlsx
- Si la elección automática fue incorrecta, se corrige únicamente via `apply_overrides.py` — no via pendientes
- Cuando se aplica un override, MZ_FINAL y LOTE_FINAL se actualizan en todas las filas del mismo ORIGEN + FECHA y se agrega una fila nueva con ELEGIDO = SI para el candidato correcto
- MZ_FINAL y LOTE_FINAL se repiten en todas las filas del mismo ORIGEN + FECHA para poder filtrar y ver de un vistazo cuál ganó
- No se borra ninguna fila — el historial de candidatos descartados es parte de la trazabilidad

---

## Hoja 3 — Pagos_multiples

Registra los pagos donde el mensaje tenía 2 o más pares MZ-LOTE. Una fila por lote.

### Columnas

| Columna | Tipo | Quién lo pone | Descripción |
|---------|------|---------------|-------------|
| FECHA | datetime | banco | fecha y hora del pago |
| ORIGEN | texto | banco | nombre del pagador — se repite por cada lote |
| MONTO_TOTAL | decimal | banco | monto total del pago del banco — se repite por cada lote |
| MENSAJE | texto | banco | mensaje original del pago — se repite por cada lote |
| MZ | texto | sistema/usuario | MZ asignada a este lote |
| LOTE | texto | sistema/usuario | LOTE asignado |
| NOMBRE | texto | planilla | nombre del dueño del lote en la planilla |
| DEUDA | decimal | planilla | deuda del lote en la planilla del mes |
| MONTO_ASIGNADO | decimal | sistema | porción del pago total asignada a este lote |
| DIFF | decimal | sistema | `MONTO_ASIGNADO - DEUDA` |
| CICLO | entero | sistema | ciclo en que se procesó |
| FECHA_CORRECCION | datetime | sistema | `datetime.now()` del momento en que se procesó |

### Reglas
- Una fila por lote — si el pago es para 3 lotes hay 3 filas con el mismo ORIGEN y FECHA
- MONTO_TOTAL es siempre el del banco — no cambia entre filas del mismo origen
- MONTO_ASIGNADO: cada lote recibe su deuda exacta. El último lote absorbe la diferencia si el total no cuadra
- El sistema asigna automáticamente — nunca llega a pendientes.xlsx
- Si la asignación fue incorrecta se corrige únicamente via `apply_overrides.py` — no via pendientes
- Se acumulan todos los pagos múltiples del mes — diff=0 y diff≠0 por igual
- No se borra ninguna fila — si hay un override se agregan filas nuevas con los valores corregidos

---

## Hoja 4 — Overrides

Registra cada cambio post-validación aplicado por `apply_overrides.py`. Una fila por override.

### Columnas

| Columna | Tipo | Quién lo pone | Descripción |
|---------|------|---------------|-------------|
| FECHA_PAGO | datetime | usuario | fecha del pago copiada desde pagos_yape.xlsx — clave de identificación |
| ORIGEN | texto | sistema | nombre del pagador — lo extrae apply_overrides.py de pagos_yape.xlsx |
| MZ_ANTERIOR | texto | usuario | MZ que tenía el pago antes del override |
| LOTE_ANTERIOR | texto | usuario | LOTE que tenía el pago antes del override |
| MZ_NUEVO | texto | usuario | MZ nueva que se asignó |
| LOTE_NUEVO | texto | usuario | LOTE nuevo que se asignó |
| MOTIVO | texto | usuario | texto libre — razón del cambio |
| FECHA_OVERRIDE | datetime | sistema | `datetime.now()` del momento en que apply_overrides.py aplicó el cambio |

### Reglas
- Una fila por override aplicado
- Nunca se borra la entrada anterior — cada override es una fila nueva
- ORIGEN lo extrae el script automáticamente desde pagos_yape.xlsx buscando por FECHA_PAGO — el usuario no lo escribe
- Si el mismo pago recibe 2 overrides en el mismo mes hay 2 filas en esta hoja
- Se puede aplicar override sobre un pago que ya tuvo override antes — sin límite

---

## Valores calculados por el sistema

| Campo | Fórmula |
|-------|---------|
| DIFF_CANDIDATO | `round(MONTO - DEUDA_CANDIDATO, 2)` |
| DIFF (multiples) | `round(MONTO_ASIGNADO - DEUDA, 2)` |
| MONTO_ASIGNADO último lote | `round(DEUDA + (MONTO_TOTAL - suma_deudas_anteriores), 2)` |
| FECHA_CORRECCION | `datetime.now()` al momento de procesar |
| FECHA_OVERRIDE | `datetime.now()` al momento de aplicar |

---

## Nombre del archivo

```python
# Se genera así en main.py Ciclo 1:
from datetime import datetime
mes = datetime.now().strftime("%Y_%m")
nombre = f"trazabilidad_{mes}.xlsx"
# Ejemplo: trazabilidad_2026_05.xlsx
```

---

## Lo que NO hace este archivo

- No reemplaza a `pagos_yape.xlsx` — ese es el output operativo, este es el historial
- No se lee en Ciclo 2+ para tomar decisiones — solo se escribe
- No contiene pagos identificados por maestro único o mensaje simple — solo los casos que necesitaron procesamiento especial
- Ambiguos y Pagos_multiples nunca pasan por `pendientes.xlsx` — el sistema los resuelve solo y se corrigen únicamente via override
- No maneja los pagos BLANCO a largo plazo — eso lo hace `Blancos/blancos_acumulados.xlsx`
