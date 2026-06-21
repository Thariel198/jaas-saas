# 1_lecturas

Módulo que sincroniza el registro del operario con el padrón reconciliado y procesa las lecturas mensuales del medidor. Produce `lecturas_planilla_YYYY-MM.xlsx`, el input limpio que alimenta a `2_planilla` para calcular el cobro.

## Qué hace

1. **Sincroniza el padrón** (`proponer_sincronizacion.py` + `aplicar_sincronizacion.py`): detecta deltas entre `0_padron/outputs/padron_reconciliado.xlsx` y `registro_operario_acumulado.xlsx`, los clasifica (AGREGADO, SIN_SERVICIO, RENAME, MOVIMIENTO) y los aplica con backup + trazabilidad. Solo MOVIMIENTO requiere autorización humana.
2. **Genera el template del operario** (`crear_template.py`): produce `registro_operario_mes.xlsx` con todos los usuarios pre-cargados desde el acumulado ya sincronizado.
3. **Procesa las lecturas** (`main.py`): valida lecturas contra el historial usando 13 reglas de anomalía, gestiona ciclos de corrección y emite `lecturas_planilla_YYYY-MM.xlsx`.
4. **Genera la orden de campo** (`pdf_orden.py`): produce `orden_verificacion_YYYY-MM.pdf` para que el operario resuelva casos bloqueantes en campo.

## Cuándo se corre

| Momento | Script | Condición |
|---|---|---|
| Inicio del mes — antes de generar template | `proponer_sincronizacion.py` | Padrón actualizado en `0_padron/02_matching/outputs/` |
| Después de que supervisor revise MOVIMIENTOs | `aplicar_sincronizacion.py` | `reporte_sincronizacion.xlsx` con AUTORIZAR resuelto |
| Inicio del mes — una sola vez | `crear_template.py` | Acumulado ya sincronizado |
| Cada ciclo de corrección | `main.py` | Operario completó MARC_ACT, M3, obs_operario |
| Al final, automático | `pdf_orden.py` (invocado por main.py) | Hay anomalías bloqueantes |

## Estructura

```
1_lecturas/
├── proponer_sincronizacion.py    # ★ Diff padrón vs operario → reporte_sincronizacion.xlsx
├── aplicar_sincronizacion.py     # ★ Único writer de registro_operario_acumulado · backup + trazabilidad
├── crear_template.py             # Genera registro_operario_mes desde acumulado
├── main.py                       # Valida lecturas · gestiona ciclos · emite lecturas_planilla
├── pdf_orden.py                  # Genera orden_verificacion.pdf para el operario en campo
├── formato_excel.py              # Helpers de formato Excel
├── config.py                     # Umbrales, paths, códigos obs_operario
├── inputs/
│   ├── registro_operario_mes.xlsx          # Operario llena MARC_ACT, M3, obs_operario
│   ├── registro_operario_acumulado.xlsx    # Historial permanente · sincronizado con padrón
│   └── backups/                            # Backups timestamped del acumulado
├── outputs/
│   ├── lecturas_planilla_YYYY-MM.xlsx      # PERMANECE — input para 2_planilla
│   ├── trazabilidad_YYYY-MM.xlsx           # PERMANECE — auditoría mensual
│   ├── correcciones_YYYY-MM.xlsx           # Cola de maquillaje · se borra al cerrar
│   ├── orden_verificacion_YYYY-MM.pdf      # Se regenera cada ciclo
│   ├── run.log
│   └── sync/
│       ├── reporte_sincronizacion.xlsx     # Deltas detectados · sticky entre ciclos
│       └── trazabilidad_sincronizacion.xlsx # Append-only · historial de mutaciones
├── tests/
└── docs/
    ├── diagrama_flujo_sincronizacion.html  # Flujo rápido del sync
    ├── diagrama_sincronizacion.html        # Diseño detallado del sync
    ├── manual_uso.md                       # Detalle de anomalías, ciclos, errores comunes
    ├── arquitectura_1_lecturas.html
    ├── contrato_*.html                     # Schema por archivo producido
    └── flujo_modulo_1_lecturas.svg
```

## Dependencias externas

| Recurso | Tipo | Quién lo gobierna |
|---|---|---|
| `0_padron/02_matching/outputs/padron_reconciliado.xlsx` | archivo (lectura) | `0_padron/` — fuente de verdad de MZ, LT, NOMBRE |
| `inputs/registro_operario_acumulado.xlsx` | archivo (escritura) | `aplicar_sincronizacion.py` — único writer; `main.py` actualiza columnas de ciclo |
| `outputs/lecturas_planilla_YYYY-MM.xlsx` | archivo (escritura) | `main.py` — consumido por `2_planilla` |

**`aplicar_sincronizacion.py` es el único script que muta MZ/LT/NOMBRE en `registro_operario_acumulado.xlsx`.** Backup timestamped antes de cada mutación, trazabilidad append-only, idempotencia vía columna APLICADO. `main.py` solo agrega columnas de ciclo (YYYY-MM.MARCACION, YYYY-MM.M3).

## Reglas clave de sincronización

Cuatro tipos de delta entre `padron_reconciliado.xlsx` y `registro_operario_acumulado.xlsx`:

| Tipo | Condición | Acción | Autorización |
|---|---|---|---|
| `AGREGADO` | (MZ, LT) en padrón, NO en operario | Agrega fila vacía en acumulado | Automática |
| `SIN_SERVICIO` | (MZ, LT) en operario, NO en padrón | Marca `SIN_SERVICIO=Si` · conserva lecturas históricas | Automática |
| `RENAME` | Mismo (MZ, LT), NOMBRE difiere (normalizado) | Actualiza nombre desde padrón (padrón manda) | Automática |
| `MOVIMIENTO` ★ | Mismo NOMBRE aparece en padrón con (MZ, LT) distinto | **Bloqueado** hasta `REVISADO=Si` + `AUTORIZAR=Si` | **Manual** |

Normalización aplicada a MZ, LT, NOMBRE: strip · UPPER · NFD Unicode · sin marcas de acento · espacios colapsados.

Columnas sticky en `reporte_sincronizacion.xlsx` (preservadas entre ciclos): `REVISADO`, `AUTORIZAR`. Columna idempotente: `APLICADO`.

## Reglas clave de lecturas

Las 13 anomalías (8 bloqueantes + 4 informativas + 1 legitimada) están documentadas en `docs/manual_uso.md`. Resumen:

- **Bloqueantes:** `MEDIDOR_INVERTIDO`, `POSIBLE_CAMBIO_MEDIDOR`, `DIFERENCIA_M3`, `EXCESIVO`, `SIN_LECTURA`, `DUPLICADO`, `USUARIO_FANTASMA`, `MARC_ACT_NO_NUMERICO`.
- **Informativas:** `SIN_HISTORIAL`, `CONSUMO_CERO`, `SALTO_HISTORICO`, `MEDIDOR_CAMBIADO`.
- **Códigos obs_operario:** `M` (medidor cambiado), `F` (fuga), `P` (predio cerrado).
- **Ciclos:** Ciclo 1 = primer pass; Ciclo 2+ = procesa `correcciones_YYYY-MM.xlsx` y mueve resueltas a trazabilidad.

## Flujo mensual

```
# PASO 0 — Sincronizar padrón (al inicio del mes)

python proponer_sincronizacion.py
   ← 0_padron/02_matching/outputs/padron_reconciliado.xlsx
   ← inputs/registro_operario_acumulado.xlsx
   ← outputs/sync/reporte_sincronizacion.xlsx  (previo, si existe)
   → outputs/sync/reporte_sincronizacion.xlsx  [deltas clasificados]

# Si hay MOVIMIENTOs: supervisor abre el reporte y marca AUTORIZAR=Si/No.
# Si no hay MOVIMIENTOs: avanzar directo.

python aplicar_sincronizacion.py
   ← outputs/sync/reporte_sincronizacion.xlsx
   ← inputs/registro_operario_acumulado.xlsx
   → inputs/registro_operario_acumulado.xlsx   [mutado]
   → inputs/backups/registro_operario_acumulado_<ts>.xlsx
   → outputs/sync/trazabilidad_sincronizacion.xlsx  [append]

# PASO 1 — Generar template del operario

python crear_template.py
   ← inputs/registro_operario_acumulado.xlsx   [ya sincronizado]
   → inputs/registro_operario_mes.xlsx         [operario completa]

# PASO 2 — Operario completa el template en campo

# PASO 3 — Procesar lecturas

python main.py
   ← inputs/registro_operario_mes.xlsx
   ← inputs/registro_operario_acumulado.xlsx
   → outputs/lecturas_planilla_YYYY-MM.xlsx    [si bloqueantes=0]
   → outputs/correcciones_YYYY-MM.xlsx         [si hay bloqueantes]
   → outputs/orden_verificacion_YYYY-MM.pdf    [si hay casos de campo]
   → outputs/trazabilidad_YYYY-MM.xlsx         [informativas]

# PASO 4 — Ciclo de corrección (si aplica)
# Repetir hasta correcciones_YYYY-MM.xlsx = 0 filas.

# PASO 5 — Listo → 2_planilla
```

## Lifecycle de outputs

| Archivo | Lifecycle |
|---|---|
| `inputs/registro_operario_acumulado.xlsx` | **PERMANECE** · mutado por sync · enriquecido por main.py |
| `inputs/backups/*.xlsx` | **PERMANECE** · backup pre-mutación · uno por ejecución de aplicar_sincronizacion |
| `outputs/sync/reporte_sincronizacion.xlsx` | Mensual — se regenera cada ciclo · sticky en REVISADO/AUTORIZAR |
| `outputs/sync/trazabilidad_sincronizacion.xlsx` | **PERMANECE** · append-only · historial de todas las mutaciones |
| `outputs/lecturas_planilla_YYYY-MM.xlsx` | **PERMANECE** · uno por mes · fuente de verdad para 2_planilla |
| `outputs/trazabilidad_YYYY-MM.xlsx` | **PERMANECE** · uno por mes · auditoría de anomalías |
| `outputs/correcciones_YYYY-MM.xlsx` | Se elimina al cerrar el mes (bloqueantes=0) |
| `outputs/orden_verificacion_YYYY-MM.pdf` | Se regenera cada ciclo |
| `outputs/run.log` | Se sobreescribe cada run |

## Lo que este módulo NO hace

- **No reconcilia el padrón** — eso lo hace `0_padron`. Aquí solo sincronizamos contra su output.
- **No calcula cobros** — eso es `2_planilla`, a partir de `lecturas_planilla_YYYY-MM.xlsx`.
- **No emite boletas** — eso es `3_boletas`.
- **No mantiene tarifas** — viven en `shared/tarifas.xlsx` y las lee `2_planilla`.
- **No corrige lecturas ya facturadas** — para eso existe `6b_override`.
- **No elimina filas del acumulado** — SIN_SERVICIO marca, no borra; las lecturas históricas se conservan siempre.

## Errores comunes

**"No se encontró `padron_reconciliado.xlsx`"**
→ Correr primero el módulo `0_padron`. El sync depende del output de la reconciliación COFOPRI + padrón agua.

**"No se encontró `registro_operario_acumulado.xlsx`"**
→ Si es la primera vez que se corre el sistema, partir copiando manualmente desde el último Excel del operario. Si ya existía, revisar si fue movido por error.

**"No se encontró `registro_operario_mes.xlsx`"**
→ Correr `crear_template.py` primero. Si es el primer mes, el template queda vacío y el operario lo completa manualmente.

**"MOVIMIENTO bloqueado — AUTORIZAR vacío"**
→ Hay deltas tipo MOVIMIENTO en `reporte_sincronizacion.xlsx` que el supervisor no revisó. Abrir el reporte, marcar `REVISADO=Si` y `AUTORIZAR=Si/No` por fila, guardar, y re-correr `aplicar_sincronizacion.py`.

**"`aplicar_sincronizacion.py` no muta nada"**
→ Revisar columna `APLICADO` en el reporte. Si todas dicen `Si`, ya se aplicaron en una corrida previa (idempotencia). Si no, revisar `AUTORIZAR` — los MOVIMIENTOs requieren `Si` explícito.

**"Columnas faltantes en `registro_operario_acumulado.xlsx`"**
→ El archivo no tiene el doble header esperado. Posiblemente se generó con una versión vieja. Restaurar desde `inputs/backups/` el último backup válido.

**"`obs_operario` tiene valor inválido: X"**
→ Solo se aceptan `M`, `F`, `P` o celda vacía. Revisar el template y corregir.

**"Anomalía registrada pero el operario dice que está bien"**
→ El operario puede legitimar con `obs` (M para retroceso, F para fuga, P para predio cerrado). Si no aplica obs y el valor es correcto, marcar `resuelto_por = acepta_original` en correcciones.

## Señales de alerta

| Señal | Diagnóstico |
|---|---|
| `reporte_sincronizacion.xlsx` con muchos MOVIMIENTOs | Revisar si el padrón tuvo una recarga masiva — confirmar con 0_padron antes de autorizar |
| `aplicar_sincronizacion.py` no muta nada | Revisar columna APLICADO en reporte — todas las filas pueden estar ya aplicadas, o AUTORIZAR vacío |
| `crear_template.py` genera filas inesperadas | El acumulado tiene MZ/LT que no corresponden — correr `proponer_sincronizacion.py` para diagnóstico |
| `lecturas_planilla` tiene menos filas que el acumulado | Hay usuarios marcados `SIN_SERVICIO=Si` — esperado, no error |
| Backup en `inputs/backups/` no se generó | `aplicar_sincronizacion.py` falló antes de mutar — revisar log; el acumulado no fue tocado |
| Anomalía operario reporta error pero no había sync previo | Faltó correr sync este mes — el padrón pudo haber cambiado un nombre o agregado un lote |

## Referencias

- **Detalle de anomalías y ciclos:** `docs/manual_uso.md`
- **Diseño del sync:** `docs/diagrama_flujo_sincronizacion.html` y `docs/diagrama_sincronizacion.html`
- **Contratos de cada archivo:** `docs/contrato_*.html`
- **Decisiones arquitectónicas:** `docs/decisiones/1_lecturas.md`
