# shared/

Capa delgada de primitivos puros y servicios transversales que **múltiples módulos** consumen. No es un módulo de negocio — no tiene main, no ejecuta nada por sí mismo. Solo expone utilidades y servicios para que la lógica de cada módulo permanezca local.

## Cuándo va algo en `shared/`

Solo si cumple **las tres**:

1. **Múltiples módulos lo usan** — no se inventa "por si acaso"; existe porque hay ≥2 callers reales.
2. **API estable** — su firma no cambia con cada release de un módulo.
3. **Sin lógica de negocio específica de un módulo** — primitivo puro o servicio horizontal (audit, lock, autorización, etc.).

Si cumple 1 y 2 pero no 3, vive en el módulo que es dueño del concepto y se importa desde otros.

## Contenido

### Servicios (con persistencia y estado)

| Archivo | Qué hace | Quién lo usa |
|---|---|---|
| `data_boletas_repo.py` | **Único writer de `DATA_boletas.xlsx`.** Encapsula backup + write + audit en cada mutación. Lecturas y escrituras pasan por este repo. | `4b_reclamos/{resolucion, aplicar_correcciones, validacion_*}.py` · agentes futuros · UI/dashboard futuro |
| `data_boletas_audit.xlsx` | Audit log centralizado de toda mutación a `DATA_boletas`. Append-only. Owner: `data_boletas_repo`. | Cualquier consumidor que quiera responder "quién cambió qué/cuándo/por qué" |
| `seguimiento_repo.py` | **Único writer de `seguimiento_pueblo.xlsx`.** Implementación transitoria de `estado_cuenta`: AGUA · MANTENIMIENTO · CORTE_RECONEXION · MULTA · ACUERDOS · CONVENIO · OTROS. | `7_cierre` (commit batch) · `2_planilla` y reportes (lectura) |
| `seguimiento_pueblo.xlsx` | Persistencia física transitoria de la cuenta completa. Append-only; el nombre se conserva mientras existan consumidores directos. | 2_planilla · 7_cierre · reportes |
| `vista_seguimiento_provicional.xlsx` | Simulación regenerable del snapshot validado sobre una copia temporal del ledger. Nunca es fuente ni escribe eventos reales. | Revisión humana durante el ciclo abierto |

### Primitivos puros (sin estado)

| Archivo | Qué hace | Quién lo usa |
|---|---|---|
| `utils_lote.py` | Lee `5_cobranza/inputs/correcciones_lote.xlsx` → dict `{(mz_orig, lt_orig): (mz_dest, lt_dest)}`. Función pura, sin side-effects. | `4b_reclamos/{main, resolucion}.py` y cualquier módulo que necesite remapear MZ/LT antes de cruzar |

### Datos derivados (no son código)

| Archivo | Qué es | Generado por |
|---|---|---|
| `blancos_acumulados.xlsx` | Acumulado mensual de blancos para validaciones cruzadas | (varía según pipeline) |
| `usuarios_id.xlsx` | Tabla maestra de usuarios | (varía según pipeline) |

## Patrón Repository — `data_boletas_repo`

`DATA_boletas.xlsx` es el padrón maestro de predios y vive físicamente en `3_boletas/inputs/`. A medida que el sistema escala a agentic SaaS, múltiples agentes y módulos van a querer leer y mutar este archivo. Si cada uno escribe directo, el audit trail se fragmenta y la disciplina (backup, log, validación) no escala.

**El repo es la única puerta a `DATA_boletas`:**

```
Callers                         Repo                            Persistencia
─────────                       ────                            ────────────
resolucion.py        ──get_predio()──▶                      ┌─▶ DATA_boletas.xlsx
validacion.py        ──read_padron()─▶ data_boletas_repo ───┼─▶ data_boletas_audit.xlsx
aplicar_correcciones ──apply_corr.()─▶                      └─▶ 3_boletas/backup/DATA_boletas/
(agente futuro)      ──apply_corr.()─▶
```

### API

| Función | Side-effects | Para qué |
|---|---|---|
| `read_padron()` | Ninguno | Devuelve el padrón completo como DataFrame |
| `get_predio(mz, lt)` | Ninguno | Devuelve la fila del predio o `None` |
| `apply_correction(mz, lt, campo, valor, *, source, audit_ref, motivo)` | Backup + write + audit append | Mutación atómica — devuelve `{valor_antes, valor_despues, ts, audit_id}` |

### Invariantes

- **Single writer:** `DATA_boletas.xlsx` solo se escribe vía `apply_correction()`. Ningún módulo lo abre con `load_workbook` para escribir directo.
- **Audit obligatorio:** no hay forma de mutar sin pasar por el audit log. Cada cambio queda con timestamp, source, audit_ref y motivo trazables.
- **Backup automático:** antes de cualquier write, snapshot completo en `3_boletas/backup/DATA_boletas/`.
- **Idempotencia:** aplicar el mismo `audit_ref` dos veces produce un solo write efectivo.
- **Validación:** el repo rechaza llamadas con `campo` que no existe en DATA_boletas o `(mz, lt)` que no existe en el padrón.

### Listo para escalar

Cuando `DATA_boletas` migre a Postgres o a un microservicio, **solo el repo cambia internamente**. Los callers siguen llamando la misma API. Eso convierte una migración de archivo → base de datos en una sustitución de implementación, no en un refactor global.

## Patrón Event-Sourced — `seguimiento_repo`

Desde agosto de 2026 todos los conceptos de deuda comparten la misma cuenta. Agua y mantenimiento
reciben cargos mensuales; corte, multa, acuerdos y convenio se amortizan según sus eventos. El repo
guarda **eventos**, no una foto mutable del saldo.

```
Emisores de eventos              Repo                              Persistencia
──────────────────               ────                              ────────────
sembrar_seguimiento_pueblo ──registrar_cargo()──▶
                                                  ┐
5_cobranza/main.py         ──registrar_pago()────┼─▶ seguimiento_repo ──▶ shared/seguimiento_pueblo.xlsx
                                                  ┘
(corrección manual)        ──registrar_ajuste()──▶

2_planilla/main.py         ──get_saldo()─────────▶ (solo lectura)
(consulta humana)          ──estado_cuenta()─────▶ (solo lectura, pivot a vista ancha)
```

**Génesis = el primer evento.** No existe un "ledger de génesis" aparte — sembrar la deuda inicial de
un predio es simplemente su primer `registrar_cargo()`. El mismo mecanismo cubre siembra, cargos nuevos
mensuales (faena/reunión/asamblea) y pagos: todo es un evento.

### API

| Función | Side-effects | Para qué |
|---|---|---|
| `registrar_cargo(mz, lt, concepto, mes, monto, *, source, audit_ref)` | Append de evento CARGO | Apertura o deuda nueva de cualquiera de los siete conceptos |
| `registrar_pago(mz, lt, concepto, mes, monto, *, source, audit_ref)` | Append de evento PAGO | 5_cobranza registra la porción de un pago que va a ese concepto |
| `registrar_ajuste(mz, lt, concepto, mes, ±monto, *, source, audit_ref, motivo)` | Append de evento AJUSTE | Corrección — nunca se edita un evento pasado |
| `get_saldo(mz, lt, concepto, mes)` | Ninguno | Suma eventos hasta ese mes — lo que 2_planilla consulta |
| `estado_cuenta(mz, lt, concepto)` | Ninguno | Pivot ancho de 1 predio (DEUDA·PAGO·SALDO por mes) — DataFrame en memoria |
| `generar_vista(ruta=None)` | Escribe `vista_seguimiento_pueblo.xlsx` (regenerable) | Pivot ancho de todos los predios, una hoja por concepto |

### Invariantes

- **Single writer:** `seguimiento_pueblo.xlsx` solo se escribe vía las 3 funciones `registrar_*`. Ningún módulo lo abre con `load_workbook` para escribir directo.
- **Append-only real:** un evento nunca se modifica ni se borra. Toda corrección es un nuevo evento `AJUSTE` con motivo obligatorio.
- **Anulación lógica:** `anulaciones_ledger.json` excluye referencias causales de las proyecciones sin borrar filas físicas; cada lote conserva motivo, fecha y detalle de eventos.
- **Saldo derivado, nunca mutable:** `SALDO(mz, lt, concepto, mes) = Σcargos − Σpagos ± Σajustes`. Sin celda que dos escrituras puedan pisarse (evita la clase de bug de B7 — ver `docs/aprendizaje/writer_unico_desincronizacion_20260701.html`).
- **Idempotencia:** mismo `(source, audit_ref)` no duplica el evento.
- **Guardar largo, mostrar ancho:** el registro es la fuente; el estado de cuenta ancho es un pivot regenerable, nunca se edita directo.
- **Escritura atómica:** cada `wb.save()` va a un archivo temporal en el mismo directorio y recién se reemplaza con `os.replace()` (atómico en Windows/POSIX). Si el proceso se corta a mitad de camino, el archivo real queda intacto — nunca un `.xlsx` corrupto a medio escribir. Verificado con test que simula el corte (`shared/tests/test_seguimiento_repo.py`) y con un fallo real en producción (`PermissionError` de Windows por antivirus/indexador bloqueando el archivo un instante — el archivo real quedó intacto). El `os.replace()` reintenta hasta 5 veces ante `PermissionError` transitorio antes de fallar de verdad.

### Cuándo se corre

| Quién llama | Cuándo | Comando / forma de invocar |
|---|---|---|
| `sembrar_seguimiento_pueblo.py` | Una sola vez — siembra la deuda inicial (julio 2026) o cuando aparezca una tanda nueva de altas | `py shared/sembrar_seguimiento_pueblo.py` (manual, no forma parte de ningún pipeline automático) |
| `5_cobranza/main.py` | Cada corrida mensual del ciclo, automático — dentro de su propio flujo | no se llama aparte; es una llamada interna de `5_cobranza` a `repo.registrar_pago()` |
| `2_planilla/main.py` | Cada corrida mensual, automático — al armar la planilla | no se llama aparte; es una llamada interna a `repo.get_saldo()` |
| Consulta humana (`estado_cuenta`, `deudores`) | Bajo demanda, sin calendario | `python -c "import seguimiento_repo as r; print(r.estado_cuenta('A','6','CONVENIO'))"` o script ad-hoc |

### Tabla de lifecycle

| Archivo | Se borra? | Se regenera? | Persistencia |
|---|---|---|---|
| `shared/seguimiento_pueblo.xlsx` | Nunca | Nunca — cada corrida solo agrega filas | Permanente |
| `shared/anulaciones_ledger.json` | Nunca | No; agrega lotes de anulación lógica auditados | Permanente |
| Vista `estado_cuenta()` (en memoria) | N/A — no se guarda a disco | Sí, cada vez que se llama | Efímera |
| `shared/genesis_inputs/medidor_saldo.xlsx` · `inscripcion_saldo.xlsx` | No — se reemplaza a mano cuando hay una fuente más nueva | Manual (copiar de nuevo desde Downloads) | Permanente hasta el próximo reemplazo |

### Lo que NO hace `seguimiento_repo`

- No decide el reparto del pago entre conceptos (agua/corte/multa/acuerdos/convenio) — eso lo calcula `5_cobranza` con el waterfall de prioridad; el repo solo registra el resultado.
- No genera la planilla ni escribe en `shared/planilla_mes` — `2_planilla` solo *lee* `get_saldo()`.
- Agosto usa el consolidado de julio una sola vez como apertura. Desde septiembre, `2_planilla` consulta AGUA, MANTENIMIENTO y CORTE_RECONEXION directamente en este repo.
- No borra ni corrige eventos existentes — toda corrección es un evento `AJUSTE` nuevo.

### Errores comunes

| Síntoma | Causa probable | Cómo revisar |
|---|---|---|
| `ValueError: concepto inválido` | Se llamó con un concepto fuera de la taxonomía de `estado_cuenta` | Revisar el caller y `CONCEPTOS_VALIDOS` |
| Un pago se registró dos veces con distinto saldo | `audit_ref` no era único para ese pago (ej. reutilizado entre corridas) | Revisar que `5_cobranza` arme `audit_ref` con una clave estable del pago (mesa+fecha o id de trazabilidad) |
| `get_saldo()` devuelve 0 para un predio que sí tiene deuda | Todavía no se corrió `sembrar_seguimiento_pueblo.py`, o el predio no estaba en las fuentes de siembra | Revisar `shared/seguimiento_pueblo.xlsx` filtrando por MZ/LT — si no hay ninguna fila, falta la siembra |
| `estado_cuenta()` muestra un mes con DEUDA y PAGO en cero pero SALDO distinto al mes anterior | No debería pasar — señal de un evento AJUSTE sin revisar | Filtrar TIPO_EVENTO=AJUSTE de ese predio/mes y verificar el motivo |

### Esquema de inputs — `sembrar_seguimiento_pueblo.py`

Los 3 archivos que lee la siembra inicial. Ninguno se copia a un `inputs/` del pipeline automático —
`DATA_boletas` ya vive en `3_boletas/inputs/`; `medidor_saldo.xlsx` e `inscripcion_saldo.xlsx` viven en
`shared/genesis_inputs/` (copiados a mano desde `Downloads/Base de datos/` — no hay fuente automática).

**1. `3_boletas/inputs/DATA_boletas.xlsx`, hoja `Data`** (usado para MULTA y ACUERDOS)

| Columna requerida | Tipo | Uso |
|---|---|---|
| `MZ` | texto | llave |
| `LT` | texto | llave |
| `Multa (faena + reunión)` | número, puede venir vacío | valor &gt; 0 → 1 evento CARGO concepto MULTA |
| `Cuota directa` | número, puede venir vacío | valor &gt; 0 → 1 evento CARGO concepto ACUERDOS |

Si falta el archivo: la siembra aborta con error explícito (no hay MULTA/ACUERDOS sin esta fuente).
Si falta una de las 2 columnas: la siembra aborta — son las únicas 2 columnas de esta fuente que usa.
La columna `Convenio` de este mismo archivo **no se usa** — el monto de convenio sale de las otras 2 fuentes.

**2. `shared/genesis_inputs/medidor_saldo.xlsx`, hoja `Cobro medidores`** (usado para CONVENIO)

| Columna requerida | Tipo | Uso |
|---|---|---|
| `MZ` | texto | llave |
| `LT` | texto | llave |
| `Saldo` | número | predio con `Saldo > 0` entra a la población de convenio; el monto se suma con inscripción |

Si falta el archivo: la siembra aborta — sin esta fuente no hay forma de saber la deuda real de medidor.

**3. `shared/genesis_inputs/inscripcion_saldo.xlsx`, hoja `NUEVAS INSTALACIONES`** (usado para CONVENIO)

| Columna requerida | Tipo | Uso |
|---|---|---|
| `MZ` | texto | llave |
| `LT` | texto | llave |
| `SALDO` | número | predio con `SALDO > 0` entra a la población de convenio; se suma con medidor si tiene ambos |

Si falta el archivo: la siembra aborta — mismo motivo que medidor.

**Predios excluidos de la siembra de CONVENIO:** B-20 · C-43 · C-35 · F1-11 · G-21 · W-2.

Estos 6 tienen un convenio de **instalación** (conexión nueva), no de medidor/inscripción — montos grandes
(283 · 346 · 300 · 776 · 50 · 50). Junio ya cargó esa deuda **completa** en `arrastre_consolidado_2026-06`
y se sigue arrastrando bien mes a mes desde ahí. Si la siembra nueva **también** les crea un evento
CONVENIO (porque por casualidad aparecen con saldo&gt;0 en `medidor_saldo`/`inscripcion_saldo`), quedarían
con la misma deuda contada dos veces — una en el consolidado, otra en `seguimiento_pueblo` — la misma
clase de bug que B7 (dos escritores para el mismo dinero), prevenida acá antes de que pase.

Verificado (2026-07-02): hoy ninguno de los 6 aparece con saldo&gt;0 en medidor ni inscripción — la
exclusión no cambia ningún número actual, es un candado para el día en que alguno sí aparezca ahí.

### Documentación

```
shared/
├── README.md                              (este archivo)
├── data_boletas_repo.py                   # módulo
├── data_boletas_audit.xlsx                # audit log (output)
├── seguimiento_repo.py                    # módulo — pendiente de codificar (Sonnet)
├── seguimiento_pueblo.xlsx                # registro de eventos (output) — pendiente de codificar
├── utils_lote.py                          # primitivo puro
└── docs/
    ├── diagrama_repo_pattern.html              # arquitectura del repo (data_boletas)
    ├── formato_data_boletas_audit.html         # contrato del audit log (data_boletas)
    ├── diagrama_seguimiento_pueblo.html        # arquitectura event-sourced (5 capas)
    ├── diagrama_flujo_seguimiento_pueblo.html  # flujo de 5 segundos (LEE/GENERA por paso)
    └── formato_seguimiento_pueblo.html         # contrato del registro de eventos
```

## Reglas para contribuir aquí

- **No agregar lógica de negocio de un módulo específico.** Si necesitas "cómo se procesa un reclamo", eso vive en `4b_reclamos`.
- **Mantén la API estable.** Cambiar la firma de una función en `shared/` rompe potencialmente varios módulos. Si necesitas cambiar algo, agrega una nueva función o un parámetro opcional retrocompatible.
- **Documenta el contrato antes del código.** Especialmente para servicios con persistencia (como el repo): primero el HTML de formato del output, después el código.
