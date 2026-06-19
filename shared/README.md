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

## Documentación

```
shared/
├── README.md                              (este archivo)
├── data_boletas_repo.py                   # módulo
├── data_boletas_audit.xlsx                # audit log (output)
├── utils_lote.py                          # primitivo puro
└── docs/
    ├── diagrama_repo_pattern.html         # arquitectura del repo
    └── formato_data_boletas_audit.html    # contrato del audit log
```

## Reglas para contribuir aquí

- **No agregar lógica de negocio de un módulo específico.** Si necesitas "cómo se procesa un reclamo", eso vive en `4b_reclamos`.
- **Mantén la API estable.** Cambiar la firma de una función en `shared/` rompe potencialmente varios módulos. Si necesitas cambiar algo, agrega una nueva función o un parámetro opcional retrocompatible.
- **Documenta el contrato antes del código.** Especialmente para servicios con persistencia (como el repo): primero el HTML de formato del output, después el código.
