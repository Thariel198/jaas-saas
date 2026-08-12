# Respaldos y recuperación — dónde está cada cosa y cómo se recupera

Referencia permanente (no un evento activo — para eso está `LEER_ANTES.md`).
Escrito el 2026-08-12, después de recuperar las mesas de julio que se habían
borrado el 26/07 y de casi perder las de agosto corriendo tests.

---

## 1. La cadena del dato — qué se puede reconstruir y qué no

```
mesa_N.xlsx  ──►  pagos_efectivo.xlsx  ──►  planilla_cobrado.xlsx  ──►  ledger
 (FUENTE)          (consolidado)             (aplicado)
 lo que el         4_pagos/efectivo          5_cobranza
 cobrador
 escribió
```

**La fuente no se puede reconstruir desde el consolidado.** `recuperar_mesas.py`
lo dice en su propio encabezado: rearma las mesas desde `pagos_efectivo.xlsx`
pero con `MONTO_EFECTIVO = MONTO` y `MONTO_YAPE = 0`, "corregir a mano después".
Se pierden:

| Dato | Vive solo en la mesa | Por qué importa |
|---|---|---|
| Split `MONTO_EFECTIVO` / `MONTO_YAPE` | sí | sin él no se puede detectar el yape que nunca entró a la cuenta |
| `COMENTARIO` del cobrador | sí | es donde el vecino dice "ya pagué", "yape a Janet", etc. |
| `COBRADOR` + `MESA` + `HOJA` por fila | parcial | a quién preguntarle cuando un pago no aparece |
| `CATEGORIA` (reclamo / compromiso) | sí | distingue visita sin cobro de cobro real |

Consecuencia práctica: **si se pierden las mesas de un ciclo, ese ciclo deja de
ser auditable** aunque la plata esté bien contabilizada.

---

## 2. Dónde vive cada ciclo

Cada ciclo cerrado se congela en su propia carpeta (el pipeline se copia entero
por mes). El código que lo resuelve es `rh.REPOS_CICLO_CERRADO` en
`4b_reclamos/reporte_historico.py`.

| Ciclo | Ruta | Git propio |
|---|---|---|
| 2026-06 | `PycharmProjects/Junio/jass_system - junio` | sí |
| 2026-07 | `PycharmProjects/Julio/jass_system - Julio` | sí |
| activo | `PycharmProjects/jass_system` | sí |

Cuando cierra un mes nuevo se agrega su línea a `REPOS_CICLO_CERRADO` — no hace
falta tocar ninguna función.

**Ojo con los nombres de archivo:** no son uniformes entre repos.
`planilla_cobrado_2026-06.xlsx` (junio) vs `planilla_cobrado_julio.xlsx` (julio,
mes en español) vs `planilla_cobrado.xlsx` (sin periodo, en el activo).
`shared/ciclo.py:resolver()` acepta las tres formas.

### Copias sueltas del repo (sin git, pero salvaron julio)

```
jass_system - 5           18/07/2026    mesas de julio, 374 filas CON fecha
jass_system - copia (3)   15/07/2026    idéntica a la anterior (mismo hash)
jass_system - copia-2     11/07/2026    373 filas, hasta el 07/07
jass_system - agosto1     31/07/2026
jass_system - agosto2     02/08/2026
```

No son basura: son el único respaldo de la fuente entre commits.

---

## 3. Inventario de respaldos automáticos

| Carpeta | Qué guarda | Quién la escribe |
|---|---|---|
| `4_pagos/efectivo/backup/` | `discrepancias_*`, migraciones, y desde hoy `mesas_pre_reset_<ts>/` | `main.py`, `utils_templates.respaldar_si_tiene_datos` |
| `4b_reclamos/backup/reclamos/` | `reclamos_<mes>_<ts>.xlsx` antes de cada regeneración | `4b_reclamos/main.py:_backup_con_timestamp` |
| `shared/backups_ledger/` | `seguimiento_pueblo_pre_*`, `registro_cortes_pre_*` | tools del ledger, antes de cada mutación |
| `3_boletas/backup/DATA_boletas/` | snapshot antes de cada `apply_correction` | `shared/data_boletas_repo` |
| `shared/planilla_mes/backups/` | planillas previas | `2_planilla` |
| `4_pagos/yape/motor_matching/Correcciones/` | `pendientes_<ts>.xlsx` | `motor_matching/main.py` |

---

## 4. El incidente del 26/07/2026 — cómo se perdieron y recuperaron las mesas de julio

```
26/07 15:33   se corre 4_pagos/efectivo/crear_templates.py para preparar agosto
              → hacía wb.save() sobre las 7 mesas, sin guarda ni backup
              → borra las 374 filas de julio escritas a mano
              → las 7 quedan en 16.774 bytes (template pristino)

              NO se perdió plata: los 442 pagos ya estaban consolidados.
              SÍ se perdió la fuente → julio dejó de ser auditable.

12/08         se detecta al correr buscar_pago: "mesas de julio = 0 filas"
              se recupera desde "jass_system - 5" (copia del 18/07)
              → 374 filas, 18 con yape, 146 con comentario
```

### Por qué git NO sirvió para julio y sí para agosto

Esta es la parte que hay que recordar: **la fuente correcta depende del ciclo.**

| Fuente | Julio | Agosto |
|---|---|---|
| git | `42dee24` (06/07) tiene 263 filas y **`FECHA` vacía en todas** — las mesas se siguieron llenando después y no se volvieron a commitear | `5e3bdf3` (06/08) es posterior a la cobranza del 01-02/08 → completo |
| copia suelta | `jass_system - 5` (18/07): 374 filas con `FECHA` | no hay copia de agosto |

**Antes de restaurar, comparar ambas fuentes.** Y después verificar que las
fechas correspondan al ciclo:

```python
# cada mesa debe tener FECHA del ciclo al que pertenece
2026-06 → todas 2026-06    2026-07 → todas 2026-07    2026-08 → 01/08 y 02/08
```

---

## 5. Peligros conocidos

### `pytest 4_pagos/efectivo/tests` escribe sobre el `inputs/` REAL

**No correr esa suite con un ciclo en curso.** No está aislada: escribe sus
fixtures sobre `4_pagos/efectivo/inputs/mesa_N.xlsx`. El 12/08 destruyó mesa_1
(59→2 filas) y mesa_2 (106→2) de agosto; se recuperaron con
`git checkout HEAD -- 4_pagos/efectivo/inputs/mesa_1.xlsx mesa_2.xlsx`.

El síntoma en el output del test es inconfundible: `assert (162 == 1)` — el test
está contando las filas reales en vez de las de su fixture.

Arreglo pendiente: monkeypatchear `INPUTS_DIR` a un `tmp_path` en el setup.

### `crear_templates.py` — ya tiene guarda (12/08)

Ahora se niega si alguna mesa tiene cobros escritos, listando cuántas filas se
perderían, y exige `--force`. Además `utils_templates.crear_mesa_vacio` respalda
siempre a `backup/mesas_pre_reset_<ts>/` antes de pisar — eso cubre también a
`7_cierre`, que es el otro que resetea mesas (ese sí legítimamente: archiva
primero y pide consentimiento).

### `git stash` con `.xlsx` abiertos en Excel

Falla a mitad de camino y deja archivos revertidos a HEAD sin avisar. Pasó el
12/08: revirtió 32 archivos. Se recuperaron desde el propio stash. **No usar
`git stash` en este repo**: para comparar contra HEAD, usar
`git show HEAD:<archivo>` sobre un archivo puntual.

### Commitear las mesas seguido

El agujero de julio existe porque entre el 06/07 y el 26/07 no hubo ningún
commit que incluyera las mesas. Si se commitean al terminar cada jornada de
cobranza, git alcanza y las copias sueltas dejan de ser críticas.

---

## 6. Cómo verificar una restauración

```python
# desde la raíz del repo activo
py -c "
import sys; sys.path.insert(0,'4b_reclamos/herramienta'); import verificar_yape as vy
for m in ['2026-06','2026-07','2026-08']:
    d = vy.pagos_de_mesas(m)
    print(m, len(d), 'filas |', int((d['YAPE']>0).sum()), 'con yape')
"
```

Y contrastar contra `4_pagos/efectivo/outputs/pagos_efectivo*.xlsx` del mismo
ciclo: la mesa debe tener **más o igual** filas que el consolidado (el
consolidado descarta blancos y filas con `MONTO=0`). Si tiene bastante menos,
falta data.

Herramienta que hace ese cruce sola: `4b_reclamos/herramienta/verificar_yape.py`.
