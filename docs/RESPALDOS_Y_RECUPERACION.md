# Respaldos y recuperación — dónde está cada cosa y cómo se recupera

Referencia permanente (no un evento activo — para eso está `LEER_ANTES.md`).
Escrito el 2026-08-12 a raíz de dos sustos, uno falso y uno real:

- **falso:** las mesas de julio estaban vacías y se leyó como pérdida de datos.
  Era el estado normal de un ciclo cerrado — `7_cierre` las había archivado en
  `7_cierre/archivo/2026-07/`, donde estuvieron todo el tiempo (§4).
- **real:** correr `pytest 4_pagos/efectivo/tests` sobrescribió mesa_1 y mesa_2
  del ciclo EN CURSO con las fixtures de los tests (§5). Ya está arreglado.

La lección común a los dos: **antes de concluir que algo se perdió, buscar dónde
el sistema ya lo guarda** — y antes de correr algo, saber si escribe.

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

### Copias sueltas del repo (sin git)

```
jass_system - 5           18/07/2026    mesas de julio, 374 filas — pero SIN las
jass_system - copia (3)   15/07/2026      correcciones manuales del 21/07
jass_system - copia-2     11/07/2026    373 filas, hasta el 07/07
jass_system - agosto1     31/07/2026
jass_system - agosto2     02/08/2026
```

Son fotos a mitad de ciclo: útiles como red, peligrosas como fuente de verdad
(ver "Restaurar mal tiene un costo propio" en §4).

No son basura, pero **no son el respaldo oficial de las mesas** — ese es
`7_cierre/archivo/<mes>/` (ver abajo). Sirven como red extra entre commits.

---

## 3. Inventario de respaldos automáticos

**El primero de la lista es el que hay que mirar antes que cualquier otro.**

| Carpeta | Qué guarda | Quién la escribe |
|---|---|---|
| **`7_cierre/archivo/<mes>/`** | **las 7 mesas + planilla_cobrado + arrastres + correcciones_lote del ciclo, tal como quedaron al cerrar** | **`7_cierre/consolidar_cierre.py:paso2_cosechar`** |
| `4_pagos/efectivo/backup/` | `discrepancias_*`, migraciones, y desde 12/08 `mesas_pre_reset_<ts>/` | `main.py`, `utils_templates.respaldar_si_tiene_datos` |
| `4b_reclamos/backup/reclamos/` | `reclamos_<mes>_<ts>.xlsx` antes de cada regeneración | `4b_reclamos/main.py:_backup_con_timestamp` |
| `shared/backups_ledger/` | `seguimiento_pueblo_pre_*`, `registro_cortes_pre_*` | tools del ledger, antes de cada mutación |
| `3_boletas/backup/DATA_boletas/` | snapshot antes de cada `apply_correction` | `shared/data_boletas_repo` |
| `shared/planilla_mes/backups/` | planillas previas | `2_planilla` |
| `4_pagos/yape/motor_matching/Correcciones/` | `pendientes_<ts>.xlsx` | `motor_matching/main.py` |

---

## 4. Las mesas vacías de un ciclo cerrado son NORMALES — no un borrado

Este apartado empezó siendo el relato de un incidente. **Era una alarma falsa**,
y la corrección vale más que el relato original.

```
26/07 15:33:22   7_cierre cierra el ciclo 2026-07 (estado_ciclo.json lo registra
                 al segundo). Su paso 4 hace exactamente lo que debe:

                   paso2_cosechar  copia las 7 mesas a 7_cierre/archivo/2026-07/
                   paso3_freeze    congela el ciclo
                   paso4_limpiar   recién entonces las resetea a template

                 y su guarda es justamente esa: solo resetea lo que verificó
                 que ya está archivado ("LIMPIAR — omitido (no hay cosecha
                 confirmada)" si falta).

12/08            al correr buscar_pago aparece "mesas de julio = 0 filas" y se
                 interpreta como pérdida de datos. NO LO ERA: el respaldo estaba
                 en 7_cierre/archivo/2026-07/, en los DOS repos (Julio y activo).
                 No se buscó ahí.
```

**Regla que sale de esto:** si las mesas de un ciclo cerrado están vacías, eso es
el estado esperado. La fuente de ese ciclo vive en `7_cierre/archivo/<mes>/`.
Antes de gritar pérdida de datos, mirar ahí.

### Restaurar mal tiene un costo propio

En la recuperación del 12/08 se restauró desde `jass_system - 5` (copia del
18/07) en vez del archivo oficial (26/07). Las dos tienen 374 filas, pero entre
el 18 y el 21/07 alguien había corregido dos filas a mano:

```
G-23   copia 18/07     MONTO=71  EFECTIVO=0   YAPE=71
       archivo 26/07   MONTO=71  EFECTIVO=22  YAPE=49
                       "Se dividió su Yape en efectivo y yape para que cuadre
                        con la segregación del reporte del banco"
```

Restaurar la copia vieja revivió el estado previo a esa corrección, y
`verificar_yape` volvió a reportar G-23 y F-14 como problemas ya resueltos.
**Restaurar de más atrás no es conservador: reintroduce trabajo ya deshecho.**

### Orden de preferencia para restaurar mesas

1. **`7_cierre/archivo/<mes>/`** — el archivo oficial del cierre. Es el estado
   final del ciclo, con las correcciones manuales aplicadas.
2. **git**, si hay un commit posterior a la última jornada de cobranza.
3. **copia suelta del repo**, como último recurso.

Sobre git: **sirve o no según el ciclo.**

| Fuente | Julio | Agosto |
|---|---|---|
| git | `42dee24` (06/07): 263 filas y **`FECHA` vacía en todas** — las mesas se siguieron llenando después y no se volvieron a commitear | `5e3bdf3` (06/08): posterior a la cobranza del 01-02/08 → completo |
| copia suelta | `jass_system - 5` (18/07): 374 filas con `FECHA`, pero **sin las correcciones del 21/07** | no hay copia de agosto |
| archivo de cierre | `7_cierre/archivo/2026-07/`: 374 filas **con** las correcciones | el ciclo aún no cerró |

Después de restaurar, verificar que las fechas correspondan al ciclo:

```python
# cada mesa debe tener FECHA del ciclo al que pertenece
2026-06 → todas 2026-06    2026-07 → todas 2026-07    2026-08 → 01/08 y 02/08
```

---

## 5. Peligros conocidos

### `pytest 4_pagos/efectivo/tests` escribía sobre el `inputs/` REAL — arreglado (12/08)

Los tres archivos de test se escribieron para correr como **script**: su `main()`
llama `_setup()` antes de cada test. Bajo pytest ese `main()` nunca corre, así
que los módulos quedaban apuntando a las rutas reales y las fixtures se escribían
sobre las mesas del ciclo en curso. El 12/08 destruyó mesa_1 (59→2 filas) y
mesa_2 (106→2) de agosto; se recuperaron con `git checkout HEAD --`.

El síntoma en el output era inconfundible y se leyó mal: `assert (162 == 1)` — el
test contaba las 162 filas reales en vez de las de su fixture. Se interpretó como
"falla preexistente ajena" en vez de "este test toca datos reales".

Arreglado con `tests/conftest.py`: un fixture `autouse` corre el
`_setup()`/`_teardown()` del propio archivo antes y después de cada test, y
**verifica que las rutas del módulo no apunten al repo real**, fallando ruidoso
si lo hacen. Un test nuevo que se agregue sin aislar falla antes de escribir.

Estado: 64 tests en verde, 0 archivos reales modificados (verificado con hashes
antes/después).

**Si escribís tests nuevos en este repo:** que no toquen datos reales. El patrón
seguro está en `4b_reclamos/tests/test_verificar_yape.py` — todo entra por
`monkeypatch` con DataFrames construidos en el propio test, sin leer ni escribir
un solo `.xlsx` del repo.

### `crear_templates.py` — el que SÍ podía borrar sin red (ya tiene guarda, 12/08)

Ojo: **este script no fue el culpable del 26/07** (fue `7_cierre`, que archivó
antes — ver §4). Pero es el único que reseteaba mesas sin archivar ni respaldar:
un `wb.save()` directo sobre las 7. Si alguien lo corría a mitad de ciclo, ahí sí
se perdía todo.

Desde el 12/08 se niega si alguna mesa tiene cobros escritos, listando cuántas
filas se perderían, y exige `--force`. Además `utils_templates.crear_mesa_vacio`
respalda siempre a `backup/mesas_pre_reset_<ts>/` antes de pisar — la guarda vive
en el primitivo compartido para que la tengan los dos que resetean mesas, no solo
el que se acuerde.

`7_cierre` no necesitaba la guarda (archiva y pide consentimiento), pero el
respaldo extra no le estorba.

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
