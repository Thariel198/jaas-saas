# RETOMAR — Abonos rezagados y revisión de H-21

**Fecha de corte de este documento:** 2026-08-15
**Propósito:** permitir que la siguiente sesión continúe sin volver a reconstruir el contexto de los abonos rezagados, el ledger `seguimiento_pueblo` y la corrida de agosto.

## Diagrama de continuidad

```text
ESTADO ACTUAL
    |
    +--> ledger vivo sin nueva corrida después de los últimos cambios
    |
    +--> manifest con confirmados, pero documentos antiguos aún dicen 18
    |
    +--> 8 filas REVISAR + 4 BLOQUEADAS
    |
    v
NO BORRAR PAGOS DEL LEDGER
    |
    v
validar por fuente: pago normal != abono rezagado
    |
    +--> H-21: conservar S/19 + aplicar S/50 separado si se confirma
    +--> N-6: revisar S/29 antes de aplicar S/70
    +--> F1-4: S/101 confirmado; descartar consolidado S/117 y candidato S/10
    +--> I-9/Q-5: no automatizar
    |
    v
backup -> 4_pagos final -> 5_cobranza --force -> 5b_validacion
    |
    v
6_corte: borrador -> publicada -> cierre
```

## 1. Objetivo de negocio

Estamos cerrando el ciclo `2026-08` y hay dinero que algunos vecinos pagaron en un ciclo anterior, pero que no llegó oportunamente a la caja de la JASS. Esos importes se recuperaron o identificaron después y se modelan como `ABONO_REZAGADO`.

El objetivo no es sustituir pagos normales ni corregir manualmente la vista. El objetivo es:

- aplicar cada abono rezagado una sola vez;
- conservar la trazabilidad de la fuente y del ciclo original;
- no duplicar pagos ya presentes en `seguimiento_pueblo.xlsx`;
- no eliminar físicamente eventos append-only;
- dejar una lista de corte de agosto defendible;
- cerrar agosto antes de hacer la reimputación histórica de `MULTA`, `ACUERDOS` y `CONVENIO`.

## 2. Regla contable central

```text
PAGO NORMAL en caja        = CLASE/SOURCE de COBRANZA
ABONO REZAGADO recuperado  = CLASE/SOURCE de ABONO_REZAGADO
REIMPUTACION de concepto   = CLASE REASIGNACION
EXONERACION                = no es plata
CORRECCION_SISTEMA         = no es plata
```

Un importe que aparece en la columna `PAGO` de una vista por concepto no necesariamente es el importe completo de una transacción. La cascada puede haber tomado un pago total y haberlo aplicado parcialmente a `ACUERDOS`, `CONVENIO`, `MULTA` u otro concepto. Por eso no se debe concluir que un pago de S/19 es el mismo dinero que un abono rezagado de S/50 solo porque ambos aparecen asociados al mismo predio.

## 3. Estado de archivos y cambios de trabajo

### Archivos principales

- `shared/abonos_rezagados.xlsx`: fuente clasificada de los abonos rezagados.
- `shared/seguimiento_pueblo.xlsx`: ledger vivo, hoja `Eventos`.
- `shared/vista_seguimiento_pueblo.xlsx`: vista por concepto y mes; no es el writer del ledger.
- `shared/vista_seguimiento_pueblo.pdf`: vista PDF generada.
- `5_cobranza/main.py`: orquestador transitorio de la cobranza y reconciliación.
- `5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json`: manifest operativo.
- `shared/reclasificar_abonos.py`: script que clasifica `Mapa_Abonos`.
- `README_PLAN_RECLAMOS_2026-08.md`: plan mayor de cierre agosto/septiembre.
- `Pendiente/abonos_rezagados_pendientes_2026-07.md`: bloqueados históricos.

### Cambios ya realizados

- El módulo `5_cobranza` dejó de depender de `BALDE` para decidir la cascada.
- La cascada calcula el pago normal y el abono juntos, pero los eventos deben permanecer separados.
- El pago normal se identifica como `COBRANZA`.
- El abono se identifica como `ABONO_REZAGADO`.
- Se agregó guard de manifest por:
  - `MZ`;
  - `LT`;
  - `MONTO`;
  - `MES_CICLO`;
  - `MES_ANO_APLICA`.
- El guard debe detener la corrida si hay una fila extra, faltante, duplicada o bloqueada.
- Se creó `shared/abonos_rezagados.xlsx` con hojas:
  - `Abonos_Raw`;
  - `Mapa_Abonos`;
  - `Categorias`.
- Se crearon pruebas para el manifest y la reconciliación.
- Pasaron previamente:

```powershell
py tests/test_abonos_manifest.py
py tests/test_reconciliacion_pueblo.py
py -m py_compile main.py
```

No declarar una nueva corrida como válida solo por esos tests: todavía falta ejecutar la cadena completa después de resolver las decisiones pendientes.

## 4. Conteos: tener cuidado con la documentación antigua

Hay una diferencia entre los documentos anteriores y el estado actual:

```text
documentos anteriores:
6 históricos de julio + 12 confirmados para agosto = 18

estado actual del mapa:
6 históricos de julio + 13 confirmados para agosto = 19
```

El registro adicional incorporado después fue `L-5 S/126`. Por eso el manifest/mapa actual puede tener 19 confirmados aunque `5_cobranza/README.md` y `README_PLAN_RECLAMOS_2026-08.md` todavía mencionen 18. Antes de cerrar la sesión siguiente hay que reconciliar ese conteo y actualizar la documentación que corresponda. No quitar `L-5` para hacer coincidir el número viejo.

Los 12 confirmados originalmente para agosto sumaban S/457 y son:

| MZ | LT | Monto | Ciclo origen | Aplica | Retenido por |
|---|---:|---:|---|---|---|
| D | 18 | S/32 | 2026-08 | 2026-08 | Yerald Romero |
| K | 12 | S/122 | 2026-08 | 2026-08 | Yerald Romero |
| X | 33 | S/89 | 2026-08 | 2026-08 | Yerald Romero |
| C | 3A | S/10 | 2026-08 | 2026-08 | Maximo Encarnacion |
| F1 | 8 | S/13 | 2026-08 | 2026-08 | Wagner Trujillo |
| P | 3 | S/33 | 2026-07 | 2026-08 | Yerald Romero |
| W | 5 | S/15 | 2026-07 | 2026-08 | Wagner Trujillo |
| F1 | 8 | S/14 | 2026-07 | 2026-08 | Wagner Trujillo |
| H1 | 15 | S/36 | 2026-07 | 2026-08 | Wagner Trujillo |
| E1 | 9 | S/3 | 2026-07 | 2026-08 | Wagner Trujillo |
| S | 2 | S/60 | 2026-07 | 2026-08 | Wagner Trujillo |
| P | 12 | S/30 | 2026-07 | 2026-08 | Wagner Trujillo |

El manifest actual además incluye `L-5 S/126`, sujeto a la verificación de que su limpieza no haya dejado un evento duplicado.

## 5. H-21: decisión importante de esta sesión

### Hechos observados en el ledger

En `shared/seguimiento_pueblo.xlsx`, hoja `Eventos`, se encontró:

```text
H-21 | ACUERDOS | 2026-06 | CARGO | S/50 | génesis
H-21 | CONVENIO | 2026-06 | CARGO | S/50 | génesis
H-21 | ACUERDOS | 2026-07 | PAGO  | S/19 | 5_cobranza
```

No se encontró todavía un evento `ABONO_REZAGADO` de S/50 para H-21 en el ledger vivo al momento de esta revisión.

La vista muestra para `H-21 / ACUERDOS`:

```text
2026-07: deuda S/50 | pago S/19 | saldo S/31
```

### Interpretación correcta

El S/19 es un pago de julio ya registrado. El S/50 del archivo de abonos rezagados es otro hecho que todavía debe validarse/aplicarse. No hay base para afirmar que el sistema “solo tomó 19 del abono”. Lo que sí se observa es que la cascada aplicó S/19 de un pago normal al concepto `ACUERDOS`.

La solución propuesta de borrar el pago S/19 **no es válida** con la evidencia actual.

La solución correcta, si se confirma que el documento de secretaria corresponde a un abono real independiente, es:

```text
conservar PAGO normal S/19 de julio
        +
registrar/aplicar ABONO_REZAGADO S/50 separado
        =
dos fuentes de dinero auditables
```

Si se demostrara que S/19 y S/50 son la misma transacción, no se debe borrar una fila directamente. Se debe registrar una reversión/ajuste enlazado a la transacción original, conservar el backup y volver a reconciliar por delta. La arquitectura del ledger es append-only.

### Evidencia adicional del reporte disponible

El reporte disponible en el workspace es `reporte_reimputacion_cascada_2026-07.pdf`, no el PDF 2026-08 adjunto. En su página de H-21 se ven referencias separadas de `ABONO REZ.` por S/50 y un pago Yape de S/36; la aplicación por concepto muestra S/19 en `ACUERDOS`. Esto refuerza que la columna por concepto no debe confundirse con el importe total de una fuente.

El PDF 2026-08 adjunto no quedó como archivo separado en el workspace. Si la siguiente sesión necesita validar una página concreta, volver a adjuntar el PDF o colocarlo en la raíz con un nombre identificable.

## 6. Casos `REVISAR`

Estos casos no deben entrar al manifest hasta que exista decisión y evidencia suficiente.

### H-21 — S/50

- El ledger tiene PAGO normal S/19 en julio aplicado a `ACUERDOS`.
- El abono S/50 no aparece aún como evento rezagado en el ledger vivo.
- Acción: conservar S/19; verificar la nota/fuente del S/50; aplicar el S/50 completo como `ABONO_REZAGADO` si es independiente.
- No borrar ni editar físicamente el S/19.

### N-6 — S/70

- La vista muestra un pago de julio de S/29 y saldo posterior en `ACUERDOS`.
- El abono rezagado de S/70 proviene de una nota de secretaria y está marcado para revisar por posible doble conteo.
- Acción: confirmar si el S/29 proviene de una cobranza real distinta del abono S/70.
- No borrar S/29 solo porque el abono sea de S/70.
- Si son dos fuentes independientes, conservar ambos y aplicar S/70 completo.

### F1-4 — cerrado el 2026-08-22

- El usuario confirmó visualmente que el pago correcto es S/101 y pertenece a agosto.
- La evidencia es `shared/abono_rezagado/5.jpeg`, `OCR_ID=IMG-005-R01`.
- El consolidado S/117 no provenía de una referencia de pago y el candidato S/10 no
  tenía comprobante independiente; ambos quedan reemplazados por una sola fila S/101.
- El evento append-only ya aplicado al ledger (`MULTA`, PAGO S/50) se conserva: S/101
  sigue cubriendo esos S/50 y no requiere reversión.

### F1-7 — S/30

- Figura como candidato de notas de secretaria.
- No hay comprobante independiente confirmado.
- Acción: confirmar con secretaria y revisar si el predio ya quedó resuelto manualmente.

### E-8 — S/50

- Figura como candidato de notas de secretaria.
- No debe asumirse que está pendiente solo porque la vista tenga saldo.
- Acción: verificar si hubo pago independiente, exoneración o resolución manual.

### I-9 — S/50

- Figura con riesgo de doble conteo por crédito/exceso previo.
- En reportes anteriores del predio aparecen abonos rezagados y movimientos que pueden confundirse con un crédito o una aplicación previa.
- Acción: revisar todos los eventos del predio y el origen de cada abono antes de aplicar este S/50.
- Prioridad alta: no automatizar.

### Q-5 — S/69 y S/45

- Tiene dos filas de abonos rezagados.
- `BALDE=mixto`.
- Hay traslape con resolución manual.
- Hay un monto sin balde aplicable.
- Acción: mantener fuera del manifest hasta definir qué parte corresponde a agua, qué parte a tanque/otro destino y qué parte ya está en el ledger.
- Prioridad máxima junto con I-9.

## 7. Casos bloqueados

Mantener fuera de la corrida automática:

| Predio | Abono | Motivo | Acción |
|---|---:|---|---|
| S-5 | S/71 | Colisión; saldo final necesita recalculo | No escribir |
| D-16 | S/85 | Colisión; saldo final necesita recalculo | No escribir |
| D1-6 | S/33 | Colisión; saldo final necesita recalculo | No escribir |
| Q-11 | S/17 | Colisión; saldo final necesita recalculo | No escribir |

Regla para resolverlos al cierre:

```text
saldo disponible = min(abono, saldo abierto real)
si saldo = 0      -> no registrar pago
si abono > saldo  -> registrar solo saldo; remanente queda pendiente
```

## 8. Casos manuales

El mapa actual tiene 13 filas `YA_RESUELTO_MANUAL`. Esos importes ya están representados en `seguimiento_pueblo.xlsx` o en decisiones manuales y no deben reaplicarse como `ABONO_REZAGADO`.

La lista incluye los casos históricos ya resueltos y `G-4`, que fue incorporado posteriormente. Los documentos que hablan de 12 manuales están desactualizados frente al `Mapa_Abonos` actual.

Regla:

```text
YA_RESUELTO_MANUAL
        |
        +--> no entra en 5_cobranza --force
        +--> no se elimina del ledger
        +--> queda pendiente integrar su resolución en la lista de corte
```

## 9. L-5 y limpieza reciente

Se trabajó `L-5 S/126` porque tenía una inconsistencia previa:

- se eliminó el par antiguo `PAGO S/8` + `AJUSTE -S/8`;
- la multa volvió a S/50;
- se agregó el abono rezagado S/126 para agosto;
- se crearon backups antes de la limpieza y antes del intento de corrida.

Backups relevantes:

- `shared/backups_ledger/seguimiento_pueblo_pre_L5_cleanup_20260815_093048.xlsx`
- `shared/backups_ledger/abonos_rezagados_pre_L5_126_20260815_093048.xlsx`
- `shared/backups_ledger/seguimiento_pueblo_pre_force_L5_126_20260815_093302.xlsx`
- `shared/backups_ledger/abonos_rezagados_pre_force_L5_126_20260815_093302.xlsx`

Antes de una nueva corrida, verificar si el intento posterior realmente dejó eventos correctos o si solo actualizó archivos de salida. No asumir que el backup anterior representa el estado actual.

## 10. Qué no hacer

- No borrar el PAGO S/19 de H-21.
- No borrar pagos parciales para “reemplazarlos” por un abono completo sin demostrar que son la misma transacción.
- No editar físicamente eventos append-only del ledger.
- No procesar `Q-5` automáticamente.
- No procesar los cuatro bloqueados.
- No procesar los 13 manuales como abonos nuevos.
- No usar `BALDE` para decidir la cascada general.
- No fusionar pago normal y abono en un solo evento.
- No correr `5_cobranza --force` mientras el manifest tenga filas inesperadas sin decisión.
- No generar lista de corte después de publicar sin cambiar la fase y revisar nuevamente.
- No ejecutar reimputación histórica antes de congelar/publicar/cerrar agosto.
- No declarar que el ledger está “limpio” solo porque los tests unitarios pasan.

## 11. Secuencia recomendada para la próxima sesión

### Paso 1 — Reconfirmar estado

```powershell
git status --short
git log -3 --oneline -- 5_cobranza/ shared/
py -m py_compile "5_cobranza/main.py"
```

Confirmar que no hay cambios nuevos del usuario que contradigan este documento. No revertir cambios ajenos.

### Paso 2 — Inspeccionar eventos por predio

Leer de `shared/seguimiento_pueblo.xlsx`, hoja `Eventos`, todos los eventos de:

```text
H-21, N-6, F1-4, F1-7, E-8, I-9, Q-5
```

Para cada uno, construir una tabla con:

```text
predio | mes | concepto | movimiento | monto | CLASE | SOURCE | AUDIT_REF | saldo posterior
```

No basta con mirar `vista_seguimiento_pueblo.xlsx`.

### Paso 3 — Cerrar decisiones humanas

Resolver en este orden:

1. H-21: confirmar S/50 como fuente independiente del S/19.
2. N-6: confirmar S/70 contra el pago S/29.
3. F1-4: separar documentalmente S/101 y S/10.
4. I-9: resolver posible crédito/exceso.
5. Q-5: separar agua, mixto y resolución manual.
6. F1-7 y E-8: obtener confirmación de secretaria/evidencia.

Registrar cada decisión en `shared/abonos_rezagados.xlsx` y, si cambia el manifest, en `5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json`.

### Paso 4 — Verificar el manifest actual

```powershell
py "shared/reclasificar_abonos.py"
py "5_cobranza/tests/test_abonos_manifest.py"
```

Revisar especialmente:

- si contiene 18 o 19 confirmados;
- si `L-5` está incluido una sola vez;
- si H-21, N-6, F1-4, F1-7, E-8, I-9 y Q-5 siguen fuera;
- si los cuatro bloqueados siguen fuera;
- si los manuales no tienen `REF_MANIFEST` activo.

### Paso 5 — Backup y corrida

Solo cuando `4_pagos` esté cerrado y las decisiones anteriores estén registradas:

```powershell
$env:PYTHONIOENCODING='utf-8'
$env:PYTHONUTF8='1'
py "5_cobranza/main.py" --force
```

Antes de la corrida, crear backups identificables de:

- `shared/seguimiento_pueblo.xlsx`;
- `shared/abonos_rezagados.xlsx`;
- outputs de `5_cobranza` que se regeneren.

### Paso 6 — Validar después de la corrida

Comprobar:

- pagos normales con `CLASE=COBRANZA`;
- abonos con `CLASE=ABONO_REZAGADO`;
- `SOURCE` y `AUDIT_REF` separados;
- exactamente una fila por abono confirmado;
- cero eventos para los cuatro bloqueados;
- cero eventos nuevos para manuales;
- H-21 con S/19 normal conservado y S/50 separado;
- no saldos negativos inesperados;
- conservación de pagos y deudas.

Luego ejecutar la validación de `5b_validacion`. Si hay alertas, detenerse antes de `6_corte`.

## 12. Pipeline posterior

```text
4_pagos final
    |
    v
5_cobranza --force
    |
    v
5b_validacion = OK
    |
    v
6_corte/generar_lista.py
    |
    v
lista BORRADOR
    |
    v
revisión y publicación
    |
    v
cierre/foto inmutable de agosto
    |
    v
reporte de reimputación post-cierre
    |
    v
asientos REASIGNACION
```

La reimputación `CONVENIO -> ACUERDOS -> MULTA` es posterior al cierre y no debe confundirse con la carga de abonos rezagados. La simulación debe conservar:

- deuda total;
- pagos totales;
- cero saldos negativos indebidos;
- trazabilidad por predio, concepto y fuente.

## 13. Estado final al cerrar esta sesión

```text
IMPLEMENTADO
  guard de manifest
  separación COBRANZA/ABONO_REZAGADO
  cascada normal + abono
  clasificación de 44 filas
  backups de L-5
  tests unitarios previos OK

NO EJECUTADO DESPUÉS DE LA ÚLTIMA REVISIÓN
  nueva corrida completa de 5_cobranza --force
  nueva validación 5b
  lista de corte final

REQUIERE DECISIÓN HUMANA
  H-21, N-6, F1-4, F1-7, E-8, I-9, Q-5

FUERA DE LA CORRIDA
  13 manuales
  S-5, D-16, D1-6, Q-11
```

## 14. Archivos de referencia y advertencias

- `5_cobranza/README.md` todavía describe el manifest como de 18 filas; confrontar con el mapa actual de 19.
- `README_PLAN_RECLAMOS_2026-08.md` todavía describe 12 confirmados y 13 manuales como pendientes de integración; mantenerlo sincronizado cuando se cierre el conteo.
- `RETOMAR_ABONOS_REZAGADOS_2026-08-14.md` contiene el checkpoint anterior; este documento es posterior y debe prevalecer para H-21 y para el conteo 18/19.
- `reporte_reimputacion_cascada_2026-07.pdf` es el PDF disponible localmente.
- El PDF 2026-08 mencionado en la conversación no está en la raíz ni identificado como archivo local al cerrar esta sesión.
