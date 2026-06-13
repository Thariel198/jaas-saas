# 4_pagos/efectivo — Módulo de pagos en efectivo

## Qué hace

Consolida los registros de cobranza en efectivo levantados en campo por los cobradores
de cada mesa. Compara los registros dentro de una misma mesa para detectar coincidencias
o discrepancias, y entrega una lista limpia de pagos confirmados lista para `5_cobranza`.

## Cuándo se corre

Una vez por mes, después de que todos los cobradores han entregado sus registros y antes
de correr `5_cobranza`. Si quedan discrepancias sin resolver, el módulo puede correrse
de nuevo hasta que `discrepancias.xlsx` desaparezca.

---

## Estructura de carpetas

```
4_pagos/efectivo/
├── inputs/
│   ├── mesa_1.xlsx          ← hojas: registro_1, [registro_2], [registro_3]
│   ├── mesa_2.xlsx
│   ├── ...
│   └── mesa_7.xlsx
├── outputs/
│   ├── pagos_efectivo.xlsx  ← resultado limpio → 5_cobranza
│   └── discrepancias.xlsx   ← temporal, desaparece al resolverse todo
├── trazabilidad/
│   ├── consolidado_YYYY-MM.xlsx     ← todo lo procesado (permanente)
│   └── incidencias_YYYY-MM.xlsx     ← anomalías del mes (permanente)
├── backup/
│   └── migracion_YYYY-MM/           ← archivos anteriores al rediseño
├── docs/
│   ├── diagrama_efectivo.html
│   ├── arquitectura_efectivo.html
│   └── formato_pagos_efectivo.html
├── tests/
├── main.py
└── crear_templates.py
```

---

## Formato de inputs — mesa_N.xlsx

Cada archivo representa una mesa física de cobranza. Puede tener 1, 2 o 3 hojas
nombradas `registro_1`, `registro_2`, `registro_3`.

Columnas de cada hoja (todas requeridas salvo COMENTARIO):

| Columna          | Tipo    | Descripción                              |
|------------------|---------|------------------------------------------|
| COBRADOR         | texto   | Nombre de quien cobra                    |
| FECHA_REGISTRO   | fecha   | Día en que se llenó el registro          |
| MZ               | texto   | Manzana del predio                       |
| LT               | texto   | Lote del predio                          |
| MONTO            | decimal | Monto cobrado en soles                   |
| FECHA            | fecha   | Fecha del pago (puede diferir del cobro) |
| COMENTARIO       | texto   | Nota libre, opcional                     |

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

## Flujo paso a paso

```bash
# 1. Asegurarse de que los archivos mesa_N.xlsx están en inputs/
#    (cada archivo puede tener 1, 2 o 3 hojas)

# 2. Correr el módulo
python main.py

# 3. Revisar outputs/discrepancias.xlsx si existe
#    Llenar columna RESOLUCION en cada fila (acepta / corrige)

# 4. Volver a correr si había discrepancias
python main.py

# 5. Cuando no hay discrepancias, outputs/pagos_efectivo.xlsx está listo
#    → pasar a 5_cobranza
```

---

## Tabla de lifecycle

| Archivo | Tipo | Cuándo se crea | Cuándo desaparece |
|---|---|---|---|
| `outputs/pagos_efectivo.xlsx` | permanente | cada corrida | nunca (se sobreescribe) |
| `outputs/discrepancias.xlsx` | temporal | si hay discrepancias | cuando todas se resuelven |
| `trazabilidad/consolidado_YYYY-MM.xlsx` | permanente | cada corrida | nunca |
| `trazabilidad/incidencias_YYYY-MM.xlsx` | permanente | si hay anomalías | nunca |
| `inputs/mesa_N.xlsx` | manual — sagrado | el cobrador lo llena | nunca se borra sin backup |

---

## Lo que este módulo NO hace

- No calcula si el monto es correcto (eso es responsabilidad de `2_planilla` + `5_cobranza`)
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
