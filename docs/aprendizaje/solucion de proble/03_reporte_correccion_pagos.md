# Problema 03 - Reporte para correccion de pagos

## Problema

Cuando se corrige un pago de usuario, la informacion esta repartida entre el
historial mensual, el ledger y las fuentes de pago. Para revisar un caso hay que
abrir varias planillas y reconstruir manualmente:

```text
usuario
   ├── deuda mensual
   ├── pago aplicado por concepto
   └── referencia del pago
```

## Lotes del caso de correccion

Estos son los 7 lotes de la interseccion `lista de corte ∩ abonos rezagados`.
La clasificacion indica que se debe revisar en el mini pipeline antes de tocar el
ledger real.

| Lote | Abono fuente | Filas | Clasificacion | Problema a corregir | Accion previa al ledger real |
|---|---:|---:|---|---|---|
| `I-9` | S/136 → **S/133** | 2 | CONFIRMADO + AJUSTE | Secretaria declara "al dia"; Wagner S/86 + Secretaria S/50 no cuadra con el saldo pendiente S/133 | Conservar ambas filas y ajustar Secretaria a S/47 |
| `L-5` | S/126 | 1 | CONFIRMADO | Revisar aplicacion entre convenio, multa y acuerdos | Probar cascada y validar saldo |
| `P-12` | S/30 | 1 | CONFIRMADO | Corregir el pago de convenio y validar que no reaparezca el bug de signo | Probar cascada y validar saldo |
| `P-3` | S/33 | 1 | CONFIRMADO | Separar abono rezagado de la deuda vigente del ciclo | Probar cascada y validar saldo |
| `Q-5` | S/114 | 2 | REVISAR | Hay dos filas; una parte requiere confirmar antes de aplicarse | Separar filas, decidir monto y recién probar |
| `S-2` | S/60 | 1 | CONFIRMADO | Revisar aplicacion al concepto de campo/acuerdos | Probar cascada y validar saldo |
| `W-5` | S/15 | 1 | CONFIRMADO | Secretaria indica convenio cancelado; revisar campo y convenio | Probar cascada y validar saldo |

```text
7 lotes identificados
        |
        +--> 5 confirmados: L-5, P-12, P-3, S-2, W-5
        +--> 1 confirmado con ajuste: I-9
        +--> 1 en revision por dos filas: Q-5
        |
        v
mini pipeline aislado → aprobacion por lote → ledger real
```

## Solucion

Crear primero un reporte individual por predio. Para cada mes se imprimen dos
filas consecutivas, sin eliminar meses:

```text
MES
  DEUDA  -> Consumo · Mant. · Mes ant. · Corte · Convenio · Multa · Acuerdos · Total
  PAGO   -> Consumo · Mant. · Mes ant. · Corte · Convenio · Multa · Acuerdos · Total
```

La referencia del pago se conserva como bloque independiente con mes, medio,
fecha/hora y monto. Si el historial y las referencias no caben en una hoja, el
bloque de referencias pasa a una segunda hoja; no se fuerza una segunda hoja
cuando el contenido cabe en la primera.

## Contrato del reporte

| Bloque | Contenido |
|---|---|
| Deuda | Importe que correspondia al predio en cada concepto y mes |
| Pago | Importe del usuario distribuido por concepto y mes |
| Referencia | Mes, medio, fecha/hora y monto de cada pago identificado |
| Saldo vigente | Saldo pendiente de multa, acuerdos y convenio al cierre |

El reporte es de solo lectura. No corrige pagos ni escribe el ledger; sirve para
auditar y preparar la correccion autorizada.

## Muestra de diseño

```text
referencias/03_reporte_predio_I-3.pdf
```

La muestra contiene un solo predio y conserva todo el historial mensual. La hoja
del predio combina deuda, pago y referencias cuando el espacio disponible lo
permite.

## Prueba

```text
PDF de muestra generado para I-3
4 paginas totales del reporte
1 hoja para el predio: historial + referencia de pagos
sin escritura en shared/seguimiento_pueblo.xlsx
```

## Estado

**RESUELTO como diseño y muestra de referencia.**

La corrida masiva queda pendiente de una aprobacion visual de esta muestra.

## Caso real de correccion de lote: I-9

El mini pipeline se usa aqui como aislamiento del lote: reproduce la cascada con
una copia temporal y no modifica el ledger real. No es una solucion del problema
de tiempo por si sola; en este caso permite aprobar los montos antes de corregir
los 7 predios en conjunto.

```text
I-9: deuda foto real S/141
        - pago ya registrado S/8
        = saldo pendiente S/133

Wagner       S/86   (abono del ciclo 2026-06)
Secretaria   S/47   (declaracion del ciclo 2026-07)
             -----
Abono total  S/133
        ↓
Saldo final esperado S/0 · CANCELADO
```

### Pasos, montos y sitios

1. **Conservar las dos filas historicas** en `shared/abonos_rezagados.xlsx`:
   - `I-9` · Wagner · `S/86` · `MES_CICLO=2026-06`.
   - `I-9` · Secretaria · cambiar solo `S/50` a `S/47` · `MES_CICLO=2026-07`.
2. **Aplicar ambas en `2026-08`**: cambiar `MES_ANO_APLICA` de las dos filas a `2026-08`; no cambiar sus ciclos originales.
3. **Actualizar el guard** en `5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json` con ambas filas, sus montos, ciclos, aplicacion y estado `CONFIRMADO`.
4. **No borrar cargos del ledger**: `I-9` conserva MULTA `S/50` y ACUERDOS `S/75` de junio.
5. **No duplicar el pago normal**: el pago real de `S/8` permanece separado del abono rezagado.
6. **Correr el mini pipeline aislado** con `5_cobranza/tests/generar_mini_corrida_abonos.py` y revisar `mini_resultado_cascada.xlsx` y `mini_ledger_predicho.xlsx`.
7. **Aceptar para I-9 solo este resultado**: deuda `S/141`, pago previo `S/8`, abonos `S/133`, pago total `S/141`, saldo `S/0`, estado `CANCELADO`.
8. **Verificar la imputacion**: `S/8` a MES_ANTERIOR, `S/75` a ACUERDOS, `S/50` a MULTA y CONVENIO `S/0`.
9. **Revisar los otros 6 predios**. Si un mini saldo no coincide, no correr el ledger real.
10. **Crear backup inmediato** de `shared/seguimiento_pueblo.xlsx`, `shared/abonos_rezagados.xlsx`, el manifest y outputs de `5_cobranza`.
11. **Ejecutar una sola reconciliacion real** mediante `5_cobranza/main.py`; no escribir pagos manualmente en el Excel.
12. **Validar el ledger**: eventos append-only, `SOURCE=abonos_rezagados`, `CLASE=ABONO_REZAGADO`, `AUDIT_REF` distinto para Wagner y Secretaria y saldo cero para I-9.
13. **Regenerar y revisar el lote de 7** con `5b_validacion`, la lista de corte y el reporte de foto real.

```text
Fuente de abonos       → shared/abonos_rezagados.xlsx
Guard                   → 5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json
Mini pipeline           → 5_cobranza/tests/generar_mini_corrida_abonos.py
Writer unico            → shared/seguimiento_repo.py
Reconciliacion          → 5_cobranza/main.py
Ledger real             → shared/seguimiento_pueblo.xlsx
Validacion              → 5b_validacion/ + 4b_reclamos/reporte_deuda_ledger.py
```

La prueba mini ya paso sin modificar archivos reales: Wagner `S/86` + Secretaria
`S/47` = `S/133`; deuda `S/141`; pago total `S/141`; saldo final `S/0`.

## Soluciones mini aprobadas: L-5 y S-2

Estas soluciones siguen el mismo contrato: primero se prueban en copia aislada y
solo despues se preparan para la reconciliacion append-only del ledger real.

### L-5

```text
deuda foto real:       S/176
abono rezagado:        S/126
orden vigente:         agua -> convenio -> acuerdos -> multa
saldo final esperado:  S/50
concepto pendiente:    MULTA
```

El `PAGO S/60` de convenio y el `PAGO S/50` de acuerdos que se habian escrito en
agosto eran una aplicacion prematura y fueron retirados del ledger. La fuente
`abonos_rezagados.xlsx` S/126 queda pendiente de la corrida real.

### S-2

```text
deuda foto real:       S/107
abono rezagado:        S/60
origen:                Yape de Wagner Trujillo, retenido por Wagner
saldo final esperado:  S/47
concepto pendiente:    ACUERDOS
```

El `PAGO S/3` de agosto tambien era una aplicacion parcial prematura y fue
retirado del ledger. La multa S/20 no se exonero: tiene un pago normal de junio
(`SOURCE=5_cobranza`) y permanece en saldo cero.

### Estado de preparacion

```text
I-9  -> S/133 -> saldo S/0 · CANCELADO
L-5  -> S/126 -> saldo S/50 · MULTA
S-2  -> S/60  -> saldo S/47 · ACUERDOS
```

Los tres resultados estan probados en mini pipeline. Todavia no se deben ejecutar
en el ledger real hasta completar la aprobacion de los otros cuatro lotes y
resolver Q-5.
