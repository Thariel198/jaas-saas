# Problema 01 - Pruebas de 20 minutos

## Problema

Cada vez que se ejecutaba `5_cobranza --force` habia que esperar aproximadamente
20 minutos para saber si una correccion de abonos rezagados habia funcionado.

El ciclo de prueba era demasiado lento:

```text
correccion
    ↓
5_cobranza --force sobre todo el universo
    ↓
espera de 20 minutos
    ↓
recién se podia revisar el resultado
```

## Causa

La correccion afectaba pocos lotes, pero la prueba procesaba todos los predios y
todos los inputs del sistema.

```text
universo real completo
    ├── lotes que no cambiaron
    ├── lotes que no participan en la prueba
    └── lotes de la correccion
```

Se estaba pagando el costo de procesar todo el universo para verificar una
correccion aislada.

## Solucion

Crear un mini-pipeline que use las mismas herramientas y calculos reales, pero
filtre los inputs al universo de los lotes que se quieren corregir.

```text
inputs reales del sistema
          ↓
filtro por (MZ, LT) objetivo
          ↓
misma carga de pagos
          ↓
misma cascada
          ↓
misma descomposicion
          ↓
proyeccion aislada
```

El mini-pipeline no escribe el ledger real ni los outputs productivos.
Cada nueva corrida respalda primero la corrida temporal anterior en
`backups_mini_corrida/` y luego reconstruye su carpeta de salida desde cero.
Esto evita perder la evidencia de ayer y evita reutilizar una proyeccion vieja.

## Que se conserva

- carga real de pagos Yape;
- carga real de pagos efectivo;
- planilla y overlays reales;
- carga de abonos rezagados;
- calculo real;
- descomposicion real;
- proyeccion de eventos del ledger.

## Resultado

```text
5_cobranza --force completo  → aproximadamente 20 minutos
mini-pipeline filtrado       → aproximadamente 2 minutos
```

Tambien permitio detectar la diferencia entre claves `(MZ, LT)` y la clave textual
`MZ-LT`.

## Ejemplo vivo: los 7 lotes de la lista de corte

El caso actual se aisla con la interseccion real:

```text
lista de corte ∩ abonos rezagados
    ↓
I-9, L-5, P-12, P-3, Q-5, S-2, W-5
```

El runner `5_cobranza/tests/generar_mini_corrida_abonos.py` ahora toma esa
interseccion automaticamente. No procesa los otros abonos.

```powershell
$env:PYTHONIOENCODING='utf-8'
py "5_cobranza/tests/generar_mini_corrida_abonos.py"
```

Resultado de la corrida del 2026-08-16:

```text
8 filas de abono fuente
7 lotes calculados
aproximadamente 2 minutos
sin escritura en el ledger real
backup previo de la corrida temporal antes de reconstruir
```

## Test: esperado de secretaria vs resultado

Las notas de secretaria son la expectativa de negocio, no una prueba de que el
ledger ya este corregido. El test compara esa expectativa contra la salida real
del mini-pipeline:

```powershell
py "5_cobranza/tests/test_lista_corte_7_vs_notas.py"
```

### Mensajes usados como expectativa

Estos son los mensajes de `notas_2026-07.xlsx` que estamos usando como ejemplo
vivo para probar los siete lotes. La columna `FECHA` esta vacia en todos estos
registros; cuando aparece una fecha dentro de una resolucion, se conserva como
contexto y no se confunde con la fecha de la nota.

| Lote | Fecha en nota | Mensaje |
|---|---|---|
| I-9 | Sin fecha | "Al dia." |
| I-9 | Sin fecha | "Verificar todos sus pagos; pago e incluso reclamo porque salio en lista de corte; ella esta al dia." |
| L-5 | Sin fecha | "Revisar su convenio, iba a pagar 20 cada mes desde enero, tambien su multa 34? y techado 100?" |
| Q-5 | Sin fecha | "Esta al dia en todo; borrar solo, ponle su consumo de este mes." |
| S-2 | Sin fecha | "Revisar su campo." |
| W-5 | Sin fecha | "[CONVENIO] Cancelo; revisa campo y convenio." |
| P-3 | No aparece | No hay nota de secretaria localizada. |
| P-12 | No aparece | No hay nota de secretaria localizada. |

La resolucion de `Q-5` menciona `28/07/2026`, pero esa fecha esta escrita dentro
de la resolucion, no en la columna `FECHA`.

Expectativas fuertes actuales:

```text
I-9 → FUERA_DE_CORTE; abono esperado >= S/50
Q-5 → FUERA_DE_CORTE; abono esperado >= S/69
L-5 → PENDIENTE_CONFIRMACION
S-2 → PENDIENTE_CONFIRMACION
W-5 → PENDIENTE_CONFIRMACION
P-3 → SIN_NOTA
P-12 → SIN_NOTA
```

La primera corrida produjo estos resultados relevantes:

```text
I-9 → abono observado S/0; saldo S/66  → NO CUMPLE
Q-5 → abono observado S/0; saldo S/36  → NO CUMPLE
```

Esto es exactamente lo que el test debe mostrar: el mini-pipeline ya esta
aislado en el universo correcto, pero la logica de aplicacion todavia no cumple
la expectativa de las notas para I-9 y Q-5. Los cinco lotes restantes quedan
como contexto abierto, no como pagos automaticamente aprobados.

## Limite de la solucion

El mini-pipeline sirve para probar y corregir rapido. Antes de cerrar el trabajo:

```text
mini-pipeline aprobado
          ↓
backup del ledger
          ↓
aplicar la misma modificacion al ledger real
          ↓
5_cobranza --force completo
          ↓
validacion de outputs
```

## Estado

**RESUELTO como herramienta de prueba.**

La solucion reduce el tiempo de iteracion, pero no reemplaza la corrida completa
ni autoriza escribir automaticamente en `shared/seguimiento_pueblo.xlsx`.
