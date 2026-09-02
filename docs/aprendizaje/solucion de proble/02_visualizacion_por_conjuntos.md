# Problema 02 - Visualizar la lista de corte y sus excepciones

## Problema

La prioridad operativa es publicar la lista de corte. Sin embargo, algunos lotes
de esa lista tienen abonos rezagados u otros precursores que pueden reducir o
eliminar su saldo.

Si cada fuente se revisa por separado, es dificil responder rapidamente:

```text
¿Que lotes estan en la lista de corte?
¿Cuales tienen abono rezagado?
¿Cuales tienen otro precursor?
¿Cuales realmente deben salir de la lista?
¿Que falta resolver en cada lote?
```

## Modelo del problema

Cada fuente se trata como un conjunto de lotes identificados por `(MZ, LT)`:

```text
C = lotes de la lista de corte
A = lotes con abono rezagado
P = lotes con otros precursores
```

La interseccion prioritaria es:

```text
                  A
              .--------.
             /          \
            /     C ∩ A  \
           /       +      \
          '--------|------'
                   |
                   v
          revisar antes de publicar
```

Con tres fuentes:

```text
                    C: lista de corte
                 .-------------------.
                /                     \
               /    C ∩ A ∩ P         \
              /       !!              \
             /   .-------------.       \
            /   /               \      \
           '---/-----------------\-----'
              A                   P
       abonos rezagados      otros precursores
```

## Clasificacion visual

Cada lote del mapa usa una forma y un estado:

```text
(( RESUELTO ))   el saldo proyectado ya no permite corte
[ EN PRUEBA ]    esta dentro del mini-pipeline
{ FALTA DATO }   falta un abono, precursor o evidencia
!! REVISION !!   requiere una decision humana
```

El mapa no es decorativo. Cada figura representa una decision pendiente o
verificada sobre un lote.

## Flujo de trabajo

```text
lista de corte C
       +
abonos A
       +
precursores P
       |
       v
calcular C ∩ (A ∪ P)
       |
       v
filtrar esos lotes en el mini-pipeline
       |
       v
comparar saldo original vs saldo proyectado
       |
       +--> saldo <= 0  → (( sale de lista ))
       |
       +--> saldo > 0   → [ permanece explicablemente ]
       |
       +--> falta dato   → { no publicar como definitivo }
```

## Interseccion real actual

La interseccion se calculo con:

```text
C = lista de corte actual: 45 lotes
A = abonos rezagados:     40 lotes unicos
C ∩ A =                   7 lotes
```

```text
                 C: LISTA DE CORTE (45)
              .---------------------------.
             /                             \
            /        C ∩ A (7 lotes)        \
           /    I-9  L-5  P-12  P-3         \
          /     Q-5  S-2  W-5                \
         '-------------------+---------------'
                    /
                   /
      A: ABONOS REZAGADOS (40 lotes unicos)
```

Los 7 lotes de `C ∩ A` son los unicos que hoy requieren revision conjunta antes
de publicar la lista de corte.

| Lote | Abono total | Filas de abono | Situacion | Saldo en lista | Ejecutar corte |
|---|---:|---:|---|---:|---|
| I-9 | S/136 | 2 | CONFIRMADO + REVISAR | S/66 | NO |
| L-5 | S/126 | 1 | CONFIRMADO | S/50 | SI |
| P-12 | S/30 | 1 | CONFIRMADO | S/23 | SI |
| P-3 | S/33 | 1 | CONFIRMADO | S/25 | SI |
| Q-5 | S/114 | 2 | REVISAR | S/36 | SI |
| S-2 | S/60 | 1 | CONFIRMADO · Yape de Wagner Trujillo; retenido por Wagner | S/47 | SI |
| W-5 | S/15 | 1 | CONFIRMADO | S/50 | NO |

### Trazabilidad por lote

La situacion del abono se debe leer junto con la nota disponible y el cobrador que
figura en la fuente. `SIN_NOTA` significa que el lote no aparece en
`4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx`; no significa que el abono
sea falso o que el lote no haya sido revisado.

| Lote | Nota de secretaria | Cobrador del registro | Evidencia del abono |
|---|---|---|---|
| I-9 | "Al dia."; "Verificar todos sus pagos... ella esta al dia." | Wagner Trujillo (S/86, junio) | `verificacion_yape_todos.xlsx`: YA_REGULARIZADO; además S/50 confirmado por nota secretaria |
| L-5 | "Revisar su convenio, iba a pagar 20 cada mes desde enero, tambien su multa 34? y techado 100?" | No identificado en la fuente del abono | Decisión confirmada; sin comprobante Yape independiente |
| P-12 | SIN_NOTA | Wagner Trujillo (S/30, 05/07/2026) | `verificacion_yape_todos.xlsx`: NO_EXISTE; el Yape anotado no aparece en la cuenta de la JASS |
| P-3 | SIN_NOTA | Yerald Romero (S/33, 05/07/2026) | `verificacion_yape_todos.xlsx`: NO_EXISTE; el Yape anotado no aparece en la cuenta de la JASS |
| Q-5 | "Esta al dia en todo -- borrar solo, ponle su consumo de este mes." | No identificado en la fuente del abono | Nota de secretaria, sin comprobante Yape independiente |
| S-2 | "Revisar su campo." | Wagner Trujillo (S/60, 05/07/2026) | `verificacion_yape_todos.xlsx`: NO_EXISTE; el Yape anotado no aparece en la cuenta de la JASS |
| W-5 | "[CONVENIO] Cancelo reviza campo y convenio" | Wagner Trujillo (S/15, 04/07/2026) | `verificacion_yape_todos.xlsx`: NO_EXISTE; el Yape anotado no aparece en la cuenta de la JASS |

En particular, P-12 no tiene mensaje de secretaria ni reclamo del vecino. Su
S/30 proviene exclusivamente del registro del cobrador Wagner Trujillo y de la
verificacion bancaria negativa; por eso la ficha debe conservar ambas cosas
separadas: `SIN_NOTA` y `NO_EXISTE`.

```text
C ∩ A
   |
   +--> 5 lotes con EJECUTAR_CORTE=SI
   |       L-5, P-12, P-3, Q-5, S-2
   |
   +--> 2 lotes ya marcados NO
           I-9, W-5
```

Los otros 33 lotes de `A` no aparecen en la lista de corte actual. No son parte
de la interseccion prioritaria de esta publicacion, aunque siguen siendo parte
del trabajo de abonos.

### Inventario comprobado de precursores

Se revisaron las fuentes de precursores de `shared/` y se cruzaron sus lotes
contra los 45 lotes de `C`. El resultado fue:

| Fuente | Naturaleza | Es dinero | Interseccion con C | Tratamiento |
|---|---|---|---:|---|
| `abonos_rezagados.xlsx` | plata pagada antes y regularizada despues | SI | 7 | entra al mini-pipeline |
| `aportes_tanque_manuales.xlsx` | aporte voluntario al tanque | SI | 0 | no es pago de deuda |
| `blancos_efectivo.xlsx` | efectivo de caja sin atribucion inicial | SI | 0 | solo aplicar si tiene lote confirmado |
| `blancos_acumulados.xlsx` | Yape sin atribucion inicial | SI | 0 | los ya aplicados no se reaplican |
| `reidentificacion.xlsx` | pago real movido de un lote a otro | SI | 0 | ya tiene destino definido |
| `reidentificacion_cargo.xlsx` | mueve un cargo al lote correcto | NO | 0 | corrige deuda, no paga |
| `devoluciones_aplicadas.xlsx` | exceso ya aplicado a otro concepto | SI | 0 | no crear un pago nuevo |
| `reasignaciones_aplicacion.xlsx` | cambio de concepto de un pago | SI | 0 | conserva el pago original |
| `genesis_tardia.xlsx` | cargo de deuda descubierto tarde | NO | 0 | aumenta deuda, no paga |
| `ajustes_cargo.xlsx` | anulacion o correccion de cargo | NO | 0 | corrige deuda, no paga |
| `deuda_correcciones.xlsx` | correccion de deuda | NO | 0 | corrige deuda, no paga |
| `exoneraciones_multa.xlsx` | decision de exoneracion | NO | 0 | no es dinero |

`deuda_directiva.xlsx` no tiene `(MZ, LT)`, solo `USER_ID`, por lo que no puede
generar una interseccion por lote. `parches_manuales_pendientes_julio.xlsx` es
un pendiente historico de W-4, pero W-4 no esta en la lista actual; queda fuera
de esta publicacion y no se cuenta como pago confirmado.

La comprobacion incluyo tambien origen y destino en `reidentificacion.xlsx` y
las columnas `MZ`/`LOTE` de los blancos. No se encontro ningun lote de la lista
actual en esas fuentes.

### Decision sobre los precursores

Los precursores se originan en movimientos de caja, pero un movimiento de caja no
se convierte automaticamente en pago:

```text
movimiento de caja
        |
        +--> TANQUE   → no entra como pago
        |
        +--> BLANCO   → si entra como pago, pero esos lotes fueron excluidos
        |                de la lista de corte por estar en reclamo
        |
        +--> DUDOSO   → revision manual; no automatizar
```

La comprobacion real confirma que los blancos identificados no generan una
interseccion con la lista actual y que el tanque no debe entrar al calculo de pago:

```text
C ∩ P_dinero_sin_abono = vacío
```

Esto no significa que todos los movimientos de caja sean irrelevantes. Significa
que solo se incorporan como pago despues de clasificar su naturaleza. Los casos
dudosos quedan fuera de la corrida hasta tener una decision.

Por tanto, la interseccion que gobierna esta publicacion queda:

```text
C ∩ (A ∪ P_pago)
       = C ∩ A
       = 7 lotes
```

### Mapa de trabajo actual

```text
C ∩ A = I-9, L-5, P-12, P-3, Q-5, S-2, W-5

I-9  [ EN PRUEBA ]   abonos S/86 + S/50
Q-5  { FALTA DATO }  mini tiene S/69; falta S/45
P-3  [ EN PRUEBA ]   abono S/33
W-5  [ EN PRUEBA ]   abono S/15
S-2  [ EN PRUEBA ]   abono S/60 · Yape de Wagner Trujillo · retenido por Wagner
P-12 [ EN PRUEBA ]   abono S/30
L-5  [ EN PRUEBA ]   abono S/126
```

## Contexto debajo del dibujo

Cada lote debe tener debajo del mapa una ficha corta:

```text
MZ-LT: I-9
lista de corte: SI
abono: S/136
otros precursores: pendiente de revisar
saldo original: S/66
saldo proyectado: pendiente de corregir
accion: lograr que ambos abonos entren al calculo
estado: EN PRUEBA
```

La ficha permite ver el motivo sin llenar el dibujo de texto. El mapa muestra la
relacion; la ficha explica la decision.

## Criterio para publicar

```text
mapa revisado
     ↓
cada C ∩ (A ∪ P) tiene estado y evidencia
     ↓
saldo proyectado calculado
     ↓
separar los que salen y los que permanecen
     ↓
publicar lista de corte revisada
```

No se publica como definitiva una fila que este en `{ FALTA DATO }` o `!! REVISION !!`
sin una decision explicita.

## Estado

**EN IMPLEMENTACION.**

La idea adoptada es usar conjuntos para encontrar intersecciones y un mapa visual
por `(MZ, LT)` para seguir el avance. El siguiente paso es convertir este mapa en
una imagen actualizable durante la revision de los lotes.

<!-- MINI_PIPELINE_GENERATED_START -->
## Actualización del aislamiento · I-9 y cascada de fuentes

La corrida aislada de los siete lotes reconstruye primero el arrastre de julio con
el mismo pipeline y no usa el `arrastre_consolidado_2026-07.xlsx` congelado que
había quedado desactualizado.

```text
pago normal del ciclo actual
    -> consumo y mantenimiento actuales

abono de ciclo cerrado
    -> deuda arrastrada del ciclo anterior
```

Para I-9:

```text
deuda agosto:       S/152
efectivo agosto:      S/8
abono Wagner:        S/86
abono Secretaria:    S/58
abono total:        S/144
pago total:         S/152
saldo final:          S/0
```

I-9 ya no es un lote bloqueante del aislamiento. La misma separación de fuentes
se verifica contra la reconciliación real: efectivo del ciclo actual contra deuda
actual y abono cerrado contra deuda arrastrada.

## Resultado mini-pipeline y cambios previstos

Fuente: `mini_resultado_cascada.xlsx`; corrida aislada sobre los 7 lotes.

| Lote | Abono | Total pagado | Saldo | Actual | Mant. | Anterior | Corte | Convenio | Acuerdos | Multa | Estado |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|
| I-9 | S/144 | S/152 | S/0 | S/0 | S/0 | S/0 | S/0 | S/0 | S/75 | S/50 | PENDIENTE_APLICAR |
| L-5 | S/126 | S/126 | S/50 | S/0 | S/0 | S/0 | S/0 | S/0 | S/0 | S/50 | PENDIENTE_APLICAR |
| P-12 | S/30 | S/63 | S/23 | S/0 | S/3 | S/0 | S/0 | S/20 | S/0 | S/0 | PENDIENTE_APLICAR |
| P-3 | S/33 | S/33 | S/25 | S/22 | S/3 | S/0 | S/0 | S/0 | S/0 | S/0 | PENDIENTE_APLICAR |
| Q-5 | S/114 | S/114 | S/16 | S/13 | S/3 | S/0 | S/0 | S/0 | S/0 | S/0 | PENDIENTE_APLICAR |
| S-2 | S/60 | S/60 | S/47 | S/24 | S/3 | S/0 | S/0 | S/0 | S/20 | S/0 | PENDIENTE_APLICAR |
| W-5 | S/15 | S/31 | S/40 | S/12 | S/3 | S/0 | S/0 | S/0 | S/37 | S/0 | PENDIENTE_APLICAR |

### Cambios previstos en el ledger real

`PENDIENTE_APLICAR`: estas filas son la proyeccion del mini-pipeline; el script no escribe el ledger real.

| Lote | Mes | Concepto | Pago previsto | Source | Estado |
|---|---|---|---:|---|---|
| L-5 | 2026-08 | CONVENIO | S/60 | `abonos_rezagados` | PENDIENTE_APLICAR |
| L-5 | 2026-08 | ACUERDOS | S/50 | `abonos_rezagados` | PENDIENTE_APLICAR |
| I-9 | 2026-08 | ACUERDOS | S/75 | `abonos_rezagados` | PENDIENTE_APLICAR |
| I-9 | 2026-08 | MULTA | S/50 | `abonos_rezagados` | PENDIENTE_APLICAR |
| P-12 | 2026-08 | CONVENIO | S/27 | `abonos_rezagados` | PENDIENTE_APLICAR |
| Q-5 | 2026-08 | CONVENIO | S/25 | `abonos_rezagados` | PENDIENTE_APLICAR |
| Q-5 | 2026-08 | ACUERDOS | S/50 | `abonos_rezagados` | PENDIENTE_APLICAR |
| Q-5 | 2026-08 | MULTA | S/3 | `abonos_rezagados` | PENDIENTE_APLICAR |
| S-2 | 2026-08 | ACUERDOS | S/3 | `abonos_rezagados` | PENDIENTE_APLICAR |

No ejecutar el ledger real ni `5_cobranza --force` desde este script.

<!-- MINI_PIPELINE_GENERATED_END -->
