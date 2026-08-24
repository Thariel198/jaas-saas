# RETOMAR — Vista provisional y migración a libro mayor (2026-08-21)

## Resumen operativo

```text
5_cobranza --force
        |
        v
snapshot_ledger_2026-08.json  fbca6b7cd300...
        |
        +--> 5b_validacion: NO SELLADO, diferencia Nivel 1a = +S/700
        |
        +--> copia temporal del ledger
                  |
                  v
shared/vista_seguimiento_provicional.xlsx
```

- Agosto sigue `ABIERTO` y no se ejecutó `7_cierre --confirmar`.
- `shared/seguimiento_pueblo.xlsx` sigue siendo el ledger oficial transitorio.
- La vista provisional no escribió eventos en el ledger oficial: aplicó el snapshot sobre
  una copia temporal y luego eliminó esa copia.
- El snapshot actual no está validado. La vista lo dice explícitamente en la hoja
  `PROVISIONAL`: `ESTADO=NO VALIDADO` y `diferencia 1a: diferencia +700.00`.
- El usuario autorizó expresamente generar esa vista marcada aunque el snapshot no estuviera
  sellado.

## Corrida real de agosto

Se ejecutó:

```powershell
py -u -X utf8 5_cobranza/main.py --force
```

Resultado principal:

```text
ciclo cobranza     31
pagos nuevos       1
snapshot hash      fbca6b7cd3008c856cbed251ae1abdd8b7700fd21ed1b6863cfaaefb70e67840
objetivos ledger   3462
candidatos corte   81
```

La corrida retroescribió `CICLO_COBRANZA=31` en una fila nueva de
`pagos_efectivo_2026-08.xlsx`. La lista de corte previa quedó desactualizada y deberá
regenerarse/revisarse antes del cierre.

Después se ejecutó:

```powershell
py -u -X utf8 5b_validacion/main.py
```

Resultado:

```text
Nivel 1a  agua 2530 + tanque 1970 = 4500
banco TE PAGÓ                         3800
diferencia                            +700  ALERTA
Nivel 1b PAGASTE                         0  OK
Nivel 2 agua vs planilla                 0  OK
efectivo                                 0  OK
```

`shared/reporte_acumulado_procesado/estado_ciclo.json` quedó con:

```text
2026-08.estado             ABIERTO
arrastre.generado          true
arrastre.validado          false
arrastre.snapshot_hash     fbca6b7cd300...
```

No forzar `validado=true` ni maquillar S/700. La causa conocida registrada habla de
reasignaciones duplicadas de tanque, pero el importador de caja nuevo deberá demostrar
por identidad qué depósitos están duplicados; un depósito solo puede existir una vez.

## Q-5 verificado en el snapshot

```text
CARGOS
  AGUA apertura julio        20
  AGUA agosto                13
  MANTENIMIENTO agosto        3

ABONO REZAGADO TOTAL        114
  AGUA                       20
  CONVENIO                   25
  ACUERDOS                   50
  MULTA                      19

SALDO PROYECTADO
  AGUA                       13
  MANTENIMIENTO               3
  TOTAL                      16
```

En `shared/vista_seguimiento_provicional.xlsx`, Q-5 quedó:

```text
AGUA             DEUDA 33  PAGO 20  SALDO 13
MANTENIMIENTO    DEUDA  3  PAGO  0  SALDO  3
CONVENIO                   PAGO 25  SALDO  0
ACUERDOS                   PAGO 50  SALDO  0
MULTA                      PAGO 19  SALDO  0
CORTE_RECONEXION           sin evento para Q-5
```

## Archivos creados o modificados en esta sesión

### Vista provisional

- `shared/generar_vista_seguimiento_provicional.py`
  - exige que el hash del snapshot coincida con su contenido;
  - exige que el hash coincida con `estado_ciclo.json`;
  - normalmente bloquea snapshots no validados;
  - acepta `--permitir-no-validado` solo para producir una vista claramente marcada;
  - toma la alerta vigente desde `5b_validacion/outputs/validacion_diferencias.xlsx`.
- `shared/seguimiento_repo.py`
  - nueva constante `VISTA_PROVISIONAL_PATH`;
  - nueva función `generar_vista_provisional(...)`;
  - copia el ledger a un temporal, aplica el mismo batch de `7_cierre`, genera la vista y
    restaura siempre la ruta del ledger real;
  - agrega la hoja inicial `PROVISIONAL` con mes, hash, validación, alerta y alcance.
- `shared/tests/test_vista_provisional.py`
  - verifica que el ledger real quede idéntico byte por byte;
  - verifica Q-5 sintético: AGUA `33-20=13` y MANTENIMIENTO `3-0=3`;
  - verifica la marca `NO VALIDADO` y su alerta.
- `shared/README.md`
  - documenta la vista como simulación regenerable, nunca fuente de verdad.
- `shared/vista_seguimiento_provicional.xlsx`
  - salida real generada desde el snapshot `fbca6b7cd300...`.

Verificaciones ejecutadas:

```text
pytest shared/tests/test_vista_provisional.py                    1 passed
py_compile seguimiento_repo.py + generador provisional          OK
diff --check                                                    OK
```

### Plan de libro mayor

- `libro_mayor/plan_migracion_libro_mayor.html`
  - inventario de lo existente y lo faltante;
  - arquitectura caja + estado_cuenta + motor;
  - checklist de ocho fases;
  - gates para no romper septiembre;
  - caso obligatorio Q-5;
  - convivencia y rollback de `seguimiento_pueblo.xlsx`;
  - enlaces a los README, cuaderno, diagrama y decisiones.

Validación del HTML:

```text
506 líneas
16 enlaces
0 enlaces/anclas rotos
estructura section/article/div balanceada
```

## Qué significa el diseño actual

```text
CAJA: HECHO DE DINERO
  MovimientoCaja · ABONO_ID · canal · referencia · dirección · balde

ESTADO_CUENTA: HECHO DE DEUDA + INTERPRETACIÓN
  Cargo · CARGO_ID
  Aplicación · ABONO_ID -> CARGO_ID
  Ajuste

MOTOR
  única pieza que ve caja + cargos
  prioridad P1-P6 + FIFO por MES_CARGO
```

El trabajo actual de `shared/seguimiento_repo.py` es una transición, no la implementación
completa del diseño del cuaderno:

```text
YA TIENE                         TODAVÍA NO TIENE
CARGO/PAGO/AJUSTE                JASS_ID
7 conceptos                     ABONO_ID
append-only                     CARGO_ID
writer único                    SUB_CONCEPTO
vista ancha                     APLICACIONES enlazadas
commit por snapshot             céntimos int
```

`seguimiento_pueblo.xlsx` no está libre para mover o reemplazar: es el store operativo
que debe conservarse mientras septiembre todavía dependa de él.

## Plan acordado para avanzar

La guía visual y checklist viven en:

```text
libro_mayor/plan_migracion_libro_mayor.html
```

Orden técnico:

```text
1 dominio puro
  taxonomia -> entidades -> identidad -> cascada -> saldo -> politica_corte

2 puertos/repos
  caja_repo.py + cuenta_repo.py

3 adapter XLSX
  solo stores/ conoce openpyxl

4 caja en sombra
  importadores -> arqueo -> vista_seguimiento_caja.xlsx

5 estado_cuenta en sombra
  cargos -> motor -> aplicaciones -> vista nueva

6 comparación completa
  predio/concepto + conservación por abono + caja global

7 cutover
  cambiar lectores solo después de comparación exacta
```

El diseño objetivo sí se construye prácticamente desde cero dentro de `libro_mayor/`,
porque hoy esa carpeta contiene especificaciones y `.gitkeep`, no código. No se rediseña
desde cero: el contrato está cerrado. Se implementa en paralelo y se cablea al final.

## Primer paso mañana

```text
ANTES DE PROGRAMAR
  AGENTS.md
    -> docs/decisiones/estado_decisiones.md
    -> LEER_ANTES.md
    -> este RETOMAR
    -> libro_mayor/README.md
    -> libro_mayor/dominio/README.md
    -> libro_mayor/caja/README.md
    -> libro_mayor/estado_cuenta/README.md
    -> libro_mayor/plan_migracion_libro_mayor.html
```

Después:

1. No tocar aún lectores de `2_planilla` ni mover `seguimiento_pueblo.xlsx`.
2. Decidir si primero se resuelve el S/700 o se implementa `dominio/` en paralelo.
   - Recomendación: implementar dominio puro en paralelo porque no lee datos reales.
   - Caja oficial queda bloqueada hasta explicar S/700 por identidad de depósito.
3. Implementar primero los seis archivos puros de `libro_mayor/dominio/` con tests.
4. Usar Q-5 como prueba de integración posterior, no como excepción hardcodeada.

## Gates antes del cutover

```text
[ ] agosto cerrado por 7_cierre
[ ] lista de corte regenerada y comprometida
[ ] caja explica S/700 sin duplicar ni eliminar dinero
[ ] cada depósito aparece una sola vez
[ ] por ABONO_ID: monto = aplicado + saldo_a_favor
[ ] saldos viejo vs nuevo comparados por MZ/LT/concepto
[ ] Q-5 = AGUA 13 + MANTENIMIENTO 3
[ ] rollback al lector viejo probado
```

## Precauciones

- No ejecutar `7_cierre --confirmar` mientras siga pendiente la lista de corte.
- No marcar manualmente el snapshot como validado.
- No convertir los `PAGO` agregados de `seguimiento_pueblo.xlsx` directamente en
  aplicaciones: no tienen `ABONO_ID`; las aplicaciones deben re-derivarse desde caja.
- No importar `aportes_tanque.xlsx` como dinero adicional si referencia depósitos ya
  presentes en Yape/efectivo; debe clasificar el mismo movimiento, no duplicarlo.
- No cambiar el nombre canónico `AGUA` por `CONSUMO`: algunos HTML antiguos dicen
  `CONSUMO`, pero el contrato vigente usa `AGUA` y `MANTENIMIENTO` separados.
- La clave vigente de efectivo es procedencia `(JASS_ID, origen_archivo, fila)`, no la
  fórmula vieja con fecha/monto/MZ/LT que todavía aparece en `4_pagos/README.md`.
- El working tree contiene muchos cambios y artefactos operativos ajenos/acumulados. No
  limpiar, revertir ni incluirlos en un commit sin revisar uno por uno.

## Estado al pausar

```text
vista provisional        GENERADA
snapshot agosto          GENERADO · fbca6b7cd300...
validación snapshot      FALLÓ · Nivel 1a +S/700
ledger oficial agosto    SIN COMMIT NUEVO
ciclo agosto             ABIERTO
cierre real              NO EJECUTADO
plan HTML libro mayor    CREADO Y VALIDADO
implementación core      PENDIENTE
```
