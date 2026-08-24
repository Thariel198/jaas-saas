# RETOMAR - Auditoria de limpieza del ledger y proxima reimputacion

Fecha de cierre: 2026-08-24

## Resumen ejecutivo

```text
ledger fisico append-only: 3,511 eventos
                  |
                  +-- 181 anulados logicamente
                  |
                  +-- 3,330 activos
                           |
                           +-- ajustes julio: 0
                           +-- ajustes agosto: 0
                           +-- ajustes junio: 112 conservados

vista Excel/PDF regenerada
agosto: PAGO S/8,904 | DECLARADO S/0 | AJUSTE S/0
siguiente sesion: definir y auditar reimputacion de pagos antes de escribir
```

La limpieza no borro filas de `shared/seguimiento_pueblo.xlsx`. Todas las
exclusiones quedaron auditadas en `shared/anulaciones_ledger.json`, con respaldo
del manifiesto antes de cada bloque. El ledger fisico sigue siendo append-only y
las vistas consumen solamente eventos activos.

## Estado final verificado

- Ledger fisico: `3,511` eventos.
- Eventos activos: `3,330`.
- Eventos anulados logicamente: `181`.
- Ajustes activos de julio: `0`.
- Ajustes activos de agosto: `0`.
- Ajustes de junio conservados por decision del usuario: `112`.
- Detalle junio conservado: `110 CONVENIO`, neto `S/-3,191`, y dos eventos de
  `MULTA F-12`, `-50/+50`, neto cero.
- Agosto permanece intacto: `PAGO S/8,904`, `DECLARADO S/0`, `AJUSTE S/0`.
- `shared/vista_seguimiento_pueblo.xlsx` y `.pdf` fueron regenerados despues de
  cerrar el Excel que inicialmente bloqueo el reemplazo atomico.
- El ultimo chequeo del Excel mostro `112` filas en la hoja `Ajustes`, todas de
  `2026-06`; julio y agosto tienen cero.

## Barrera de seguridad implementada

Se agrego una barrera global para impedir que pruebas destructivas escriban en
datos operativos:

- `pytest.ini` y `conftest.py` activan la proteccion.
- `test_safety/` contiene el runner seguro.
- `AGENTS.md` documenta que pytest y scripts de prueba solo se ejecutan mediante
  `py -m test_safety.run ...`.
- La causa inmediata del incidente de agosto fue un pytest que sobrescribio las
  mesas de efectivo y produjo un snapshot provisional falso.

## Recuperacion del snapshot corrupto de agosto

```text
mesas reales archivadas
        |
        v
snapshot recuperado 64375c9dd5eb183c26e18510da3bbcf78dd381d848e5bb6ab8d50897e9e7eebe
        |
        v
44 PAGO provisionales + 44 AJUSTE compensatorios anulados
        |
        v
agosto vuelve a PAGO S/8,904 y AJUSTE S/0
```

Lote: `ANULACION-PARES-SNAPSHOT-CORRUPTO-2026-08`.

- Se anularon 88 eventos: 44 pagos y 44 contraajustes.
- Cada lado sumaba `S/1,612`; el impacto neto de los pares era cero, pero ambos
  lados eran ruido y no debian permanecer activos.
- El snapshot corrupto era
  `f29d58dfcefaa18ab4e59417dc227a1e9fd08fd5d04f7cca871837b238ffc705`.
- La recuperacion se hizo desde `7_cierre/archivo/2026-08/mesa_*.xlsx`, no
  inventando montos ni forzando el ledger.

## Declarados sin evidencia y exoneraciones

Lote: `ANULACION-DECLARADOS-SECRETARIA-Y-EXONERACIONES-2026-07`.

```text
24 PAGO declarados sin referencia verificable: S/1,063
47 AJUSTE causales asociados
 3 exoneraciones directas
                         |
                         v
74 eventos anulados | deuda restaurada S/1,055
```

La regla cerrada aplicada fue: la referencia de pago demuestra que el dinero
existe; un `SOURCE` tecnico no es evidencia. No se forzo ningun monto para hacer
coincidir reportes.

Precursores que quedaron pendientes en `shared/ajustes_cargo.xlsx` para una
decision posterior, sin convertirlos en pagos:

- `C-19`: `S/50`.
- `F1-10`: `S/30`.
- `R-5`: `S/12`.
- `O-2`: `S/30`.

## Limpieza de MULTA: D1-6, L-4 y O-2

Lote: `ANULACION-AJUSTES-D1-6-L-4-O-2-2026-07`.

- Se anularon cuatro ajustes y se restauraron `S/45` de deuda.
- `D1-6`: el pago real en efectivo fue `S/33`, distribuido como agua `S/16` y
  multa `S/17`; queda `MULTA SALDO S/13`.
- `L-4`: el pago real en efectivo fue `S/41`, distribuido como agua `S/24` y
  multa `S/17`; queda `MULTA SALDO S/3`.
- `O-2`: no se invento un pago; queda deuda `S/30` y un precursor unico pendiente
  en `shared/ajustes_cargo.xlsx`.

## Lote fantasma E-14A

Lote: `ANULACION-LOTE-FANTASMA-E-14A-2026-07`.

- Se anularon los dos eventos que componian el lote fantasma: `CARGO +75` y
  `AJUSTE -75`.
- `E-14A` deja de existir en la proyeccion activa.
- `E-14B` permanece como el lote real: deuda `S/75`, pago `S/9`, saldo `S/66`.

## Convenios A-4 y P-6

Lote: `ANULACION-RUIDO-CONVENIO-A-4-P-6-2026-07`.

```text
A-4: CARGO 75 + cinco eventos tecnicos anulados -> SALDO 75
P-6: CARGO 58 + cuatro eventos tecnicos anulados -> SALDO 58
                                              deuda restaurada: S/133
```

### A-4

Se conservaron unicamente los `S/75` de deuda real de convenio. Se anularon dos
pagos de `S/75`, dos ajustes de `S/-75` y un estabilizador de `S/+225`. La cadena
provenia de mezclar un yape normal de `S/136` con un aporte al tanque de `S/100`
y de correcciones superpuestas posteriores.

### P-6

La fuente contractual
`obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx`
confirma total `S/1,250`, pagos `600 + 300 + 292 = S/1,192` y saldo real `S/58`.
Se conservo el cargo neto `S/58` y se anularon el par tecnico `-300/+300`, una
aplicacion duplicada de `S/58` y su estabilizador `S/+58`.

## Pares de ajuste MULTA de julio

Lote: `ANULACION-PARES-AJUSTE-MULTA-D1-3-D-1-2026-07`.

- `D1-3`: `-18/+18`.
- `D-1`: `-30/+30`.
- Cuatro eventos anulados, impacto neto cero.
- Los dos ajustes `-50/+50` de `F-12` son de junio y se conservaron por alcance
  explicito del usuario.

## Auditoria de abonos rezagados

`shared/abonos_rezagados.xlsx` contiene 22 precursores de julio por `S/1,157`.
Al auditar el ledger no se encontro ningun evento de julio con
`CLASE/SOURCE=ABONO_REZAGADO` que demostrara su aplicacion. No se cargaron ni se
reimputaron en esta sesion; quedan para revision individual de evidencia.

Este conjunto no debe confundirse automaticamente con el plan CA1 historico de
`RETOMAR_reimputacion_cascada_ca1_2026-08-13.md`, que mide 88 predios y
`S/3,003.50`. La siguiente sesion debe fijar primero cual de los dos alcances se
va a ejecutar y cruzar cada pago con su referencia.

## Backups creados el 2026-08-24

- `shared/backups_ledger/anulaciones_ledger_pre_limpieza_declarados_20260824.json`.
- `shared/backups_ledger/ajustes_cargo_pre_limpieza_declarados_20260824.xlsx`.
- `shared/backups_ledger/anulaciones_ledger_pre_limpieza_D1_6_L4_O2_20260824.json`.
- `shared/backups_ledger/ajustes_cargo_pre_limpieza_D1_6_L4_O2_20260824.xlsx`.
- `shared/backups_ledger/anulaciones_ledger_pre_limpieza_E_14A_20260824.json`.
- `shared/backups_ledger/anulaciones_ledger_pre_limpieza_A4_P6_20260824.json`.
- `shared/backups_ledger/anulaciones_ledger_pre_ajustes_multa_D1-3_D-1_20260824.json`.

## Verificaciones ejecutadas

- `py -m test_safety.run pytest shared/tests/test_anulaciones_ledger.py -q`:
  `1 passed`.
- `py -m test_safety.run script shared/tests/test_seguimiento_repo.py`:
  todos los checks pasaron.
- Las nueve referencias de A-4/P-6 y las cuatro referencias D1-3/D-1 quedaron
  ausentes del conjunto activo.
- Los hashes almacenados en cada lote confirmaron que la limpieza logica no
  modifico el ledger fisico.
- La vista final contiene ajustes solo de junio.

## Siguiente sesion: reimputar pagos

```text
elegir alcance
   |
   +-- 22 precursores julio / S/1,157
   |
   +-- plan CA1: 88 predios / S/3,003.50
   v
referencia de pago -> origen -> aplicacion previa -> concepto correcto
   v
propuesta por predio y monto
   v
autorizacion humana
   v
eventos auditables + validacion de caso afectado y caso control
```

Orden obligatorio para retomar:

1. Leer `AGENTS.md`, `docs/decisiones/estado_decisiones.md`, `LEER_ANTES.md` y
   este documento.
2. Leer `RETOMAR_reimputacion_cascada_ca1_2026-08-13.md` si el alcance elegido
   es CA1, o `shared/abonos_rezagados.xlsx` mediante una lectura estructural
   acotada si el alcance son los 22 precursores.
3. No ejecutar `5_cobranza --force` ni cargar pagos como primer paso.
4. Confirmar evidencia y evitar doble aplicacion contra pagos ya incluidos en
   cargos netos, como ocurrio con P-6.
5. Presentar diferencias entre fuente y salida; nunca maquillarlas con ajustes.

## Archivos de verdad

- `shared/seguimiento_pueblo.xlsx`: ledger fisico append-only.
- `shared/anulaciones_ledger.json`: exclusiones logicas vigentes.
- `shared/vista_seguimiento_pueblo.xlsx` y `.pdf`: proyecciones regeneradas.
- `shared/ajustes_cargo.xlsx`: precursores de ajustes/exoneraciones pendientes.
- `shared/abonos_rezagados.xlsx`: declaraciones pendientes de evidencia y posible
  reimputacion.
- `docs/decisiones/estado_decisiones.md`: decisiones cerradas.
- `LEER_ANTES.md`: eventos operativos que prevalecen sobre el flujo mensual.
