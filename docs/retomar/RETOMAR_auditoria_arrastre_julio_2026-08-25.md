# RETOMAR - Auditoria del arrastre de julio hacia agosto

Fecha de continuacion: 2026-08-25

```text
planilla julio corregida
          |
          v
ajustes de cargo + pagos realmente procesados
          |
          v
saldo correcto al cierre de julio
          |
          v
comparar con apertura del ledger de agosto
          |
          +-- auditoria Excel de 565 predios
          +-- dry-run append-only, sin escribir ledger
```

## Alcance aprobado

- Auditar julio completo, no corregir solo I-9.
- Crear un Excel trazable con los 565 predios.
- Preparar una propuesta de asientos append-only en modo dry-run.
- No escribir en `shared/seguimiento_pueblo.xlsx` sin una revision y aprobacion
  posterior del usuario.

## Fuentes verificadas

- Fuente correcta de julio:
  `C:/Users/wilde/PycharmProjects/Julio/jass_system - Julio/shared/planilla_mes/planilla_2026-07.xlsx`.
- Foto usada por cobranza:
  `C:/Users/wilde/PycharmProjects/Julio/jass_system - Julio/5_cobranza/outputs/planilla_cobrado_julio.xlsx`.
- Arrastres generados:
  `arrastre_deuda_2026-07.xlsx` y `arrastre_consolidado_2026-07.xlsx`, en el mismo
  directorio de outputs.
- Apertura y movimientos comprometidos de agosto:
  `shared/seguimiento_pueblo.xlsx`.
- Correcciones administrativas:
  `shared/ajustes_cargo.xlsx`.
- Referencias monetarias posteriores:
  `shared/abonos_rezagados.xlsx`.

## Hallazgos de hoy

- La planilla fuente y `planilla_cobrado` tienen 565 filas cada una.
- Hay 33 diferencias en campos de agua/corte, distribuidas en 29 predios.
- Las diferencias positivas brutas suman S/557.50, pero esa cifra NO es deuda
  perdida: incluye importes pagados, exonerados o corregidos por otra via.
- La primera reconciliacion dejo seis candidatos de apertura incompleta por
  S/139.00:

| Predio | Diferencia candidata |
|---|---:|
| F1-4 | S/61.00 |
| M-7 | S/24.00 |
| B-19 | S/20.00 |
| A-6 | S/17.00 |
| I-9 | S/11.00 |
| I-2B | S/6.00 |
| **Total** | **S/139.00** |

- F1-4 no tiene S/101 de deuda candidata: S/40 de corte fueron exonerados en
  `ajustes_cargo.xlsx`; quedan S/61 por reconciliar.
- Las diferencias negativas observadas en una comparacion preliminar no prueban
  deuda excedente. Mezclan pagos rezagados aplicados recien en agosto y saldos
  pueblo que ya vivian en el ledger. No proponer asientos a partir de ellas.

## Caso I-9 confirmado

```text
julio: consumo 5 + mantenimiento 3 + mes anterior 11 = agua S/19
                                   |
                                   v
arrastre_deuda julio = S/8
                                   |
                                   v
apertura agosto AGUA = S/8
                                   |
                                   v
S/11 no ingresaron al ledger
```

- En agosto hay S/141 de pagos asentados, no simulados:
  S/133 con `CLASE=ABONO_REZAGADO`, `SOURCE=abonos_rezagados`, mas S/8 con
  `CLASE=COBRANZA`, `SOURCE=5_cobranza`.
- Por eso la fila mensual de agosto debe seguir diciendo `PAGO`.
- El defecto visual pendiente es que las referencias de S/86 y S/47 aparecen
  como no asentadas aunque el ledger ya contiene su aplicacion agregada de
  S/133.

## Metodo para cerrar la auditoria

1. Partir de los cargos correctos de `planilla_2026-07.xlsx`.
2. Aplicar solo correcciones de `ajustes_cargo.xlsx` cuyo ciclo de aplicacion sea
   julio.
3. Usar los pagos realmente procesados por `5_cobranza`; no usar `deuda - saldo`
   del reporte ni overlays como si hubieran estado asentados en julio.
4. Reproducir la cascada vigente en ese cierre y calcular el saldo esperado por
   concepto.
5. Comparar ese saldo con el arrastre y con la apertura comprometida de agosto.
6. Explicar cada diferencia de los 29 predios; ningun asiento entra al dry-run si
   no tiene fuente, pago y ajuste reconciliados.
7. Generar un contrato dry-run con `MZ`, `LT`, concepto, monto, saldo antes,
   saldo despues, fuente, motivo y `AUDIT_REF` determinista.
8. Validar suma global, suma por predio, cero saldos negativos y cero escrituras
   sobre el ledger real.

## Archivos de trabajo

- `4b_reclamos/herramienta/reporte_historico.py`: reporte actual; tiene cambios
  sin confirmar en Git.
- `4b_reclamos/herramienta/auditar_pago_vs_ledger.py`: auditor existente; no
  asumir que cubre el arrastre stale sin revisar su contrato.
- `4b_reclamos/tests/test_reporte_provisional.py`: pruebas dirigidas del reporte.
- `docs/decisiones/estado_decisiones.md`: D-005 mantiene agosto como inicio de
  `estado_cuenta` sin backfill automatico.
- `LEER_ANTES.md`: prevalece para abonos rezagados, reimputacion CA1 y orden de
  operaciones.

## Estado Git al cerrar

- El worktree tiene muchos cambios operativos y de otros alcances.
- El commit de este cierre debe incluir solamente este RETOMAR.
- No ejecutar `git init`: el repositorio ya existe.
