# RETOMAR - Reimputacion CA1 aplicada

Fecha: 2026-08-24

```text
pagos COBRANZA junio/julio con cascada vieja
              |
              v
70 predios | 80 movimientos | S/2,266.50
              |
              +-- MULTA -> CONVENIO:       S/639.00
              +-- ACUERDOS -> CONVENIO:    S/510.00
              +-- MULTA -> ACUERDOS:     S/1,117.50
              v
80 precursores + 160 AJUSTE REASIGNACION
```

## Estado verificado

- La reimputacion CA1 fue aplicada a `shared/seguimiento_pueblo.xlsx`.
- Los 160 asientos tienen `CLASE=REASIGNACION`, `SOURCE=reimputacion_ca1`,
  `MES=2026-08`, motivo completo y 80 `AUDIT_REF` deterministas con dos lados.
- `shared/reasignaciones_aplicacion.xlsx` recibio 80 precursores con `MES_ANO`
  vacio; no pueden ser reaplicados por una corrida mensual.
- Se excluyeron 10 instalaciones y una reactivacion: B-20, B-29, C-35, C-43,
  C-45, F1-11, G-21, H1-2, P-6, W-2 y M-12.
- Abonos rezagados quedaron fuera del alcance.
- B-12A, I-5 y Q-14 no se movieron porque el pago precede al cargo destino.
- Convenios que conservan saldo tras la reimputacion: B-11 S/5, G-4 S/50,
  I-16 S/68, K-9 S/45 y T-7 S/45.

## Invariantes finales

- Deuda global: S/20,415.50 antes y despues.
- Pagos globales: S/28,566.50 antes y despues.
- Ajuste neto de la reimputacion: S/0 global y S/0 por predio.
- Saldos negativos nuevos: 0.
- Predios excluidos tocados: 0.
- Vista Excel: 160 filas `REASIGNACION` en la hoja `Ajustes`.
- `shared/vista_seguimiento_pueblo.xlsx` y `.pdf` regenerados.

## Evidencia

- Contrato: `4b_reclamos/outputs/contrato_reimputacion_ca1_20260824_141914.xlsx`.
- Backups: `shared/backups_ledger/seguimiento_pueblo_pre_reimputacion_ca1_*.xlsx`
  y `reasignaciones_aplicacion_pre_reimputacion_ca1_*.xlsx`.
- `py -m test_safety.run pytest shared/tests/test_anulaciones_ledger.py -q`: 1 passed.
- `py -m test_safety.run script shared/tests/test_seguimiento_repo.py`: todos los
  checks pasaron.

## Siguiente alcance separado

Auditar los abonos rezagados y sus referencias de pago. No mezclarlos con esta
reimputacion ya cerrada.
