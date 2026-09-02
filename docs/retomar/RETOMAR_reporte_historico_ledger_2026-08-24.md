# RETOMAR - Reporte historico desde el ledger real

Fecha: 2026-08-24

```text
ledger activo comprometido
        |
        v
DEUDA | PAGO | AJUSTE | SALDO
        |
        +-- PDF con el formato historico
        +-- Excel auditable
```

## Estado

- `4b_reclamos/herramienta/reporte_historico.py` fue rehecho para leer solo
  eventos activos hasta el ultimo ciclo comprometido.
- No proyecta ciclos abiertos ni usa abonos, precursores o planillas para
  calcular deuda, pagos, ajustes o saldos.
- Junio/julio se marcan como cobertura parcial; agosto es cuenta completa.
- Los ajustes muestran `CLASE`, `SOURCE`, `AUDIT_REF` y `MOTIVO`.
- Las referencias reales de Yape/efectivo aparecen en pagina y hoja separadas;
  son evidencia externa y no modifican el ledger.

## Salidas verificadas

- `4b_reclamos/outputs/reporte_historico_ledger_A-4_2026-08.pdf`.
- `4b_reclamos/outputs/reporte_historico_ledger_A-4_2026-08.xlsx`.
- A-4 termina con agua S/5, mantenimiento S/3, multa S/30, acuerdos S/45 y
  convenio S/0; total S/83.
- Sus cuatro eventos de reasignacion aparecen como `AJUSTE`, no como pago.

## Pruebas

- `test_reporte_provisional.py`: 5 passed.
- `test_buscar_pago.py`: 43 passed.

## Uso

```powershell
py -u -X utf8 4b_reclamos/herramienta/reporte_historico.py MZ LT
py -u -X utf8 4b_reclamos/herramienta/reporte_historico.py --con-deuda 2026-08
py -u -X utf8 4b_reclamos/herramienta/reporte_historico.py --todos 2026-08
```
