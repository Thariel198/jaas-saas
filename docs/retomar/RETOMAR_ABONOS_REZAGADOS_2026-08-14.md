# Retomar — abonos rezagados

## Flujo

```text
checkpoint limpio
        |
        v
5_cobranza --force
        |
        v
eventos COBRANZA y ABONO_REZAGADO separados
        |
        v
incorporar 6 abonos seguros de julio + 12 abonos confirmados para la corrida de agosto
        |
        v
mantener 4 colisiones en Pendiente/
```

## Estado

- `shared/seguimiento_pueblo.xlsx` fue restaurado al checkpoint `seguimiento_pueblo_pre_bugsigno_20260803_164602.xlsx`.
- La corrida mezclada/interrumpida fue descartada.
- `5_cobranza/main.py` compila.
- El codigo ahora suma normal + abono solo para calcular la cascada.
- El ledger debe escribir el pago normal como `COBRANZA` y el abono como `ABONO_REZAGADO`, con `SOURCE` y `AUDIT_REF` separados.
- La fuente de abonos sigue siendo solo `shared/abonos_rezagados.xlsx`.
- No retirar pagos manuales, mixtos ni especiales.

## Primera hora

1. Validar sintaxis:

   ```powershell
   py -m py_compile "5_cobranza/main.py"
   ```

2. Regenerar agosto desde el checkpoint limpio:

   ```powershell
   $env:PYTHONIOENCODING='utf-8'; $env:PYTHONUTF8='1'; py "5_cobranza/main.py" --force
   ```

3. Reportar y aplicar solo los abonos de julio sin colision:

   ```powershell
   $env:PYTHONIOENCODING='utf-8'; $env:PYTHONUTF8='1'; py "5_cobranza/reportar_colisiones_abonos_julio.py"
   ```

4. Validar el ledger:

   - Los pagos normales deben tener `SOURCE=5_cobranza` y `CLASE=COBRANZA`.
   - Los abonos de agosto deben tener `SOURCE=abonos_rezagados` y `CLASE=ABONO_REZAGADO`.
    - Los seis abonos seguros de julio y los doce abonos confirmados para la corrida de agosto deben quedar separados de los pagos normales.
   - No debe existir `SOURCE=rebuild_abonos_normales` antes de aplicar la corrida nueva.
   - Los cuatro casos de `Pendiente/abonos_rezagados_pendientes_2026-07.md` no deben escribirse.

## Pendientes bloqueados

```text
S-5  MULTA     previsto S/30.00  saldo ledger S/0.00
D-16 ACUERDOS  previsto S/25.00  saldo ledger S/0.00
D1-6 MULTA     previsto S/5.00   saldo ledger S/1.00
Q-11 ACUERDOS  previsto S/17.00  saldo ledger S/0.00
```

Se recalculan al cierre del ciclo contra el saldo final. Si el saldo es cero, no se registra pago; si es menor que el abono, se registra solo el saldo disponible y el remanente queda pendiente.

## No hacer

- Los 12 abonos confirmados ayer entran en la corrida actual; no volver a tratarlos como casos históricos pendientes. La lista exacta está en `README_PLAN_RECLAMOS_2026-08.md`.
- No eliminar pagos manuales.
- No procesar `BALDE=mixto` ni destinos especiales en esta etapa.
- No registrar un abono como `COBRANZA`.
