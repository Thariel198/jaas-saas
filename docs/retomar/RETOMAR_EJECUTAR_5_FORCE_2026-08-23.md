# RETOMAR — Ejecutar 5_cobranza --force y cerrar agosto

Fecha de pausa: 2026-08-22

```text
ledger sin abonos viejos
        |
        v
5_cobranza --force
        |
        v
snapshot f29d58dfcefa...
        |
        v
5b: +S/700 reproducible y aceptado
        |
        v
sellar hash exacto -> 7_cierre -> verificar ledger
```

## Estado al pausar

- Agosto sigue `ABIERTO` y `arrastre.validado=false`.
- No se comprometió el snapshot al ledger y no se ejecutó la limpieza de cierre.
- El ledger real no contiene eventos de agosto con `SOURCE=abonos_rezagados`.
- `5_cobranza --force` se ejecutó por última vez el 22/08 a las 21:03.
- Snapshot regenerado: `f29d58dfcefaa18ab4e59417dc227a1e9fd08fd5d04f7cca871837b238ffc705`.
- Manifest, planilla y snapshot coinciden en 18 predios y `S/1,074` de abonos.
- `5b_validacion` volvió a reproducir únicamente `Nivel 1a TE PAGO: +S/700`.
- El usuario autorizó sellar esta diferencia reproducible el 22/08/2026.

## Explicación del gap aceptado

```text
4 reasignaciones de tanque contadas también como agua   +S/750
Nelson Mon* sin mensaje ni maestro                      -S/50
                                                        -------
Nivel 1a reproducido                                    +S/700
```

La evidencia previa está en `RETOMAR_CIERRE_AGOSTO_LISTA_CORTE_2026-08-17.md` y el detalle permanece en `estado_ciclo.json`. No corregir montos para hacer cuadrar el reporte.

## Mañana

Desde la raíz:

```powershell
py -u -X utf8 5_cobranza/main.py --force
py -u -X utf8 5b_validacion/main.py
```

Verificar antes de sellar:

- el snapshot conserva exactamente el hash `f29d58dfcefa...`;
- `5b_validacion` conserva únicamente el gap aceptado `+S/700`;
- `shared/reporte_acumulado_procesado/estado_ciclo.json` sigue ligado a ese hash.

Después, aplicar el sello explícitamente aceptado al hash exacto y continuar con los pasos existentes de `7_cierre`: gate, cosecha, commit, freeze y limpieza. El flujo normal de `7_cierre` vuelve a ejecutar `5_cobranza --force` y deja `validado=false`; por eso el sello debe aplicarse después de esa última preparación y antes del gate, sin regenerar nuevamente.

## Verificación final requerida

- agosto queda `CERRADO` con el mismo `snapshot_hash`;
- ledger: 40 aplicaciones `ABONO_REZAGADO`, total `PAGO S/1,074`, `AJUSTE S/0`;
- F1-4: `AGUA S/51 + MULTA S/50 = S/101`;
- manifest = planilla_cobrado = snapshot = ledger por predio y concepto;
- vistas del ledger regeneradas;
- repetir `5_cobranza/validacion_planilla_cobrado.py` con los archivos actuales.

Backup principal: `shared/backups_ledger/reinicio_abonos_20260822_20260822_204442/`.
