# RETOMAR - P-12: mini pipeline y correccion de convenio · 2026-08-17

Este documento es el punto de entrada de la siguiente sesion. La sesion debe
empezar por **P-12**. No ejecutar el ledger real ni `5_cobranza --force` hasta
cerrar el mini resultado y revisar la aplicacion del convenio.

---

## Estado general de la correccion de los 7

```text
C ∩ A = I-9, L-5, P-12, P-3, Q-5, S-2, W-5
        |
        +--> I-9  mini aprobado: abonos S/133 -> saldo S/0
        +--> L-5  mini aprobado: abono S/126 -> queda MULTA S/50
        +--> S-2  mini aprobado: abono S/60 -> queda ACUERDOS S/47
        +--> P-12 siguiente caso
        +--> P-3, W-5 pendientes
        +--> Q-5 bloqueado por dos filas y decision pendiente
```

Fuente formal de las soluciones: `docs/aprendizaje/solucion de proble/03_reporte_correccion_pagos.md`.
El `RETOMAR_CORRECCION_LOTE_7_PREDIOS_2026-08-16.md` conserva el orden operativo
y las restricciones de seguridad.

### Cambios reales ya hechos

```text
L-5  se retiraron del ledger dos pagos prematuros:
      CONVENIO S/60 + ACUERDOS S/50
      backup: shared/backups_ledger/seguimiento_pueblo_pre_remove_L5_invalid_20260816.xlsx

S-2  se retiro del ledger el pago prematuro ACUERDOS S/3
      backup: shared/backups_ledger/seguimiento_pueblo_pre_remove_S2_invalid_20260816.xlsx

I-9  no se escribio correccion real
      la fuente Secretaria S/50 sigue sin cambiar a S/47
```

La fuente `shared/abonos_rezagados.xlsx` de L-5 y S-2 sigue presente. Por eso no
se debe correr `--force`: podria volver a aplicar filas cuya solucion aun no fue
aprobada para el ledger real.

---

## Primer objetivo: P-12

### Abono fuente

```text
Predio:              P-12
Monto:               S/30
MES_CICLO:           2026-07
MES_ANO_APLICA:      2026-08
Origen declarado:    Yape de Wagner Trujillo
Retenido por:        Wagner Trujillo
Evidencia:           verificacion_yape_todos.xlsx · NO_EXISTE
Motivo:              Yape de julio no rendido; recuperado/aplicado como abono
```

La fila vive en `shared/abonos_rezagados.xlsx`. No fusionarla con otros pagos.

### Ledger actual de P-12

```text
2026-06  ACUERDOS  CARGO   S/50
2026-06  CONVENIO  CARGO   S/100
2026-06  ACUERDOS  PAGO    S/50       SOURCE=5_cobranza
2026-06  CONVENIO  PAGO    S/25       SOURCE=5_cobranza
2026-06  CONVENIO  AJUSTE -S/25       SOURCE=correccion_genesis_formula
2026-07  CONVENIO  PAGO    S/50       SOURCE=5_cobranza
2026-07  CONVENIO  AJUSTE -S/50       SOURCE=5_cobranza
2026-08  CONVENIO  PAGO    S/27       SOURCE=abonos_rezagados
```

⚠ El `PAGO S/27` de agosto ya esta en el ledger activo y debe auditarse antes
de cualquier nueva accion. El resultado mini de P-12 no autoriza por si solo a
conservarlo: el mini se calcula desde planilla y fuente de abono, mientras el
ledger contiene el historial de correcciones de signo.

### Mini pipeline ejecutado hasta ahora

Salida temporal:
`C:\Users\wilde\AppData\Local\Temp\opencode\mini_corrida_lista_corte_7_20260816\outputs\mini_resultado_cascada.xlsx`

Resultado observado para P-12:

```text
PAGO_YAPE:       S/0
PAGO_EFECTIVO:   S/33
ABONO_REZAGADO:  S/30
TOTAL_DEUDA:     S/86
TOTAL_PAGADO:    S/63
SALDO:           S/23
ABONO CONVENIO:  S/27
ABONO ACUERDOS:  S/0
ABONO MULTA:     S/0
```

Esto significa que el abono S/30, despues de la cascada de agua/corte, deja
S/27 aplicados a convenio. **Todavia no es la decision final de P-12.** Hay que
compararlo contra la foto ledger-only y resolver si el resultado esperado es
convenio S/0, convenio S/23 u otra distribucion por el historial de signo.

---

## Bug conocido que hay que investigar primero

P-12 aparece en `3_boletas/inputs/reclamos_2026-08-01/README.md` como caso
`BUG_SIGNO`. La secuencia relevante es:

```text
PAGO convenio junio S/25
AJUSTE correccion_genesis_formula -S/25
PAGO convenio julio S/50
AJUSTE correccion de julio -S/50
PAGO abono rezagado agosto S/27
```

El riesgo es contar dos veces una reversa o aceptar un saldo negativo generado por
el signo viejo. No borrar eventos append-only sin una decision expresa; primero
reconstruir el saldo por `(MZ, LT, CONCEPTO, MES, TIMESTAMP)` y separar:

```text
plata real recibida
  ≠ pago normal ya registrado
  ≠ ajuste de correccion de sistema
  ≠ abono rezagado nuevo
```

---

## Orden exacto de trabajo

```text
1. Leer P-12 en la foto ledger-only y en Eventos.
2. Auditar el PAGO S/27 ya escrito: de donde salio y si corresponde a S/30.
3. Ejecutar el mini pipeline solo como copia temporal.
4. Comparar deuda ledger, planilla, pago normal y abono S/30.
5. Definir saldo final y concepto pendiente de P-12.
6. Registrar la decision en 03_reporte_correccion_pagos.md.
7. Solo despues revisar P-3 y W-5.
8. Q-5 no se procesa sin separar sus dos filas y obtener decision.
```

### Comandos de la siguiente sesion

```powershell
git status --short
py 5_cobranza/tests/generar_mini_corrida_abonos.py
py 5_cobranza/tests/test_abonos_rezagados_mini.py
py 5_cobranza/tests/test_abonos_manifest.py
py -m py_compile 5_cobranza/main.py
py -m py_compile 4b_reclamos/reporte_historico.py 4b_reclamos/reporte_deuda_ledger.py
```

No ejecutar todavía:

```powershell
py 5_cobranza/main.py --force
```

---

## Reglas de seguridad

```text
NO modificar abonos_rezagados.xlsx sin decisión documentada.
NO modificar el manifest para simular aprobación.
NO borrar cargos ni pagos históricos para “dejar bonito” el saldo.
NO escribir manualmente eventos nuevos en seguimiento_pueblo.xlsx.
NO aplicar P-12 al ledger real hasta cerrar el mini y validar el signo.
SIEMPRE crear backup antes de una futura reconciliacion real.
```

## Criterio de cierre de P-12

P-12 queda cerrado solo cuando exista una ficha con:

```text
abono fuente S/30
pago normal S/33 separado
saldo final por concepto
AUDIT_REF y SOURCE definidos
ningun saldo negativo nuevo
resultado mini reproducible
decision registrada en 03_reporte_correccion_pagos.md
```
