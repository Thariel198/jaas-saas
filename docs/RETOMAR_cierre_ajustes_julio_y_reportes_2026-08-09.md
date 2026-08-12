# RETOMAR — cierre de los 10 AJUSTE de julio, abono rezagado de Wagner, y 4 bugs del reporte (2026-08-09)

## ⚡ PRIMER PASO al retomar

**Auditar los 10 predios que quedaron con `AJUSTE sin MOTIVO` en el ledger y que
NO se tocaron hoy**, mismo método que se usó todo el día (ver abajo):

```
C-19 · C-29A · C1-17 · F-12 · F-3B · F1-10 · Q-16 · Q-5 · R-5 · S-14
```

Son 16 filas de `AJUSTE` (de un total de 29 sin motivo en todo el ledger — las
otras 13 son las de hoy, con el valor ya verificado, solo falta backfillear el
motivo en las filas viejas). Varios ya tienen contexto en otras secciones de
`LEER_ANTES.md` (hay títulos "F-12", "C-29A", "S-14/Q-13-Q-16" más abajo en el
archivo) — cruzar contra eso antes de investigar de cero.

Para ver el conteo actualizado en cualquier momento:
```python
import pandas as pd
df = pd.read_excel('shared/seguimiento_pueblo.xlsx', sheet_name='Eventos', header=1, dtype=str)
aj = df[df['TIPO_EVENTO']=='AJUSTE']
sin_motivo = aj[aj['MOTIVO'].isna() | (aj['MOTIVO'].str.strip()=='')]
```

---

## Qué se cerró hoy (verificado, no solo revisado)

### 1. El hallazgo central: el abono rezagado de Wagner nunca entró a julio

`shared/abonos_rezagados.xlsx` (8 filas, 7 de Wagner Trujillo + 1 de Yanet
Villanueva, `MES_ANO_APLICA=2026-07`) se creó el 03-06/08 — **después** de que
las 3 corridas de `5_cobranza` de julio (08/07, 20/07, 31/07) ya habían
pasado. Esa plata real nunca se aplicó al ledger.

```
D-16   abono 85 + efectivo 34 = 119  → ACUERDOS y CONVENIO quedan en 0
D1-6   abono 33 + efectivo 33 = 66   → MULTA corregida a 0, ACUERDOS ya estaba bien (30)
L-4    abono 58 + efectivo 41 = 99   → ACUERDOS corregido a 20 (antes 45)
T-12   abono 155                     → MULTA y ACUERDOS pagados completo (125 de 155)
I-9    abono 86                      → MULTA pagada, ACUERDOS de 75 a 58
F-9    abono 52 + efectivo 53 = 105  → CONVENIO pagado completo (antes 25)
F1-4   abono 101 (Yanet)             → SIN CAMBIOS (agua+corte ya consumen todo)
S-5    abono 71 + efectivo 74 = 145  → ver caso especial abajo
```

Método: CARGO real de `shared/planilla_mes/planilla_2026-07.xlsx` (verificado
exacto contra `DATA_boletas.xlsx`) + cascada P1→P5 (agua→corte→multa→acuerdos→
convenio) replicada a mano con el total real. Backups en
`shared/backups_ledger/seguimiento_pueblo_pre_*_20260809_*.xlsx`.

### 2. S-5 — dos historias distintas en el mismo predio, no confundirlas

- El "PAGO 30" de ACUERDOS (31/07, `CLASE=DECLARACION_SECRETARIA`) **no es el
  abono** — es un crédito histórico real (multa pagada en ene/feb 2026, antes
  del ledger, nunca reconocida) aplicado contra ACUERDOS. Confirmado en
  `notas_2026-07.xlsx`.
- El corte de S-5 (`registro_cortes.xlsx`, CORTADO desde 2026-06) **sí era el
  mismo abono** — Wagner tenía el pago retenido, por eso salió en corte sin
  deber. Se exoneró: `registro_cortes.xlsx` → `EXONERADO`,
  `ajustes_cargo.xlsx` nueva fila (aplica en agosto, julio ya cerrado),
  ACUERDOS AJUSTE −40 en el ledger. Queda "al día" en pueblo.

### 3. C-21 y J-6 — condonación institucional mal modelada

Estaban usando `CLASE=DECLARACION_SECRETARIA` (mecanismo de "el vecino dijo
que ya pagó", que arrastra un pendiente de investigar) para algo que en
realidad es una exención permanente (bienes del pueblo, `registro_cortes.xlsx`
`ESTADO=EXONERADO` desde 2026-02). Se reescribió como `AJUSTE` único con
`CLASE=EXONERACION`, y se agregó la fila que faltaba en `ajustes_cargo.xlsx`
(backfill para el libro mayor, mismo patrón que J-1/O-2).

### 4. 4 bugs de código en `4b_reclamos/reporte_historico.py`

Descubiertos al pedir que el reporte "reflejara fielmente el ledger":

1. Consumo/Mant/Mes_ant/Corte leían de `5_cobranza/outputs/planilla_cobrado.xlsx`,
   congelado de una corrida vieja de julio (de antes de una corrección de
   `2_planilla` del 18/07 que nunca se repropagó). Ahora lee de
   `shared/planilla_mes/planilla_<mes>.xlsx`.
2. MES_ACTUAL+MANTENIMIENTO se sumaban en un solo número, "Mant." quedaba
   siempre en 0 para junio/julio en adelante. Separados.
3. MULTA/ACUERDOS/CONVENIO del historial solo sumaban `PAGO`, ignorando lo
   saldado vía `AJUSTE` (justo las condonaciones de hoy) — L-4 mostraba 71 en
   vez de 99 real. Fórmula nueva: `PAGO + max(0, −DEBIA)`.
4. El corte exonerado se seguía contando como "pagado" en la cascada de agua,
   doble-contando la misma plata ya acreditada a ACUERDOS. `_datos_ciclo()`
   ahora consulta `ajustes_cargo.xlsx` y pone `corte_debido=0` si hay una fila
   `CORTE_RECONEXION` para ese predio.

Los 4 fixes verificados exacto contra la plata real en L-4/D-16/D1-6/S-5.
Reporte completo (208 predios) re-corrido: `4b_reclamos/outputs/reporte_reimputacion_cascada_2026-07.pdf`
+ `.xlsx`, 16:2x del 09/08 — 0 saldos negativos, 0 predios donde la deuda no
se conserva.

También se agregó la caja **"SALDO PENDIENTE"** a la página de cada predio
(usa `MULTA_DESPUES`/`ACUERDOS_DESPUES`/`CONVENIO_DESPUES` de
`calcular_tabla()`, no el `TOTAL PAGADO` del historial — esa caja es la que
hay que mirar para saber si alguien debe algo).

---

## Pendiente, sin tocar

```
① 10 predios con AJUSTE sin motivo, no auditados hoy (ver PRIMER PASO arriba)
② 22 declaraciones de la secretaria en abonos_rezagados.xlsx (BALDE=mixto/
   otros_conceptos) — mecanismo distinto al de Wagner, no se re-verificaron
③ D-16 MULTA (S/50, exonerada por asistencia a faena) + corte de F1-4/V-14/S-5
   → ya registrados en ajustes_cargo.xlsx, MES_ANO_APLICA=2026-08 — se
   aplican solos cuando corra 5_cobranza para agosto, no hace falta nada más
④ P-4/W-4 (Lourdes Melchor / Vicki Masias) — pago duplicado entre mesa_2 y
   mesa_4, pendiente de preguntarle a Wagner/Yerald directamente (arrastra
   de la sesión de ayer, nunca se retomó)
⑤ Nelson Mon* (S/50) · Lisbeth De* (S/15) · Jaime Hue* (S/39) en
   pendientes.xlsx Sin_identificar — sin candidato confiable por nombre ni
   monto (arrastra de ayer)
```

## Cómo se cierra este RETOMAR

Cuando los 10 predios del primer paso queden auditados (corregidos o
confirmados sin problema) y no quede ningún `AJUSTE sin MOTIVO` sospechoso,
borrar este archivo.
