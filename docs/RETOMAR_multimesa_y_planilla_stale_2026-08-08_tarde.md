# RETOMAR — pago_multi_mesa, motor de yape, y planilla de agosto con saldos viejos · 2026-08-08 tarde

Continúa el `RETOMAR_verificacion_lotes_efectivo_2026-08-08.md` de la mañana (esa parte quedó
prácticamente cerrada — ver sección 1).

---

## ⚡ PRIMER PASO al retomar

**Hablar con Wagner y/o Yerald sobre 2 pagos duplicados entre mesas** (mismo predio, mismo
monto, mismo día, registrados por los dos):

```
P-4  Lourdes Melchor Albornoz   S/42   01/08   mesa_2 (Yerald) Y mesa_4 (Wagner)
W-4  Vicki Masias Cusihuaman    S/24   01/08   mesa_2 (Yerald) Y mesa_4 (Wagner)
```

Evidencia encontrada (`4_pagos/recursos/Wagner_Trujillo/Fotos_1/01-08-2026-parte 1.jpeg`): los
2 montos de Wagner tienen desglose propio que cuadra EXACTO contra la boleta (P-4: consumo 39 +
mant 3 = 42, dejando fuera adrede el mes_anterior de 36 por "exceso a revisar"; W-4: consumo 21 +
mant 3 = 24). Yerald los tiene como número plano, sin desglose. No se puede decidir solo con las
fotos — hay que preguntar directo. Quedan en `4_pagos/efectivo/outputs/discrepancias.xlsx` hoja
`pago_multi_mesa`, columna OK para marcar cuál vale.

---

## 1. Verificación de lotes en efectivo (mañana) — CERRADA salvo 1 punto

Los 4 candidatos de `verificar_lotes.py` quedaron resueltos:

```
G-14 → E-14   CONFIRMADO y aplicado en mesa_3.xlsx (boleta S/12 exacto)
O-17 → O-27   CONFIRMADO y aplicado en mesa_3.xlsx (boleta S/17 exacto)
A-1  → A-2    DESCARTADO — es abono parcial real de A-1 (debe 41, pagó 21 "a cuenta",
              nota en la hoja de papel de Wagner). No se tocó mesa_4.xlsx.
H1-15 → H1-16 DESCARTADO — la hoja confirma H1-15 sin confusión de lote. Pero el monto
              (S/58) no cuadra con la boleta de H1-15 (Patricia Tarazona Carbajal, debe S/94)
              y queda SIN EXPLICAR. No es error de lote — puede ser abono parcial sin nota,
              o algo distinto. Pendiente de revisar con Yerald si aparece la oportunidad.
```

`entregas_hoja.xlsx` completado para los 4 cobradores (Wilder, Yerald, Maximo, Wagner — 2 fechas
c/u). Maximo y Yerald quedaron como ASUMIDO (sin declaración "Recibí X" legible); Wilder y Wagner
tienen declaración real que cuadra exacto.

`4_pagos` corrido completo — ancla correcta (`2026-07_procesado.xlsx` → 20/07/2026 22:48:47),
0 discrepancias en mesa_1..4 salvo el `pago_multi_mesa` de arriba.

---

## 2. Motor de matching de Yape — 2 fixes de código, sin commitear

Archivos tocados: `4_pagos/yape/motor_matching/main.py`,
`4_pagos/yape/motor_matching/exportar_motor.py`.

### 2a. CONCEPTO=tanque por mensaje (arreglo de fondo ① de `LEER_ANTES.md`)

Si el mensaje del yape dice "tanque"/"tanke", ahora se marca `CONCEPTO=tanque` automáticamente
— antes quedaba vacío y 5_cobranza/5b_validacion lo contaban como agua (el mismo bug de
C-15/A-4/P-17 de julio). Función nueva `detectar_concepto_tanque()`. Con esto,
`aportes_tanque.xlsx` pasó de 0 a 9 filas · S/1,220 este ciclo.

### 2b. 3 patrones de regex nuevos para mensajes que no se leían

```
"Mz Z-7."            → antes fallaba (sin palabra "lt", punto final) → ahora Z-7
"A-6. Julia Juarez…"  → antes fallaba (punto después del código) → ahora A-6
"Mz V..Lt 14"         → antes fallaba (dos puntos seguidos) → ahora V-14
```

Importante: estos 2 primeros casos el sistema los tenía **adivinados mal** (candidato "E-7" por
nombre del remitente) — el regex nuevo los corrigió contra el predio real, no solo los desambiguó.

Quedó **sin resolver por regex, a propósito** (demasiado ambiguo para un patrón seguro):
`"Antonio Espinoza Sifuentes MK.LT.2"` — se resolvió a mano (K-2, es Antonio mismo). Y
`"No A lt 5 tanque"` — el "No" antes de la letra rompe todos los patrones; se resolvió a mano
(A-5, tanque, Julia Agama Cuzco S/200) porque el usuario lo tenía claro por el monto (S/200)
y la palabra "tanque".

**Pendiente para otra sesión, mismo patrón de bug:** `K-17` — mensaje `"mzklt17Marcial Sánchez
Araoz"` (sin espacio entre "17" y el nombre) se leyó como lote `"17M"` (agarra la M de
"Marcial") y salió "no existe en planilla". K-17 sí existe. No se agregó regex para este caso
todavía — decidir si vale la pena (mensajes "número pegado al nombre, sin espacio" pueden ser
frecuentes o un caso aislado).

Tests: 13 fallos preexistentes (verificado con `git stash`, ya rotos antes de esta sesión,
no relacionados). 0 fallos nuevos. Pipeline real corrido 3 veces con estos cambios, sin errores.
**Nada de esto está commiteado.**

---

## 3. Planilla de agosto con 11 predios en negativo — causa raíz encontrada, sin decidir el arreglo

Se investigó un hallazgo del usuario (K-17 con "deuda negativa") y escaló a un problema real de
**11 predios** con `CONVENIO`, `MULTA` o `ACUERDOS_ASAMBLEA` negativos en
`shared/planilla_mes/planilla_2026-08.xlsx`:

```
CONVENIO(-)   A-8 · B-5 · C-1 · C-7 · E-12 · I-11 · J-3 · K-2 · K-17 · P-12
MULTA(-)      I-16
ACUERDOS(-)   B-5 · C-1   (ya contados arriba)
```

**Causa raíz confirmada** (rastreado contra el ledger y sus backups, 5 casos verificados uno por
uno, mismo patrón exacto en todos):

```
06/07 14:08-14:14   PAGO fantasma (bug ya documentado: 5_cobranza leyó yape de junio
                     contra deuda de julio)
31/07 18:08-18:20   se revierte con el signo invertido → SALDO queda negativo
06/08               esas filas SE BORRARON del ledger en vivo (limpieza ya hecha y
                     documentada en LEER_ANTES.md, sección "el ledger perdió 52 filas")
                     → shared/seguimiento_pueblo.xlsx HOY: correcto, en positivo

03/08 16:12          planilla_2026-08.xlsx SE GENERÓ ANTES de esa limpieza (mismo día,
                     pero por otro motivo — parche de E-14B) → quedó con el snapshot
                     viejo, negativo, y 2_planilla no se volvió a correr desde entonces
```

No es un bug nuevo de código — es un **snapshot desactualizado**. El ledger ya está bien; el
arreglo es re-correr `2_planilla` para que jale los saldos frescos.

**No ejecutado — falta decidir:**

```
¿Ya se imprimieron/entregaron boletas de agosto con estos 11 predios?
  Si sí: re-correr 2_planilla cambia CONVENIO/MULTA/ACUERDOS de negativo a positivo →
         hay que avisar a esos 11 vecinos, no solo corregir el archivo.
  Si no: re-correr 2_planilla es seguro, sin aviso pendiente.

3_boletas se regeneró la última vez el 30-31/07 (ver LEER_ANTES.md) — ANTES del 03/08.
No está confirmado si corrió de nuevo después del 03/08. Verificar antes de re-correr
2_planilla.
```

---

## 4. Pendiente sin tocar

```
6 filas Sin_identificar en pendientes.xlsx (PLIN sin mensaje reconocible, K-17 arriba)
15 filas NO CUADRA de verificar_lotes.py sin candidato (del RETOMAR de la mañana, sección 3③)
```
