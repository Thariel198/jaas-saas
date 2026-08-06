# RETOMAR — pagos que no eran pagos · 2026-08-06 (tarde)

Segunda sesión del mismo día. La de la mañana cerró el ruido que el sistema se había hecho
a sí mismo (`docs/diario/2026-08-06_limpieza_del_ledger.html`); esta encontró el problema de
fondo: **varias cosas distintas se estaban leyendo como "pago"**.

Bitácora con el razonamiento: `docs/diario/2026-08-06_tarde_pagos_que_no_eran_pagos.html`.

---

## ⚡ PRIMER PASO al retomar

1. **Cerrar los 3 AJUSTE de F-12** — el texto ya existe, solo hay que escribirlo en el
   `MOTIVO` del ledger. Está en `shared/reasignaciones_aplicacion.xlsx`, fila de F-12:
   *"asistió a reunión y faena, la directiva no registró su asistencia, por eso NO
   corresponde la MULTA que aparece pagada; se redirige el pago a CONVENIO/MEDIDOR"*.
   Los tres audit_ref: `F-12-MULTA-redirigido-31072026` · `F-12-MULTA-estabilizador-31072026`
   · `F-12-CONVENIO-correccion-etiqueta-31072026`.
2. **Arreglar el `EXCESO` del reporte de bug** (`4b_reclamos/reporte_correccion_bug.py`):
   hoy es `Σ PAGO − Σ CARGO` e **ignora la columna AJUSTE**, por eso difamó a F-12. Es lo
   que va a seguir generando falsos positivos mientras se cierran los 35 restantes.
3. **Verificar C-15, P-7 y P-17 en el ledger** — mismo lote de aporte al tanque que A-4.
   Dan 0 en `planilla_cobrado`, pero A-4 también daba 0 y estaba mal.

---

## 1. Lo que se cerró

### 1a. A-4 — el aporte al tanque le pagó el convenio (commit `8d95608`)

```
07/07 07:20:05  yape S/136  "mz A lt 4"          → deuda normal
07/07 07:20:55  yape S/100  "mz A lt 4 tanque"   → aporte al tanque, NO es deuda
        │
        └─ motor_matching identificó el lote pero dejó CONCEPTO vacío
           ⇒ los 236 entraron juntos y la cascada se los comió como deuda

con los 136 REALES: 136 − 31 agua − 30 multa − 75 acuerdos = 0 → CONVENIO 75 impago

planilla_cobrado  SALDO 75 · PARCIAL   ✔ siempre tuvo razón (5_cobranza:1422 ya lee
                                          aportes_tanque_manuales antes de la cascada)
ledger            CONVENIO 0           ✗ se había escrito antes de esa segregación
```

Se borraron las 5 filas posteriores a la siembra (2 `PAGO` fantasma + 3 `AJUSTE` con que el
sistema se corregía a sí mismo). Queda el `CARGO`: **saldo 75**. Ledger 1491 → 1486, cero
saldos negativos en todo el pueblo.

**La secretaria ya había reclamado que A-4 debía y se le ignoró** porque el sistema mostraba
el pago de S/236. Backup: `shared/backups_ledger/seguimiento_pueblo_pre_A4_convenio_20260806_153740.xlsx`.
El error está en **todos** los backups (se escribió el 08/07; el más viejo es del 03/08) —
documentado en la sección nueva al tope de `LEER_ANTES.md`.

### 1b. La vista dejó de mentir (commits `5cf8c18` y `f790109`)

```
antes   │ DEUDA │ PAGO │ SALDO │                        DEUDA = CARGO + AJUSTE
ahora   │ DEUDA │ PAGO │ DECLARADO │ AJUSTE │ SALDO │   DEUDA = CARGO

DEUDA      solo lo que se le CARGÓ
PAGO       CLASE ∈ CLASES_SUMAN_CAJA  → plata real, sumable para un balance
DECLARADO  CLASE fuera                → ámbar: salda deuda, NO es caja
AJUSTE     ROJO si ninguno de esa celda tiene MOTIVO
·          gris = no pasó nada   ≠   0 = pasó algo que se cancela solo

hoja "Ajustes" nueva: 151 filas, una por evento, con CLASE · SOURCE · MOTIVO completo
```

Dos números que justifican el corte:

```
Σ PAGO del ledger                 S/ 21,025.50
Σ PAGO que entró a caja           S/ 19,768.50
                                  ─────────────
                                  S/  1,257.00   28 filas DECLARACION_SECRETARIA

570 celdas con PAGO:  542 solo caja · 23 solo declarado · 5 MIXTAS
   B1-12 · S-1 · S-5 · S-8 · T-7 (julio) — S-1 mostraba "PAGO 75" siendo 20 + 55
```

**El ledger no se tocó.** La separación ya existía ahí como *atributo* (`CLASE`); al pivotear
a formato ancho la celda agrega varios eventos y no puede llevar el atributo adentro, así que
la dimensión se vuelve columna. `estado_cuenta()` alineado (no tenía consumidores en
producción). Contrato `formato_vista_seguimiento_pueblo.html` a **v1.4**.

### 1c. Otros

```
7 commits de backlog (e46c84b → 6889499)  186 entradas sin commitear desde el 10/07
                                          libro_mayor/ · obligaciones/ · backfill_ledger/
                                          entran a git por primera vez
vista vacía diagnosticada y regenerada    25fd566 — causa NO determinada (el ledger estaba
                                          completo 16 min antes); es derivada, se regeneró
reporte_correccion_bug.py                 e2165f7 — 6 formas del bug + referencia de pago
                                          contra el crudo de yape/efectivo
README raíz                               sumado backfill_ledger/ (faltaba, Regla 7)
```

---

## 2. Dos hallazgos que valen para el diseño, no solo para el caso

### `_saldo_previo` no ve la génesis tardía

```
30/07 11:47  CARGO   58  mes 06  saldo  58   ← génesis tardía a un mes YA cerrado
30/07 11:55  PAGO    58  mes 07  saldo −58   ← pagó y el saldo se va a negativo
30/07 14:04  AJUSTE +58  mes 07  saldo   0   ← parche manual

_saldo_previo ordena ["MES","TIMESTAMP"] y toma la última fila → gana una fila de
julio cuyo SALDO se calculó el 27/07, cuando el cargo de junio todavía no existía.
```

**El patrón:** una génesis tardía escribe un CARGO en un mes pasado, pero las filas ya
escritas en meses posteriores conservan su `SALDO` viejo — nadie las recalcula. La deuda
retroactiva queda invisible para el saldo corriente. P-6 llegó a 0 por un parche manual, no
porque el ledger lo calculara. **No está arreglado.**

### Una fórmula mecánica sobre columnas agregadas difama

`EXCESO = Σ PAGO − Σ CARGO` acertó en A-4 por casualidad (no había ajustes que compensaran) y
señaló como fantasma a F-12, donde `−25` de fórmula y `+50` de corrección de etiqueta hacían
que el neto real fuera 0. Cualquier regla que se escriba sobre la vista tiene que mirar las
cuatro columnas, no dos.

---

## 3. Pendientes por riesgo

```
⚠ ANTES DE QUE CORRA AGOSTO
   5_cobranza/main.py:2320 — el AJUSTE de reversión sigue con el signo invertido.
   Sin decidir desde el 06/08 mañana. Ver 3_boletas/inputs/reclamos_2026-08-01/README.md
   § BUG_SIGNO. Es lo único que puede volver a fabricar saldos negativos.

   motor_matching no marca CONCEPTO=tanque cuando el mensaje lo dice y el lote
   matchea directo (solo se llena a mano vía pendientes.xlsx). Sin eso, el bug de
   A-4 se repite este mes. Cubre C-15, A-4 y P-17; NO cubre P-7 (se confirmó de palabra).

AVISAR A VECINOS
   A-4      debe S/75 de convenio — si ya se le dijo que estaba al día, lo va a notar
   E-12 · L-5 · F-4 · W-5   boleta del 01/08 corta: en septiembre sube sin consumo

DESALINEACIONES CONOCIDAS, NO URGENTES
   F-12: planilla_cobrado dice CONVENIO 50, el ledger y DATA_boletas dicen 0. El
     redirect se ejecutó con 5_cobranza --force en el repo de Julio y se sincronizó a
     mano. La boleta salió bien; un balance leído de planilla_cobrado lo vería debiendo.
   validar_vista_boletas.py FALLA con 70 desajustes de MULTA — ya fallaba antes de esta
     sesión (verificado contra HEAD): desfase DATA_boletas (agosto) vs ledger (cierre junio).

DECISIONES DE NEGOCIO ABIERTAS (vienen del 04/08 y del 06/08 mañana)
   los 28 DECLARACION_SECRETARIA: ¿exceso ya en caja (→ DECLARACION) o pago nuevo
   (→ abonos_rezagados → ABONO_REZAGADO)? Ahora son visibles en la columna DECLARADO.
   El PDF de re-imputación de la cascada sigue esperando la charla con los compañeros.
```

---

## 4. Los 35 AJUSTE sin MOTIVO — dónde quedó el mapa

`4b_reclamos/outputs/reporte_correccion_bug_2026-07.pdf` (+ `.xlsx`) los agrupa por forma.
Correr con `py 4b_reclamos/reporte_correccion_bug.py`.

```
1. PAGO FANTASMA    1 par  ·  1 fila  ← solo F-12, y es falso positivo (ver §2)
2. CARGO ANULADO    8 pares·  9 filas    C-19 · C-29A · C1-17 · F-3B · F1-10 · Q-5 · S-14 ×2
3. PAGO PARCIAL     3 pares·  4 filas    D-16 · L-4 · R-5
4. PAR QUE NETEA 0  4 pares· 10 filas    D1-3 · F-12 MULTA · Q-4 ×2
5. GENESIS TARDIA   2 pares·  4 filas    P-6 · Q-16
6. DEUDA REABIERTA  4 pares·  7 filas    C-21 · D1-6 · J-6 ×2   → S/176 de saldo vivo
```

Ojo con dos: **F-12 MULTA y F-12 CONVENIO son las dos mitades de una reasignación** — cerrar
una sin la otra rompe el par. Y **D1-6 tiene saldo 1.00**, un céntimo suelto de `30 − 17 − 12`
que nadie había visto.

Los 5 "duplicados" del 31/07 (C-19, C-29A, C1-17, S-14 ×2) tienen el CARGO **una sola vez** en
el ledger: "duplicado" significa que la deuda ya estaba contada en otro archivo. Para escribir
un MOTIVO honesto hay que abrir `4b_reclamos/pendientes_secretaria/notas_2026-07.xlsx` y decir
dónde estaba la otra copia.

---

## 5. Archivos tocados

```
COMMITEADO Y PUSHEADO (11 commits, e46c84b → f790109)
  e46c84b  libro_mayor/ + obligaciones/ + backfill_ledger/ + docs de diseño · retira 7b_historial_pagos
  6a5bd49  6_corte · candado Día 0 + lista_sin_servicio
  5e3bdf3  4_pagos · arqueo de efectivo + entregas
  33684e3  4b_reclamos · 6 reportes + insumos secretaria
  615fb24  3_boletas · recibos MP- + corregidas del 01/08
  a027f9e  shared · overlays del ledger        ⚠ incluyó la vista VACÍA por error
  6889499  docs · 22 RETOMAR + .gitignore/.idea
  25fd566  regenerar la vista
  e2165f7  reporte_correccion_bug.py
  8d95608  A-4 · el aporte al tanque le pagó el convenio
  5cf8c18  vista · AJUSTE fuera de DEUDA + hoja Ajustes
  f790109  vista · PAGO partido en caja/DECLARADO

REGRESIÓN CORRIDA
  shared/tests/test_seguimiento_repo.py        TODOS LOS CHECKS PASARON (actualizado al layout nuevo)
  5_cobranza/tests/test_reconciliacion_pueblo  TODOS LOS CHECKS PASARON
  2_planilla/tests/test_publicar_shared        3/3
  (los tests de este repo se corren como script: py test_x.py, no con pytest)
```
