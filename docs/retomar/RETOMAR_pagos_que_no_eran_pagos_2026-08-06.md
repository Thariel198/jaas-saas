# RETOMAR — pagos que no eran pagos · 2026-08-06 (tarde)

Segunda sesión del mismo día. La de la mañana cerró el ruido que el sistema se había hecho
a sí mismo (`docs/diario/2026-08-06_limpieza_del_ledger.html`); esta encontró el problema de
fondo: **varias cosas distintas se estaban leyendo como "pago"**.

Bitácora con el razonamiento: `docs/diario/2026-08-06_tarde_pagos_que_no_eran_pagos.html`.

---

## ⚡ PRIMER PASO al retomar

**① ARREGLAR EL SIGNO DEL AJUSTE DE REVERSIÓN — antes de cargar los pagos de agosto y
antes de generar la lista de corte.** Decidido al cierre del 06/08. Detalle abajo en §0.

Recién después:

2. **Cerrar los 3 AJUSTE de F-12** — el texto ya existe, solo hay que escribirlo en el
   `MOTIVO` del ledger. Está en `shared/reasignaciones_aplicacion.xlsx`, fila de F-12:
   *"asistió a reunión y faena, la directiva no registró su asistencia, por eso NO
   corresponde la MULTA que aparece pagada; se redirige el pago a CONVENIO/MEDIDOR"*.
   Los tres audit_ref: `F-12-MULTA-redirigido-31072026` · `F-12-MULTA-estabilizador-31072026`
   · `F-12-CONVENIO-correccion-etiqueta-31072026`.
3. **Arreglar el `EXCESO` del reporte de bug** (`4b_reclamos/reporte_correccion_bug.py`):
   hoy es `Σ PAGO − Σ CARGO` e **ignora la columna AJUSTE**, por eso difamó a F-12. Es lo
   que va a seguir generando falsos positivos mientras se cierran los 35 restantes.
4. **Verificar C-15, P-7 y P-17 en el ledger** — mismo lote de aporte al tanque que A-4.
   Dan 0 en `planilla_cobrado`, pero A-4 también daba 0 y estaba mal.

---

## 0. EL SIGNO — por qué va primero (analizado el 06/08 al cierre)

### Por qué importa más de lo que parece: el ledger es plata, no reporte

```
2_planilla/main.py:148   _join_saldo_pueblo()
   MULTA · ACUERDOS_ASAMBLEA · CONVENIO  ←  repo.get_saldos_bulk(concepto, mes_ant)
                                                      │
   ledger  →  planilla del mes siguiente  →  boleta  →  lo que se le cobra al vecino
```

Un saldo torcido en el ledger **no queda en el ledger**: se factura el mes siguiente.

### Cuándo dispara (y cuándo no)

`5_cobranza/main.py:2372` es el **único** `registrar_ajuste` de producción. Necesita
`delta < 0`, o sea `pagado_fresco < ya`, donde `ya` es lo que esta misma reconciliación ya
acreditó **del mes que corre**.

```
primera corrida de agosto:  el ledger no tiene ningún PAGO con MES=2026-08
                            ⇒ ya = 0 ⇒ delta ≥ 0 ⇒ SOLO PAGO, nunca AJUSTE

dispara SOLO en una re-corrida donde el insumo encogió
   (un pago que estaba y ya no está · un --force tras corregir el crudo)
   es exactamente lo que pasó el 06/07 y el 31/07
```

### Qué hace mal, y en qué dirección

```
delta < 0 significa "acredité de más, hay que DEVOLVER la deuda"
   correcto:  ajuste +75  →  el saldo SUBE, el vecino vuelve a deber
   hoy:       ajuste −75  →  el saldo BAJA otra vez

   queda 2 × el monto por debajo de la verdad  →  LA DEUDA DESAPARECE
```

Se le cobra de menos y **nadie reclama por una boleta más barata**. Es la brecha
caja↔deuda, en silencio.

### ⚠ El chequeo de "0 saldos negativos" NO lo detecta

En julio se descubrió porque el saldo quedó negativo — pero eso pasó porque la deuda era
chica. Con deuda suficiente el error se absorbe y queda positivo:

```
cargo 200 · run 1 acredita 75   → saldo 125
run 2, el pago ya no está       → saldo  50    ✗ debería ser 200
                                   positivo, sin alarma, invisible
```

**El detector correcto** es contar los AJUSTE con `source="5_cobranza"` y `MES` del ciclo:
si hay alguno después de una corrida, mirarlo uno por uno.

### El arreglo son DOS líneas acopladas, no una

Por eso quedó sin hacer — invertir solo el signo escrito rompe la idempotencia.

```
hoy                                    verificado en pizarra
  ya  = pago_registrado + ajuste_reconciliado
  esc = delta                (−75)     run 2: 75 − 75 + (−75) = −75   ✗

arreglo — las dos mitades juntas       (main.py:2349-2350 y 2374)
  ya  = pago_registrado − ajuste_reconciliado
  esc = −delta               (+75)     run 2: 75 − 75 + 75 = 75       ✓ deuda restaurada
                                       run 3: ya = 75 − 75 = 0 → delta 0 → no escribe ✓
                                       run 4: reaparece el pago → PAGO 75 → saldo 0    ✓
```

**Test que hay que escribir** (no existe): la secuencia completa
`corrida → el insumo encoge → re-corrida → el pago reaparece`, comprobando el saldo en cada
paso y que la re-corrida sin cambios no escriba nada.

### Si por lo que sea se decide NO arreglarlo antes de correr

Protocolo mínimo, en este orden:

```
① backup del ledger ANTES de 5_cobranza
     shared/backups_ledger/seguimiento_pueblo_pre_agosto_<ts>.xlsx
② 5_cobranza UNA sola vez, con 4_pagos ya definitivo
③ contar AJUSTE nuevos con source="5_cobranza" y MES=2026-08
     0   → no se disparó, limpio
     >0  → PARAR. Cada uno está al revés; revertirlo a mano ANTES de que
           2_planilla lo lea para septiembre
④ NO correr 5_cobranza --force por ningún motivo hasta que el signo esté arreglado
```

### La lista de corte NO depende de esto

```
6_corte NO lee el ledger — lee planilla_cobrado (verificado con grep)
```

Por eso la boleta de A-4 salió bien aunque el ledger estuviera mal. Los riesgos de la lista
de corte son otros dos, los dos en `LEER_ANTES.md`: el hueco `4b_reclamos → 6_corte` (un
reclamo RESUELTO no llega a `resolucion_reclamos_YYYY-MM.xlsx` y el vecino reaparece como
`EJECUTAR_CORTE=SI` — revisar a mano `trazabilidad_reclamos.xlsx` antes de publicar) y el
candado del Día 0 (una vez que corre `aplicar_penalidad.py`, `generar_lista` se niega a
regenerar: generar la lista solo cuando `5_cobranza` esté definitivo).

### Orden sugerido para el día

```
1. arreglar el signo + su mitad `ya`, con el test de la secuencia      ← Sonnet, ~30 min
2. cargar pagos · 4_pagos completo
3. 5_cobranza UNA vez
4. chequeo: ¿AJUSTE nuevos con source=5_cobranza? (deberían ser 0)
5. cruzar trazabilidad_reclamos.xlsx RESUELTO contra la lista
6. 6_corte generar_lista → publicar
```

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
   5_cobranza/main.py:2372 — el AJUSTE de reversión con el signo invertido.
   DECIDIDO 06/08: se arregla PRIMERO, antes de cargar pagos. Ver §0 arriba
   (dispara solo en re-corrida · el ledger se factura vía 2_planilla ·
    "0 saldos negativos" NO lo detecta · son 2 líneas acopladas).
   Contexto previo en 3_boletas/inputs/reclamos_2026-08-01/README.md § BUG_SIGNO.

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
