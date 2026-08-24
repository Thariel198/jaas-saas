# RETOMAR — auditoría junio+julio cerrada · reporte de re-imputación listo · 2026-08-04

Sesión larga, tres bloques: **(A)** se auditaron los ciclos junio y julio contra el
banco y quedó descartado que el sistema pierda pagos · **(B)** se sincronizó el
checklist de los reclamos del 01/08 · **(C)** se construyeron dos reportes nuevos
para decidir el ajuste masivo de la cascada.

**Nada se escribió en el ledger. Nada se commiteó.**

---

## ⚡ PRIMER PASO al retomar

1. **Mirar el PDF de re-imputación con los compañeros** —
   `4b_reclamos/outputs/reporte_reimputacion_cascada_2026-07.pdf`. Es la decisión
   grande que quedó lista y sin tomar. La primera página es la validación.
2. **Decidir `DECLARACION` vs `ABONO_REZAGADO`** para los 8 casos de "ya pagué"
   (sección 4 de acá). Bloquea 8 de las 13 correcciones pendientes del 01/08.
3. Lo mecánico que se puede hacer sin decidir nada: la fila de E-1 en
   `shared/reasignaciones_aplicacion.xlsx` (sección 3) y correr
   `4_pagos/consolidar_tanque.py` (sección 5).

---

## 1. Lo que se cerró — auditoría de junio y julio

**El detalle completo está en `LEER_ANTES.md`**, en las tres secciones nuevas que
abren el archivo. No repetirlo acá; acá solo el veredicto:

```
                          junio              julio
banco crudo ↔ procesado   1:1 exacto         1:1 exacto
mesas ↔ pagos_efectivo    6 difs con causa   0 difs
5b_validacion             3 alertas          1 alerta (+550)
```

```
De TODOS los descuadres, el único error real es uno solo, repetido:

   junio  S/100  (E-1)                 pagos al tanque que matchearon lote
   julio  S/550  (C-15·P-7·A-4·P-17)   directo → motor_matching nunca les llena
                                        CONCEPTO → se cuentan como agua Y como
                                        tanque a la vez

   afecta el REPARTO agua/tanque en los reportes, NO el total de plata recibida
```

**Conclusión operativa (confirmada por el usuario):** los reclamos de "ya pagué el
mes anterior" **no vienen de un error del sistema**. Quedan dos causas y solo una se
resuelve con papel:

```
① el pago SÍ se anotó y la cascada lo consumió en otra cosa
   caso probado: F-1, yape S/41 del 17/06 → agua 17 + mant 3 + mes ant 21 = 41
   la MULTA de 20 quedó viva. El recibo físico no aporta nada acá.

② el pago nunca se anotó  ← la causa dominante
   → CONSULTAR LAS HOJAS FÍSICAS DE REGISTRO / los recibos de los vecinos
```

---

## 2. Checklist de los reclamos del 01/08 — sincronizado

`3_boletas/inputs/reclamos_2026-08-01/README.md` estaba desactualizado: G-4 ya se
había cerrado el 03/08 (precursor + ledger) y seguía en `[ ]`.

```
ANTES  37 · hechas 22 · pendientes 15
AHORA  37 · hechas 24 · pendientes 13
```

```
pendientes 13 = YA_PAGO 10       K-9 · T-14 · K-8 · F-10 · F-1 · F-7 · D-6 · B-8×3
              + REASIGNACION 3   G-14 (CONVENIO · MULTA · ACUERDOS)
```

- **G-14 sigue bloqueado**: la directiva dio el resultado (MULTA 0→50, ACUERDOS
  21→50) pero no el origen de esos montos. Confirmar con ella antes de escribir.
  Su ledger sigue intacto (CONVENIO 38 · MULTA 0 · ACUERDOS 21).

---

## 3. E-1 — la única fila lista para escribir, sin decisión pendiente

```
shared/reasignaciones_aplicacion.xlsx   (una fila)

  MZ=E · LT=1 · CONCEPTO_ORIGEN=AGUA · CONCEPTO_DESTINO=TANQUE · MONTO=100
  MES_ANO = (vacío)  → SOLO REGISTRO, no lo aplica ninguna corrida
  MOTIVO  = el yape del 10/06 decía "E1 1 tanque"; tepago lo dejó sin concepto y
            5_cobranza lo aplicó como agua. Su agua (S/8) ya estaba pagada en
            efectivo → el 100 es 100% tanque.
  REF_TRANSACCION = pagos_yape_tepago 10/06/2026 19:08:13
```

**NO se re-corre junio.** Regla que quedó fijada esta sesión y vale como
metodología: *un ciclo cerrado solo se lee; se corrige con un asiento fechado hoy y
efecto declarado en el ciclo de origen. Lo único que se puede recalcular libremente
es lo que no escribe en el libro (el "lente", ej. `5b_validacion`).*

---

## 4. LA DECISIÓN ABIERTA — `DECLARACION` vs `ABONO_REZAGADO`

Los dos cierran la deuda igual. La diferencia es **si esa plata suma a la caja**.

```
                        deuda del vecino      caja de la JASS
  DECLARACION              → 0                  sin cambio
  ABONO_REZAGADO           → 0                  + el monto
```

Hay una **tensión real entre dos documentos del repo** y hay que resolverla:

```
LEER_ANTES.md:570  (30/07)          seguimiento_repo.py:67  (03/08)
"se cargan igual como abono          "DECLARACION no está en CLASES_SUMAN_CAJA
 rezagado, sin exigir comprobante     a propósito: esa plata ya entró y ya se
 — el respaldo real es el pool de     contó (como exceso sin atribuir)"
 EXCESO no reclamado"
        │                                        │
        └────────────── MISMO argumento ─────────┘
                    conclusiones OPUESTAS sobre la caja

   si el respaldo es "ya está en el pool de exceso", entonces esa plata YA
   está en la recaudación → marcarla ABONO_REZAGADO la cuenta dos veces.
   Es el mismo "Riesgo real: doble conteo" que esa sección advierte.
```

**Criterio que resolvería esto con un número** (propuesto, no ejecutado): medir el
pool de EXCESO sin atribuir de junio+julio.

```
pool ≥ ~300  → la regla de la secretaria se sostiene → DECLARACION, sin doble conteo
pool ≈ 0     → no hay contrapartida → ABONO_REZAGADO, como G-4
```

Precedente ya aplicado: **G-4**, cuyo `MOTIVO` dice textual *"no aparece en mesas,
yape, blancos ni en el pool de exceso […] es plata real que nunca entró al registro:
por eso ABONO_REZAGADO"*.

**Casos y su tratamiento propuesto:**

```
K-8 · D-6 · F-1 · F-7 · F-10 · B-8(ACUERDOS) · B-8(CONVENIO)   → 7 filas, decidir clase
K-9 (S/100) · T-14 (S/75), efectivo en el local                → PREGUNTAR A LA
     son los únicos con monto y lugar concretos, y sobran 25      SECRETARIA:
     en los dos → si confirma que lo recibió, ABONO_REZAGADO      "¿recibiste ese efectivo?"
B-8 MES_ANTERIOR (46)  → es agua, va a shared/ajustes_cargo.xlsx, no al ledger
```

---

## 5. Reportes nuevos — dos scripts, dos PDF

### 5a. `4b_reclamos/reporte_deuda_ledger.py`
Padrón completo de deuda viva. Salida en la ruta que pidió el usuario:
`outputs/reporte_convenio_multa_referencias_2026-07.pdf` (216 pág).
El PDF anterior de convenio quedó respaldado en
`outputs/backups/reporte_convenio_multa_referencias_2026-07_PRE_deuda_ledger_20260804.pdf`.

```
MULTA        95 predios   S/  3,823.00
ACUERDOS    128 predios   S/  6,856.50
CONVENIO     72 predios   S/  5,800.00
208 predios distintos     S/ 16,479.50
```

### 5b. `4b_reclamos/reporte_reimputacion_cascada.py`  ← **la pieza para decidir**
Simula la cascada **CA1** (`CONVENIO → ACUERDOS → MULTA`) sobre los pagos de
**feb–jul 2026**, contra el orden que aplica el código hoy
(`MULTA → ACUERDOS → CONVENIO`). La multa cede porque es lo único que se cubre con
faena o se exonera.

```
① solo si el convenio es MEDIDOR:  MULTA → CONVENIO, y si no alcanza ACUERDOS → CONVENIO
② para los 208 sin excepción:      lo que sobre de MULTA → ACUERDOS

clases (las 3 últimas no reciben en convenio, pero sí entran al paso ②)
   MEDIDOR        62   convenio ≤100 y fuera de las listas del Excel
   SIN_CONVENIO  136   ← los "ya pagué techado y campo" que nunca debieron medidor
   INSTALACION     9   6 por monto >100 + 3 de NUEVAS INSTALACIONES / ANTERIOR DIRECTIVA
   REACTIVACION    1   M-12
lista fuente: obligaciones/inputs/SEGUMIENTO INSTALACIONES Y CAMBIO DE MEDIDORES.xlsx
```

**Validación (todas OK, está en la pág. 1 del PDF):**

```
                              ANTES        DESPUÉS
predios con deuda              208    →       208     nunca sube ✔
deuda MULTA                3,823.00   →  6,213.50     absorbe
deuda ACUERDOS             6,856.50   →  6,000.00
deuda CONVENIO             5,800.00   →  4,266.00
DEUDA TOTAL               16,479.50   → 16,479.50     IDÉNTICA ✔
pagos MULTA (feb-jul)      3,156.00   →    765.50
pagos ACUERDOS             2,371.50   →  3,228.00
pagos CONVENIO             5,104.00   →  6,638.00
PAGOS TOTALES             10,631.50   → 10,631.50     IDÉNTICOS ✔
deuda no conservada: 0 predios · saldo negativo: 0 predios ✔
```

**Qué resuelve:**

```
MEDIDOR (62)           35 medidor SALDADO · 7 parcial · 20 sin efecto
ACUERDOS (128 deben)   23 techado y campo SALDADO · 24 parcial
87 de 208 se mueven · 65 terminan debiendo SOLO multa
```

También hay `.xlsx` con las 208 filas (antes · movimientos · después) para filtrar o
mandar aparte.

**Pendiente:** el usuario adelantó que este reporte probablemente se use después
para **aplicar un ajuste masivo al ledger**. No está diseñado ese paso todavía — el
script de hoy es solo simulación, no escribe nada.

---

## 6. Deuda técnica que dejó la sesión (nada urgente, todo anotado)

```
① motor_matching: llenar CONCEPTO=tanque cuando el MENSAJE lo diga, aunque el lote
   matchee directo. Cubre C-15, A-4, P-17 y el E-1 de junio. NO cubre P-7 (se
   confirmó de palabra) — ese siempre va a necesitar el precursor manual.

② 5b_validacion, tres arreglos del "lente" (no tocan ningún dato de un ciclo cerrado):
   · lista blanca de CONCEPTO en vez de "excluye todo lo no vacío" (main.py:358)
   · sumar los blancos de efectivo asignados en la sección EFECTIVO
   · restar aportes_tanque_manuales.xlsx (canal yape) del bucket «agua» del Nivel 1a
   · leer los asientos existentes para marcar EXPLICADA en vez de ALERTA

③ consolidar_tanque.py no se corre desde el 25/07 → O-16 (100) y H1-3 (100) están
   en aportes_tanque_manuales.xlsx pero NO en aportes_tanque.xlsx. Cuando corra, la
   dif de Nivel 1a de julio pasa de 550 a 750 (mismo origen, no es nuevo).

④ estado_ciclo.json de 2026-07 sigue con gap_conocido "causa raiz sin resolver".
   Ya está resuelta (es la sección de julio de LEER_ANTES) — retirar ese texto
   cuando se cierren los arreglos de ②.

⑤ S/200 "falta ubicar lote" al pie de mesa_2 (junio, Yerald Romero): sin rastro en
   pagos_efectivo, blancos_mes ni blancos_acumulados. El usuario los da por
   explicados por conocimiento propio. Si un balance no cierra por ~200, empezar acá.
```

---

## 7. Archivos tocados hoy

```
MODIFICADOS
  LEER_ANTES.md                                        3 secciones nuevas al tope
  3_boletas/inputs/reclamos_2026-08-01/README.md       checklist 22→24, G-4 cerrado

NUEVOS
  4b_reclamos/reporte_deuda_ledger.py
  4b_reclamos/reporte_reimputacion_cascada.py
  4b_reclamos/outputs/reporte_reimputacion_cascada_2026-07.pdf  (+ .xlsx)
  docs/retomar/RETOMAR_auditoria_junio_julio_y_reimputacion_2026-08-04.md   (este archivo)

REGENERADOS
  4b_reclamos/outputs/reporte_convenio_multa_referencias_2026-07.pdf
  5b_validacion/outputs/  de los repos junio y Julio (solo outputs, nada del libro)

SIN COMMITEAR — todo lo de arriba.
```
