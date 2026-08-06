# RETOMAR — ciclo activo + limpieza del ledger · 2026-08-05/06

Sesión larga de dos días en un solo chat. Empezó mirando **un pago fantasma de K-9** en el
PDF de re-imputación y terminó con **la causa raíz de los 15 pagos fantasma de julio**
cerrada, el **mes del ciclo dejando de ser una adivinanza** y **96 filas de ruido fuera del
ledger**.

Bitácora completa, con el razonamiento: `docs/diario/2026-08-05_pagos_de_junio_cobrados_en_julio.html`
y `docs/diario/2026-08-06_limpieza_del_ledger.html`.

---

## ⚡ PRIMER PASO al retomar

1. **Commitear lo que quedó suelto** (nada roto, solo sin commit):
   `shared/seguimiento_repo.py` · `5_cobranza/main.py` · `5_cobranza/tests/test_reconciliacion_pueblo.py` · `docs/diario/` (untracked)
2. **Seguir con los 38 AJUSTE sin MOTIVO** — el trabajo quedó cortado justo ahí, con el
   análisis hecho y el plan por bloques (sección 4 de este documento).
3. Antes de que corra el ciclo de agosto, decidir el **signo del AJUSTE de reversión**
   (`5_cobranza/main.py:2320`). Es lo único que puede volver a fabricar saldos negativos.

---

## 1. Lo que se cerró — el incidente del 06/07

```
06/07 14:08-14:13  UNA corrida de 5_cobranza cobró el ciclo 2026-07 leyendo
                   pagos_yape_tepago.xlsx que todavía era el de junio
                   (el run.log lo prueba: "Pagos Yape → 65 filas · 1 sin identificar"
                    = 66 = exactamente el archivo de junio; el de efectivo sí era de julio)
                        │
                   los 15 afectados pagaron por YAPE en junio · ninguno por efectivo
                   monto de cada fantasma = pago de junio − agua de julio, topado por
                   la deuda → cuadra exacto en 14 de 15 (C-1 difiere 5: tiene dos
                   exoneraciones de multa cargadas el 23/07)
                        │
31/07 18:08-18:20  5_cobranza los revierte… con el signo invertido → saldos negativos
03/08              se restauran 12 a mano (+2 × el monto revertido)
06/08              se borran las 52 filas del bug (ver sección 2)
```

---

## 2. Lo que se ejecutó

### 2a. Ciclo activo declarado (05/08, commit `c29de82`)

```
plantilla del operario (columna MES_ANO)
   → 1_lecturas exige que todas las filas sean del mismo mes y lo declara
      → shared/ciclo_activo.json
         → 2_planilla y 5_cobranza lo leen (se acabó el sorted()[-1])

outputs de 4_pagos con periodo en el nombre + frontera explícita:
   PRIMER_CICLO_CON_PERIODO = "2026-08"   (shared/ciclo.py)
   ciclo ≥ 2026-08 → solo el nombre con periodo; si 4_pagos no corrió, 5_cobranza CORTA
   ciclo ≤ 2026-07 → acepta el nombre pelado; re-correr julio sigue funcionando
```

También: 3 archivos del repo de junio renombrados a `_2026-06`, y los reportes de
`4b_reclamos` resolviendo por helper (`ciclo.resolver`) en vez de rutas a mano — dos
apagones silenciosos el mismo día venían de ahí.

### 2b. Limpieza del ledger (06/08, commit `9ce2a3a`)

```
1589 → 1491 eventos · 0 saldos finales negativos

52 filas · incidente del 06/07     19 pagos fantasma + 18 reversiones + 15 restauraciones
44 filas · ruido de declaraciones  22 pares (AJUSTE de 5_cobranza + estabilizador manual)

criterio, para reusarlo:
   SE BORRA     corrige un error del sistema sobre sí mismo · neto 0 (o el neto ES el error)
   SE CONSERVA  corrige un dato del negocio · su efecto es parte de la verdad de hoy

cuatro saldos se arreglaron solos al borrar:
   E-12 ACUERDOS 0→25 · L-5 MULTA 34→50 · F-4 MULTA 48→50 · W-5 ACUERDOS 37→47
   ⚠ sus boletas del 01/08 salieron cortas → en septiembre suben sin consumo. AVISARLES.
```

Backups versionados en git (el `LEER_ANTES.md` apunta a ellos):
`pre_eliminar_bug0607_20260806_081824` · `pre_motivo109_20260806_082105` ·
`pre_pares_condonacion_20260806_093522` · `pre_clase_declaracion_20260806_093613`.

### 2c. Clasificación y motivos

```
CLASE nueva DECLARACION_SECRETARIA   28 pagos · S/ 1,257.00 · con MOTIVO
   la secretaria dijo que ya pagó; el pago vale y salda la deuda
   ⚠ NO resuelve de dónde salió esa plata → sigue fuera de CLASES_SUMAN_CAJA

MOTIVO escrito en 109 AJUSTE de correccion_genesis_formula
   bug de fórmula de abril: el Saldo de génesis omitía los pagos de abril,
   el convenio salió sobrecobrado, ya comunicado con 109 boletas reimpresas
```

### 2d. Arreglos de código (SIN COMMITEAR)

```
pago_registrado(..., source=None)      seguimiento_repo.py:361
   filtra por SOURCE — simétrico con ajuste_reconciliado, que ya lo hacía
   5_cobranza pasa source="5_cobranza" ⇒ los PAGO manuales le son invisibles
   ⇒ el ruido de la sección 2b NO puede regenerarse

aviso de exceso en _registrar          el writer único, pasan todos por ahí
   un PAGO que deja el saldo bajo 0 avisa por log y devuelve {"exceso": N}
   SIN TOPE, decidido: recortar el monto borraría la única señal de que el vecino
   pagó de más (ese exceso no llega a arrastre_devolucion, que se arma del SALDO
   de la planilla, y la planilla creía que la deuda existía)

test_reconciliacion_pueblo             11/11 · se le sembró el CARGO que le faltaba
   (reconciliaba contra un concepto inexistente, imposible en producción) + caso 5
```

---

## 3. Regresión corrida

```
shared/test_seguimiento_repo.py          TODOS LOS CHECKS PASARON (incluye idempotencia)
5_cobranza/test_reconciliacion_pueblo    11/11
2_planilla/test_publicar_shared          3/3
1_lecturas (3 tests)                     todos OK
4_pagos/efectivo                         16/16 integración · 27/27 unitarios

fallos PREEXISTENTES, verificados contra HEAD (no los introdujo esta sesión):
   5_cobranza/test_cobranza          "trazabilidad · filas 6 vs 5"
   motor_matching/test_integracion y test_fixes
   4b_reclamos/test_reclamos         lee el archivo real en vez de su fixture
ojo: los tests de este repo se corren como script (py test_x.py), no con pytest
```

---

## 4. PENDIENTE INMEDIATO — los 38 AJUSTE sin MOTIVO

Análisis ya hecho (23 predio-concepto). El plan, de lo más seguro a lo que necesita decisión:

```
①  solo redacción, fuente documentada — 9 filas
     duplicados (cargo sembrado dos veces): C-29A −20 · C1-17 −30 · Q-16 −75 · S-14 ×2
     fix race condition yape (RETOMAR 27/07): A-4 +75 neto · L-4 −3 neto · P-6 +58 neto
     relabeling F-3B → F-3A: −50

②  leer notas_2026-07.xlsx para contarlo bien
     revertir condonación fallida: C-21 +50 · J-6 MULTA +50 · J-6 ACUERDOS +75
     Q-4 MULTA y Q-4 ACUERDOS (neto 0, secuencia condonación→revertir→recondonar)

③  correcciones puntuales de la secretaria — hay que ver qué dijo en cada una
     C-19 MULTA −50 · F1-10 MULTA −30 · R-5 MULTA −12

④  investigación real, acá sí hay que decidir hecho vs ruido
     D-16 ACUERDOS −25 y D1-6 MULTA −12   ← los del abono rezagado, siguen pendientes
     Q-5 CONVENIO +25 −50                 ← "corrección de signo", entender qué se corrigió
     F-12 MULTA −50/+50 y F-12 CONVENIO +50 ← son las dos mitades de una reasignación:
                                              borrar solo la de MULTA rompe el par
     D1-3 MULTA −18/+18                   ← par simple, candidato a ruido
```

---

## 5. Otros pendientes, por riesgo

```
⚠ ANTES DE QUE CORRA AGOSTO
   5_cobranza/main.py:2320 — el AJUSTE de reversión sigue con el signo invertido.
   Dos caminos, ninguno decidido: ver 3_boletas/inputs/reclamos_2026-08-01/README.md
   § BUG_SIGNO. Es lo único que puede volver a fabricar saldos negativos.

DECISIONES DE NEGOCIO ABIERTAS
   los 28 DECLARACION_SECRETARIA: ¿exceso ya en caja (→ DECLARACION) o pago nuevo
   (→ sembrar en abonos_rezagados → ABONO_REZAGADO)? Es la misma decisión que el
   04/08 dejó abierta para los 8 casos de "ya pagué".
   El PDF de re-imputación de la cascada sigue esperando la charla con los compañeros
   (4b_reclamos/outputs/reporte_reimputacion_cascada_2026-07.pdf).

TRABAJO ANOTADO
   avisar a E-12, L-5, F-4 y W-5 que su boleta de septiembre sube (boleta corta el 01/08)
   cruzar los avisos de exceso contra planilla_cobrado para separar el exceso real
     del negativo por defecto del sistema
   24 filas con SALDO negativo transitorio en el ledger — entre ellas A1-12, A1-13,
     C1-13, C1-15, H1-13, H1-36: PAGO con saldo −75, o sea pagos contra un CARGO
     que no estaba. Patrón distinto, sin investigar.
   tanda B parte 2: planilla_cobrado y trazabilidad todavía sin periodo en el nombre
   3 overlays que faltan en la "referencia de pago" del PDF: devoluciones_aplicadas,
     reidentificacion (tiene dos lados, hay que decidir el signo) y aportes_tanque_manuales
     (es plata que NO salda deuda)
   netear los AJUSTE de reversión en el historial del PDF — depende del signo (§5, primer punto)
```

---

## 6. Archivos tocados

```
COMMITEADO Y PUSHEADO
  c29de82  shared/ciclo.py (nuevo) · ciclo_activo.json · 1_lecturas · 2_planilla ·
           5_cobranza · 4b_reclamos/reporte_historico.py y reporte_referencias_pago.py
  9ce2a3a  LEER_ANTES.md (entra a git por primera vez) · seguimiento_pueblo.xlsx ·
           seguimiento_repo.py (CLASE nueva) · vista_seguimiento_pueblo.xlsx/.pdf ·
           4 backups del día

SIN COMMITEAR
  shared/seguimiento_repo.py                    filtro por source + aviso de exceso + TOL_SALDO
  5_cobranza/main.py                            usa el filtro + su propio aviso
  5_cobranza/tests/test_reconciliacion_pueblo.py  cargo sembrado + caso 5
  docs/diario/ (untracked)                      2 entradas nuevas + index actualizado
  docs/RETOMAR_…_2026-08-06.md                  este archivo

FUERA DE GIT (otro repositorio)
  Junio/jass_system - junio  →  3 archivos renombrados a _2026-06
```
