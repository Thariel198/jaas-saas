# RETOMAR — V-14 y B-8: parche manual con candado · decisión ① cerrada · 2026-08-02

Cierre de la decisión ① del plan del 02/08
(`docs/diario/2026-08-02_plan_reconciliacion_reclamos.html`, sección ③).

Las otras 23 correcciones del 01/08 **no dependen de esto** — van por el ledger
(`registrar_ajuste` con `MES=2026-07`) y no necesitan nada de este archivo.

---

## ⚡ PRIMER PASO al retomar

1. **Nada de esto está ejecutado todavía.** Es la decisión cerrada, no el trabajo hecho.
2. Antes de escribir la primera fila: **leer la fila existente de F1-4 en
   `shared/ajustes_cargo.xlsx`** (`header=1`) para copiar el esquema exacto de columnas.
   Este documento NO inventa los nombres de columna — el archivo real manda.
3. Antes de escribir la fila de B-8: hacer la **verificación de composición** de la
   sección 4. Si se salta, el saldo puede quedar negativo.

---

## 1. La decisión (tomada por el usuario, 02/08)

**Parche manual, sin re-correr `5_cobranza` de julio — pero con candado para que no
se pierda en corridas futuras.**

```
V-14   CORTE_RECONEXION   40 → 0    "no corresponde" (secretaria, 01/08)
B-8    MES_ANTERIOR       46 → 0    "no debe mes anterior" (atendido en persona)
```

Estas 2 filas son las únicas de las 25 del 01/08 que **no viven en el ledger de
pueblo**: `seguimiento_repo._validar_concepto` solo acepta MULTA · ACUERDOS ·
CONVENIO y tira `ValueError` con cualquier otro concepto. Para agua y corte
todavía **no existe ledger** — eso llega con `libro_mayor`.

---

## 2. Por qué "parche + candado" y no una sola cosa

El problema es que el dato vive en un archivo **congelado** que nadie re-deriva:

```
arrastre_consolidado_2026-07.xlsx      ← lo generó 5_cobranza de julio y quedó quieto
        │
        │  2_planilla lo LEE para armar agosto
        ▼                                  (DEUDA_AGUA → MES_ANTERIOR · CORTE → CORTE)
planilla_2026-08.xlsx
```

Ninguna opción sola alcanza:

```
SOLO parchar el .xlsx
   ✔ planilla_2026-08 sale correcta ya
   ✗ si algún día corre 5_cobranza julio --force, el archivo se regenera
     de cero y el parche desaparece SIN AVISO
     (es exactamente por esto que W-4, Q-5, O-16 y Q-11 siguen siendo deuda)

SOLO fila en ajustes_cargo.xlsx
   ✔ durable: 5_cobranza lo re-lee en cada corrida, nunca se pierde
   ✗ NO toca el archivo ya congelado — la planilla de agosto seguiría
     mostrando corte=40 y mes_anterior=46 hasta que corra 5_cobranza

LAS DOS  ← lo decidido
   ✔ el parche arregla el HOY (el archivo congelado)
   ✔ la fila de ajustes_cargo arregla el MAÑANA (cualquier regeneración)
```

**No hay doble conteo.** El parche vive en el *output* ya generado; el overlay
actúa *durante la generación*. Nunca se aplican los dos sobre el mismo cálculo: si
julio se re-corre, el archivo parchado se descarta entero y lo reemplaza el
resultado del overlay.

---

## 3. Las filas exactas

### 3a. Parche en el archivo congelado — en las DOS carpetas

```
C:\Users\wilde\PycharmProjects\jass_system\5_cobranza\outputs\arrastre_consolidado_2026-07.xlsx
C:\Users\wilde\PycharmProjects\jass_system - Julio\5_cobranza\outputs\arrastre_consolidado_2026-07.xlsx

   fila V-14 · columna CORTE   → 0     (valor actual esperado: 40)
   fila B-8  · columna DEUDA_AGUA → 0  (valor actual esperado: 46)
```

⚠ Ojo con `TOTAL_ARRASTRE`: el contrato del archivo exige
`suma(P1..P5) == TOTAL_ARRASTRE`. Al bajar un componente hay que bajar el total
en el mismo monto, o la validación del consolidado falla.

### 3b. Candado en `shared/ajustes_cargo.xlsx` — en las DOS carpetas

```
MZ=V   LT=14   CONCEPTO=CORTE_RECONEXION   MONTO=40   MES_ANO_APLICA=2026-07
MZ=B   LT=8    CONCEPTO=MES_ANTERIOR       MONTO=46   MES_ANO_APLICA=2026-07
```

- `MES_ANO_APLICA=2026-07` y no `2026-08`: el cargo nació y se anuló **en julio**.
  Si se pusiera 2026-08, la corrección se aplicaría en la corrida de agosto y
  además el parche del 3a la duplicaría visualmente.
- El mapa de conceptos válidos está en `5_cobranza/main.py:514`
  (`_CONCEPTO_DEVOLUCION_A_CAMPO`): acepta `CORTE_RECONEXION` y `MES_ANTERIOR`.
- `REF_AUDIT` de V-14 debe apuntar a su fila `APLICADO` en
  `6_corte/outputs/audit_penalidad.xlsx` — así el backfill del ledger puede contar
  la historia completa: nace el CARGO, nace el AJUSTE que lo anula.
- **Precedente ya resuelto así: F1-4**, en este mismo archivo. Copiar su formato.

### 3c. Registro del parche

```
shared/parches_manuales_pendientes_julio.xlsx
   una fila por cada uno, con estado PENDIENTE y el motivo
   (mismo formato que las filas de W-4 y R-5)
```

---

## 4. Verificación obligatoria ANTES de escribir la fila de B-8

`ajustes_cargo` hace `u[campo] -= monto` **sin piso**. Si el `MES_ANTERIOR` real de
B-8 en el ciclo de julio es menor a 46, restarle 46 lo deja **negativo** y se
reintroduce exactamente la clase de bug que estamos limpiando en el bloque 5.

```
abrir  shared/planilla_mes/planilla_2026-07.xlsx  →  fila B-8
   ¿cuánto vale su columna MES_ANTERIOR?

   MES_ANTERIOR >= 46   → la fila del 3b es correcta tal cual
   MES_ANTERIOR  < 46   → los 46 NO son todos arrastre; parte es consumo de
                           julio impago. Hay que partir el ajuste entre
                           CONCEPTO=MES_ANTERIOR y CONCEPTO=AGUA según la
                           composición real, o el saldo queda negativo.
```

V-14 no necesita esta verificación: su 40 es la penalidad completa de corte, y el
audit lo confirma.

---

## 5. Riesgos conocidos

| Riesgo | Mitigación |
|---|---|
| El parche del 3a se pierde en un `--force` de julio | La fila de `ajustes_cargo` lo reproduce sola. Es el candado. |
| Las dos carpetas divergen | Escribir siempre en `jass_system` y en `jass_system - Julio`. Hoy `arrastre_consolidado_2026-07` es idéntico en las dos (31/07 20:20:33). |
| `TOTAL_ARRASTRE` queda descuadrado | Bajar el total en el mismo monto que el componente (ver 3a). |
| B-8 queda con `MES_ANTERIOR` negativo | Verificación de la sección 4. |
| Alguien cree que hay doble conteo y borra una de las dos cosas | Este archivo. Las dos son necesarias y no se solapan. |

---

## 6. Cómo se cierra este RETOMAR

Cuando se cumplan **las tres**:

1. Las 2 filas escritas en `ajustes_cargo.xlsx` de las dos carpetas.
2. El parche aplicado en `arrastre_consolidado_2026-07.xlsx` de las dos carpetas,
   con `TOTAL_ARRASTRE` cuadrado.
3. `2_planilla` re-corrido y verificado: en `planilla_2026-08.xlsx`,
   V-14 con `CORTE_RECONEXION=0` y B-8 con `MES_ANTERIOR=0`, coincidiendo con la
   boleta impresa el 01/08.

Recién ahí la validación del paso 6 del plan puede dar **556/556** en vez de 554/556.
Borrar este archivo cuando `libro_mayor` absorba agua y corte — ahí el parche deja
de existir y pasa a ser un `registrar_ajuste` normal.

---

## 7. Lo que sigue abierto (no es parte de esta decisión)

- **② K-9 (S/100) y T-14 (S/75):** ¿el efectivo entró a la caja de la JASS?
  Sí → `abonos_rezagados.xlsx`. No → `registrar_ajuste`.
- **③ E-14B · G-14 · B-5:** confirmar que se asienta lo impreso, con la pregunta de
  origen abierta en `notas_2026-07.xlsx`.
- Las 33 correcciones del ledger de pueblo (bloques 1-5) siguen sin ejecutar.
