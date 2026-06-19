# 6b_corte_multas — Pendiente de implementación

**Fecha de diseño:** 16/06/2026  
**Estado:** Carpeta creada, scripts pendientes de codificar  
**Ruta del módulo:** `jass_system/6b_corte_multas/`

---

## Qué hace este módulo

Genera una lista de corte físico de agua para usuarios que tienen deuda pendiente en **MULTA** y/o **ACUERDOS_ASAMBLEA** de la planilla — columnas que representan la uniformización de deuda acordada en asamblea.

Es paralelo a `6_corte` (deuda mensual de agua) pero con criterios distintos y población distinta.

---

## Archivos de entrada

| Archivo | Ruta | Qué se lee |
|---------|------|-----------|
| `planilla_cobrado.xlsx` | `5_cobranza/outputs/` | MULTA (col 15), ACUERDOS_ASAMBLEA (col 16), SALDO (col 23), MZ, LT, NOMBRE |
| `mesa_1.xlsx` … `mesa_7.xlsx` | `4_pagos/efectivo/inputs/` | COMENTARIO (col 9) — tag "compromiso"/"compromizo", MZ, LT, MONTO |
| `lista_corte.xlsx` | `6_corte/outputs/` | MZ, LT, EJECUTAR_CORTE — para excluir los que ya están en corte de agua |

---

## Lógica de filtrado (en orden)

```
CANDIDATOS = usuarios donde (MULTA > 0 OR ACUERDOS_ASAMBLEA > 0) AND SALDO > 0
           → 252 usuarios en ciclo 2026-06

EXCLUIR si COMENTARIO contiene "compromis" (cubre "compromiso" y "compromizo") en cualquier mesa
           → 21 usuarios firmaron compromiso de pago al 05/07/2026 y pagaron parcial
           → OJO: también existe "compromiso + exoneracion" — igualmente excluido

EXCLUIR si aparece en lista_corte.xlsx con EJECUTAR_CORTE = SI
           → ~36 usuarios ya reciben corte por deuda de agua + S/20 de penalidad
           → No repetir: S/20 es único aunque la deuda tenga dos orígenes

RESULTADO ≈ 195 usuarios van a lista_corte_multas.xlsx
```

---

## Archivos de salida

| Archivo | Ruta | Descripción |
|---------|------|-------------|
| `lista_corte_multas.xlsx` | `6b_corte_multas/outputs/` | Lista de campo para supervisor |
| `audit_penalidad_multas.xlsx` | `6b_corte_multas/outputs/` | Audit trail separado del de agua |

---

## Columnas de lista_corte_multas.xlsx

Mismo patrón visual que `lista_corte.xlsx` de `6_corte`. Secciones sugeridas:

| Sección | Columnas |
|---------|---------|
| Identidad | MZ, LT, NOMBRE |
| Deuda uniformización | MULTA, ACUERDOS_ASAMBLEA, SALDO |
| Penalidad | PENALIDAD (S/20), TOTAL_A_PAGAR |
| ¿Ejecutar corte? | EJECUTAR_CORTE (SI/NO), MOTIVO_NO_EJECUTAR |

`EJECUTAR_CORTE = NO` cuando: compromiso firmado (motivo: "Compromiso pago 05/07/2026") o ya en corte agua (motivo: "Ya en corte agua").

---

## Scripts a crear

### `generar_lista_multas.py`

Pasos:
1. Leer `planilla_cobrado.xlsx` — extraer candidatos (MULTA|ACUERDO > 0 AND SALDO > 0)
2. Leer todos los `mesa_*.xlsx` — construir `set_compromisos` con (MZ, LT) donde COMENTARIO contiene "compromis"
3. Leer `lista_corte.xlsx` — construir `set_corte_agua` con (MZ, LT) donde EJECUTAR=SI
4. Filtrar: excluir compromisos y corte_agua
5. Exportar `lista_corte_multas.xlsx` con formato visual igual al de `6_corte`

### `aplicar_penalidad_multas.py`

**Mismo patrón de reconciliación bidireccional que `aplicar_penalidad.py` de `6_corte`:**

```
SET_DEBE  = (MZ, LT) con EJECUTAR=SI en lista_corte_multas.xlsx
SET_TIENE = net APLICADO - REVERTIDO en audit_penalidad_multas.xlsx para el ciclo

DEBE - TIENE  → aplica +S/20 en CORTE_RECONEXION de planilla, escribe fila APLICADO en audit
TIENE - DEBE  → revierte -S/20, escribe fila REVERTIDO en audit
DEBE ∩ TIENE  → skip (idempotente)
```

Columnas del audit: MZ, LT, NOMBRE, MES_ANO, PENALIDAD_APLICADA, CORTE_RECON_DESPUES, FECHA_APLICACION, SOURCE, ACCION  
SOURCE fijo: `"6b_corte_multas/aplicar_penalidad_multas"`

---

## Config a crear: `config.py`

```python
from pathlib import Path

ROOT = Path(__file__).parent.parent

PLANILLA_COBRADO_PATH = ROOT / "5_cobranza/outputs/planilla_cobrado.xlsx"
LISTA_CORTE_AGUA_PATH = ROOT / "6_corte/outputs/lista_corte.xlsx"
PAGOS_MESAS_DIR       = ROOT / "4_pagos/efectivo/inputs"
LISTA_CORTE_PATH      = ROOT / "6b_corte_multas/outputs/lista_corte_multas.xlsx"
PLANILLA_MES_DIR      = ROOT / "2_planilla/outputs"
AUDIT_PATH            = ROOT / "6b_corte_multas/outputs/audit_penalidad_multas.xlsx"
BACKUP_DIR            = ROOT / "6b_corte_multas/backup"
OUTPUTS_DIR           = ROOT / "6b_corte_multas/outputs"

PENALIDAD        = 20.0
TOL              = 0.005
TAG_COMPROMISO   = "compromis"   # cubre "compromiso" y "compromizo"
```

---

## Datos reales del ciclo 2026-06

| Métrica | Valor |
|---------|-------|
| Candidatos (MULTA/ACUERDO y SALDO>0) | 252 |
| Con solo MULTA | 163 |
| Con solo ACUERDOS | 85 |
| Con ambas | 165 |
| Excluidos por compromiso | 21 |
| Excluidos por ya estar en corte agua (EJECUTAR=SI) | ~36 |
| Resultado esperado en lista | ~195 |
| Tags de compromiso en mesas | "compromiso" (14x) + "compromizo" (8x) + "compromiso + exoneracion" (1x) |

---

## Notas de diseño

- **S/20 único:** si un usuario ya tiene S/20 aplicado por `6_corte`, este módulo NO le aplica otro. Los 36 que están en ambas listas ya recibieron su penalidad.
- **Audit separado:** `audit_penalidad_multas.xlsx` es independiente de `audit_penalidad.xlsx` de agua — trazabilidad diferenciada por origen.
- **Trigger distinto:** `6_corte` corre mensual. `6b_corte_multas` corre cuando la asamblea lo decide (puede ser una vez o varias).
- **Leer HTMLs antes de codificar:** crear `docs/formato_lista_corte_multas.html` y `docs/formato_audit_penalidad_multas.html` antes de escribir código Excel.
- **Copiar patrón de `6_corte`:** `generar_lista.py` y `aplicar_penalidad.py` son la referencia directa. Mismo estilo, mismas funciones auxiliares (`_norm_mz`, `_norm_lt`, `_c`, `_gh`, etc.).

---

## Orden de trabajo para la sesión

1. Cambiar a **Opus** para diseño y código
2. Crear `docs/formato_lista_corte_multas.html` (contrato visual)
3. Crear `docs/formato_audit_penalidad_multas.html`
4. Crear `config.py`
5. Codificar `generar_lista_multas.py`
6. Correr y validar contra los ~195 esperados
7. Codificar `aplicar_penalidad_multas.py`
8. Correr y validar audit + planilla
