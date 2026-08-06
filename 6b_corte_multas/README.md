# 6b_corte_multas

## ⚠ DISEÑO POST-LEDGER — este módulo SE DISUELVE en `6_corte` (Fase 1 cerrada · 2026-07-16)

Bajo el ledger, cortar por multa **no es un módulo aparte**: es el **trigger `multa`** del motor
de corte unificado en `6_corte`. Hace las mismas 4 operaciones que el corte de agua (riesgo =
query de saldo · `dominio/politica_corte.evaluar()` · emite CARGO `corte_reconexion` ·
`registro_cortes[MOTIVO]` · seguimiento @ T2), difiriendo solo en 4 parámetros de la config del
tenant. Por el lente de escala (política ≠ arquitectura, ver `docs/lente_escala.md`) un módulo
espejo no se justifica.

**El spec destino manda vive en `6_corte/README.md`** (tabla de triggers). Este archivo describe
el código **pre-ledger que aún corre** como módulo separado.

Lo que 6b aporta al motor como trigger `multa`:
```
CONCEPTOS_SALDO       [MULTA, ACUERDOS]   (CONVENIO excluido)
UMBRAL_CENTIMOS       0                   → elegibilidad = saldo>0
PROTEGE_PAGO_PARCIAL  false               (salvado exige saldar todo)
PENALIDAD_CENTIMOS    {inicial:2000, escala:4000}   (20→40; el S/40 directo fue solo junio 2026)
SALVADO_CUANDO        saldo_conceptos_cero          (saldar TODA la deuda, no solo la penalidad)
EN_REVISION protege   sí
```
El `DEUDA_MULTA = (MULTA+ACUERDOS) − max(0, pagado−cargo_agua)` de abajo **desaparece**: ese
"excedente de agua se abona a multa" es la cascada P1→P3→P4 — lo deriva el motor. El writer
`+40 a PENALIDAD_MULTA` pasa a ser el CARGO `corte_reconexion`.

**Estado:** módulo separado en el código pre-ledger (sigue corriendo single-motivo). La fusión
en `6_corte` es trabajo de Fase 2 (código). Ver `docs/RETOMAR_dominio_saldo_unico_2026-07-13.md`
§11 y `docs/arquitectura_pipeline_futuro.html`.

---

## [PRE-LEDGER] Módulo espejo de `6_corte` — código actual

Módulo espejo de `6_corte` que ejecuta el ciclo de penalidad por deuda de multas y acuerdos de asamblea: genera la lista de usuarios con deuda pendiente en multas, aplica la penalidad de S/40, gestiona la ventana de gracia de 2 días y clasifica el resultado final en salvados, cortados físicamente y arrastre al mes siguiente.

> **Hermano de 6_corte.** Misma arquitectura de 3 scripts + phase gate. La diferencia está en la condición de elegibilidad (multa real, no saldo de consumo), la columna que se penaliza (`PENALIDAD_MULTA`) y la condición de salvado (saldar toda la deuda, no solo S/20).

## Qué hace

1. **Genera la lista de multas** (`generar_lista_multas.py`): filtra `planilla_cobrado.xlsx` por `DEUDA_MULTA > 0`, donde `DEUDA_MULTA = (MULTA + ACUERDOS_ASAMBLEA) − max(0, pagado − cargo_agua)`. CONVENIO excluido. Produce `lista_multas.xlsx`.
2. **Aplica la penalidad** (`aplicar_penalidad_multas.py`): suma `+40` a `PENALIDAD_MULTA` en `shared/planilla_mes/planilla_YYYY-MM.xlsx`. Genera audit log para idempotencia; re-correr no duplica.
3. **Espera ventana de gracia** (2 días): el usuario puede saldar toda su deuda (agua + multa + acuerdos) para salvarse del corte físico.
4. **Clasifica el resultado** (`seguimiento_multas.py`): cruza lista_multas con `planilla_cobrado.xlsx` ciclo 2 y separa en: salvados, corte físico, arrastre.

## Diferencias clave vs 6_corte

| Aspecto | 6_corte | 6b_corte_multas |
|---|---|---|
| Elegibilidad | `SALDO > 0 AND MES_ANT ≥ 8` | `(MULTA + ACUERDOS) − max(0, pagado − agua) > 0` |
| Columna penalizada | `CORTE_RECONEXION` | `PENALIDAD_MULTA` |
| Monto penalidad (normal) | S/20 → escala a S/40 | S/20 → escala a S/40 con corte físico |
| Monto penalidad (junio 2026) | S/40 directo | S/40 directo |
| Condición SALVADO | `pagado_corte ≥ 20` | `pagado_total ≥ cargo_agua + MULTA + ACUERDOS` |
| CONVENIO | aplica | **excluido** — se paga por cuotas separadas |
| registro_cortes | `6_corte/inputs/` | `shared/` — compartido con 6_corte |

## Cuándo se corre

| Momento | Script | Condición |
|---|---|---|
| Día 0 — después de 5_cobranza ciclo 1 | `generar_lista_multas.py` | `planilla_cobrado.xlsx` ciclo 1 disponible |
| Día 0 — inmediatamente después | `aplicar_penalidad_multas.py` | `lista_multas.xlsx` generado |
| Día 2 — después de re-correr 4_pagos + 5_cobranza | `seguimiento_multas.py` | `planilla_cobrado.xlsx` ciclo 2 disponible |

## Estructura

```
6b_corte_multas/
├── generar_lista_multas.py       # Filtra planilla_cobrado ciclo 1 → lista_multas.xlsx
├── aplicar_penalidad_multas.py   # ★ Suma +40 a PENALIDAD_MULTA en shared/planilla_mes
├── seguimiento_multas.py         # Clasifica resultado post-ventana → 4 outputs
├── config.py                     # Paths, reglas de negocio, tolerancias
├── inputs/                       # Vacío — lee de 5_cobranza y shared/
├── outputs/
│   ├── lista_multas.xlsx                  # Usuarios elegibles (Día 0)
│   ├── audit_penalidad_multas.xlsx        # Registro de penalidades aplicadas (idempotencia)
│   ├── pagaron_penalidad_multas.xlsx      # Salvados: saldaron toda su deuda
│   ├── corte_fisico_multas.xlsx           # Para el operario: cortar físicamente
│   └── arrastre_multa_YYYY-MM.xlsx       # Para 2_planilla del mes siguiente
├── backup/
├── tests/
└── docs/
    ├── diagrama_flujo_6b_corte_multas.html   # Flujo rápido (cajas + flechas)
    ├── diagrama_6b_corte_multas.html          # Detallado: reglas, I/O, acoplamientos
    ├── arquitectura_6b_corte_multas.html
    ├── formato_lista_multas.html
    ├── formato_pagaron_penalidad_multas.html
    ├── formato_corte_fisico_multas.html
    └── formato_arrastre_multa.html
```

## Dependencias externas

| Recurso | Tipo | Quién lo gobierna |
|---|---|---|
| `5_cobranza/outputs/planilla_cobrado.xlsx` | archivo (lectura) | `5_cobranza/` — ciclo 1 y ciclo 2 |
| `shared/planilla_mes/planilla_YYYY-MM.xlsx` | archivo (escritura) | `aplicar_penalidad_multas.py` — único writer de `PENALIDAD_MULTA` |
| `shared/registro_cortes.xlsx` | archivo (lectura + append) | compartido con `6_corte` — ambos módulos registran cortes aquí |
| `4_pagos/yape/.../pagos_yape_tepago.xlsx` | archivo (lectura) | `seguimiento_multas.py` — trazabilidad del pago |
| `4_pagos/efectivo/outputs/pagos_efectivo.xlsx` | archivo (lectura) | `seguimiento_multas.py` — trazabilidad del pago |

**`aplicar_penalidad_multas.py` es el único script del sistema que escribe sobre `PENALIDAD_MULTA` en `shared/planilla_mes`.** Lo hace con backup automático, audit log e idempotencia.

## Reglas clave

- **Elegibilidad:** `DEUDA_MULTA = (MULTA + ACUERDOS_ASAMBLEA) − max(0, pagado − cargo_agua) > 0`. El excedente de pago de agua se abona primero a multas; si aun así queda saldo, entra en lista.
- **CONVENIO excluido:** no se cuenta como deuda de multa. El convenio se gestiona por separado.
- **Penalidad:** S/40 sumada a `PENALIDAD_MULTA`. _(Junio 2026: S/40 directo desde Día 0. Normal: S/20 Día 0, escala a S/40 tras corte físico en Día 2.)_
- **Ventana de gracia:** 2 días desde la generación de lista_multas.
- **Condición SALVADO (estricta):** `pagado_total ≥ cargo_agua + MULTA + ACUERDOS`. No basta pagar solo la penalidad — debe saldar toda la deuda.
- **Escalada de penalidad para cortados:** arrastre = S/40 menos lo que pagó de PENALIDAD_MULTA.
- **registro_cortes en shared/:** el estado de quién está cortado (agua o multa) vive en un solo lugar. `generar_lista_multas.py` lo lee para no listar a usuarios ya cortados en multas.
- **Idempotencia:** `aplicar_penalidad_multas.py` chequea `audit_penalidad_multas.xlsx` antes de sumar. Re-correr no duplica.

## Flujo mensual

```
# DÍA 0 — después del ciclo 1 de 5_cobranza

python generar_lista_multas.py
   ← 5_cobranza/outputs/planilla_cobrado.xlsx   (ciclo 1)
   ← shared/registro_cortes.xlsx                (excluye ya-cortados)
   ← 4b_reclamos/outputs/reclamos_YYYY-MM.xlsx
   → outputs/lista_multas.xlsx                  [elegibles + DEUDA_MULTA + PENALIDAD=40]

python aplicar_penalidad_multas.py
   ← outputs/lista_multas.xlsx
   ← shared/planilla_mes/planilla_YYYY-MM.xlsx
   → planilla actualizada (+40 en PENALIDAD_MULTA)
   → outputs/audit_penalidad_multas.xlsx
   → backup/planilla_YYYY-MM_<ts>.xlsx

# VENTANA DE GRACIA — 2 días
# El usuario puede saldar toda su deuda (agua + multa + acuerdos) para salvarse.

# DÍA 2 — pasos manuales previos al seguimiento:
#   1. Descargar reporte banco actualizado
#   2. python 4_pagos/...    → pagos_yape_tepago.xlsx actualizado
#   3. python 5_cobranza/... → planilla_cobrado.xlsx ciclo 2

python seguimiento_multas.py
   ← outputs/lista_multas.xlsx
   ← 5_cobranza/outputs/planilla_cobrado.xlsx   (ciclo 2)
   ← 4_pagos/.../pagos_yape_tepago.xlsx         (trazabilidad)
   ← 4_pagos/efectivo/.../pagos_efectivo.xlsx   (trazabilidad)
   → outputs/pagaron_penalidad_multas.xlsx       [salvados + trazabilidad]
   → outputs/corte_fisico_multas.xlsx            [para operario]
   → outputs/arrastre_multa_YYYY-MM.xlsx        [para 2_planilla mes siguiente]
   → shared/registro_cortes.xlsx                [append nuevos CORTADO_MULTA]
```

## Lifecycle de outputs

| Archivo | Lifecycle |
|---|---|
| `lista_multas.xlsx` | Mensual — se regenera en Día 0; base del ciclo completo |
| `audit_penalidad_multas.xlsx` | Mensual — crece por ciclo; garantiza idempotencia de aplicar_penalidad_multas |
| `pagaron_penalidad_multas.xlsx` | Mensual — output final del ciclo; referencia para cobranza siguiente |
| `corte_fisico_multas.xlsx` | Mensual — entregado al operario; se archiva después del corte |
| `arrastre_multa_YYYY-MM.xlsx` | Mensual → insumo de `2_planilla` del mes siguiente |

## Lo que este módulo NO hace

- No calcula la deuda de consumo — eso lo hace `5_cobranza`.
- No modifica `planilla_cobrado.xlsx` — solo lo lee (ciclo 1 y ciclo 2).
- No gestiona el CONVENIO — se paga por separado en cuotas acordadas.
- No duplica el arrastre de deuda de agua — `arrastre_deuda.xlsx` lo produce `5_cobranza`.
- No escribe sobre archivos de otros módulos, excepto `shared/planilla_mes` (con backup + audit) y `shared/registro_cortes.xlsx` (solo append).

## Señales de alerta

| Señal | Diagnóstico |
|---|---|
| `lista_multas.xlsx` tiene 0 filas | Verificar si `MULTA` y `ACUERDOS_ASAMBLEA` están poblados en `planilla_cobrado.xlsx`; verificar filtro DEUDA_MULTA |
| `aplicar_penalidad_multas.py` reporta "ya aplicado" en todos | Audit existente — normal si se re-corre el mismo día; problema si es día diferente |
| `seguimiento_multas.py` clasifica todos como SALVADO sin haberlos cobrado | Verificar que `planilla_cobrado.xlsx` ciclo 2 tiene `PENALIDAD_MULTA` actualizado; verificar que `aplicar_penalidad_multas.py` corrió antes |
| `arrastre_multa_YYYY-MM.xlsx` tiene valores negativos | `pagado_penalidad > 40` — revisar regla: `arrastre = max(0, 40 − pagado_penalidad_multa)` |
| Usuario aparece en lista_multas y en lista_corte el mismo mes | Normal — son deudas distintas. Pero verificar que no reciba doble corte físico (registro_cortes en shared/ previene esto) |
| `corte_fisico_multas.xlsx` tiene usuarios que ya están CORTADO en agua | `generar_lista_multas.py` debería haberlos excluido via `shared/registro_cortes.xlsx` — revisar join |
