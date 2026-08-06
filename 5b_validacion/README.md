# 5b_validacion — Cuadre del dinero contra el reporte bancario

Valida que **la plata que el sistema registró cuadre con la plata real** del reporte
bancario crudo (Yape) del mes. Es el control de tesorería: si el sistema dice que entró
S/X pero el banco dice S/Y, este módulo lo detecta antes del cierre.

> Módulo dependiente de `5_cobranza` (el `b`). Detalle de diseño y contratos visuales
> en `docs/validacion_diseno.html` + `docs/manual_uso.md`.

---

## ⚠ DISEÑO POST-LEDGER — este módulo SE DISUELVE (Fase 1 cerrada · 2026-07-14)

Bajo `libro_mayor`, `5b_validacion` se reemplaza por **dos tools**:

| Hoy (5b, ~500 líneas · 8 archivos) | Nuevo |
|---|---|
| sumar lo registrado por balde | `arqueo_caja()` — query pura sobre caja (`Σ MONTO GROUP BY BALDE`) |
| comparar registrado vs crudo bancario | `conciliar_caja(crudo)` — arqueo vs banco → cuadra / discrepancias |

`conciliar_caja` **no es un gate en un solo lugar** — es una tool invocada en **N momentos**
del ciclo (la capacidad "estado a-una-fecha"): en el corte (T1) y en el cierre (T2), como
**queries point-in-time** sobre el ledger append-only. El "5b corre 2 veces" deja de ser
dos corridas de un batch → pasa a ser dos llamadas a las mismas tools sobre el ledger que
fue creciendo. El `estado_ciclo.validado` deja de gatear a `5_cobranza` (disuelto) → gatea
el **asiento** del mes.

**Estado:** diseño cerrado, ledger **sin implementar**. El código descrito abajo sigue
corriendo hasta entonces. Ver `docs/RETOMAR_dominio_saldo_unico_2026-07-13.md` §11.

---

## Implementación actual (pre-ledger, transitoria)

---

## Qué hace

Compara dos mundos por **dirección de dinero** y por **balde**, con tolerancia de
redondeo `S/0.005`:

```
TE PAGÓ (entró plata)          vs   lo que el sistema registró como ingreso
PAGASTE (salió plata)          vs   lo que el sistema registró como egreso
```

Cada dirección se mueve en su **propia ventana de fechas** (TE PAGÓ y PAGASTE no van en
el mismo rango), y el cuadre se hace sumando los baldes que componen cada lado.

## Cuándo se corre

Después de `5_cobranza` (necesita `planilla_cobrado.xlsx`) y con el reporte crudo del
mes ya en `shared/reporte_mes_crudo/`. Su resultado (`estado_ciclo.json` con
`validado:true`) es lo que habilita a `5_cobranza` a exportar el `arrastre_consolidado`.

---

## Qué lee

```
shared/reporte_mes_crudo/                        reporte bancario crudo (TE PAGÓ / PAGASTE)
5_cobranza/outputs/planilla_cobrado.xlsx         agua registrada por el sistema
4_pagos/yape/motor_matching/outputs/
   pagos_yape_tepago.xlsx · _retorno · _devolucion · _pagaste · blancos_mes.xlsx
4_pagos/efectivo/outputs/pagos_efectivo.xlsx
4_pagos/outputs/aportes_tanque.xlsx              aporte al tanque (canal-agnóstico)
shared/reporte_acumulado_procesado/estado_ciclo.json
```

## Cómo cuadra (los baldes)

```
NIVEL 1a — TE PAGÓ (Yape):   agua + blancos + tanque + otros_conceptos  = crudo TE PAGÓ
              otros_conceptos = Σ TE PAGÓ Yape con CONCEPTO ≠ vacío/tanque
                                (captura deuda_directiva, etc.)
NIVEL 1b — PAGASTE:          devolución + retorno + gasto              = crudo PAGASTE
NIVEL 2  — neteado:          aplica devoluciones sobre agua para la vista consolidada
```

Genera la hoja de diferencias (`docs/formato_validacion_diferencias.html` es su contrato):
cada fila muestra reporte vs sistema y por qué cuadra o no.

---

## Reglas de negocio

- **Tolerancia `S/0.005`** — diferencias por debajo se consideran cuadradas (redondeo).
- **TE PAGÓ y PAGASTE son ventanas independientes** — no se asume que entraron y
  salieron en el mismo rango de fechas.
- **El pago original y su devolución son eventos separados** — en Nivel 1a el agua queda
  cruda (el pago fue legítimo en TE PAGÓ); la devolución es un evento PAGASTE aparte que
  se reconcilia en Nivel 1b. Netear temprano ocultaría el descuadre.
- **`aportes_tanque.xlsx` es canal-agnóstico** — reemplaza al viejo `pagos_yape_tanque.xlsx`;
  alimenta el balde tanque de Nivel 1a.

---

## Pendiente Fase 2 — el ledger vuelve esto una query

Bajo `libro_mayor/caja`, cada movimiento ya lleva `DIRECCION` + `BALDE` explícitos
(decisión ⑩). La validación deja de ser 8 lecturas con filtros dispersos y pasa a ser
una **query por balde** sobre una sola fuente:

```
ENTRÓ = Σ MONTO donde DIRECCION=INGRESO      SALIÓ = Σ MONTO donde DIRECCION=EGRESO
NETO  = ENTRÓ − SALIÓ            desglose:   GROUP BY BALDE
```

Reproduce exactamente los niveles 1a/1b/2 desde el ledger, y desaparecen los falsos
descuadres por plata sin balde. En el backlog del ledger esta capacidad es
`arqueo_caja` (reemplaza a este módulo). Ver `libro_mayor/caja/README.md` (⑩).

> Correcciones de dominio relacionadas (2026-07-13): "tanque **comunitario**" es un
> adjetivo (el tanque es de la comunidad), **no** el `CONCEPTO=comunitario` de
> `motor_matching` (eso es segregación). `deuda_directiva` es balde caja-only, se
> reconcilia acá como "otros conceptos" (Nivel 1a) y nunca cruza a deuda de predio.

---

## Lo que NO hace

- **No registra ni identifica pagos** — solo valida los ya registrados por `4_pagos`/`5_cobranza`.
- **No corrige** — reporta el descuadre; la corrección la hace quien corresponda
  (reclamo, re-identificación de pago, etc.).
- **No toca la deuda del predio** — es control de caja, no de cuenta corriente.

## Errores comunes

- Correr antes de que el reporte crudo del mes esté completo en `shared/reporte_mes_crudo/`
  → descuadre falso por pagos aún no en el crudo.
- Confundir la ventana de TE PAGÓ con la de PAGASTE.
- Esperar que el `arrastre_consolidado` de `5_cobranza` salga sin haber corrido esta
  validación (necesita `estado_ciclo.json` con `validado:true`).
