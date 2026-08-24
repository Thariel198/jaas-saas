# 6_corte

Módulo que ejecuta el ciclo de corte de servicio **por múltiples motivos** (agua y
multa/acuerdos): genera la lista de predios en mora elegibles para corte por cada motivo,
emite la penalidad de reconexión, gestiona la ventana de gracia de 2 días y clasifica el
resultado en salvados, cortados físicamente y arrastre al mes siguiente. En el destino
post-ledger **absorbe a `6b_corte_multas`**.

---

## ⚠ DISEÑO POST-LEDGER — corte unificado multi-motivo (absorbe 6b_corte_multas) · reconciliado con `dominio/politica_corte` (PC1-PC7) · 2026-07-17

`6_corte` **sobrevive reshaped** (a diferencia de `5_cobranza`/`5b`, que se disuelven) y
**absorbe a `6b_corte_multas`**: bajo el ledger, cortar por agua y cortar por multa son la
**misma máquina** — mismas 4 operaciones — que difieren solo en parámetros de política. No son
dos módulos; son dos **triggers** de un único motor de corte config-driven (ver
`docs/lente_escala.md`: política ≠ arquitectura).

### Las 4 operaciones (idénticas para todo motivo)

| Responsabilidad | Bajo el ledger |
|---|---|
| decisión "¿a quién corto?" | `riesgo_corte` (query de saldo POR CONCEPTO en `estado_cuenta`) + `dominio/politica_corte.evaluar()` invocado **una vez por trigger** |
| penalidad de reconexión | emite UN CARGO `corte_reconexion` (sin SUB_CONCEPTO, `SOURCE=6_corte`), monto del trigger — reemplaza el overlay-hack de agua y el `+40 a PENALIDAD_MULTA` de 6b |
| corte físico | `registro_cortes.xlsx` (shared) + nueva columna `MOTIVO` — un corte físico = una fila, gatíllelo agua o multa |
| seguimiento Día 2 | query de aplicaciones a `corte_reconexion` @ T2 + `evaluar()` de nuevo con `pago_ventana` real — el truco algebraico desaparece |
| phase gate BORRADOR→PUBLICADA→COMPROMETIDA | estado del workflow sobre la lista multi-motivo; inmutabilidad del ledger append-only (`lista = riesgo_corte@T1`) |

### Triggers de corte — manifiesto del tenant

**Firma cerrada** (`dominio/politica_corte`, PC1-PC7 — verdad única en
`docs/decisiones/ledger_fase1_decisiones.md` Bloque A):

```
evaluar(saldo, meses_impagos, ya_cortado, en_revision, pago_ventana, cfg)
   → (motivo, penalidad, salvado)
```

El trigger de multa es un caso **degenerado** del mismo `evaluar` (PC6): misma
función, otra `cfg`. La diferencia vive en los parámetros que `6_corte` pasa por
cada trigger, declarados en la config del tenant:

| Campo del trigger | `agua` | `multa` |
|---|---|---|
| `CONCEPTOS_SALDO` (insumo de `saldo()`/`meses_impagos()`, no de `evaluar`) | `[AGUA, MANTENIMIENTO]` | `[MULTA, ACUERDOS]` — CONVENIO excluido |
| `umbral_meses` (PC1) | `2` — **conductual**, no monto | `0` → cualquier multa impaga ya dispara (PC6) |
| `permite_salvarse` (PC1b · protege el conteo, no el corte) | `sí` — un mes con pago parcial NO cuenta como impago, frena el conteo | `no` — umbral 0 no da margen: 1 sola multa impaga ya elegible |
| `penalidad_base` (PC4) | `2000` céntimos → el MOTOR escala a `4000` si no paga en la ventana | `2000` céntimos → escala a `4000` (misma escalada que agua, PC6 — la multa no trae penalidad propia) |
| `EN_REVISION` protege | sí | sí |

**Umbral CONDUCTUAL, no de monto (PC1).** El `MES_ANTERIOR ≥ 8` viejo era un
proxy — la planilla independiente solo veía 1 mes, sin historia de pagos. El
ledger cuenta directo: **N meses seguidos en S/0** dispara elegibilidad (agua
N=2). **Precondición:** `dominio/saldo.meses_impagos()` deriva ese conteo del
ledger por FIFO de `MES_CARGO` — sin ese conteo confiable, el umbral conductual
no se sostiene (ver `dominio/README.md` §5).

**Salvado @ T2 = cubrir la penalidad, regla universal para todo trigger (PC2).**
No es "cualquier pago" (el gate Día-0 viejo, retirado) ni "saldar toda la deuda"
— es `pago_ventana ≥ penalidad`. Un parcial menor NO salva, sea agua o multa:

```
Día 0 · arma la lista (umbral por trigger, tabla arriba)
   evaluar(meses_impagos=2, pago_ventana=0,  cfg_agua)  → (motivo="agua", penalidad=20, salvado=False)

Día 2 · ¿pagó la penalidad dentro de la ventana de gracia?
   evaluar(meses_impagos=2, pago_ventana=25, cfg_agua)  → salvado=True   (pagó ≥ 20 → NO se corta)
   evaluar(meses_impagos=2, pago_ventana=5,  cfg_agua)  → salvado=False  (pagó < 20 → CORTADO, escala a 40)
```

Misma función, mismo `cfg_agua` — el MOTOR (que sabe fechas) la llama dos veces
con distinto `pago_ventana` (PC3: `evaluar` no sabe qué día es).

### `ya_cortado` — 3 estados, no 2 (PC5)

Un cortado se quedó sin agua; un exonerado sigue con agua pero se le **perdonó**
el corte por una razón — no son el mismo estado:

```
activo │ cortado │ EXONERADO
              ├─ mensual    → CADUCA al mes · motivo obligatorio (enfermedad/reclamo/verificación)
              └─ permanente → no caduca · lo decide la junta (ej. vejez)
```

`registro_cortes.xlsx` gana esta distinción explícita (`ESTADO` deja de ser un
string suelto `CORTADO`/`EXONERADO` — la exoneración guarda `{tipo, motivo,
periodo}` y caduca sola si es mensual).

### `registro_cortes.xlsx` (shared) — gana la columna `MOTIVO`

```
Grupo "¿Quién es el usuario?" : MZ · LT · NOMBRE
Grupo "Período del corte"     : MES_INICIO_CORTE · MES_REACTIVACION
Grupo "Estado actual"         : ESTADO · MOTIVO          ← NUEVA: agua | multa
Grupo "Trazabilidad"          : OBSERVACIONES · FECHA_REGISTRO · SOURCE
```
`SOURCE` deja de discriminar módulo (siempre `6_corte`); `MOTIVO` es el discriminador. El
corte FÍSICO es único por predio: si ya está `CORTADO` por cualquier motivo, no se re-corta —
la deuda del otro motivo se marca, no genera una 2da visita.

### Qué se absorbe de `6b_corte_multas`

Su lógica se vuelve el trigger `multa` (arriba). Lo que **desaparece** por el ledger:
- `DEUDA_MULTA = (MULTA+ACUERDOS) − max(0, pagado−cargo_agua)`: ese "excedente de agua se
  abona a multa" **es la cascada** P1→P3→P4 — lo deriva el motor. Query directa del saldo.
- El writer `+40 a PENALIDAD_MULTA` en la planilla: pasa a ser el CARGO `corte_reconexion`.

**Contrato ⑧:** el `CARGO corte_reconexion` lo emite `6_corte` (evento de corte, cualquier
motivo), **no** `2_planilla` (que queda solo con agua + mantenimiento).

**Estado:** diseño destino **reconciliado con `dominio/politica_corte`** (PC1-PC7,
`docs/decisiones/ledger_fase1_decisiones.md` Bloque A · spec consolidado en
`libro_mayor/dominio/README.md` §4), el schema de CARGO del contrato ⑧ y las
columnas reales de `registro_cortes`. Ledger **sin implementar**: hoy el código
pre-ledger corre single-motivo (agua, `MES_ANTERIOR≥8` por monto — el proxy que
PC1 reemplaza) y `6b_corte_multas` es un módulo separado que sigue corriendo; la
unificación es trabajo de Fase 2 (código). Ver
`docs/retomar/RETOMAR_dominio_saldo_unico_2026-07-13.md` §11 (histórico) y
`docs/arquitectura_pipeline_futuro.html`.

---

## Qué hace

1. **Genera la lista de corte** (`generar_lista.py`): filtra `planilla_cobrado_YYYY-MM.xlsx` (ciclo 1) por `SALDO > 0 AND MES_ANTERIOR ≥ 8` y produce `lista_corte.xlsx` con `PENALIDAD = S/20` y `TOTAL_A_PAGAR = SALDO + 20`.
2. **Aplica la penalidad** (`aplicar_penalidad.py`): suma `+20` a `CORTE_RECONEXION` en `shared/planilla_mes/planilla_YYYY-MM.xlsx` para cada usuario en lista_corte. Genera audit log para idempotencia; re-correr no duplica.
3. **Espera ventana de gracia** (48 h): el usuario puede pagar S/20 por Yape o efectivo y salvarse del corte físico.
4. **Clasifica el resultado** (`seguimiento.py`): cruza lista_corte con `planilla_cobrado_YYYY-MM.xlsx` ciclo 2 (post-ventana) y separa en tres grupos: pagaron penalidad, corte físico, arrastre.

## Cuándo se corre

| Momento | Script | Condición |
|---|---|---|
| Día 0 — al cierre del ciclo 1 de 5_cobranza | `generar_lista.py` | `planilla_cobrado.xlsx` ciclo 1 disponible |
| Día 0 — inmediatamente después | `aplicar_penalidad.py` | `lista_corte.xlsx` generado |
| Día 2 — después de re-correr 4_pagos + 5_cobranza | `seguimiento.py` | `planilla_cobrado.xlsx` ciclo 2 disponible |

## Estructura

```
6_corte/
├── generar_lista.py          # Filtra planilla_cobrado ciclo 1 → lista_corte.xlsx
├── aplicar_penalidad.py      # ★ Suma +20 a CORTE_RECONEXION en shared/planilla_mes
├── seguimiento.py            # Clasifica resultado post-ventana → 3 outputs
├── inputs/                   # Vacío — lee de 5_cobranza y shared/
├── outputs/
│   ├── lista_corte.xlsx               # Usuarios elegibles para corte (Día 0)
│   ├── audit_penalidad.xlsx           # Registro de penalidades aplicadas (idempotencia)
│   ├── pagaron_penalidad.xlsx         # Salvados: pagaron ≥ S/20 en ventana
│   ├── corte_fisico.xlsx              # Para el operario: cortar físicamente
│   └── arrastre_corte_YYYY-MM.xlsx   # Para 2_planilla del mes siguiente
├── backup/                   # Backups automáticos de planilla_mes antes de aplicar_penalidad
├── tests/
└── docs/
    ├── diagrama_flujo_6_corte.html    # Flujo rápido (cajas + flechas)
    ├── diagrama_6_corte.html          # Detallado: reglas, I/O, acoplamientos
    ├── formato_lista_corte.html
    ├── formato_pagaron_penalidad.html
    ├── formato_corte_fisico.html
    └── formato_arrastre_corte.html
```

## Dependencias externas

| Recurso | Tipo | Quién lo gobierna |
|---|---|---|
| `5_cobranza/outputs/planilla_cobrado_YYYY-MM.xlsx` | archivo (lectura) | `5_cobranza/` — ciclo 1 y ciclo 2 |
| `shared/planilla_mes/planilla_YYYY-MM.xlsx` | archivo (escritura) | `6_corte/aplicar_penalidad.py` — único writer de `CORTE_RECONEXION` |

**`aplicar_penalidad.py` es el único script del sistema que escribe sobre `shared/planilla_mes`.** Lo hace con backup automático, audit log (`audit_penalidad.xlsx`) e idempotencia — re-correr sobre la misma lista no suma el +20 dos veces.

## Reglas clave

- **Elegibilidad para corte:** `SALDO > 0` AND `MES_ANTERIOR ≥ 8`. Usuarios con menos de 8 meses de antigüedad no entran en lista de corte.
- **Penalidad inicial:** `S/20` sumada a `CORTE_RECONEXION` en la planilla del mes. `TOTAL_A_PAGAR = SALDO + 20`.
- **Ventana de gracia:** 48 horas desde la generación de lista_corte. Basta pagar S/20 para salvarse, aunque quede saldo mayor pendiente.
- **Clasificación post-ventana:** leer `CORTE_RECONEXION` en `planilla_cobrado.xlsx` ciclo 2:
  - `pagado ≥ 20` → **SALVADO** → `pagaron_penalidad.xlsx`
  - `pagado < 20` → **CORTADO** → `corte_fisico.xlsx` + `arrastre_corte.xlsx`
- **Escalada de penalidad para cortados:** la penalidad total sube a S/40 (S/20 penalidad + S/20 reconexión). `arrastre_corte = 40 − pagado`.
- **La deuda original no se toca aquí:** el saldo de consumo sigue en `arrastre_deuda.xlsx` de 5_cobranza. Este módulo solo gestiona el componente de corte/reconexión.
- **Idempotencia en todos los scripts:** re-correr con los mismos inputs produce el mismo output. `aplicar_penalidad.py` chequea `audit_penalidad.xlsx` antes de sumar.

## Flujo mensual

```
# DÍA 0 — después del ciclo 1 de 5_cobranza

python generar_lista.py
   ← 5_cobranza/outputs/planilla_cobrado.xlsx  (ciclo 1)
   → outputs/lista_corte.xlsx                  [usuarios elegibles + PENALIDAD=20]

python aplicar_penalidad.py
   ← outputs/lista_corte.xlsx
   ← shared/planilla_mes/planilla_YYYY-MM.xlsx
   → planilla actualizada (+20 en CORTE_RECONEXION)
   → outputs/audit_penalidad.xlsx
   → backup/planilla_YYYY-MM_<ts>.xlsx

# VENTANA DE GRACIA — 48 horas
# El usuario puede pagar S/20 para salvarse del corte físico.

# DÍA 2 — pasos manuales previos al seguimiento:
#   1. Descargar reporte banco actualizado
#   2. python 4_pagos/...     → pagos_yape_tepago.xlsx actualizado
#   3. python 5_cobranza/...  → planilla_cobrado.xlsx ciclo 2

python seguimiento.py
   ← outputs/lista_corte.xlsx
   ← 5_cobranza/outputs/planilla_cobrado.xlsx  (ciclo 2)
   → outputs/pagaron_penalidad.xlsx             [salvados]
   → outputs/corte_fisico.xlsx                 [para operario]
   → outputs/arrastre_corte_YYYY-MM.xlsx       [para 2_planilla mes siguiente]
```

## Lifecycle de outputs

| Archivo | Lifecycle |
|---|---|
| `lista_corte.xlsx` | Mensual — se regenera en Día 0; base del ciclo completo |
| `audit_penalidad.xlsx` | Mensual — crece por ciclo; garantiza idempotencia de aplicar_penalidad |
| `pagaron_penalidad.xlsx` | Mensual — output final del ciclo; referencia para cobranza siguiente |
| `corte_fisico.xlsx` | Mensual — entregado al operario; se archiva después del corte |
| `arrastre_corte_YYYY-MM.xlsx` | Mensual → insumo de `2_planilla` del mes siguiente |

## Lo que este módulo NO hace

- No calcula la deuda de consumo — eso lo hace `5_cobranza`.
- No modifica `planilla_cobrado.xlsx` — solo lo lee (ciclo 1 y ciclo 2).
- No procesa pagos — depende de que `4_pagos` + `5_cobranza` hayan corrido entre el Día 0 y el Día 2.
- No duplica el arrastre de deuda — `arrastre_deuda.xlsx` lo produce `5_cobranza`; este módulo solo arrastra el componente corte/reconexión.
- No escribe sobre `DATA_boletas.xlsx` ni sobre archivos de otros módulos, excepto `shared/planilla_mes` (con backup + audit).

## Señales de alerta

| Señal | Diagnóstico |
|---|---|
| `lista_corte.xlsx` tiene 0 filas | Revisar si `planilla_cobrado.xlsx` ciclo 1 fue generado correctamente; verificar filtros SALDO y MES_ANTERIOR |
| `aplicar_penalidad.py` reporta "ya aplicado" en todos | El script detectó audit existente — normal si se re-corre el mismo día; problema si es día diferente |
| `audit_penalidad.xlsx` tiene duplicados de (MZ, LT) | `aplicar_penalidad.py` falló a mitad y re-corrió sin limpiar — revisar lógica idempotente |
| `seguimiento.py` produce `corte_fisico.xlsx` con 0 filas cuando hay mora alta | Posible que `planilla_cobrado.xlsx` ciclo 2 no tenga la columna CORTE_RECONEXION actualizada — verificar que `aplicar_penalidad.py` corrió antes |
| `arrastre_corte_YYYY-MM.xlsx` tiene valores negativos en arrastre | `pagado > 40` — revisar regla: `arrastre_corte = max(0, 40 − pagado)` |
| `corte_fisico.xlsx` y `pagaron_penalidad.xlsx` suman menos que `lista_corte.xlsx` | Hay usuarios sin columna CORTE_RECONEXION en planilla_cobrado ciclo 2 — cruce incompleto |
