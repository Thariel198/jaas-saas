# Decisión de diseño — arqueo (sub-módulo de 4_pagos/efectivo)

Fecha: 2026-07-07
Estado: Aprobado en conversación · Fase 2.0.6

---

**Problema:**
Cada día el cobrador anota en `mesa_N.xlsx` lo que recaudó (efectivo y yape) y le entrega físicamente a la tesorera el efectivo y le reenvía el yape. Nadie valida que lo del papel = lo que la tesorera recibió. Hay que cuadrar, por día y por cobrador, ambos flujos, y saber a quién reclamar cuando falta plata.

**Criterios:**
- Auditable por una persona no técnica, fila por fila (una fila = un cobrador en un día).
- El descuadre dice el monto y a quién reclamar — desglose por cobrador, no total agregado.
- Regenerable — el arqueo se recalcula cada corrida; ninguna columna humana en el output.
- Escala a agentic SaaS — el input nuevo es un log append-only (= un POST de la tesorera desde su teléfono).

**Enfoque elegido:**
Sub-módulo de `4_pagos/efectivo` (el output es control/auditoría, no lo consume ningún módulo downstream).
- **Ledger `inputs/entregas.xlsx`** — log append-only event-sourced, escritor único `entregas_repo.py`. Nunca se edita a mano. El arqueo pliega sus eventos (`RECIBIDO = Σ filas` por `(FECHA, COBRADOR)`); una corrección es un evento delta. Misma familia que `seguimiento_pueblo` (afecta un cálculo — el cuadre — no es trazabilidad plana), con menor alcance (alimenta solo la vista arqueo, no otros módulos).
- **Captura — flujo B (Excel físico + import):** la tesorera llena `inputs/entregas_hoja.xlsx` a mano (staging, como una mesa); `importar_entregas.py` vuelca cada fila al ledger vía `registrar_entrega()`, idempotente por `(source, audit_ref)`. Se descartó el flujo C (Excel a mano directo sobre el ledger) porque haría editable el ledger y se perdería la propiedad event-sourced. En SaaS, un endpoint POST reemplaza al import y llama a la misma función — cambia quién invoca, no el ledger. `registrar_entrega.py` queda como CLI alterno para entradas sueltas.
- **Script nuevo `arqueo.py`** — tool self-documenting. Lee las mesas crudas (no `pagos_efectivo.xlsx`, que dedupe cross-cobrador y distorsionaría el total por cobrador), agrupa `MONTO_EFECTIVO` y `MONTO_YAPE` por `(FECHA, COBRADOR)`, compara contra `entregas.xlsx`.
- **Output nuevo `outputs/arqueo_YYYY-MM.xlsx`** — vista regenerable: `FECHA · COBRADOR · EFECTIVO_PAPEL · EFECTIVO_RECIBIDO · DIF_EFECTIVO · YAPE_PAPEL · YAPE_RECIBIDO · DIF_YAPE · ESTADO`.
- `ESTADO`: `CUADRA` · `DESCUADRE` (con el ± en las columnas DIF) · `SIN_ENTREGA` (hay papel, la tesorera no declaró) · `SIN_MESA` (declaró, no hay papel).

**Decisiones de negocio (cerradas en conversación):**
1. Agrupa por `FECHA` (día de cobro), no `FECHA_REGISTRO`. Razón de negocio (confirmada
   con datos reales jul-2026): una hoja de cobro se puede registrar en varios días por
   falta de tiempo → `FECHA_REGISTRO` NO es única para un cobro. `FECHA` (día que se cobró)
   sí es única. Costo: los cobradores deben llenar `FECHA` en la mesa; si queda vacía, la
   fila no entra al cuadre y `arqueo.py` emite un warning con el conteo (no se subcuenta en
   silencio).
2. Efectivo esperado = Σ `MONTO_EFECTIVO` **bruto**, sin netear. `CONCEPTO=gasto` no se modela: la plata llega primero a la tesorera, nadie gasta antes de entregar. No ha pasado nunca en efectivo.
3. Yape fluye usuario → cuenta del cobrador → cuenta de la tesorera. Se cuadra **por cobrador** (`DIF_YAPE≠0` = a quién reclamar), no a nivel día-total.

**Alternativas descartadas:**
- *Parsear el estado de cuenta del banco para el yape* — el comentario del banco es texto libre del usuario → ambigüedad alta. La tesorera declara el monto que le llegó por cada cobrador.
- *Columnas de confirmación en `mesa_N.xlsx`* — mezcla dos actores (cobrador + tesorera) en un archivo; la tesorera recibe por cobrador-día, no por mesa.
- *Hoja mutable que la tesorera edita in-place* — pierde el rastro de quién declaró qué y cuándo; el log append-only es la forma que migra 1:1 al SaaS.
- *Leer `pagos_efectivo.xlsx` en vez de las mesas crudas* — el dedupe cross-cobrador de `main.py` mezcla registros y distorsiona el total físico por cobrador.

**Señal de alerta:**
El arqueo marca `SIN_ENTREGA` de forma sistemática para un cobrador → la tesorera no está declarando (proceso roto en la fuente, no en el código). Si aparecen muchos `SIN_MESA` → los cobradores no están registrando en la mesa lo que entregan. Si `DIF` nunca es 0 ni en casos que deberían cuadrar → revisar que la clave de agrupación sea `(FECHA, COBRADOR)` con FECHA normalizada (mismo formato de fecha en mesa y en entregas).
