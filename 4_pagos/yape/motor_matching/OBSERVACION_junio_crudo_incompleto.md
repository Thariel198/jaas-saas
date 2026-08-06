# ⚠ OBSERVACIÓN — Junio cerrado con crudo incompleto (pagos 16-21/06 sin procesar)

Fecha: 2026-07-08
Estado: EN OBSERVACIÓN — pendiente de corregir al retomar. Boletas de julio YA emitidas (no se re-emiten; el saldo se corrige en el reproceso de junio).

---

## Qué pasó

Junio se **cerró el 2026-07-03** con un reporte crudo del banco que solo llegaba al **15/06**
(`2026-06_banco.xlsx` archivado: MAX 15/06/2026 21:13). Después se cargó un crudo nuevo que llega
al 21/06, pero el ciclo ya estaba cerrado → esos movimientos no entraron.

Resultado: **4 pagos TE PAGÓ (S/421) fechados 16-21/06 no se procesaron** y por lo tanto no
entraron al arrastre que se generó para julio (2026-07-06).

| Fecha | Monto | Referencia | Lote / estado |
|---|---|---|---|
| 17/06 20:32 | S/41  | PLIN - Elias Agapito  | sin mensaje → pendiente de match |
| 18/06 21:01 | S/300 | PLIN - Anali Quineche | sin mensaje → pendiente de match |
| 21/06 12:03 | S/44  | "mz J lote 1 comedor" (Janet Eva*) | J-1 COMEDOR POPULAR CLUB DE MADRES |
| 21/06 12:08 | S/36  | "MZ C lote 43" (Janet Eva*)         | C-43 JANET EVARISTO COCHACHIN |

**CONFIRMADO en las boletas emitidas** (`3_boletas/inputs/DATA_boletas.xlsx`), no es inferencia:

- **C-43 Janet Evaristo:** boleta con `MES ANTERIOR = 16`, `IMPORTE A PAGAR = 24`. Pagó S/36 el
  21/06 sin acreditar → con el pago el arrastre de 16 se cancela (36 > 16) y hasta quedaría saldo a
  favor. Se le cobró 24 de más.
- **J-1 Comedor Club de Madres:** boleta con `MES ANTERIOR = 65`, `IMPORTE A PAGAR = 245`. Pagó S/44
  el 21/06 sin acreditar → el arrastre de 65 debió bajar a 21.
- El campo `MES ANTERIOR` de la boleta = la deuda arrastrada; en ambos sigue entero, sin descontar
  el pago del 21/06.
- Los 2 PLIN (Elias Agapito S/41, Anali Quineche S/300) quedaron sin acreditar a ningún lote.

Las boletas de julio ya se emitieron con estos montos. No se re-emiten; el saldo se corrige al
reprocesar junio y se propaga al próximo ciclo.

## Causa raíz

**Crudo incompleto al momento de cerrar junio** — NO el bug del ancla. El ancla usada por junio
(09/05, fin de mayo) era correcta; el ancla fija el piso del corte, no el techo. El techo lo puso
un crudo que aún no tenía los movimientos del 16 al 21/06.

## Bug SEPARADO del ancla — FIX APLICADO (2026-07-08)

`obtener_ancla()` en `main.py` usaba `sorted(glob("*_procesado.xlsx"))[-1]` (alfabético). Cuando
`2026-06_procesado.xlsx` coexiste con los legacy `reporte_..._procesado.xlsx`, el `2026-06` queda
PRIMERO alfabéticamente ('2' < 'r') y `[-1]` agarraba el legacy viejo (max 09/05) en vez del real
(15/06). En julio esto habría reprocesado ~74 filas de junio como "nuevas".

**Fix aplicado:** en vez de `sorted()[-1]`, se filtran los archivos del formato nuevo
`AAAA-MM_procesado.xlsx` (regex) y se elige el ciclo `(AAAA, MM)` más alto; los legacy quedan fuera
de la selección (siempre son más viejos). Si no hay ningún `AAAA-MM` (era manual pura), cae al
último legacy. **Validado:** `obtener_ancla()` ahora devuelve `2026-06-15 21:13:40`.

## Plan DECIDIDO — NO reprocesar junio

Reprocesar junio (re-abrir ciclo cerrado + regenerar arrastre + re-emitir boletas) es churn y
riesgo. El pago es un evento durable en el reporte; no se pierde. La corrección la hace julio solo:

1. **Julio:** procesar con el reporte del banco. El **ancla ya queda en 15/06** (fix de
   `obtener_ancla` aplicado y validado el 2026-07-08) → julio lee del 16/06 en adelante y acredita
   los 4 pagos, sin recontar mayo-15/junio.
2. Los 2 con lote (C-43, J-1) se acreditan solos; los 2 PLIN (sin mensaje) caen en pendientes →
   matchear a mano en julio.
3. **Comunicar** a los afectados que su pago está registrado y se refleja en julio; que NO vuelvan a
   pagar la boleta (si no, sobre-pagan). La boleta de papel no se corrige; la cuenta en el sistema sí.

Junio queda intacto y cerrado. No hace falta tocar `2026-06_procesado` ni el arrastre.
