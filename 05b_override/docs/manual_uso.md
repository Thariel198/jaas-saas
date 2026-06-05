# Manual de uso — 05b_override

**Para quién es este manual:** la persona que detecta que un pago fue
registrado al lote incorrecto y necesita corregirlo sin volver a correr
módulos anteriores.

---

## Qué hace este módulo

A veces un pago llega asignado al lote equivocado — ya sea en Yape/TePagó
o en los pagos en efectivo. Este módulo mueve ese pago del lote incorrecto
al lote correcto y actualiza `cobranza_final.xlsx` en consecuencia:

```
overrides.xlsx  (llenas tú)
                    ↓
              05b_override
                    ↓
  pagos_yape_tepago.xlsx  ← MZ/LOTE corregidos
  pagos_efectivo.xlsx     ← MZ/LT corregidos
  cobranza_final.xlsx     ← SALDO y ESTADO actualizados
  trazabilidad_overrides_YYYY-MM.xlsx  ← registro del cambio
  backup/                 ← copia de los 3 archivos antes del cambio
```

---

## Cuándo usar este módulo

**Después** de que `04_cobranza` ya corrió y generaste `cobranza_final.xlsx`.
Si todavía estás en el proceso de cobranza, es más fácil corregir
directamente en `pagos_efectivo.xlsx` antes de correr 04.

Casos típicos:
- El banco asignó el Yape a "MZ C LT 3" pero el pago era de "MZ C LT 13"
- Cobraste en efectivo a "MZ A LT 5" pero lo cargaste en "MZ A LT 6"

---

## Paso 1 — Llenar overrides.xlsx

El archivo está en `inputs/overrides.xlsx`. Una fila por corrección:

| Columna | Qué poner |
|---|---|
| **FECHA** | Fecha del pago (solo para YAPE — formato como aparece en el reporte del banco) |
| **MEDIO** | `YAPE` o `EFECTIVO` |
| **MZ_ANTERIOR** | Manzana donde está registrado el pago hoy (el incorrecto) |
| **LOTE_ANTERIOR** | Lote incorrecto |
| **MZ_NUEVO** | Manzana correcta |
| **LOTE_NUEVO** | Lote correcto |
| **MOTIVO** | Texto libre — explica por qué se corrige (para el registro) |
| **ESTADO** | Escribir `pendiente` — el módulo lo cambia a `aplicado` al terminar |

**Nota YAPE:** La búsqueda usa FECHA + MZ + LOTE. Si no coincide exactamente
con el reporte del banco, el override se saltará con una advertencia.

**Nota EFECTIVO:** No hay fecha en el registro de efectivo. La búsqueda usa
solo MZ + LT. Si hay varias filas con ese MZ/LT, todas se mueven.

---

## Paso 2 — Correr el módulo

Abrir una terminal en la carpeta `05b_override/` y ejecutar:

```
python main.py
```

El sistema mostrará el progreso:

```
═══════════════════════════════════════════════════════
  05b_override — Corrección de pagos mal asignados
═══════════════════════════════════════════════════════

[1/4] Validando inputs...
[2/4] Cargando overrides pendientes...
[3/4] Aplicando 2 override(s)...

  → YAPE C-3 → C-13
     S/50.00 movidos — aplicado

  → EFECTIVO A-5 → A-6
     S/30.00 movidos — aplicado

[4/4] Guardando trazabilidad...

═══════════════════════════════════════════════════════
  2 override(s) aplicados correctamente
  Trazabilidad: trazabilidad_overrides_2025-05.xlsx
  Backup en: backup/
═══════════════════════════════════════════════════════
```

Si un override no encontró coincidencia:

```
  → YAPE C-3 → C-13
     SIN COINCIDENCIA — override saltado (revisar FECHA/MZ/LOTE)
```

---

## Revisar los outputs

### trazabilidad_overrides_YYYY-MM.xlsx

Registro de todos los cambios aplicados este mes. Una fila por override:

| Columna | Descripción |
|---|---|
| APLICADO_EN | Fecha y hora en que se aplicó el cambio |
| MEDIO | YAPE o EFECTIVO |
| FECHA_PAGO | Fecha del pago original (solo YAPE; `—` para EFECTIVO) |
| MZ_ANTERIOR / LOTE_ANTERIOR | De dónde se movió el pago |
| MZ_NUEVO / LOTE_NUEVO | A dónde se movió |
| MONTO_MOVIDO | Importe en soles |
| MOTIVO | Texto que escribiste en overrides.xlsx |

### backup/

Antes de aplicar cualquier cambio, el módulo guarda copias con marca de tiempo:

```
backup/
  pagos_yape_tepago_2025-05-15_143022.xlsx
  pagos_efectivo_2025-05-15_143022.xlsx
  cobranza_final_2025-05-15_143022.xlsx
```

Un juego de backups por ejecución. Si algo salió mal, reemplaza los tres
archivos originales con los del backup y vuelve a correr.

---

## Qué hacer si hay un error

1. Revisar `outputs/run.log` — el log muestra exactamente qué falló
2. Restaurar desde `backup/` si fue necesario modificar manualmente
3. Corregir `overrides.xlsx` (la fila del error quedará en `pendiente`)
4. Volver a correr — el módulo solo procesa filas con ESTADO `pendiente`

---

## Referencia rápida

```
PRE-CONDICIONES
  04_cobranza ya corrió → cobranza_final.xlsx existe

INPUTS
  inputs/overrides.xlsx          ← llenas tú antes de correr

OUTPUTS
  pagos_yape_tepago.xlsx         ← actualizado (en 04_cobranza/inputs/pagos_yape/)
  pagos_efectivo.xlsx            ← actualizado (en 04_cobranza/inputs/pagos_efectivo/)
  cobranza_final.xlsx            ← actualizado (en 04_cobranza/outputs/)
  outputs/trazabilidad_overrides_YYYY-MM.xlsx
  outputs/run.log
  backup/  ← copias pre-cambio con timestamp

DESPUÉS DE CORRER
  Verificar en cobranza_final que los saldos quedaron correctos
  El módulo puede correrse varias veces — solo procesa filas "pendiente"
```
