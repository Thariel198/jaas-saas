# 6_corte

Módulo que ejecuta el ciclo de corte de servicio: genera la lista de usuarios en mora elegibles para corte, aplica la penalidad de S/20, gestiona la ventana de gracia de 2 días y clasifica el resultado final en pagados, cortados físicamente y arrastre al mes siguiente.

## Qué hace

1. **Genera la lista de corte** (`generar_lista.py`): filtra `planilla_cobrado.xlsx` (ciclo 1) por `SALDO > 0 AND MES_ANTERIOR ≥ 8` y produce `lista_corte.xlsx` con `PENALIDAD = S/20` y `TOTAL_A_PAGAR = SALDO + 20`.
2. **Aplica la penalidad** (`aplicar_penalidad.py`): suma `+20` a `CORTE_RECONEXION` en `shared/planilla_mes/planilla_YYYY-MM.xlsx` para cada usuario en lista_corte. Genera audit log para idempotencia; re-correr no duplica.
3. **Espera ventana de gracia** (48 h): el usuario puede pagar S/20 por Yape o efectivo y salvarse del corte físico.
4. **Clasifica el resultado** (`seguimiento.py`): cruza lista_corte con `planilla_cobrado.xlsx` ciclo 2 (post-ventana) y separa en tres grupos: pagaron penalidad, corte físico, arrastre.

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
| `5_cobranza/outputs/planilla_cobrado.xlsx` | archivo (lectura) | `5_cobranza/` — ciclo 1 y ciclo 2 |
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
