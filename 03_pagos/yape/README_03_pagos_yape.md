# 03_pagos / yape

Identifica a qué MZ-LOTE corresponde cada pago Yape recibido en el mes.
Compuesto por dos módulos: uno construye el maestro de usuarios, el otro lo usa cada mes.

---

## Estructura de carpetas

```
03_pagos/yape/
├── construir_maestro/         → corre solo cuando hay cambios en usuarios
│   ├── Inputs/
│   │   └── Planilla/          → planillas históricas del banco (.xlsx)
│   ├── Outputs/
│   │   └── maestro_yape.xlsx  → tú lo copias a shared/ manualmente
│   └── main.py
│
├── motor_matching/            → corre cada mes
│   ├── Inputs/
│   │   ├── Maestro/           → copias maestro_yape.xlsx desde shared/
│   │   ├── Planilla_anterior/ → planilla del mes pasado (ancla de corte)
│   │   ├── Planilla_mes/      → planilla del mes actual con deudas
│   │   └── Reporte_mes/       → ReporteTransacciones.xlsx del banco
│   ├── Correcciones/          → temporal · se borra en Ciclo 1
│   │   └── pendientes.xlsx    → 1 hoja: Sin_identificar
│   ├── Outputs/               → temporal · se borra en Ciclo 1
│   │   ├── pagos_yape.xlsx    → resultado final → 04_cobranza
│   │   └── resumen_validacion.txt
│   ├── Trazabilidad/          → PERMANENTE · nunca se borra
│   │   └── trazabilidad_YYYY_MM.xlsx
│   ├── Blancos/               → PERMANENTE · nunca se borra
│   │   └── blancos_acumulados.xlsx → lo lee cargar_planilla en 04_cobranza
│   ├── main.py
│   └── README.md              → reglas detalladas de motor_matching
│
└── arquitectura_03_pagos_yape.html

```

> Ver `motor_matching/README.md` para las reglas de negocio detalladas de motor_matching.

---

## Recursos compartidos — shared/

Estos archivos viven en `shared/` y son leídos directamente desde ahí:

| Archivo | Quién lo produce | Quién lo lee |
|---------|-----------------|--------------|
| `usuarios_id.xlsx` | manual | construir_maestro · motor_matching |
| `maestro_yape.xlsx` | construir_maestro → tú lo copias | motor_matching |
| `blancos_acumulados.xlsx` | motor_matching | 04_cobranza · 05b_override |

---

## Módulo 1 — construir_maestro

### ¿Qué hace?
Aprende de planillas históricas del banco qué origen de pago corresponde a qué MZ-LOTE.
Construye el maestro de usuarios con nivel de confianza por origen.

### ¿Cuándo se corre?
Solo cuando hay usuarios nuevos o cambios en los orígenes de pago.
Meses normales no se toca — el `maestro_yape.xlsx` en `shared/` ya está listo.

### Flujo
```
1. Agregar planillas nuevas en Inputs/Planilla/
2. Correr: python main.py
3. Copiar Outputs/maestro_yape.xlsx → shared/maestro_yape.xlsx
```

### Inputs
- Planillas históricas del banco (.xlsx) en `Inputs/Planilla/`
- `shared/usuarios_id.xlsx` — MZ y LOTE por usuario

### Output
- `Outputs/maestro_yape.xlsx` — orígenes de pago con nivel de confianza
- Tú lo copias manualmente a `shared/` — en el futuro será automático

---

## Módulo 2 — motor_matching

### ¿Qué hace?
Cada mes identifica a qué MZ-LOTE corresponde cada pago Yape del banco.
Corre en ciclos hasta que todos los pagos estén identificados (pendientes = 0).

### Prioridad de identificación

El sistema intenta identificar cada pago en este orden:

```
1. Corrección acumulada (ciclo 2+)   → prioridad máxima · ya fue validado manualmente
2. Múltiples en mensaje              → "MzE Lt7, MzP Lt11A" · automático
3. Mensaje simple                    → regex extrae 1 MZ-LOTE · automático
4. Maestro ambiguo                   → 2+ candidatos · elige menor diff · automático
5. Maestro único                     → 1 solo lote en maestro · automático
P. Sin identificar                   → va a pendientes.xlsx · intervención manual
```

> Ambiguos y Pagos_multiples se resuelven solos — nunca van a pendientes.
> Solo Sin_identificar requiere que tú completes MZ/LOTE o escribas BLANCO.

### Ciclos

| Ciclo | Señal | Qué hace |
|-------|-------|----------|
| Ciclo 1 | `pagos_yape.xlsx` NO existe | Mes nuevo · limpia Correcciones/ y Outputs/ · Trazabilidad/ y Blancos/ intactos |
| Ciclo 2+ | `pagos_yape.xlsx` SÍ existe | Lee correcciones · actualiza trazabilidad · regenera pendientes con los que faltan |

### Flujo mensual

```
1. Copiar inputs del mes nuevo
2. python main.py → Ciclo 1
3. Completar pendientes.xlsx hoja Sin_identificar
4. python main.py → Ciclo 2
5. Repetir 3 y 4 hasta pendientes = 0
6. pagos_yape.xlsx listo → pasa a 04_cobranza
```

### Archivos que produce

| Archivo | Tipo | Destino |
|---------|------|---------|
| `Outputs/pagos_yape.xlsx` | temporal · se borra en Ciclo 1 del mes siguiente | 04_cobranza |
| `Outputs/resumen_validacion.txt` | temporal · su ausencia indica Ciclo 1 | solo lectura |
| `Correcciones/pendientes.xlsx` | temporal · se borra en Ciclo 1 | tú lo completas |
| `Trazabilidad/trazabilidad_YYYY_MM.xlsx` | permanente · nunca se borra | auditoría |
| `Blancos/blancos_acumulados.xlsx` | permanente · crece indefinidamente | 04_cobranza · 05b_override |

---

## Lo que NO hace este módulo

- **No hace override** — las correcciones de usuario por confusión de MZ-LOTE después de que el dinero cuadró se hacen en `05b_override/` — después de corte y antes de boletas
- **No carga la planilla** — eso es responsabilidad de `04_cobranza/cargar_planilla/`
- **No valida que el dinero cuadre** — eso lo hace `04b_validacion/`

---

## Errores comunes

**"No se encontró maestro_yape.xlsx"**
→ Verificar que copiaste el archivo a `Inputs/Maestro/` desde `shared/`.

**"No hay archivo en Planilla_anterior"**
→ El sistema corre sin ancla de corte — incluirá todos los pagos sin filtro de fecha.

**"MZ-LOTE no existe en planilla"**
→ El MZ-LOTE que escribiste en pendientes.xlsx no existe en la planilla del mes.

**"Ningún archivo válido en Reporte_mes"**
→ El archivo del banco no tiene la columna "Tipo de Transacción" o no hay filas "TE PAGÓ".
