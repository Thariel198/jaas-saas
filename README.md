# jass_system

Sistema de gestión mensual de cobros de agua para la JASS. Automatiza el ciclo completo: desde la lectura de medidores hasta el cierre del mes con arrastres para el siguiente ciclo.

---

## Pipeline de módulos

```
0_padron → 1_lecturas → 2_planilla → 3_boletas
→ [4_pagos + 4b_reclamos]
→ [5_cobranza + 5b_validacion]
→ [6_corte + 6b_corte_multas]
→ 7_cierre
```

| Módulo | Qué hace | Estado |
|---|---|---|
| `0_padron` | Limpia, reconcilia y mantiene el padrón de socios (MZ+LT). Sub-módulos: `01_limpieza`, `02_matching`, `03_llenado`. | Operativo |
| `1_lecturas` | Sincroniza el registro del operario con el padrón y procesa las lecturas mensuales del medidor. Produce `lecturas_planilla_YYYY-MM.xlsx`. | Operativo |
| `2_planilla` | Genera la planilla mensual de cobro consolidando lecturas con deudas arrastradas del mes anterior. | Operativo |
| `3_boletas` | Genera las boletas de cobro por socio. | En desarrollo |
| `4_pagos` | Registra pagos recibidos — efectivo y Yape — y los aplica a la planilla. | Operativo |
| `4b_reclamos` | Gestiona el ciclo de vida de reclamos detectados en cobros de efectivo: detección, clasificación, corrección y cierre con auditoría. | Operativo |
| `5_cobranza` | Consolida la cobranza del mes: planilla cobrado, correcciones, arrastres y devoluciones. | Operativo |
| `5b_validacion` | Valida que el dinero cobrado cuadre con lo registrado en el sistema. | Operativo |
| `6_corte` | Genera la lista de usuarios en mora elegibles para corte, aplica penalidad de S/20, gestiona ventana de gracia de 2 días. Ciclo: `BORRADOR → PUBLICADA → COMPROMETIDA`. | Operativo |
| `6b_corte_multas` | Espejo de `6_corte` para deuda de multas y acuerdos de asamblea. Penalidad de S/40. | Operativo |
| `7_cierre` | Consolida decisiones del mes (reclamos + cortes) y produce arrastres finales para que `2_planilla` del próximo mes los pre-cargue. | Diseñado — pendiente de implementación |

Los módulos `b` son dependientes del principal: `4b` valida pagos, `5b` cuadra el dinero, `6b` gestiona multas paralelas al corte.

---

## Cómo correr el ciclo mensual

Cada módulo tiene su propio `README.md` con el flujo paso a paso y los comandos exactos.
El orden del pipeline es el de la tabla de arriba.

```bash
# Ejemplo: correr un módulo
cd 1_lecturas
python main.py
# El resultado queda en outputs/ junto con run.log
```

---

## Estructura general de cada módulo

```
MODULO/
├── inputs/          ← archivos de entrada del ciclo
├── outputs/         ← archivos generados + run.log
├── docs/            ← diagramas HTML y contratos de formato
├── tests/           ← tests sintéticos (donde aplica)
├── backup/          ← respaldos automáticos pre-corrida
├── config.py        ← paths y constantes del módulo
├── main.py          ← punto de entrada principal
└── README.md        ← fuente de verdad del módulo
```

---

## Recursos compartidos

```
shared/          ← funciones puras reutilizables entre módulos
docs/            ← metodología, arquitectura del sistema, skill tracker
recursos/        ← archivos de referencia (tarifas, padrón base, etc.)
```

---

## Instalación

```bash
pip install -r requirements.txt
```

Requiere Python 3.10+, pandas, openpyxl.
