# Decisión — Rediseño módulo efectivo: de 2 personas a N mesas

**Fecha:** 2026-06-09
**Módulo afectado:** `4_pagos/efectivo`

---

## Problema

El diseño original comparaba dos registros fijos (`registro_mia.xlsx` vs `registro_amiga.xlsx`):
una cobra, otra anota — cross-check de 2 personas en una sola sesión. En junio 2026 se
incorporaron 7 mesas físicas de cobranza, cada una con hasta 3 cobradores. El modelo de
2 archivos no escala: no hay forma de saber qué par comparar ni de cuál mesa proviene
cada registro.

## Criterios de éxito

1. El código soporta de 1 a 7 mesas sin cambiar estructura.
2. Cada mesa puede tener 1, 2 o 3 cobradores (cross-check opcional, no requerido).
3. Cuando solo hay 1 cobrador por mesa, el sistema acepta el registro sin bloquear.
4. Si un usuario aparece en más de una mesa (pago dividido), se detecta y se registra.
5. Toda incidencia queda en `trazabilidad/` — auditable mes a mes.
6. El trabajo manual (datos ya ingresados) nunca se borra sin respaldo explícito.

## Enfoque elegido — 1 mesa = 1 archivo, hasta 3 hojas internas

```
inputs/
  mesa_1.xlsx   ← hojas: registro_1, [registro_2], [registro_3]
  mesa_2.xlsx
  ...
  mesa_7.xlsx
```

Cada archivo representa una mesa física. Las hojas dentro del archivo son los registros
de cada cobrador en esa mesa. El cross-check ocurre **dentro del mismo archivo** (misma
mesa). Mesas distintas → registros distintos por diseño.

Columnas de cada hoja: `COBRADOR | FECHA_REGISTRO | MZ | LT | MONTO | FECHA | COMENTARIO`

### Lógica de consolidación por mesa

| Situación | Resultado |
|---|---|
| 1 sola hoja | `solo_un_cobrador` — se acepta como verdad sin comparar |
| 2-3 hojas coinciden | `confirmado` |
| 2-3 hojas con diferencia | `discrepancia` — requiere revisión |
| Mismo (MZ, LT) en mesas distintas | `pago_multi_mesa` — alerta, se traza |

### Regla de mayoría en discrepancias

Si 2 de 3 cobradores coinciden → se toma la mayoría como verdad.
La minoría queda registrada en `trazabilidad/incidencias_YYYY-MM.xlsx`.

## Alternativas descartadas

| Alternativa | Por qué se descartó |
|---|---|
| Mantener `registro_mia / registro_amiga` | No escala a 7 mesas. Semántica confusa: "mia" y "amiga" no son roles, son personas. |
| Un archivo único con columna MESA por fila | Ruido visual. El usuario llenaría mal la columna MESA. Más difícil de validar. |
| 1 archivo por cobrador (14 archivos) | Pierde el agrupamiento por mesa. No refleja la realidad operativa. |
| 4+ cobradores por mesa | Ineficiente operativamente. 3 es el máximo trazable con cruce manual. |

## Migración junio 2026

Los `registro_01.xlsx … registro_07.xlsx` actuales (ya llenados) se mueven a
`backup/migracion_2026_06/`. El usuario re-llena los datos manualmente en los nuevos
`mesa_1.xlsx … mesa_7.xlsx`. Este mes solo hay 1 cobrador por mesa → todas las filas
saldrán como `solo_un_cobrador`, lo cual es correcto y esperado.

## Señal de alerta

Si más del 30% de los registros salen como `solo_un_cobrador` durante **2 meses seguidos**,
es señal de que la metodología operativa no se está aplicando en campo.
Acción: revisar con los cobradores el procedimiento de doble registro por mesa.
