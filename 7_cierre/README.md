# 7_cierre

Verifica que el ciclo mensual cerró correctamente y prepara los inputs para el mes siguiente.

> **Estado:** placeholder · no implementado todavía.
> Se construye cuando 1–6 estén funcionando con datos reales.

## Cuándo correr

Último paso del ciclo, después de que `6_corte` y `6b_override` cerraron OK.

```
... → 5b_validacion → 6_corte → 6b_override → 7_cierre → (próximo mes)
```

## Qué hace

**No calcula nada.** Solo verifica y coordina:

1. Confirma que existen `arrastre_deuda_YYYY-MM.xlsx` (de 5b) y `arrastre_corte_YYYY-MM.xlsx` (de 6)
2. Copia ambos a `2_planilla/inputs/` con el nombre del próximo mes
3. Crea tag git del ciclo (`v2026-MM`)
4. Actualiza `docs/CHANGELOG.md` con el resumen del mes

## Por qué existe

Reemplaza al viejo `07_nueva_planilla` que era redundante: los arrastres ya los emiten `5b_validacion` y `6_corte`. Aquí solo se verifica que están listos y se transfieren al siguiente ciclo.

## Lo que NO hace

- No calcula deudas (lo hace 5b_validacion)
- No calcula cortes (lo hace 6_corte)
- No genera planilla del próximo mes (lo hace 2_planilla del próximo mes)
