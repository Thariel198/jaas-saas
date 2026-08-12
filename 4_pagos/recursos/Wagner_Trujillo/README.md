# Wagner Trujillo — hojas de cobro del 01 y 02 de agosto 2026

**Son 2 hojas de papel, no 3.** Las dos fotos del 01/08 son **la misma hoja**,
fotografiada dos veces con el encuadre corrido: ninguna de las dos entra entera.

```
01-08-2026-parte 1.jpeg          01-08-2026-parte 2.jpeg
┌──────────────────────┐         ┌ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─┐   ← corta las 2 primeras filas
│ I-6    yapeó a Yanet │              (no se ven)
│ I-9    faltó cargo   │
├──────────────────────┤         ├──────────────────────┤
│ G-12 · P-13 · P-19   │         │ G-12 · P-13 · P-19   │
│ X-21 · P-4  · W-4    │  ═══>   │ X-21 · P-4  · W-4    │   ← ZONA COMÚN, idéntica
│ V-16 · T-12 · S-7    │  igual  │ V-16 · T-12 · S-7    │      en las dos fotos
│ C1-3 · G-1  · M-15   │         │ C1-3 · G-1  · M-15   │
│ K-8  · O-12 · D-11   │         │ K-8  · O-12 · D-11   │
│ B-6                  │         │ B-6                  │
├ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─┤         ├──────────────────────┤
      (no se ve)                 │ H1-35  23  revisión  │   ← solo aparece en la parte 2
└──────────────────────┘         └──────────────────────┘

parte 1 = completa por ARRIBA        parte 2 = completa por ABAJO
```

Para transcribir esta hoja hay que **leer las dos juntas**: la parte 1 aporta `I-6` e
`I-9`, la parte 2 aporta `H1-35`, y el resto de las filas se lee en cualquiera de las dos.

## ⚠ Ninguna de las dos llega al pie de la hoja

Debajo de `H1-35` la parte 2 todavía muestra un renglón suelto — *"(9 soles) revisión"* con
un *"pago (60.00) consumo y campo"* — sin que se vea a qué MZ-LT pertenece, y el total del
día **no aparece en ninguna de las dos fotos**.

Consecuencia: el 01/08 de Wagner se transcribió **sin verificar**, a diferencia de las otras
mesas, donde los subtotales del papel confirmaron cada bloque. Hace falta una foto que llegue
al pie, o el total declarado de ese día, para validarlo.

## Las dos hojas y su día

| Archivo | Día | Cómo se determinó | Estado |
|---|---|---|---|
| `01-08-2026-parte 1.jpeg` + `parte 2.jpeg` | 01/08/2026 | por descarte: es la hoja que no es la del 02 | ⚠ sin total que la valide |
| `02-08-2026.jpeg` | 02/08/2026 | lo dice el pie: *"Wagner/Julio 2/8/26 — recibe completo 639.00"* | ✅ verificada |

La del 02/08 cierra exacta: las 17 filas suman **639** (626 en efectivo + 13 de yape de
`F1-8`), y ese 639 es justo lo que dice el pie. El único monto ilegible de esa hoja, el de
`H-10`, se dedujo de esa resta: **8 soles**.

## Dónde se cargaron

`4_pagos/efectivo/inputs/mesa_4.xlsx`, hoja `registro_1`:

```
filas  4-20   02/08/2026   17 filas   639  (626 efectivo + 13 yape)   ✅ verificada
filas 21-39   01/08/2026   19 filas   367  (355 efectivo + 12 yape)   ⚠ provisional
```
