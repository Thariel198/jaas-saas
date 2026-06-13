# backup/ — código en transición

Esta carpeta es temporal. Contiene código y archivos que fueron construidos
antes de que existiera la arquitectura formal del sistema y que no pertenecen
a su ubicación actual.

---

## Por qué existe

`cargar_planilla` fue construido de forma independiente antes de que se definiera
el pipeline completo de `jass_system`. Como resultado acumuló deuda técnica:

- Tenía sus propios `Inputs/` en lugar de consumir los outputs de otros módulos
- Vivía en `03_pagos/yape/` cuando su responsabilidad corresponde a `04_cobranza/`
- Su estructura no sigue los estándares de arquitectura definidos

En lugar de borrar todo sin revisión, se guarda aquí lo que puede tener
lógica útil para reutilizar cuando se construya la versión correcta.

---

## Qué hay aquí

| Archivo | Por qué se guarda | Cuándo se usa |
|---------|-------------------|---------------|
| `cargar_planilla/main.py` | Puede tener lógica de carga de planilla reutilizable | Al construir `04_cobranza/cargar_planilla/` desde cero |

---

## Reglas de esta carpeta

- **No se ejecuta nada de aquí** — es solo referencia
- **No se importa desde ningún módulo** — no forma parte del sistema activo
- **Se elimina completa** cuando el módulo correcto esté construido y validado
- Si se agrega algo nuevo aquí, documentarlo en la tabla de arriba con su motivo

---

## Cuándo se elimina

Cuando `04_cobranza/cargar_planilla/` esté construido, probado y funcionando
correctamente dentro de la arquitectura formal. En ese momento esta carpeta
`backup/` se borra completa — ya no tiene razón de existir.

---

## Nota

Guardar aquí no significa que el código sea correcto o reutilizable tal cual.
Significa que puede contener ideas o lógica que vale la pena revisar antes
de construir la versión limpia.
