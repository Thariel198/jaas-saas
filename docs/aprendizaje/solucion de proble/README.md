# Soluciones de problemas

Indice de problemas operativos y de las soluciones verificadas o en curso.

```text
README indice
      |
      +--> 01_tiempo_20_minutos.md
      |
      +--> 02_visualizacion_por_conjuntos.md
      |
      +--> 03_reporte_correccion_pagos.md
```

## Problemas documentados

| Archivo | Problema | Estado |
|---|---|---|
| `01_tiempo_20_minutos.md` | Cada prueba con `5_cobranza --force` tardaba 20 minutos | RESUELTO |
| `02_visualizacion_por_conjuntos.md` | La lista de corte, los abonos y otros precursores no se veian juntos | EN IMPLEMENTACION |
| `03_reporte_correccion_pagos.md` | No existia un reporte completo por usuario para corregir pagos | RESUELTO · diseño de referencia |

## Regla de uso

Cada problema tiene su propio documento. El documento debe conservar:

```text
problema → causa → solucion → diagrama → prueba → estado
```

No mezclar dos problemas en el mismo archivo. Si aparece un problema nuevo, se
agrega otro archivo y se incorpora una linea a esta tabla.
