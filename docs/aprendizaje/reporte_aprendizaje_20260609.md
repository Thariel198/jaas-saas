# Reporte de Aprendizaje — 09 Junio 2026

---

## Lo que se construyó esta sesión

### Parte 1 — Bug fix motor_matching: correcciones manuales perdidas en Ciclo 1

Se identificó y corrigió un bug crítico: al volver a correr Ciclo 1, el código borraba
`correcciones/` con `shutil.rmtree` en el paso [1], **antes** de que el paso [10] leyera
las correcciones preservadas. El trabajo manual desaparecía silenciosamente.

Corrección aplicada:
- Paso [1] ya no borra `correcciones/` — solo limpia `outputs/`
- Se agregó `cleanup_temporales()` para que lo llame `7_cierre` después del git tag

### Parte 2 — Rediseño completo del módulo 4_pagos/efectivo

Se rediseñó el módulo de 2 personas fijas (`registro_mia / registro_amiga`) a sistema
de N mesas con hasta 3 cobradores cada una. Proceso completo siguiendo la metodología:

**Fase 2 completa:**
- ADR en `docs/decisiones/4_pagos_efectivo.md`
- `4_pagos/efectivo/docs/diagrama_efectivo.html`
- `4_pagos/efectivo/README.md`
- `4_pagos/efectivo/docs/arquitectura_efectivo.html`
- 3 contratos de formato: `formato_registro.html`, `formato_pagos_efectivo.html`, `formato_discrepancias.html`

**Fase 3:**
- `crear_templates.py` reescrito → genera `mesa_1…7.xlsx` con 3 hojas cada uno
- `main.py` reescrito → cross-check por mesa, regla de mayoría, pago_multi_mesa, resolución de discrepancias

**Metodología v2.0:**
- Pasos 2.1, 2.3 y 2.6 actualizados: los artefactos de diseño de módulo viven en
  `MODULO/docs/`, no en la carpeta `docs/` global del sistema

---

## Términos técnicos aprendidos

### cleanup_temporales() — patrón profesional de limpieza

**Qué es:** Cada módulo expone una función `cleanup_temporales()` que borra sus archivos
temporales del mes. Esta función NO se auto-llama al inicio del módulo. La invoca
`7_cierre` después de que el git tag del ciclo se creó exitosamente.

**De dónde sale:** El debate sobre cuándo limpiar en motor_matching. Dos opciones:
- Limpiar al inicio del Ciclo 1 (lo que había) → destruye trabajo manual en re-corridas
- Limpiar al final del pipeline completo (patrón profesional) → preserva todo durante el mes

**Por qué al final:** El mes puede requerir múltiples re-corridas. Si cada re-corrida
borra el estado anterior, el trabajo manual (correcciones, resoluciones) desaparece.
El cleanup solo tiene sentido cuando el ciclo completo ya se guardó en Git.

**Regla:** Cada módulo hace lo mismo:
```python
def cleanup_temporales():
    # Borra temporales del mes. NO se autollama.
    # La invoca 7_cierre después del git tag exitoso.
    ...
```

**En tu sistema:** `motor_matching/main.py` y `4_pagos/efectivo/main.py` — ambos tienen
esta función. `7_cierre` las orquestará cuando se implemente.

---

### Mesa como unidad de cross-check

**Qué es:** Una mesa física de cobranza se modela como 1 archivo Excel con hasta 3 hojas.
El cross-check ocurre dentro del archivo (misma mesa), no entre archivos (diferentes mesas).

**De dónde sale:** El diseño original comparaba 2 archivos globales (`mia` vs `amiga`).
Con 7 mesas y 3 personas por mesa, ese modelo no escala — no hay forma de saber qué par
comparar ni a qué mesa pertenece cada registro.

**Regla de diseño:** La unidad de cross-check es la mesa física, no el cobrador individual.
`mesa_N.xlsx` → registro_1, registro_2, registro_3 (máx 3).

**Alternativas descartadas:**
- 1 archivo por cobrador (14 archivos) → pierde el agrupamiento por mesa
- Un archivo único con columna MESA por fila → ruido visual, error humano probable
- Nombres fijos mia/amiga → no escala, no refleja realidad operativa

---

### Regla de mayoría 2/3

**Qué es:** Si una mesa tiene 3 cobradores y 2 coinciden pero 1 difiere, se toma
la mayoría como verdad. La fila discrepante queda en `trazabilidad/incidencias`.

**Por qué:** No bloquea el proceso por un error de un cobrador. La traza permanece
para auditoría. Si la señal del 30% se activa (2 meses seguidos >30% solo_un_cobrador),
es momento de revisar la metodología operativa.

**Implementación:**
```python
# Si 2 de 3 hojas coinciden en MZ+LT+MONTO → mayoria_aplicada
# La minoría va a trazabilidad/incidencias_YYYY-MM.xlsx
```

---

### Protección de trabajo manual — "sagrado"

**Qué es:** Los archivos que llenan los usuarios a mano nunca se borran sin respaldo
explícito previo. Antes de cualquier migración o rediseño: backup primero.

**De dónde sale:** Al redesignar efectivo, los `registro_NN.xlsx` de junio ya estaban
llenos. Moverlos a `backup/migracion_2026_06/` antes de generar los nuevos templates
asegura que los datos nunca se pierdan, incluso si el re-llenado toma días.

**Regla:** `inputs/` de módulos manuales = sagrado. Antes de cualquier cambio:
```
1. Copiar a backup/migracion_YYYY-MM/
2. Verificar que el backup está completo
3. Recién entonces crear los nuevos templates
```

---

### Docs de módulo en carpeta propia (metodología v2.0)

**Qué es:** Los artefactos de diseño de un módulo (diagrama, arquitectura, formatos)
viven en `MODULO/docs/`, no en la carpeta `docs/` global del sistema.

**De dónde sale:** Al crear el diagrama de efectivo, la metodología v1.9 decía
`docs/diagrama_MODULO.html` (global). El módulo 2_planilla ya tenía su propio
`2_planilla/docs/`. Se generalizó la regla.

**Regla:** La carpeta `docs/` global es para documentación del sistema completo
(arquitectura general, decisiones cross-módulo, skill tracker, metodología).
Cada módulo tiene su propio `MODULO/docs/` para sus contratos de diseño.

---

## Errores cometidos — y lo que revelan

### No leer HTMLs antes de codificar

El usuario tuvo que recordarme explícitamente: "acuerdate de usar todos los html"
antes de que empezara a codificar. Estaba a punto de escribir código sin releer
los contratos `formato_*.html` que acababa de crear.

**Lo que revela:** Los contratos HTML son la especificación. Si el código no los
implementa exactamente, el formato aprobado y el código divergen. El error es
silencioso — el código funciona pero no replica lo que se diseñó.

**Corrección:** Guardado en memoria. Antes de codificar cualquier exportación Excel,
leer explícitamente cada `formato_*.html` del módulo. No asumir que se recuerda.

---

## Lo que hiciste bien — nivel profesional

- **Detectaste el patrón profesional de cleanup.** Cuando se presentaron las dos opciones
  (limpiar al inicio vs al final), identificaste inmediatamente que limpiar al final
  del pipeline es lo correcto: "eso debería ser el flujo profesional".

- **Corregiste el diseño de los docs antes de codificar.** Al notar que los artefactos
  de diseño estaban apuntando a la carpeta global, pediste actualizar la metodología
  primero. El código vino después.

- **Aplicaste la metodología completa para efectivo, de corrido.** Fase 2 entera: ADR →
  diagrama → README → arquitectura → 5 contratos de formato → carpetas físicas → código.
  Sin saltear pasos.

- **Preguntaste antes de codificar en toda la sesión.** Cumpliste la regla
  "no decidir solo" consistentemente — cada vez que había una decisión de diseño,
  preguntaste y esperaste confirmación.

---

## Pendientes

1. Correr `main.py` de efectivo con los datos de junio llenados (datos ya están en `inputs/`)
2. Correr `main.py` de motor_matching para terminar de identificar pagos Yape pendientes
3. Borrar `registro_01…07.xlsx` de `inputs/` (ya están en backup, ya se re-llenaron en `mesa_N.xlsx`)
4. Borrar `registro_diseno.html`, `discrepancias_diseno.html`, `pagos_efectivo_diseno.html` de `docs/`

---

## Resumen

Sesión dividida en dos mitades. La primera resolvió un bug sutil en motor_matching
(correcciones manuales que desaparecían en re-corridas de Ciclo 1) y estableció
el patrón profesional de cleanup: cada módulo expone `cleanup_temporales()`, el pipeline
la llama al cerrar el mes. La segunda aplicó la metodología completa para redesignar
el módulo efectivo de 2 personas a N mesas, produciendo 5 contratos HTML, README,
arquitectura y código nuevo. El feedback más importante: los contratos HTML se escriben
para ser leídos antes de codificar, no solo para ser creados.
