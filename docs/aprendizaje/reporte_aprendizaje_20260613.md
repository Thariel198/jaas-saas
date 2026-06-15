# Reporte de Aprendizaje — 13 Junio 2026

---

## Lo que se construyó esta sesión

### Parte 1 — GitHub push de jass_system

Se realizó el primer push del proyecto jass_system a GitHub. El desafío no fue técnico sino
de gitignore: el módulo `3_boletas` tiene una carpeta `Outputs - copia/` con 160MB+ de PDFs
que exceden el límite de GitHub (100MB por archivo).

Soluciones aplicadas:
- `.gitignore` actualizado: `outputs/` (lowercase), `venv/`, `.venv/`, `**/Outputs - copia/`, `3_boletas/Outputs - copia/`
- La regla `**/Outputs - copia/` es más robusta que enumerar cada módulo
- Push exitoso después de resolver el bloqueador de archivos grandes

**Lección:** Antes de hacer el primer push de un proyecto existente, revisar si hay carpetas de outputs grandes que no deben viajar a GitHub. Los PDFs de boletas son outputs regenerables — no pertenecen al repositorio.

### Parte 2 — Análisis de colisión en padrón catastral (00_padron)

Consulta: si se agrega C1-16 NELLY MORENO LICITO a COFOPRI (lote ya asignado a MARIA COSME JULCA), ¿habrá colisión en `padron_reconciliado.xlsx`?

Análisis del código existente:
- `multi_lot_cof`: detecta personas con 2+ lotes COFOPRI y asigna todos explícitamente — NELLY MORENO LICITO tendría entonces 2 lotes (C1-16 + el original), se convierte en caso multi-lote
- MARIA COSME JULCA en C1-16 se detecta como conflicto y aparece en `reporte_conflictos.xlsx`
- El operario decide (AUTORIZAR=Si) → `aplicar_conflictos.py` reemplaza en reconciliado
- **Resultado: no hay crash, el sistema maneja esto**

### Parte 3 — Diseño del patrón de preservación para 00_padron

Objetivo: aplicar en 00_padron el mismo patrón de preservación de 4_pagos para que el trabajo
manual de autorización no se pierda al re-correr los módulos.

**Decisiones de arquitectura tomadas:**
- Enfoque elegido: thin layer of shared primitives (no duplicación total, no módulo grande)
- Ubicación shared: `00_padron/shared/utils_preservacion.py`
- Auto-apply: si AUTORIZAR=Si de sesión anterior → se aplica automáticamente sin aparecer en el nuevo reporte
- Backup: `02_matching/outputs/backup/` con timestamp
- Columna REVISADO: distingue "no visto" (rojo) de "visto, sin cambio" (neutro) de "autorizado" (auto-apply)

**Estado al cerrar sesión:** arquitectura decidida, implementación pendiente.

---

## Términos técnicos aprendidos

### Thin layer of shared primitives

**Qué es:** El punto medio profesional entre "duplicar todo en cada módulo" y "crear un módulo compartido grande".

**El problema con el módulo compartido grande:**
Cuando varios módulos comparten el 80% de lógica pero el 20% difiere, crear `preservar.py` con
todos los pasos resulta en funciones con 8 parámetros y callbacks porque cada módulo tiene sus quirks.
Termina siendo más difícil de mantener que la duplicación.

**La solución: primitivos puros en shared, orquestación en main.py:**
```
shared/utils_preservacion.py          02_matching/main.py
───────────────────────────           ──────────────────────────────
backup_con_timestamp(path)   →  usa  backup_con_timestamp(CONFLICTOS_PATH)
leer_xlsx_a_mapa(path, key)  →  usa  leer_xlsx_a_mapa(path, key=("MZ","LT","NOMBRE"))
normalize(s)                 →  usa  normalize(s)
```

**Test rápido para decidir:**
- Si la función necesita `if modulo == "X"` adentro → no compartir, vive en cada módulo
- Si es genuinamente la misma cosa sin importar el llamador → compartir en `shared/`
- La orquestación (orden, columnas llave, cuándo aplicar) → siempre en `main.py`

**Regla del Tres:** No abstraer hasta tener 3+ usos reales de lógica genuinamente idéntica.

**Dónde aparece:** Diseño de 00_padron — tres módulos (limpiar, matching, pendientes) comparten
`backup_con_timestamp` y `leer_xlsx_a_mapa` pero cada uno tiene su flujo de orquestación propio.

**Por qué importa:** Este patrón lo usan kernel Linux, React y pandas. No es "evitar duplicación a toda costa" —
es "duplicar lo que diverge, compartir solo lo que es genuinamente idéntico". La localidad (cada módulo
contiene su lógica) vence al DRY agresivo cuando los módulos tienen distintas reglas de negocio.

---

### Columna REVISADO en archivos de autorización

**El problema:** Un archivo de autorización tiene columna AUTORIZAR. Cuando está vacía, hay dos
interpretaciones posibles pero visualmente idénticas:

1. El operario nunca vio esta fila (es nueva desde el último run)
2. El operario la vio y decidió no autorizar (mantener lo que había)

Al re-correr el módulo, el sistema no puede distinguir cuál es cuál.

**La solución:**

| REVISADO | AUTORIZAR | Significado | Acción |
|---|---|---|---|
| vacío | vacío | Nueva fila — no vista | Resaltar rojo (acción requerida) |
| Si | vacío | Vista, decide conservar | Fondo neutro (sin cambio) |
| Si | Si | Vista, autoriza cambio | Auto-apply en próxima corrida |

**Por qué importa:** Permite que el sistema re-corra sin obligar al operario a re-revisar filas
que ya decidió. Solo las filas ROJAS (REVISADO vacío) requieren atención. Las demás ya fueron
procesadas aunque AUTORIZAR esté vacío.

**Dónde aparece:** Diseño de `reporte_conflictos.xlsx` en 00_padron/02_matching.

---

### Gitignore avanzado para proyectos con outputs grandes

**El problema:** Un proyecto Python que genera PDFs, Excels o imágenes como outputs puede tener
carpetas con cientos de MB que no deben ir a GitHub.

**Patrones útiles en .gitignore:**
```gitignore
# Bloquear cualquier carpeta "Outputs - copia/" en cualquier nivel
**/Outputs - copia/

# Módulo específico
3_boletas/Outputs - copia/

# Outputs generados (regenerables)
outputs/
Outputs/

# Ambientes virtuales
.venv/
venv/
```

**Regla:** si un archivo es regenerable (lo produce el código), no va al repositorio. Si es un
input manual (lo llena el usuario), SÍ va al repositorio.

---

## Errores cometidos — y lo que revelan

### Querer diseñar el módulo compartido completo desde el principio

La primera propuesta fue un `preservar.py` con todos los pasos de preservación compartidos.
El problema: cada módulo de 00_padron tiene columnas llave distintas, claves de deduplicación
distintas, y condiciones de auto-apply distintas. Un módulo compartido con esas diferencias
requeriría parámetros para todo.

**Lo que revela:** DRY agresivo mal aplicado crea acoplamiento. El código que "no se repite"
pero necesita 8 parámetros para funcionar igual en todos lados es más difícil de mantener
que la "duplicación" controlada con primitivos compartidos.

**Corrección:** Thin layer — compartir solo lo genuinamente idéntico. La orquestación es territorio
de cada módulo.

---

## Lo que hiciste bien — nivel profesional

- **Validaste la intuición sobre shared/**: Cuando la IA propuso "todo separado", vos dijiste
  "me imagino que la parte que se repite se puede implementar en una carpeta shared y de ahí
  que compartan los 3 módulos el mismo código." Eso es exactamente el patrón correcto — tu
  intuición llegó antes que la solución técnica.

- **Conectaste el problema de negocio con el diseño técnico**: La pregunta sobre NELLY MORENO
  LICITO no fue "¿falla el código?" sino "¿habrá colisión con el trabajo ya hecho?" — pensaste
  en los datos que ya existen, no solo en el código abstracto.

- **Pediste explicar el diagrama antes de aceptar**: Cuando la IA explicó las 3 opciones,
  pediste un diagrama visual para entender POR QUÉ la opción C era mejor. Eso es criterio
  técnico aplicado — no aceptar la primera sugerencia sin entender el razonamiento.

---

## Tres conceptos que aparecieron sin ser nombrados

### DRY vs localidad — cuándo NO abstraer

**Qué es:** DRY (Don't Repeat Yourself) es una regla de ingeniería, no una ley. Cuando el código
que "se repite" difiere en un 20% por razones de negocio válidas, abstraerlo crea acoplamiento
innecesario.

**La localidad vence al DRY cuando:**
- Cada instancia tiene reglas de negocio propias (distintas columnas llave, distintos flujos)
- El código "compartido" necesitaría parámetros para cada diferencia
- Un cambio en un módulo no necesariamente aplica a los otros

**Dónde apareció:** Diseño de 00_padron. Los tres módulos comparten el CONCEPTO de preservación
pero no los DETALLES de cómo preservar.

---

### Análisis de colisión como pensamiento de producto

**Qué es:** Antes de agregar datos al sistema, pensar en cómo afecta al estado EXISTENTE —
no solo si el código corre sin error.

**Preguntas de análisis de colisión:**
- ¿Este nuevo dato coincide con alguno existente?
- ¿El sistema detecta y maneja esa coincidencia?
- ¿Qué le pasará al operario cuando vea el output?

**Dónde apareció:** C1-16 NELLY MORENO LICITO — el análisis no fue "¿el código falla?" sino
"¿MARIA COSME JULCA aparece en el reporte de conflictos? ¿El operario puede resolver esto?"
Eso es pensar como product manager, no solo como programador.

---

### gitignore como decisión de arquitectura

**Qué es:** El .gitignore no es solo "ignorar archivos temporales" — es una decisión sobre
qué es parte del sistema (código, inputs manuales, config) vs qué es output regenerable.

**El criterio:**
- Código fuente → siempre en Git
- Inputs manuales (llenan los usuarios) → en Git si no son grandes
- Outputs regenerables → fuera de Git
- PDFs/imágenes grandes generados → fuera de Git, siempre

**Dónde apareció:** El primer push de jass_system falló porque 3_boletas/Outputs - copia/ tenía
PDFs de 160MB+. La solución fue reconocer que son outputs regenerables, no parte del sistema.

---

## Pendientes

1. Implementar `00_padron/shared/utils_preservacion.py` con `backup_con_timestamp` y `leer_xlsx_a_mapa`
2. Agregar columna REVISADO a `reporte_conflictos.xlsx` (en `padron_matching.py`)
3. Agregar preservación por clave a `padron_matching.py` (leer REVISADO+AUTORIZAR antes de regenerar)
4. Implementar auto-apply de AUTORIZAR=Si en siguiente corrida de `padron_matching.py`
5. Aplicar mismo patrón en `limpiar_padrones.py` y módulo pendientes

---

## Resumen

Sesión en tres partes. La primera resolvió el primer push de jass_system a GitHub con gitignore
avanzado para excluir PDFs de 160MB+. La segunda analizó la colisión potencial al agregar NELLY
MORENO LICITO a COFOPRI — resultado: el código ya maneja multi-lote y conflictos, no hay crash.
La tercera diseñó la arquitectura de preservación para 00_padron: thin layer of shared primitives
(`shared/utils_preservacion.py` para primitivos puros) + columna REVISADO para distinguir filas
no vistas de filas ya decididas + auto-apply de AUTORIZAR=Si. El aprendizaje central: DRY agresivo
crea acoplamiento; compartir solo lo genuinamente idéntico y dejar la orquestación en cada módulo
es el patrón que usan los sistemas más robustos.
