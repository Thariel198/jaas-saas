si d# RETOMAR — 4b_reclamos: RECLAMO desincronizado + filas duplicadas · Handoff sesión 2026-07-10

Dónde nos quedamos y qué sigue. Leer de arriba a abajo antes de tocar nada.
**Próxima sesión: Opus** (es diseño de solución, no código mecánico — ver regla de asignación de modelo en CLAUDE.md).

---

## ⚡ TL;DR — lo PRIMERO al retomar

1. **Bug 1 (RECLAMO desincronizado) → ✅ YA RESUELTO esta sesión.** No re-hacer.
2. **Bug 2 (filas duplicadas en `reclamos_2026-07.xlsx`) → ⚠ DIAGNOSTICADO, SIN CODIFICAR.**
   Es la tarea de la próxima sesión: diseñar e implementar un módulo de validación de duplicados
   para el canal efectivo. Ver sección "Lo que sigue" abajo — ya hay una pregunta de arquitectura
   concreta esperando decisión.
3. Antes de correr `4b_reclamos/main.py`, **cerrar `reclamos_2026-07.xlsx` en Excel** — se bloqueó
   el guardado dos veces esta sesión por tenerlo abierto.

---

## Bug 1 — RECLAMO se congelaba con el placeholder "reclamo" (✅ RESUELTO)

### Qué pasaba
El cobrador a veces escribe primero la palabra suelta `"reclamo"` en `COMENTARIO` (nota rápida al
cobrar) y completa el detalle real después, en otra visita a `mesa_N.xlsx`. `4b_reclamos` auto-puebla
`RECLAMO = COMENTARIO` la primera vez que detecta el reclamo — pero `_aplicar_manual()` (línea ~324
de `4b_reclamos/main.py`) trataba cualquier valor no-vacío de `RECLAMO` como "trabajo manual del
supervisor" y lo preservaba para siempre. Como el placeholder "reclamo" nunca está vacío, quedaba
congelado aunque el cobrador después escribiera el detalle real en el input.

Confirmado: 32 filas en `reclamos_2026-07.xlsx` (mesa_1: 4, mesa_4: 28) tenían el placeholder en vez
del detalle real, con el detalle real ya disponible en `mesa_N.xlsx` / `pagos_efectivo.xlsx`.

### Qué se hizo
- **Código** (`4b_reclamos/main.py` línea ~324): se quitó la preservación de `RECLAMO` en
  `_aplicar_manual()` — ahora siempre refleja el `COMENTARIO` fresco del input. Lo que sigue
  preservándose (trabajo manual real del supervisor, no derivable de ningún input): `TIPO_RECLAMO`,
  `RESOLUCION`, `ESTADO`, `FECHA_RESOLUCION`.
- **Header actualizado**: la sección "Reclamo — llenar a mano" pasó a "Reclamo y resolución" (ya no
  se llena a mano, se llena sola desde el input; lo manual es la resolución).
- **Datos**: se corrió `4b_reclamos/main.py --mes 2026-07` y se validó 0 placeholders restantes en
  las 32 filas afectadas. Backups automáticos en `4b_reclamos/backup/reclamos/`.
- **Aprendizaje documentado**: `docs/aprendizaje/Aprendizaje html/placeholder_disfrazado_de_manual_20260710.html`
  — explica el error (auto-poblado vs trabajo manual real), la solución, y de paso qué significa el
  guion bajo en `_aplicar_manual` (convención Python de función privada al módulo).
- **Excepciones sin tocar** (a propósito, confirmado con el usuario):
  - mesa_3 Q-16, D-5: el cobrador todavía no escribió el detalle real en el input — nada que sincronizar.
  - mesa_4 A-1, D-6, S-8: el cobrador ya los reclasificó a `compromiso`/`exoneracion` en el input
    (ya no son `reclamo`), pero siguen arrastrados como reclamo activo en el output. Usuario dijo
    "nada por ahora" — **queda pendiente decidir si se dan de baja** (no es parte de este bug, es
    un tema de que el detector no reacciona cuando `CATEGORIA` cambia después de la primera detección).
  - mesa_4 H1-1: esa fila ya no existe en `mesa_4.xlsx` (se borró o se editó) — no hay con qué cruzar.

**Sin commitear:** `4b_reclamos/main.py` (el fix). Revisar y commitear cuando se cierre también el
bug 2, para no mezclar commits a medias.

---

## Bug 2 — Filas duplicadas por predio en `reclamos_2026-07.xlsx` (⚠ DIAGNOSTICADO, sin codificar)

### Cómo se descubrió
Validando el fix del Bug 1, aparecieron 22 predios con 2 filas idénticas en contenido pero con
`FECHA_COBRO` distinto. Ejemplo real, `mesa_4` G-12 (verificado: en `mesa_4.xlsx` hay **un solo**
pago para G-12, así que es un duplicado real, no dos cobros legítimos):

```
('G','12','mesa_4','Wagner Trujillo', datetime(2026,7,4), 9, '2026-07','2026-07', None, 'Revizar convenio', None, 'PENDIENTE', None)
('G','12','mesa_4','Wagner Trujillo', None,                9, '2026-07','2026-07', None, 'Revizar convenio', None, 'PENDIENTE', None)
```

### Causa raíz
`_pres_key()` (línea 129-132 de `4b_reclamos/main.py`) identifica un "evento reclamo" como
`(MESA, MZ, LT, FECHA_COBRO)`. `FECHA_COBRO` viene del campo `FECHA` de `mesa_N.xlsx` (canal
efectivo), que un humano puede corregir o completar después de la primera detección. Cuando eso
pasa, la clave cambia entre corridas: la fila vieja (con la fecha vieja/vacía) queda como
"arrastre" para siempre, y la detección fresca (con la fecha nueva) entra como fila "nueva" — dos
filas para el mismo predio, ninguna se vuelve a fusionar.

**No es un problema de idempotencia que crece sin límite** (verificado a mano, no se pudo re-correr
por el archivo bloqueado pero se rastreó la lógica): la fila con fecha correcta se sigue actualizando
en su lugar en cada corrida; la fila con fecha vieja/vacía se arrastra sin cambios porque nada en el
input actual vuelve a tener esa fecha. El conteo se queda estable en 2, no sube a 3/4/5. Es un
"fork que nunca se vuelve a unir", no una multiplicación descontrolada.

**Por qué solo pasa en efectivo, nunca en yape:** el reporte de yape (motor_matching) nunca se edita
a mano — es 100% generado. `mesa_N.xlsx` (efectivo) sí lo edita el cobrador día a día, y está sujeto
a errores humanos de tipeo (fecha vacía, fecha corregida después, etc.). El bug es estructural del
canal efectivo, no de yape.

### Por qué NO se puede simplemente sacar FECHA_COBRO de la clave
Primer intento de solución propuesto (sacar `FECHA_COBRO` de `_pres_key`, matchear solo por
`MESA+MZ+LT`) — **el usuario lo rechazó**: cobran en varios días distintos, así que dos reclamos
reales del mismo predio en fechas distintas del mismo mes son un caso legítimo que existe (aunque
hoy, verificado, 0 predios tienen ese caso). Sacar la fecha del todo arriesgaría fusionar dos
reclamos reales en uno. La solución tiene que distinguir "la fecha cambió porque es un evento nuevo"
de "la fecha cambió porque alguien la corrigió/completó después".

### Predios duplicados hoy (22) — para no tener que re-derivarlos
```
mesa_1: M-15, G-1, V-6
mesa_3: Q-16, D-5
mesa_4: G-12, N-3, Q-3, F1-10, F1-1, G-17, L-3, D-1, T-20, K-9, W-6, L-6, Z-13, P-4, B-6, Q-1, W-5
```
(query para reproducir: agrupar `reclamos_2026-07.xlsx` por `(MESA,MZ,LT)` y filtrar count>1)

---

## Lo que sigue — diseño pendiente (sesión Opus)

Instrucción del usuario, textual: *"El problema es que si alguien se equivoca en poner la fecha
mañana y corrió varias veces se duplica [...] La solución es validar que no haya duplicados
comparando con su input y si los hay eliminarlos. Ahora lo implementamos en el código, lo hacemos
su propio módulo. Automático eliminamos o preguntamos en consola, le hacemos su trazabilidad para
que quede trazado lo que se hizo. Cuál escala a agentic SaaS, cuál es lo profesional."*

Preguntas de diseño a resolver (en orden, cada una condiciona la siguiente):

1. **Qué significa "validar comparando con su input"** — ¿el detector de duplicados cruza
   `reclamos_2026-07.xlsx` contra `mesa_N.xlsx` (fuente original) para confirmar que hay un solo
   pago real detrás de las N filas encontradas? (así se confirmó G-12 a mano esta sesión). Definir
   el criterio exacto de "es un duplicado real" vs "son dos eventos legítimos" — probablemente:
   mismo `(MESA,MZ,LT)` + mismo `RECLAMO`/`COMENTARIO` + el input solo tiene 1 pago con ese comentario.
2. **Módulo propio o función dentro de `4b_reclamos`** — el usuario pidió explícitamente que sea
   su propio módulo. Definir nombre, dónde vive, cuándo se invoca (¿parte del pipeline de
   `4b_reclamos/main.py`, o un script aparte que se corre antes/después?).
3. **Automático vs pregunta en consola** — el usuario planteó ambas opciones sin decidir. Pensar el
   trade-off: automático es más rápido pero arriesga borrar un caso legítimo no previsto; preguntar
   es más seguro pero no escala si hay muchos duplicados. Podría ser automático solo cuando el
   criterio de "duplicado real" es inequívoco (input confirma 1 solo pago) y preguntar en los casos
   ambiguos.
4. **Trazabilidad** — registrar qué se eliminó, cuándo, por qué criterio, para poder auditar después
   (mismo patrón que ya usa el módulo en `4b_reclamos/trazabilidad/`).
5. **"Cuál escala a agentic SaaS, cuál es lo profesional"** — el usuario quiere que la decisión de
   arquitectura tenga en cuenta no solo el fix puntual de jass_system sino qué patrón sería el
   correcto si esto se convirtiera en un producto (SaaS) usado por otras JASS. Vale la pena pensar
   en el patrón como "reconciliación con detección de duplicados + journal de decisiones" en vez de
   un parche ad-hoc — pero evaluar sin sobre-ingeniería (Regla del Tres, no diseñar para
   hipotéticos — ver `feedback_no_sobreingenieria_edge_case_raro` en memoria).

**No empezar a codificar sin antes proponer el diseño en consola y esperar aprobación** (Regla 2 de
CLAUDE.md — cambio a un módulo existente/nuevo con dimensión de arquitectura).

---

## Estado de archivos al cerrar

```
4b_reclamos/main.py                 M   fix del Bug 1 aplicado, SIN commitear (esperar Bug 2)
4b_reclamos/outputs/reclamos_2026-07.xlsx   regenerado (gitignored, no aparece en git status)
4b_reclamos/backup/reclamos/*.xlsx  varios backups automáticos de hoy (gitignored)
docs/aprendizaje/Aprendizaje html/placeholder_disfrazado_de_manual_20260710.html   nuevo, sin commitear (?? en git status)
docs/retomar/RETOMAR_reclamos_duplicados_2026-07-10.md   este archivo
```

## SIGUIENTE_ACCION — orden sugerido

1. **[Opus]** Resolver las 5 preguntas de diseño de la sección "Lo que sigue" con el usuario,
   proponer en consola, esperar aprobación.
2. **[Opus/Sonnet]** Implementar el módulo de validación de duplicados aprobado.
3. **[Sonnet]** Limpiar las 22 filas duplicadas actuales de `reclamos_2026-07.xlsx` con el módulo
   nuevo (no a mano — así se prueba el módulo con el caso real que lo motivó).
4. **[Sonnet]** Decidir mesa_4 A-1/D-6/S-8 (reclasificados a compromiso/exoneracion, siguen
   arrastrados como reclamo activo) — pendiente suelto del Bug 1, no bloquea nada.
5. **[Sonnet]** Commitear en bloque: fix Bug 1 + módulo de duplicados + limpieza de datos + doc de
   aprendizaje, una vez todo validado junto.
