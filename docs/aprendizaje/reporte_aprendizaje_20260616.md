# Reporte de Aprendizaje — 16 Junio 2026

---

## Lo que se construyó esta sesión

- **MESA + COBRADOR** agregados a `lista_corte.xlsx` — trazabilidad del pago efectivo desde `pagos_efectivo.xlsx`
- **Estado EXONERADO** en `registro_cortes.xlsx` — tercer estado de exclusión para locales comunales y casos especiales
- **Reconciliación bidireccional** diseñada para `aplicar_penalidad.py` v2 — SET_DEBE vs SET_TIENE
- **Verificación de fuente** de `pagos_efectivo.xlsx` — confirmado que usa solo MONTO_EFECTIVO

---

## ¿Cómo lo descubrimos? — La historia del descubrimiento

### Historia 1: El error de los "22 pagos parciales" — y la verificación que lo resolvió

**El setup:** El usuario pidió agregar MESA + COBRADOR a `lista_corte.xlsx`. Antes de codificar, exploré `pagos_efectivo.xlsx` y `planilla_cobrado.xlsx` para entender el cruce.

**El supuesto incorrecto:** Calculé cuántos de los 40 usuarios en lista_corte tenían pagos en `pagos_efectivo.xlsx`. Encontré 22. Concluí que "22 de 40 pagaron parcialmente este ciclo" — y lo reporté como un hallazgo preocupante.

**El momento de quiebre:** El usuario interrumpió:
> *"No te entiendo, yo solo te pedí que agregues la mesa donde se hizo el cobro y el cobrador. Pero me sales con que 22 de los 40 pagaron parcial."*

Y luego explicó algo clave que yo no sabía:
> *"el yape no se toca porque eso se incluye en el reporte de pagos_yape_tepago. Sale en mesa el pago yape en MONTO_YAPE como un informativo nada más."*

**La corrección:** Escribimos `verificar_pagos2.py` — un script de exploración que buscó los 6 usuarios con pago mixto (efectivo > 0 Y yape > 0 en los registros de mesa) y comparó el MONTO en `pagos_efectivo.xlsx` contra MONTO_EFECTIVO vs MONTO_TOTAL.

Resultado para los 4 casos que SÍ aparecían en pagos_efectivo:
- K-7: efe=19, total=49 → pe.MONTO=**19** (= EFE correcto)
- F1-2: efe=8, total=33 → pe.MONTO=**8** (= EFE correcto)
- O-24: efe=66, total=266 → pe.MONTO=**66** (= EFE correcto)
- O-23: efe=16, total=216 → pe.MONTO=**16** (= EFE correcto)

Mi análisis anterior estaba equivocado. Los "22 pagos parciales" no existían.

**Por qué importa:** Si hubiéramos diseñado la lógica de cruce asumiendo que pagos_efectivo podía tener el total, habríamos descartado usuarios válidos o clasificado mal los pagos. Un script de verificación de 50 líneas antes de diseñar evita un error de arquitectura.

---

### Historia 2: El audit con 42 entradas cuando la lista tiene 38 — y el patrón que nació de eso

**El setup:** El usuario reportó que `audit_penalidad.xlsx` parecía desactualizado respecto a `lista_corte.xlsx`. Corrí una verificación.

**El supuesto incorrecto:** Pensé que "desactualizado" significaba que faltaban entradas en el audit — que no se habían aplicado penalidades nuevas. Pero al leer el audit, tenía **42 entradas** para 2026-06, mientras lista_corte tiene 40 usuarios (38 con EJECUTAR=SI).

**El momento de quiebre:** Ví los nombres en las 42 entradas:

```
A   4   YOLANDA ESPINOZA JAIMES   2026-06   S/ 20
A  11   FELIX POZO CASTRO         2026-06   S/ 20
D1  1   FAUSTINA POMALLANQUI      2026-06   S/ 20
X  31   OLINDA LOPEZ VERDE        2026-06   S/ 20
```

Esos 4 son exactamente los usuarios en `registro_cortes.xlsx` con ESTADO=CORTADO. La penalidad se les aplicó **antes** de que implementáramos la exclusión de cortados. El audit guardó la historia de lo que pasó, y esa historia ya no coincidía con el estado actual del sistema.

También aparecían locales comunales como PUESTO DE SALUD, CAPILLA, CENTRO EDUCATIVO — que el usuario luego confirmaría como EXONERADOS.

**La pregunta del usuario que lo cambió todo:**
> *"Lo ves profesional el cargar y revertir cambio y si no lo es como podríamos solucionarlo?"*

**La corrección — patrón de reconciliación bidireccional:**
En vez de un script separado de reversión (que el usuario tendría que recordar correr), diseñamos que `aplicar_penalidad.py` compara en cada corrida:
- **SET_DEBE** = quiénes deben tener penalidad (lista_corte EJECUTAR=SI)
- **SET_TIENE** = quiénes ya la tienen en el audit (APLICADO menos REVERTIDO)

La diferencia en ambas direcciones se maneja sola:
- `DEBE - TIENE` → aplica nuevos
- `TIENE - DEBE` → revierte sobrantes (los 10 casos de hoy)
- `DEBE ∩ TIENE` → skip

**Por qué importa:** Es la diferencia entre un sistema que "no duplica" y un sistema que "siempre está correcto". El primero protege contra un error en una dirección. El segundo no importa cuántas veces se corra ni cuánto cambie la lista — siempre converge al estado correcto.

---

### Historia 3: El estado EXONERADO — cuando CORTADO no es la palabra correcta

**El setup:** El usuario mencionó:
> *"registro_cortes considera que los exonerados no deben salir tampoco en corte ya que están exonerados su consumo por ser muchas veces locales del pueblo o enfermedad del usuario."*

**El supuesto incorrecto:** El sistema tenía dos estados: CORTADO y REACTIVADO. Mi primera reacción fue "puedo manejar esto con el estado CORTADO existente". Pero hay una diferencia fundamental:

- **CORTADO** = el usuario puede pagar agua pero no lo hizo — se le cortó el servicio como consecuencia
- **EXONERADO** = el usuario no paga agua por su naturaleza (local comunal, enfermedad) — nunca debió estar en el flujo

Usar CORTADO para EXONERADO contamina el historial: un PUESTO DE SALUD aparecería como "cortado por deuda" cuando nunca tuvo deuda.

**El momento de quiebre:** Al leer `registro_cortes.xlsx`, el supervisor ya había cargado 4 entradas con ESTADO=EXONERADO:
- J-6 CAPILLA DEL PUEBLO
- J-14 CENTRO EDUCATIVO 20902
- Z-6 AREA COMUNAL DEL PUEBLO
- C-21 PUESTO DE SALUD TUPAC AMARU

El supervisor ya sabía que necesitaba un estado distinto. El código todavía no lo soportaba.

**La corrección:** EXONERADO como tercer estado con semántica propia. El código excluye CORTADO y EXONERADO por igual, pero el historial diferencia por qué cada uno está excluido. Cuando se revise en 6 meses, se sabe exactamente: este fue cortado por deuda, ese fue exonerado porque es la capilla del pueblo.

**Por qué importa:** Los estados en un sistema persistente son lenguaje. Si el lenguaje no refleja la realidad del negocio, el historial acumula ambigüedad. En 2 años, alguien va a leer el registro y necesita entender de un vistazo por qué cada usuario está excluido — sin llamar a nadie para preguntar.

---

### Historia 4: MESA + COBRADOR — la pregunta que nadie podía responder

**El setup:** El usuario describió un problema operacional concreto:
> *"me dicen tal MZ pagó pero no sé en qué mesa pagó y quién cobró. Más adelante va a ser fundamental. Ahí estamos ciegos."*

**El supuesto incorrecto (del sistema, no nuestro):** El sistema registraba SI o NO había pago, pero perdía la trazabilidad de origen. Para auditorías, reclamos o disputas, saber que "MZ B, LT 12 pagó S/ 20" no es suficiente — el supervisor necesita saber "pagó en mesa_2, lo cobró Yerald Romero el día X".

**El momento de quiebre:** No fue un error técnico — fue un dolor operacional real. Alguien en campo dice "yo pagué" y la JASS no puede verificar ni dónde ni quién recibió ese pago.

**La solución:** `pagos_efectivo.xlsx` ya tenía columnas MESA y COBRADOR por fila — construidas cuando el módulo 4_pagos consolidó los registros de mesa. Solo faltaba cruzar ese archivo por (MZ, LT) y agregar las columnas a `lista_corte`.

**El resultado:** 15 de los 40 usuarios en lista_corte tienen pago efectivo registrado con su mesa y cobrador. Para los 25 restantes, las columnas quedan vacías — confirma que no hay registro de pago efectivo para ellos en este ciclo.

**Por qué importa:** El sistema ya tenía los datos. Lo que faltaba era exponerlos en el lugar donde alguien los necesita. La trazabilidad no es costosa de agregar cuando los datos de origen están bien diseñados.

---

## Términos técnicos aprendidos

### Reconciliación bidireccional

**Qué es:** Patrón de diseño para scripts que aplican cambios a un archivo compartido. En vez de solo "no duplicar" (idempotencia unidireccional), compara dos sets y corrige en ambas direcciones.

```
SET_DEBE = quiénes DEBEN tener el cambio aplicado ahora
SET_TIENE = quiénes YA lo tienen (según audit log)

NUEVOS    = DEBE - TIENE  →  aplicar
SOBRANTES = TIENE - DEBE  →  revertir
CORRECTOS = DEBE ∩ TIENE  →  skip
```

**Cuándo usarlo:** Cuando la lista de beneficiarios puede cambiar entre corridas y se necesita que el archivo compartido siempre refleje el estado actual de esa lista.

**Dónde aparece:** `aplicar_penalidad.py` v2 — los 10 cargos erróneos (4 CORTADOS + 4 EXONERADOS + 2 con reclamo bloqueante) se revertirán automáticamente al re-correr, sin intervención manual.

---

### Estado persistente con semántica de negocio

**Qué es:** Un archivo de estado persistente (como `registro_cortes.xlsx`) donde cada valor en la columna ESTADO tiene una semántica de negocio distinta — no solo "activo/inactivo" sino el motivo real de cada estado.

| Estado | Semántica | Efecto en el sistema |
|--------|-----------|---------------------|
| CORTADO | Deuda impaga — servicio físicamente cortado | Excluido de lista_corte |
| EXONERADO | Local comunal o caso especial — nunca debió pagar | Excluido de lista_corte |
| REACTIVADO | Volvió al flujo normal (pagó deuda o terminó exoneración) | Incluido normalmente |
| EJEMPLO | Fila guía para el supervisor — ignorada por el código | Ignorada |

**Por qué importa:** El estado almacena una decisión, no solo un valor. Esa decisión tiene una razón de negocio que debe ser legible en el historial, no solo inferible del código.

---

## Lo que hiciste bien — nivel profesional

- **Interrumpiste cuando algo no tenía sentido:** Cuando el análisis derivó hacia "22 pagos parciales", paraste y reorientaste: "no te entiendo, yo solo te pedí la mesa y el cobrador". No dejaste que el análisis innecesario bloqueara la tarea real.

- **Distinguiste semántica de conveniencia:** Podrías haber dicho "usa CORTADO para los exonerados también, total excluye igual". En cambio, insististe en que los locales comunales merecen su propio estado porque la razón de exclusión importa para el historial.

- **Preguntaste por el diseño, no solo por la solución inmediata:** "Lo ves profesional el cargar y revertir cambio?" es una pregunta de arquitectura, no de implementación. Antes de codificar la solución fácil (script de reversión manual), preguntaste si el diseño era correcto.

- **Ya tenías el archivo lleno antes del código:** El supervisor cargó los 4 EXONERADOS en `registro_cortes.xlsx` antes de que el código los soportara. Eso es dominio anticipando la herramienta — el negocio iba adelante de la implementación.

---

## Errores cometidos — y lo que revelan

### Asumir que pagos_efectivo tenía el MONTO total

**Qué pasó:** Sin verificar la fuente, concluí que pagos_efectivo podría estar sumando efectivo + yape y lo reporté como hallazgo.

**Lo que revela:** En sistemas donde múltiples archivos manejan montos relacionados (efectivo, yape, total), siempre verificar qué contiene exactamente cada uno antes de cruzarlos. La semántica del nombre del archivo (`pagos_efectivo`) no es suficiente — puede haber matices.

**Corrección aplicada:** Script de verificación que compara el MONTO en pagos_efectivo contra MONTO_EFECTIVO y MONTO_TOTAL de los registros de mesa para casos mixtos. Confirmó que el sistema es correcto.

---

### Agregar PAGOS_YAPE_PATH a config sin confirmación

**Qué pasó:** Al agregar PAGOS_EFECTIVO_PATH al config, también agregué PAGOS_YAPE_PATH "por completitud" — sin que el usuario lo pidiera.

**Lo que revela:** En sistemas donde cada módulo tiene responsabilidades bien definidas, agregar rutas a archivos de otros módulos sin necesidad real crea acoplamiento innecesario. El Yape no pertenece a 6_corte.

**Corrección aplicada:** PAGOS_YAPE_PATH removido de config.py.

---

## Pendientes

1. Implementar `aplicar_penalidad.py` v2 con reconciliación bidireccional (SET_DEBE vs SET_TIENE + columna ACCION en audit)
2. Re-correr `generar_lista.py` para excluir los 4 EXONERADOS del archivo
3. Re-correr `aplicar_penalidad.py` v2 para revertir los 10 cargos erróneos en planilla_mes
4. Actualizar `formato_audit_penalidad.html` con la nueva columna ACCION

---

## Resumen

Sesión centrada en trazabilidad y corrección de estado. El punto de partida fue una petición simple: "agrega MESA y COBRADOR a lista_corte". Eso derivó en descubrir que la fuente (pagos_efectivo) solo registra MONTO_EFECTIVO — verificado con script exploratorio. Al revisar el audit, aparecieron 42 entradas cuando debería haber 38, revelando que 4 CORTADOS y 4 EXONERADOS fueron cargados antes de que el sistema los excluyera. Esto llevó al diseño del patrón de reconciliación bidireccional: en vez de un script de reversión manual, `aplicar_penalidad.py` comparará SET_DEBE vs SET_TIENE en cada corrida y corregirá en ambas direcciones automáticamente. El aprendizaje central: un sistema bien diseñado no necesita scripts de corrección adicionales — el mismo script que aplica también revierte, porque sabe exactamente qué debería estar y qué no.
