# Intake: Captura OCR auditada de abonos rezagados

## Pedido original

Wilder solicita reemplazar el ingreso manual no auditado de
`shared/abonos_rezagados.xlsx` por un flujo profesional que permita:

- capturar comprobantes mediante OCR;
- evitar que una misma imagen o registro se procese dos veces;
- someter los datos extraídos a verificación humana;
- registrar el pago aprobado en el sistema sin duplicarlo;
- confirmar de forma auditable que el pago quedó cargado.

El alcance inicial confirmado es un MVP completo de punta a punta. Los comprobantes
llegarán principalmente por WhatsApp y una persona los descargará o arrastrará a una
bandeja de entrada. Un único revisor humano podrá corregir y aprobar cada extracción;
su identidad, decisión y momento deben quedar auditados.

Resultado observable esperado: para cada imagen debe poder demostrarse qué archivo se
recibió, qué leyó el OCR, qué corrigió o decidió el revisor, si se creó un pago y con qué
identidad se confirmó su presencia en el sistema.

## Evidencia disponible

### Hechos observados

- `shared/abonos_rezagados.xlsx` es la fuente operativa actual y declara que su writer
  es humano. Contiene las hojas `Abonos_Raw`, `Mapa_Abonos` y `Categorias`.
- La inspección del 2026-08-22 encontró 43 registros en `Mapa_Abonos`. El archivo guarda
  datos como predio, monto, ciclo, canal, retenido por, evidencia, situación y estado del
  ledger, pero no conserva el archivo original del comprobante ni una identidad de imagen.
- `shared/reclasificar_abonos.py` clasifica registros mediante conjuntos escritos en
  código: `CONFIRMADO`, `BLOQUEADO`, `YA_RESUELTO_MANUAL` y `REVISAR`.
- `5_cobranza/main.py::_validar_abonos_manifest` evita filas activas inesperadas,
  faltantes, duplicadas o bloqueadas comparando `MZ + LT + MONTO + MES_CICLO +
  MES_ANO_APLICA` contra un manifest.
- `5_cobranza/main.py::_cargar_abonos_rezagados` agrupa los abonos aprobados por predio
  y ciclo. El control actual comienza después de que los datos ya fueron transcritos al
  Excel; no demuestra el origen documental ni la revisión humana.
- En `shared/` no existe actualmente una colección de imágenes de comprobantes asociada
  a `abonos_rezagados.xlsx`; solo se encontró la vista PDF del estado de cuenta.
- El diseño objetivo de `jass_system` es multi-tenant, con `JASS_ID`, identidades
  deterministas, eventos append-only, PostgreSQL y Docker. El Excel actual debe tratarse
  como adapter transitorio, no como arquitectura final.
- Agosto de 2026 permanece abierto y la validación de caja tiene una diferencia conocida
  de S/700. Este cambio no autoriza cargar pagos reales, cerrar agosto ni ocultar esa
  diferencia.

### Supuestos por validar durante exploración

- Los comprobantes de WhatsApp tienen calidad y formatos suficientemente representativos
  para evaluar OCR con una muestra real.
- Una misma imagen puede reenviarse, renombrarse, recortarse o recomprimirse; por ello la
  duplicidad exacta y la similitud visual probablemente requerirán controles distintos.
- El revisor dispondrá de una identidad autenticable y de una interfaz local para revisar
  la imagen junto con los campos extraídos.
- La confirmación final deberá enlazar la captura con una identidad de abono o evento, no
  limitarse a marcar una celda como procesada.

## Preguntas iniciales

### Confirmado

- Alcance: MVP completo `captura -> OCR -> revisión -> registro -> confirmación`.
- Entrada inicial: descarga o arrastre manual de imágenes recibidas por WhatsApp.
- Autoridad operativa: un revisor humano puede corregir, aprobar o rechazar.
- Autoridad SDD: Wilder elige opciones y aprueba cada gate; el agente recomienda y aporta
  evidencia, pero no aprueba decisiones en su nombre.
- Restricción: no modificar pagos ni ledgers reales durante diseño o pruebas. La primera
  ejecución real requerirá autorización explícita y superar los controles vigentes.

### Decisiones abiertas para las fases de exploración y opciones

- Motor OCR: local, servicio externo o enfoque híbrido.
- Campos mínimos que constituyen un candidato de abono y cuáles puede corregir el revisor.
- Reglas para duplicado exacto, documento visualmente similar y posible pago duplicado.
- Estados, transiciones y tratamiento de reapertura, rechazo y corrección posterior.
- Almacenamiento inmutable del original, retención, acceso y protección de datos personales.
- Interfaz de revisión humana y mecanismo de autenticación del revisor.
- Integración transitoria con `abonos_rezagados.xlsx` y el manifest, sin acoplar el dominio
  nuevo a Excel.
- Evidencia necesaria para declarar `REGISTRADO` y luego `CONFIRMADO`, considerando que el
  libro mayor objetivo todavía no está implementado y la caja oficial sigue bloqueada por
  la diferencia de S/700.
- Muestra representativa de comprobantes reales que podrá usarse para medir precisión,
  duplicados y tiempo de revisión sin exponer datos fuera del entorno autorizado.
