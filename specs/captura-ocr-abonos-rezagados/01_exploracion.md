# Exploración: Captura OCR auditada de abonos rezagados

## Sistema actual

### Flujo ejecutable observado

```text
declaración o comprobante externo
        |
        v
transcripción humana sin archivo original enlazado
        |
        v
shared/abonos_rezagados.xlsx
  Abonos_Raw -> Mapa_Abonos -> Categorias
        |
        +-> shared/reclasificar_abonos.py
        |      clasificaciones escritas en código
        |
        v
5_cobranza/inputs/abonos_rezagados_manifest_2026-08.json
        |
        v
5_cobranza.main._validar_abonos_manifest()
  comparación exacta MZ + LT + MONTO + MES_CICLO + MES_ANO_APLICA
        |
        v
_cargar_abonos_rezagados() -> agrupación por predio y ciclo
        |
        v
_aplicaciones_por_fuente() -> reparto agregado por concepto
        |
        v
snapshot_ledger_YYYY-MM.json -> 5b_validacion -> 7_cierre
        |
        v
shared/seguimiento_pueblo.xlsx
```

### Fuente y clasificación

- `shared/abonos_rezagados.xlsx` declara un writer humano. No existe un importador de
  imágenes ni una bandeja documental antes de este archivo.
- `Abonos_Raw` tiene 43 filas. Sus columnas registran predio, monto, balde, canal,
  fecha real, ciclo, canal de origen, persona que retuvo el dinero, evidencia textual,
  mes de aplicación, motivo y respaldo.
- Las 43 filas tienen texto en `EVIDENCIA`; dos no tienen valor en `RESPALDO`. El texto
  describe la evidencia, pero no referencia un archivo original mediante hash o ID.
- La distribución actual de `Mapa_Abonos` es: 20 `CONFIRMADO`, 13
  `YA_RESUELTO_MANUAL`, 6 `REVISAR` y 4 `BLOQUEADO`.
- `shared/reclasificar_abonos.py` contiene conjuntos de claves escritos en código. Al
  evaluar las 43 filas actuales sin escribir el Excel, 35 fueron clasificables y 8
  produjeron `Fila sin clasificación`; el mapa no es reproducible íntegramente con el
  script presente.

### Manifest y control previo al cálculo

- El manifest actual contiene 28 entradas: 24 `CONFIRMADO` y 4 `BLOQUEADO`, sin claves
  duplicadas según su contrato.
- Para `2026-08`, fuente y manifest sí coinciden: 20 filas en ambos, cero inesperadas,
  cero confirmadas faltantes y cero bloqueadas activas.
- Para `2026-07`, la fuente conserva 22 filas históricas mientras el manifest tiene 8;
  hay 14 filas no manifestadas y 4 bloqueadas activas. Esto no rompe agosto porque el
  lector filtra por `MES_ANO_APLICA`, pero el manifest no es un registro canónico global.
- `_validar_abonos_manifest()` falla cerrado ante una fila activa inesperada, faltante,
  repetida o bloqueada. Es el control de integridad más fuerte del flujo actual.

### Transformación y pérdida de identidad

- `_cargar_abonos_rezagados()` suma las filas por `(MZ, LT)` y separa únicamente monto
  de ciclo cerrado y monto de ciclo vigente.
- `_aplicaciones_por_fuente()` combina esos totales con la deuda y obtiene cuánto se
  aplica a cada concepto. Desde ese punto ya no existe relación 1:1 con `Abonos_Raw`.
- `_objetivos_ledger()` produce objetivos por `(source, MZ, LT, concepto)`. El snapshot
  tiene hash criptográfico, pero no `EVIDENCIA_ID` ni `ABONO_ID` por pago.
- `7_cierre` solo compromete el snapshot cuando el mismo hash fue validado por
  `5b_validacion`. El commit es idempotente a nivel de snapshot y objetivo agregado.
- En el ledger transitorio se observaron 19 eventos con `SOURCE=abonos_rezagados`: 18
  `ABONO_REZAGADO` y una `CORRECCION_SISTEMA`, por S/668 aplicados. Todos tienen
  `AUDIT_REF`, pero representan porciones por concepto, no el movimiento completo ni
  el comprobante original.

### Consumidores adicionales

- `4b_reclamos/herramienta/comun.py::_abonos_rezagados_predio()` lee directamente
  `Abonos_Raw` para sumar dinero por predio y mes en reportes.
- `6_corte/generar_lista.py` consume `_cargar_abonos_rezagados()` para considerar estos
  importes en la decisión operativa de corte.
- Los tests y mini-pipelines de `5_cobranza` copian el Excel y el manifest para proyectar
  casos sin escribir el ledger real.

### Arquitectura objetivo ya decidida

```text
HECHO DE DINERO                 INTERPRETACIÓN
libro_mayor/caja               estado_cuenta
MovimientoCaja / ABONO_ID ---> Aplicacion / CARGO_ID
          \________________ motor único ______________/
```

- `libro_mayor/` es el bounded context acordado para caja y estado de cuenta; el nuevo
  system-of-record no pertenece a `shared/`.
- El contrato de caja exige `JASS_ID`, céntimos enteros, writer único, eventos
  append-only y `ABONO_ID` determinista.
- Para efectivo, la identidad acordada se ancla a procedencia estable, no a fecha, monto
  o predio editables. El contrato ya contempla un gatillo humano de casi-duplicado.
- El ledger objetivo todavía no está implementado. `abonos_rezagados.xlsx` es un
  precursor durable y el futuro adapter deberá convertir cada abono nuevo en
  `registrar_movimiento`, sin duplicar el motor de aplicación.

## Hallazgos

| ID | Hallazgo | Consecuencia observable |
|---|---|---|
| H-01 | El control exacto comienza después de la transcripción al Excel. | Una foto repetida o una segunda transcripción con campos distintos puede entrar a la fuente antes de ser detectada. |
| H-02 | `EVIDENCIA` es texto libre, no una entidad documental. | No se demuestra integridad del original ni se detecta el mismo archivo renombrado. |
| H-03 | No hay registro de extracción OCR. | No se conserva propuesta, confianza, versión, errores ni correcciones humanas. |
| H-04 | `SITUACION` no guarda historial de decisiones. | No consta estructuradamente quién confirmó o rechazó, cuándo, qué cambió y por qué. |
| H-05 | La clave del manifest es control de contenido, no identidad de pago. | No enlaza imagen, candidato, decisión, movimiento y confirmación. |
| H-06 | El cálculo agrega por predio y concepto. | El ledger transitorio no puede confirmar si un abono individual específico ya fue cargado. |
| H-07 | El hash del snapshot confirma una proyección completa. | Demuestra integridad del lote, no procedencia individual de cada comprobante. |
| H-08 | El clasificador no reconstruye 8 de 43 filas. | Excel, script y decisiones humanas ya divergen. |
| H-09 | Agosto tiene fuente y manifest alineados exactamente. | La transición debe preservar este fail-closed hasta aportar evidencia equivalente. |
| H-10 | Caja oficial permanece bloqueada por una diferencia de S/700. | El MVP debe probarse en sombra; no hay autorización de cutover. |

### Distinciones necesarias

```text
archivo recibido != extracción OCR != candidato revisado
                 != movimiento registrado != aplicación a deuda
                 != confirmación de cierre
```

- Duplicado de archivo y duplicado de pago no son lo mismo. Dos imágenes distintas
  pueden mostrar la misma operación; una imagen puede contener más de una operación.
- Aprobar que los campos fueron bien leídos no demuestra que el dinero existe o que no
  fue registrado previamente.
- Registrar un movimiento en caja y aplicarlo a conceptos son responsabilidades distintas.
  El OCR no debe decidir la cascada de deuda.

## Restricciones

### Decisiones y contratos cerrados

- D-001: la referencia de pago confirma si el pago existe. Sin referencia no se reparte
  en el historial.
- D-002: un `SOURCE` técnico no constituye evidencia de pago.
- D-005: la cuenta completa empieza en agosto de 2026 sin backfill anterior.
- Caja y estado de cuenta forman un solo bounded context con consistencia fuerte y un
  único motor de aplicación.
- `JASS_ID` participa en identidades y eventos. Política por tenant entra por
  configuración; no se crean forks por JASS.
- El dinero nuevo se representa en céntimos enteros. El Excel con `float` es adapter.
- Correcciones de evidencia, decisiones o ledger se expresan como eventos posteriores;
  no se borran ni sobrescriben hechos auditables.

### Compatibilidad y migración

- `5_cobranza`, `4b_reclamos` y `6_corte` consumen hoy el Excel o su loader. El MVP no
  puede cambiar su contrato silenciosamente.
- El registro canónico debe existir antes del adapter legacy. La bandeja OCR no puede
  escribir directamente `abonos_rezagados.xlsx`.
- Caja oficial se alimenta después del cierre, no en caliente. Antes puede existir
  captura, revisión y staging, pero `CONFIRMADO_EN_CAJA` debe corresponder al commit real.
- Agosto sigue abierto, el snapshot no está validado y `7_cierre --confirmar` no está
  autorizado. Las pruebas usarán temporales, datos sintéticos o modo sombra.

### Seguridad y operación

- Los comprobantes contienen datos personales y financieros. No deben salir del entorno
  autorizado sin decisión explícita sobre proveedor, retención y tratamiento.
- El OCR nunca puede aprobar ni publicar dinero. El revisor humano es obligatorio.
- Una coincidencia aproximada debe bloquear o alertar para revisión; no descartar
  evidencia automáticamente.
- Reintentos, reinicios y doble clic del operador deben ser idempotentes.
- No se afirmará precisión OCR ni umbrales de confianza antes de medir imágenes reales.
- Esta fase no selecciona OCR local, nube, híbrido, interfaz ni store; esas alternativas
  corresponden a `03_opciones.md`.

## Incertidumbres

- El 2026-08-22 se incorporó una primera muestra de 8 JPEG manuscritos en
  `shared/abono_rezagado/`. La extracción asistida produjo 30 candidatos: 23 con
  confianza alta, 4 media y 3 baja; una fila no contiene monto identificable.
- Una imagen contiene hasta 10 candidatos, por lo que la relación evidencia 1:N queda
  confirmada para esta muestra.
- Doce candidatos recibieron alertas por repetición entre imágenes, coincidencia o mismo
  predio en `shared/abonos_rezagados.xlsx`. Esto confirma que duplicado documental y
  posible duplicado de pago requieren controles separados y revisión humana.
- El borrador verificable está en
  `shared/abono_rezagado/abonos_rezagados_ocr_borrador.xlsx`; tiene hashes SHA-256 para
  las 8 imágenes, enlaces al original y 30 decisiones humanas en `PENDIENTE`. No fue
  importado al manifest ni al ledger.
- La muestra todavía es pequeña y solo contiene notas manuscritas; no permite medir
  precisión para capturas bancarias, PDFs, reenvíos, recortes o recompresiones variadas.
- Falta definir qué referencia de negocio satisface D-001 para cada canal y tipo de abono.
- El capítulo didáctico `Ingenioro de IA/07_caso_abonos_rezagados.html` propone evidencia
  1:N candidatos y conservación permanente. La cardinalidad 1:N ya tiene evidencia; la
  retención permanente sigue siendo una hipótesis hasta decidir la política aplicable.
