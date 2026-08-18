# jass_system: instrucciones para agentes

## Arranque obligatorio

```text
AGENTS.md
  -> docs/decisiones/estado_decisiones.md
  -> LEER_ANTES.md (si existe)
  -> ultimo RETOMAR relevante al alcance
  -> README.md raiz + README.md del modulo
  -> entrypoint/config/lectores exactos que participan
```

- Lee tambien `docs/lente_escala.md` antes de recomendar arquitectura: el destino es
  multi-tenant (`JASS_ID`), config-driven, PostgreSQL + Docker; politica distinta no
  justifica duplicar motores o modulos.
- Declara mentalmente alcance, decisiones `CERRADAS`, `ABIERTAS` y `BLOQUEADAS`.
  Un `RETOMAR` conserva estado operativo, pero no cambia decisiones cerradas.
- `LEER_ANTES.md` documenta eventos activos que atraviesan modulos. Sus pasos y orden
  prevalecen sobre flujos mensuales normales hasta que el propio archivo los cierre.

## Barrera de cambios

- No edites codigo sin una instruccion explicita que identifique cambio y alcance.
  Antes de editar ejecuta el preflight indicado por el entorno con el `TargetPath`
  exacto y lee todos los archivos `CONTEXT` que reporte.
- Una decision solo pasa de `ABIERTA` a `CERRADA` con confirmacion explicita del
  usuario. No edites `docs/decisiones/estado_decisiones.md` para cerrarla por cuenta
  propia.
- `docs/decisiones/guard_codigo.json` bloquea targets mientras su estado sea
  `CERRADO`; no modifiques ni desbloquees esos archivos.
- Si el usuario pide correr, regenerar o mostrar un output, ejecuta primero el codigo
  existente. Una diferencia entre fuente y salida se informa e investiga; no autoriza
  cambiar reglas, inputs, generadores ni formato.
- Ante un pedido ambiguo o que implique ejecutar/generar algo y no sea 100% inequivoco,
  entra en Plan Mode antes de actuar: redacta el prompt/plan exacto que vas a correr,
  espera aprobacion explicita del usuario, recien entonces ejecuta. No alcanza con
  preguntar "lo hago?" en texto suelto -- usa el mecanismo de Plan Mode para que el
  usuario vea y valide el plan antes de la ejecucion.

## Arquitectura real

```text
0_padron -> 1_lecturas -> 2_planilla -> 3_boletas
         -> [4_pagos + 4b_reclamos]
         -> [5_cobranza + 5b_validacion]
         -> [6_corte + 6b_corte_multas] -> 7_cierre

libro_mayor/ = substrato permanente, no una etapa numerada
```

- Los modulos numerados son independientes y poseen `inputs/`, `outputs/`, `config.py`,
  entrypoints y README propios. Limita lecturas y pruebas al modulo afectado y a sus
  consumidores reales.
- El codigo operativo actual es pre-ledger. `libro_mayor/` y la disolucion futura de
  modulos descrita en los README siguen pendientes; no confundas specs de destino con
  comportamiento implementado. Si prosa y codigo difieren, verifica el flujo ejecutable.
- Sigue siempre `fuente -> lector -> transformacion -> calculo -> writer -> reporte`.
  Un nombre parecido, un `SOURCE` tecnico o un output derivado no determina la fuente
  de verdad.
- `shared/` contiene recursos entre modulos y archivos persistentes. Antes de tocar un
  primitivo o writer compartido, localiza todos sus importadores/consumidores.

## Ejecucion y pruebas

- No existe manifiesto, lockfile, CI ni configuracion global de pytest en la raiz.
  `README.md` menciona `requirements.txt`, pero el archivo no existe actualmente; no
  presentes `pip install -r requirements.txt` como setup verificado.
- Ejecuta desde la raiz con UTF-8 cuando el README/RETOMAR no exija otro contexto:
  `py -u -X utf8 MODULO/main.py`. Los flujos con varios scripts y pasos humanos se
  toman del README del modulo; no asumas que `main.py` cubre todo.
- No hay suite global confiable. Verificacion enfocada: `py -m pytest ruta/test.py -q`
  o `py ruta/test.py` cuando el test define setup en `main()` y la documentacion lo
  declara standalone. Inspecciona el test/conftest antes de elegir.
- Tests y scripts pueden escribir `.xlsx`. Deben redirigir toda ruta a temporales;
  compara `git status --short` antes y despues. Nunca uses datos reales como fixtures.
- Tras un fix, ejecuta el caso afectado y uno no afectado; luego ejecuta los consumidores
  de cada output modificado. Para cambios en `shared/`, verifica todos sus consumidores.

## Integridad del dominio

- Ledger y trazabilidades son append-only: corrige con eventos inversos auditables,
  no borrando filas. Respeta writers unicos y preserva backups/decisiones humanas al
  regenerar Excel.
- No fuerces `PAGO`, `SALDO`, `TOTAL_PAGADO` ni repartos para coincidir con un resultado
  esperado. Prohibidos overrides por usuario, lote, fecha, PDF o caso puntual.
- Una correccion señalada por MZ, lote o registro aplica al bloque involucrado, salvo que
  el usuario autorice expresamente una excepcion. Valida el caso citado y otro del bloque.
- No reviertas ni limpies cambios ajenos: el worktree suele contener datos y artefactos
  operativos no rastreados.
