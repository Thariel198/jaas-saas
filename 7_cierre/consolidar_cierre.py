"""7_cierre/consolidar_cierre.py — Commit único y transición de período.

7_cierre NO genera arrastres (ya los hace 5_cobranza) ni copia nada a
2_planilla/inputs (2_planilla lee en vivo, Opción A). Su único trabajo es
transicionar el mes:

  PASO 0 PREPARAR  · corre la proyección final de 5_cobranza y 5b_validacion
  PASO 1 GATE      · verifica el hash exacto del snapshot validado
  PASO 2 COSECHAR  · copia BALDE 2 (canónicos + fuentes de pago) a archivo/{mes}/
  ── seguro ── de acá para abajo es irreversible, requiere --confirmar + "SI" ──
  PASO 3 COMMIT    · aplica el snapshot al ledger en un batch atómico e idempotente
  PASO 4 FREEZE    · estado_ciclo[mes].estado = CERRADO
  PASO 5 LIMPIAR   · reset a template SOLO las fuentes manuales (mesa_*, correcciones)
                     — solo resetea lo que YA se verificó cosechado en el PASO 2

Sin --confirmar: prepara, valida y cosecha, pero no compromete el ledger ni cierra.
Con --confirmar: además pide escribir "SI" (consentimiento humano explícito)
antes de sellar y resetear.

No hace `git commit` — persistir es un paso separado (imprime el comando).
Ver README.md y docs/diagrama_consolidador_cierre.html.

Uso:
    python consolidar_cierre.py --mes 2026-06                # dry-run del cierre
    python consolidar_cierre.py --mes 2026-06 --confirmar     # cierre real
"""
import logging
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import config

sys.path.insert(0, str(Path(__file__).parent.parent / "shared"))
import utils_estado_ciclo as repo_estado  # noqa: E402
import seguimiento_repo as repo            # noqa: E402
import utils_lote                          # noqa: E402
import utils_templates                     # noqa: E402

log = logging.getLogger(__name__)


class CicloNoValidadoError(Exception):
    pass


def paso0_preparar(mes: str) -> None:
    """Regenera la proyección final y valida sus fuentes antes del gate."""
    activo = config.ciclo.activo(default=None, path=config.SHARED_DIR / "ciclo_activo.json")
    if activo != mes:
        raise CicloNoValidadoError(
            f"El cierre solicitado es {mes}, pero el ciclo activo es {activo or 'indefinido'}")
    subprocess.run([sys.executable, "-u", "-X", "utf8", str(config.COBRANZA_MAIN), "--force"],
                   cwd=config.ROOT.parent, check=True)
    subprocess.run([sys.executable, "-u", "-X", "utf8", str(config.VALIDACION_MAIN)],
                   cwd=config.ROOT.parent, check=True)


def _leer_snapshot(mes: str) -> tuple[dict, str]:
    ruta = config.snapshot_ledger_path(mes)
    if not ruta.exists():
        raise CicloNoValidadoError(f"Falta snapshot ledger para {mes}: {ruta}")
    documento = json.loads(ruta.read_text(encoding="utf-8"))
    snapshot_hash = documento.pop("snapshot_hash", "")
    normalizado = json.dumps(documento, ensure_ascii=True, sort_keys=True,
                             separators=(",", ":")).encode("utf-8")
    calculado = hashlib.sha256(normalizado).hexdigest()
    if calculado != snapshot_hash or documento.get("mes") != mes:
        raise CicloNoValidadoError(f"Snapshot inválido para {mes}: {ruta}")
    return documento, snapshot_hash


def paso1_gate(mes: str) -> tuple[dict, str]:
    snapshot, snapshot_hash = _leer_snapshot(mes)
    if not repo_estado.ciclo_validado(
            mes, ruta=config.ESTADO_CICLO_PATH, snapshot_hash=snapshot_hash):
        raise CicloNoValidadoError(
            f"El snapshot {snapshot_hash[:12]} de {mes} no está validado."
        )
    log.info(f"PASO 1 · GATE — snapshot {snapshot_hash[:12]} validado")
    return snapshot, snapshot_hash


def paso2_cosechar(mes: str) -> int:
    destino = config.archivo_mes_dir(mes)
    destino.mkdir(parents=True, exist_ok=True)

    fuentes = {
        **config.canonicos_a_cosechar(mes),
        **config.fuentes_manuales_a_resetear(),
        **config.fuentes_auto_a_cosechar(),
    }
    copiados = 0
    for nombre, origen in fuentes.items():
        if not origen.exists():
            log.warning(f"COSECHAR — no encontrado, se omite: {origen}")
            continue
        shutil.copy2(origen, destino / nombre)
        copiados += 1
    log.info(f"PASO 2 · COSECHAR — {copiados}/{len(fuentes)} archivos → {destino}")
    return copiados


def paso3_commit(mes: str, snapshot: dict, snapshot_hash: str) -> int:
    if repo_estado.ledger_comprometido(
            mes, ruta=config.ESTADO_CICLO_PATH, snapshot_hash=snapshot_hash):
        log.info(f"PASO 3 · COMMIT — {mes} ya estaba comprometido con este hash")
        return 0
    if repo_estado.ledger_comprometido(mes, ruta=config.ESTADO_CICLO_PATH):
        raise CicloNoValidadoError(
            f"El ledger de {mes} ya fue comprometido con un snapshot diferente")
    resultado = repo.reconciliar_objetivos_batch(
        mes, snapshot_hash, snapshot["objetivos"], snapshot.get("cargos", []))
    repo_estado.marcar_ledger_comprometido(
        mes, snapshot_hash, resultado["eventos"], ruta=config.ESTADO_CICLO_PATH)
    repo.generar_vista()
    repo.exportar_vista_pdf()
    log.info(f"PASO 3 · COMMIT — {resultado['eventos']} eventos aplicados")
    return resultado["eventos"]


def paso4_freeze(mes: str) -> None:
    repo_estado.marcar_cerrado(mes, ruta=config.ESTADO_CICLO_PATH)
    log.info(f"PASO 4 · FREEZE — estado_ciclo[{mes}].estado = CERRADO")


def _fuentes_listas_para_resetear(mes: str) -> dict[str, Path]:
    """Capa 2 del seguro: solo resetea lo que YA está confirmado en archivo/{mes}/.
    Si la cosecha de un archivo falló (PASO 2 lo omitió), NO se resetea — mejor
    dejarlo con datos viejos que perder el único registro que quedaba de junio."""
    destino = config.archivo_mes_dir(mes)
    listas = {}
    for nombre, ruta in config.fuentes_manuales_a_resetear().items():
        if (destino / nombre).exists():
            listas[nombre] = ruta
        else:
            log.warning(f"LIMPIAR — omitido (no hay cosecha confirmada): {nombre}")
    return listas


def paso5_limpiar(mes: str) -> None:
    for nombre, ruta in _fuentes_listas_para_resetear(mes).items():
        if nombre == "correcciones_lote.xlsx":
            utils_lote.escribir_correcciones_lote(ruta, [])
        else:
            utils_templates.crear_mesa_vacio(ruta)
        log.info(f"PASO 5 · LIMPIAR — reset a template: {nombre}")

    out_dir = config.COBRANZA_DIR / "outputs"
    borrados = 0
    for patron in config.PATRONES_BASURA:
        for f in out_dir.glob(patron):
            f.unlink()
            borrados += 1
    log.info(f"PASO 5 · LIMPIAR — {borrados} archivo(s) de basura borrado(s)")


def _pedir_consentimiento(mes: str) -> bool:
    """Capa 1 (interactiva) del seguro: muestra exactamente qué se va a resetear
    y exige escribir SI. --confirmar habilita el gatillo; esto es apretarlo."""
    listas = _fuentes_listas_para_resetear(mes)
    print(f"\nSe va a COMPROMETER el snapshot, SELLAR {mes}=CERRADO y RESETEAR:")
    for nombre in listas:
        print(f"  - {nombre}  (ya cosechado en archivo/{mes}/ — a salvo)")
    resp = input(f"\nEscribí SI para confirmar el cierre de {mes}: ").strip().upper()
    return resp == "SI"


def main(mes: str, confirmar: bool = False) -> None:
    config.OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)s  %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(config.OUTPUTS_DIR / "run.log", encoding="utf-8"),
        ],
        force=True,
    )
    print("=" * 60)
    print(f"  7_cierre/consolidar_cierre.py --mes {mes}" + ("  --confirmar" if confirmar else "  (dry-run)"))
    print("=" * 60)

    paso0_preparar(mes)
    snapshot, snapshot_hash = paso1_gate(mes)
    n_cosechados = paso2_cosechar(mes)

    if not confirmar:
        print(f"\n{n_cosechados} archivo(s) cosechados en archivo/{mes}/ — revisalos.")
        print(f"Snapshot validado: {snapshot_hash}")
        print("COMMIT, FREEZE y LIMPIAR NO corrieron (falta --confirmar).")
        print(f"Para cerrar de verdad: python consolidar_cierre.py --mes {mes} --confirmar\n")
        return

    if not _pedir_consentimiento(mes):
        print("Cancelado — no se selló ni se reseteó nada.")
        return

    n_eventos = paso3_commit(mes, snapshot, snapshot_hash)
    paso4_freeze(mes)
    paso5_limpiar(mes)

    print("\n" + "=" * 60)
    print(f"  Cierre de {mes} completado — {n_eventos} eventos · {n_cosechados} archivos")
    print("  Para persistir (paso manual/agente separado):")
    print(f"    git add 7_cierre/archivo/{mes}/ shared/reporte_acumulado_procesado/estado_ciclo.json")
    print(f"    git commit -m \"cierre ciclo {mes}\"")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    if "--mes" not in sys.argv:
        print("Uso: python consolidar_cierre.py --mes YYYY-MM [--confirmar]")
        sys.exit(1)
    mes_arg = sys.argv[sys.argv.index("--mes") + 1]
    main(mes_arg, confirmar="--confirmar" in sys.argv)
