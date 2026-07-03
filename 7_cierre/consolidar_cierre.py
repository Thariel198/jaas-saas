"""7_cierre/consolidar_cierre.py — Transición de período (freeze · foto · limpieza)

7_cierre NO genera arrastres (ya los hace 5_cobranza) ni copia nada a
2_planilla/inputs (2_planilla lee en vivo, Opción A). Su único trabajo es
transicionar el mes:

  PASO 1 GATE      · aborta si el ciclo no está validado (5b_validacion)
  PASO 2 COSECHAR  · copia BALDE 2 (canónicos + fuentes de pago) a archivo/{mes}/
  ── seguro ── de acá para abajo es irreversible, requiere --confirmar + "SI" ──
  PASO 3 FREEZE    · estado_ciclo[mes].estado = CERRADO
  PASO 4 LIMPIAR   · reset a template SOLO las fuentes manuales (mesa_*, correcciones)
                     — solo resetea lo que YA se verificó cosechado en el PASO 2

Sin --confirmar: corre GATE+COSECHAR (de solo lectura/copia, reversible) y
muestra qué haría FREEZE+LIMPIAR — no toca estado_ciclo.json ni resetea nada.
Con --confirmar: además pide escribir "SI" (consentimiento humano explícito)
antes de sellar y resetear.

No hace `git commit` — persistir es un paso separado (imprime el comando).
Ver README.md y docs/diagrama_consolidador_cierre.html.

Uso:
    python consolidar_cierre.py --mes 2026-06                # dry-run del cierre
    python consolidar_cierre.py --mes 2026-06 --confirmar     # cierre real
"""
import logging
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import config

sys.path.insert(0, str(Path(__file__).parent.parent / "shared"))
import utils_estado_ciclo as repo_estado  # noqa: E402
import utils_lote                          # noqa: E402
import utils_templates                     # noqa: E402

log = logging.getLogger(__name__)


class CicloNoValidadoError(Exception):
    pass


def paso1_gate(mes: str) -> None:
    if not repo_estado.ciclo_validado(mes, ruta=config.ESTADO_CICLO_PATH):
        raise CicloNoValidadoError(
            f"Ciclo {mes} no está validado en estado_ciclo.json — "
            f"correr 5b_validacion antes de cerrar el período."
        )
    log.info(f"PASO 1 · GATE — ciclo {mes} validado, sigue el cierre")


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


def paso3_freeze(mes: str) -> None:
    repo_estado.marcar_cerrado(mes, ruta=config.ESTADO_CICLO_PATH)
    log.info(f"PASO 3 · FREEZE — estado_ciclo[{mes}].estado = CERRADO")


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


def paso4_limpiar(mes: str) -> None:
    for nombre, ruta in _fuentes_listas_para_resetear(mes).items():
        if nombre == "correcciones_lote.xlsx":
            utils_lote.escribir_correcciones_lote(ruta, [])
        else:
            utils_templates.crear_mesa_vacio(ruta)
        log.info(f"PASO 4 · LIMPIAR — reset a template: {nombre}")

    out_dir = config.COBRANZA_DIR / "outputs"
    borrados = 0
    for patron in config.PATRONES_BASURA:
        for f in out_dir.glob(patron):
            f.unlink()
            borrados += 1
    log.info(f"PASO 4 · LIMPIAR — {borrados} archivo(s) de basura borrado(s)")


def _pedir_consentimiento(mes: str) -> bool:
    """Capa 1 (interactiva) del seguro: muestra exactamente qué se va a resetear
    y exige escribir SI. --confirmar habilita el gatillo; esto es apretarlo."""
    listas = _fuentes_listas_para_resetear(mes)
    print(f"\nSe va a SELLAR estado_ciclo[{mes}] = CERRADO y RESETEAR a template:")
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

    paso1_gate(mes)
    n_cosechados = paso2_cosechar(mes)

    if not confirmar:
        print(f"\n{n_cosechados} archivo(s) cosechados en archivo/{mes}/ — revisalos.")
        print("FREEZE y LIMPIAR NO corrieron (falta --confirmar). Nada irreversible pasó.")
        print(f"Para cerrar de verdad: python consolidar_cierre.py --mes {mes} --confirmar\n")
        return

    if not _pedir_consentimiento(mes):
        print("Cancelado — no se selló ni se reseteó nada.")
        return

    paso3_freeze(mes)
    paso4_limpiar(mes)

    print("\n" + "=" * 60)
    print(f"  Cierre de {mes} completado — {n_cosechados} archivos en archivo/{mes}/")
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
