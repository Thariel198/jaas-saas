"""7_cierre/config.py — paths para consolidar_cierre.py.

7_cierre no genera arrastres (los hace 5_cobranza) ni tiene schema propio —
solo transiciona el período: cosecha una foto inmutable + resetea slots
mutables. Ver README.md y docs/diagrama_consolidador_cierre.html.
"""
import sys
from pathlib import Path

ROOT        = Path(__file__).parent
ARCHIVO_DIR = ROOT / "archivo"      # TRACKEADO en git — la foto por período
OUTPUTS_DIR = ROOT / "outputs"      # solo run.log (gitignored)

SHARED_DIR   = ROOT.parent / "shared"
sys.path.insert(0, str(SHARED_DIR))
import ciclo  # noqa: E402
COBRANZA_DIR = ROOT.parent / "5_cobranza"
EFECTIVO_DIR = ROOT.parent / "4_pagos" / "efectivo"
YAPE_DIR     = ROOT.parent / "4_pagos" / "yape" / "motor_matching" / "outputs"

ESTADO_CICLO_PATH = SHARED_DIR / "reporte_acumulado_procesado" / "estado_ciclo.json"
SEGUIMIENTO_PATH  = SHARED_DIR / "seguimiento_pueblo.xlsx"
COBRANZA_MAIN     = COBRANZA_DIR / "main.py"
VALIDACION_MAIN   = ROOT.parent / "5b_validacion" / "main.py"


def snapshot_ledger_path(mes: str) -> Path:
    return COBRANZA_DIR / "outputs" / f"snapshot_ledger_{mes}.json"


def archivo_mes_dir(mes: str) -> Path:
    return ARCHIVO_DIR / mes


def canonicos_a_cosechar(mes: str) -> dict[str, Path]:
    """Derivados de 5_cobranza — BALDE 2, solo cosecha (nadie los resetea,
    julio no vuelve a leer el nombre de junio)."""
    out = COBRANZA_DIR / "outputs"
    planilla = ciclo.resolver(
        out, "planilla_cobrado", mes,
        legacy_sin_periodo=ciclo.acepta_legacy(mes),
    )
    return {
        f"planilla_cobrado_{mes}.xlsx":     planilla,
        f"snapshot_ledger_{mes}.json":       snapshot_ledger_path(mes),
    }


def fuentes_manuales_a_resetear() -> dict[str, Path]:
    """Inputs tipeados a mano — BALDE 2, cosechar Y resetear a template
    (nadie más los vacía; si no se resetean, julio hereda datos de junio)."""
    inputs = EFECTIVO_DIR / "inputs"
    d = {f"mesa_{i}.xlsx": inputs / f"mesa_{i}.xlsx" for i in range(1, 8)}
    d["correcciones_lote.xlsx"] = COBRANZA_DIR / "inputs" / "correcciones_lote.xlsx"
    return d


def fuentes_auto_a_cosechar() -> dict[str, Path]:
    """Outputs de motor_matching — BALDE 2, solo cosecha. NO se resetean:
    motor_matching los sobreescribe solo al correr sobre el crudo de julio."""
    return {
        "pagos_yape_tepago.xlsx":     YAPE_DIR / "pagos_yape_tepago.xlsx",
        "pagos_yape_devolucion.xlsx": YAPE_DIR / "pagos_yape_devolucion.xlsx",
        "pagos_yape_retorno.xlsx":    YAPE_DIR / "pagos_yape_retorno.xlsx",
    }


# BALDE 3 — basura one-time de 5_cobranza/outputs a borrar (no cruza a julio)
PATRONES_BASURA = ["trazabilidad_cobranza_pre_dedup_*.xlsx"]
