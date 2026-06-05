import sys
import importlib.util
from pathlib import Path

ROOT = Path(__file__).parent


def _cargar(alias: str, ruta: Path):
    spec = importlib.util.spec_from_file_location(alias, ruta)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[alias] = mod
    spec.loader.exec_module(mod)
    return mod


def _banner(texto: str):
    print("\n" + "═" * 55)
    print(f"  {texto}")
    print("═" * 55)


def _paso(n: int, total: int, texto: str):
    print(f"\n[{n}/{total}] {texto}")
    print("─" * 55)


def main():
    _banner("PROCESAMIENTO DE PAGOS — 03_pagos")

    origenes   = _cargar("pagos_origenes",   ROOT / "yape/construir_maestro/crear_origenes/main.py")
    maestro    = _cargar("pagos_maestro",    ROOT / "yape/construir_maestro/crear_maestro/main.py")
    motor      = _cargar("pagos_motor",      ROOT / "yape/motor_matching/main.py")
    validacion = _cargar("pagos_validacion", ROOT / "yape/validacion/main.py")
    efectivo   = _cargar("pagos_efectivo",   ROOT / "efectivo/main.py")

    # ── Pipeline Yape ───────────────────────────────────────
    _paso(1, 5, "Yape — Construyendo orígenes")
    origenes.main()

    _paso(2, 5, "Yape — Construyendo maestro")
    maestro.main()

    _paso(3, 5, "Yape — Motor de matching")
    motor.main()

    _paso(4, 5, "Yape — Validación")
    validacion.main()

    # ── Pipeline Efectivo ───────────────────────────────────
    _paso(5, 5, "Efectivo — Procesando cobros")
    efectivo.main()

    _banner("03_pagos completado → entregar outputs a 04_cobranza")


if __name__ == "__main__":
    main()
