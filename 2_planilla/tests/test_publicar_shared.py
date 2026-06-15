"""
test_publicar_shared.py
Verifica que publicar_a_shared() copia planilla_{mes}.xlsx a shared/planilla_mes/
sobreescribiendo el archivo del mismo mes sin tocar archivos de otros meses.

Patron 3.6c metodologia: estado minimo sintetico + assert con mensaje claro.
"""

import shutil
import sys
from pathlib import Path

THIS = Path(__file__).resolve()
sys.path.insert(0, str(THIS.parent.parent))

import config  # noqa: E402

TEST_ROOT = THIS.parent / "_tmp_publicar"


def _setup():
    if TEST_ROOT.exists():
        shutil.rmtree(TEST_ROOT)
    (TEST_ROOT / "outputs").mkdir(parents=True)
    (TEST_ROOT / "shared" / "planilla_mes").mkdir(parents=True)
    config.OUTPUTS_DIR = TEST_ROOT / "outputs"
    config.SHARED_PLANILLA_DIR = TEST_ROOT / "shared" / "planilla_mes"


def _crear_planilla(mes: str, contenido: bytes):
    src = TEST_ROOT / "outputs" / f"planilla_{mes}.xlsx"
    src.write_bytes(contenido)
    return src


def test_copia_a_shared(mes="2026-07"):
    _crear_planilla(mes, b"contenido_julio")
    # Re-importar despues del setup para que tome los paths patcheados
    if "main" in sys.modules:
        del sys.modules["main"]
    import main as mod
    mod.publicar_a_shared(mes)

    dest = config.SHARED_PLANILLA_DIR / f"planilla_{mes}.xlsx"
    assert dest.exists(), f"No se copio a shared: {dest}"
    assert dest.read_bytes() == b"contenido_julio", "Contenido no coincide tras copiar"
    print(f"  OK test_copia_a_shared: {dest.name}")


def test_sobreescribe_mismo_mes(mes="2026-07"):
    _crear_planilla(mes, b"v1")
    if "main" in sys.modules:
        del sys.modules["main"]
    import main as mod
    mod.publicar_a_shared(mes)

    # Re-publicar con contenido nuevo
    _crear_planilla(mes, b"v2_actualizado")
    mod.publicar_a_shared(mes)

    dest = config.SHARED_PLANILLA_DIR / f"planilla_{mes}.xlsx"
    assert dest.read_bytes() == b"v2_actualizado", "No sobreescribio el del mismo mes"
    print(f"  OK test_sobreescribe_mismo_mes")


def test_no_toca_otros_meses(mes="2026-07"):
    # Pre-existente del mes anterior en shared
    junio = config.SHARED_PLANILLA_DIR / "planilla_2026-06.xlsx"
    junio.write_bytes(b"junio_existente")

    _crear_planilla(mes, b"julio_nuevo")
    if "main" in sys.modules:
        del sys.modules["main"]
    import main as mod
    mod.publicar_a_shared(mes)

    assert junio.exists(), "Borro el archivo de otro mes!"
    assert junio.read_bytes() == b"junio_existente", "Modifico el archivo de otro mes!"
    print(f"  OK test_no_toca_otros_meses: junio intacto, julio publicado")


def main():
    _setup(); test_copia_a_shared()
    _setup(); test_sobreescribe_mismo_mes()
    _setup(); test_no_toca_otros_meses()
    print("\n[OK] 3/3 tests pasaron")
    shutil.rmtree(TEST_ROOT)


if __name__ == "__main__":
    main()
