"""Setup para los tests del submódulo sin_servicio.

1) Agrega 1_lecturas/ y jass_system/ a sys.path para que los imports
   `from sin_servicio.X import Y` y `from shared.X import Y` resuelvan.

2) Forza --basetemp a tests/.tmp/ — workaround para Windows donde
   %TEMP%/pytest-of-* da PermissionError en algunos perfiles.
"""
import sys
from pathlib import Path

_ROOT_1_LECTURAS = Path(__file__).parent.parent.parent
_ROOT_JASS = _ROOT_1_LECTURAS.parent
sys.path.insert(0, str(_ROOT_1_LECTURAS))
sys.path.insert(0, str(_ROOT_JASS))

_LOCAL_TMP = Path(__file__).parent / ".tmp"


def pytest_configure(config):
    if not config.getoption("--basetemp"):
        _LOCAL_TMP.mkdir(exist_ok=True)
        config.option.basetemp = str(_LOCAL_TMP)
