"""
conftest.py — hace que estos tests NO toquen los datos reales cuando corren bajo
pytest.

EL PROBLEMA

Los tres archivos de test de este módulo se escribieron para correr como script:

    py test_integracion.py      -> su main() hace _setup() antes de CADA test
                                   y _teardown() al final

Bajo pytest eso no pasa: pytest recolecta las funciones `test_*` directamente y
el `main()` del final del archivo nunca se ejecuta. Sin `_setup()`, los módulos
quedan apuntando a las rutas REALES:

    efectivo.INPUTS_DIR  -> 4_pagos/efectivo/inputs/     (mesas del ciclo en curso)
    efectivo.OUTPUTS_DIR -> 4_pagos/efectivo/outputs/

...y los tests que crean mesas de fixture las escriben ahí. El 12/08/2026 correr
`pytest 4_pagos/efectivo/tests` sobrescribió mesa_1.xlsx (59 filas de cobro
reales -> 2) y mesa_2.xlsx (106 -> 2) del ciclo de agosto. Se recuperaron con
`git checkout HEAD --`, pero solo porque estaban commiteadas.

El síntoma en el output era inconfundible y se leyó mal: `assert (162 == 1)` —
el test contaba las 162 filas reales en vez de las de su fixture.

LA SOLUCIÓN

Un fixture autouse que corre `_setup()` antes y `_teardown()` después de cada
test, replicando lo que hace el runner del script. Se pone acá y no en cada
archivo para no tocar los tests: siguen funcionando igual como script.

GUARDA EXTRA

Además se verifica, después del setup, que las rutas del módulo NO apunten al
repo real. Si un test futuro se agrega sin que su archivo tenga `_setup`, o si
alguien cambia los nombres, el fixture falla ruidosamente en vez de escribir
sobre el ciclo en curso.
"""

import sys
from pathlib import Path

import pytest

THIS = Path(__file__).resolve()
MODULO_DIR = THIS.parent.parent               # 4_pagos/efectivo
sys.path.insert(0, str(MODULO_DIR))

# Rutas reales que NUNCA deben ser el destino de un test.
_PROHIBIDAS = {
    (MODULO_DIR / "inputs").resolve(),
    (MODULO_DIR / "outputs").resolve(),
    (MODULO_DIR / "trazabilidad").resolve(),
    (MODULO_DIR / "backup").resolve(),
}


def _rutas_del_modulo(mod):
    """Los Path que el módulo usa como directorio de trabajo."""
    return {nombre: valor for nombre in
            ("BASE_DIR", "INPUTS_DIR", "OUTPUTS_DIR", "TRAZAB_DIR", "BACKUP_DIR")
            if isinstance(valor := getattr(mod, nombre, None), Path)}


def _verificar_aislamiento(mod, nombre_mod: str) -> None:
    for attr, ruta in _rutas_del_modulo(mod).items():
        if ruta.resolve() in _PROHIBIDAS:
            pytest.fail(
                f"AISLAMIENTO ROTO: {nombre_mod}.{attr} apunta al repo real ({ruta}).\n"
                f"  El test iba a escribir sobre los datos del ciclo en curso.\n"
                f"  Revisar que el archivo de test defina _setup()/_teardown() y que "
                f"redirija esa ruta a su carpeta temporal.",
                pytrace=False)


@pytest.fixture(autouse=True)
def aislar_del_repo_real(request):
    """Corre el _setup()/_teardown() del archivo de test y verifica el aislamiento.

    autouse=True: aplica a TODOS los tests de esta carpeta sin que haya que
    pedirlo — que es justamente lo que fallaba (había que acordarse)."""
    modulo = request.module
    setup = getattr(modulo, "_setup", None)
    teardown = getattr(modulo, "_teardown", None)

    if callable(setup):
        setup()

    # Verificar DESPUÉS del setup, sobre los módulos que el test haya importado.
    for nombre in ("efectivo", "vl", "main", "verificar_lotes"):
        mod = getattr(modulo, nombre, None)
        if mod is not None and hasattr(mod, "__file__"):
            _verificar_aislamiento(mod, nombre)

    yield

    if callable(teardown):
        teardown()
