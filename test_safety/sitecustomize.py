"""Bootstrap automatico cargado por Python dentro del runner seguro."""

import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

if os.environ.get("JASS_TEST_MODE") == "1":
    from test_safety.fs_guard import install

    sys.dont_write_bytecode = True
    install(os.environ.get("JASS_TEST_WORKSPACE", ROOT))
