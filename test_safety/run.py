"""Lanza pytest o un test standalone con el workspace en solo lectura."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent


def _usage() -> int:
    print(
        "Uso:\n"
        "  py -m test_safety.run pytest [argumentos]\n"
        "  py -m test_safety.run script RUTA_TEST.py [argumentos]",
        file=sys.stderr,
    )
    return 2


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] not in {"pytest", "script"}:
        return _usage()

    mode = sys.argv[1]
    args = sys.argv[2:]
    if mode == "script" and not args:
        return _usage()

    env = os.environ.copy()
    env["JASS_TEST_MODE"] = "1"
    env["JASS_TEST_WORKSPACE"] = str(ROOT)
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    bootstrap = str(ROOT / "test_safety")
    current_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = os.pathsep.join(
        part for part in (bootstrap, str(ROOT), current_pythonpath) if part
    )

    python = [sys.executable, "-u", "-X", "utf8"]
    if mode == "pytest":
        command = python + ["-m", "pytest", *args]
    else:
        script = (ROOT / args[0]).resolve()
        if not script.is_relative_to(ROOT):
            print(f"Test fuera del workspace no permitido: {script}", file=sys.stderr)
            return 2
        command = python + [str(script), *args[1:]]
    return subprocess.run(command, cwd=ROOT, env=env, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
