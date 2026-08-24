"""Barrera global: pytest nunca puede modificar el workspace real."""

import os
import sys
from pathlib import Path

from test_safety.fs_guard import install


WORKSPACE = Path(__file__).resolve().parent
os.environ["JASS_TEST_MODE"] = "1"
os.environ["JASS_TEST_WORKSPACE"] = str(WORKSPACE)
sys.dont_write_bytecode = True
install(WORKSPACE)


def pytest_report_header():
    return "JASS test guard: workspace read-only; writes allowed only outside the repository"
