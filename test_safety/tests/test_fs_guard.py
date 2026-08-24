import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


def _guarded_python(code: str) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["JASS_TEST_MODE"] = "1"
    env["JASS_TEST_WORKSPACE"] = str(ROOT)
    env["PYTHONPATH"] = os.pathsep.join((str(ROOT / "test_safety"), str(ROOT)))
    return subprocess.run(
        [sys.executable, "-c", code],
        cwd=ROOT,
        env=env,
        text=True,
        capture_output=True,
        check=False,
    )


def test_bloquea_escritura_en_workspace():
    target = ROOT / "_test_guard_no_debe_existir.txt"
    result = _guarded_python(f"open({str(target)!r}, 'w').write('x')")
    assert result.returncode != 0
    assert "TEST BLOQUEADO" in result.stderr
    assert not target.exists()


def test_permite_escritura_fuera_del_workspace(tmp_path):
    target = tmp_path / "permitido.txt"
    result = _guarded_python(f"open({str(target)!r}, 'w').write('ok')")
    assert result.returncode == 0, result.stderr
    assert target.read_text() == "ok"


def test_bloquea_openpyxl_en_workspace():
    target = ROOT / "_test_guard_no_debe_existir.xlsx"
    code = (
        "from openpyxl import Workbook; "
        f"Workbook().save({str(target)!r})"
    )
    result = _guarded_python(code)
    assert result.returncode != 0
    assert "TEST BLOQUEADO" in result.stderr
    assert not target.exists()


def test_bloquea_borrado_en_workspace():
    protected = ROOT / "pytest.ini"
    before = protected.read_bytes()
    result = _guarded_python(
        f"from pathlib import Path; Path({str(protected)!r}).unlink()"
    )
    assert result.returncode != 0
    assert "TEST BLOQUEADO" in result.stderr
    assert protected.read_bytes() == before


def test_runner_rechaza_script_fuera_del_workspace(tmp_path):
    result = subprocess.run(
        [sys.executable, "-m", "test_safety.run", "script", str(tmp_path / "test.py")],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    assert result.returncode == 2
    assert "fuera del workspace" in result.stderr
