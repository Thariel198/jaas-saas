"""Bloquea mutaciones del workspace mientras se ejecutan tests."""

from __future__ import annotations

import os
import sys
from pathlib import Path


_INSTALLED = False
_WORKSPACE: Path | None = None


class UnsafeTestWriteError(PermissionError):
    """Un test intento modificar el workspace real."""


def _as_path(value) -> Path | None:
    if isinstance(value, int) or value is None:
        return None
    try:
        return Path(os.fsdecode(value)).resolve(strict=False)
    except (TypeError, ValueError, OSError):
        return None


def _is_under(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def _allowed_workspace_temp(path: Path) -> bool:
    """Conserva los fixtures legacy documentados bajo tests/_tmp*."""
    assert _WORKSPACE is not None
    try:
        parts = path.relative_to(_WORKSPACE).parts
    except ValueError:
        return False

    if "__pycache__" in parts or ".pytest_cache" in parts:
        return True
    for index, part in enumerate(parts[:-1]):
        if part == "tests" and parts[index + 1].startswith(("_tmp", ".tmp")):
            return True
    return False


def _deny(path_value, event: str) -> None:
    path = _as_path(path_value)
    if path is None or _WORKSPACE is None or not _is_under(path, _WORKSPACE):
        return
    if _allowed_workspace_temp(path):
        return
    raise UnsafeTestWriteError(
        f"TEST BLOQUEADO: {event} intento modificar el workspace real: {path}. "
        "Use tmp_path/TemporaryDirectory o ejecute fuera del modo test."
    )


def _open_is_write(mode, flags) -> bool:
    if isinstance(mode, str) and any(char in mode for char in "wax+"):
        return True
    if isinstance(flags, int):
        write_flags = os.O_WRONLY | os.O_RDWR | os.O_CREAT | os.O_TRUNC | os.O_APPEND
        return bool(flags & write_flags)
    return False


def _audit(event: str, args: tuple) -> None:
    if event == "open" and args and _open_is_write(
        args[1] if len(args) > 1 else None,
        args[2] if len(args) > 2 else None,
    ):
        _deny(args[0], event)
        return

    single_path_events = {
        "os.remove", "os.rmdir", "os.mkdir", "os.chmod", "os.utime",
        "shutil.rmtree",
    }
    if event in single_path_events and args:
        _deny(args[0], event)
        return

    if event in {"os.rename", "os.link", "os.symlink", "shutil.move"}:
        if args:
            _deny(args[0], event)
        if len(args) > 1:
            _deny(args[1], event)
        return

    if event in {"shutil.copyfile", "shutil.copymode", "shutil.copystat"} and len(args) > 1:
        _deny(args[1], event)


def install(workspace: str | Path | None = None) -> None:
    global _INSTALLED, _WORKSPACE
    requested = Path(
        workspace or os.environ.get("JASS_TEST_WORKSPACE") or Path.cwd()
    ).resolve()
    if _INSTALLED:
        if requested != _WORKSPACE:
            raise RuntimeError(f"Test guard ya instalado para {_WORKSPACE}, no para {requested}")
        return

    _WORKSPACE = requested
    sys.addaudithook(_audit)
    _INSTALLED = True
    os.environ["JASS_TEST_GUARD_ACTIVE"] = "1"


def is_active() -> bool:
    return _INSTALLED
