"""Bridge for shared writers loaded as top-level modules by legacy entrypoints."""

import sys
from pathlib import Path


_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

def require_authoritative_writes(*paths):
    """Require an active token only when this workspace enables Work Guard."""
    if not (_ROOT / ".workguard.json").exists():
        return
    from work_guard.runtime import require_write

    for path in paths:
        target = Path(path).resolve()
        try:
            relative = target.relative_to(_ROOT)
        except ValueError:
            continue
        if "tests" in relative.parts and any(part.startswith("_tmp") for part in relative.parts):
            continue
        require_write(target)
