"""
Punto de entrada histórico: la fuente de verdad es `legacy/alexa/horarios.py`.

Ejecutá desde la raíz talento/:
    python legacy/doc/horarios.py
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
_LEGACY = _ROOT / "legacy"
if _LEGACY.is_dir() and str(_LEGACY) not in sys.path:
    sys.path.insert(0, str(_LEGACY))

from alexa.horarios import *  # noqa: F403

if __name__ == "__main__":
    from alexa.horarios import listar_turnos

    listar_turnos()
