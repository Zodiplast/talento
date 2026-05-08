"""
Punto de entrada histórico: la fuente de verdad es `biometrico/horarios.py`.

Ejecutá desde la raíz talento/:
    python doc/horarios.py
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from biometrico.horarios import *  # noqa: F403

if __name__ == "__main__":
    from biometrico.horarios import listar_turnos

    listar_turnos()
