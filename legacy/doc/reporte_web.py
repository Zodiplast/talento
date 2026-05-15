"""
reporte_web.py — CLI que genera el HTML del reporte biométrico.

La lógica vive en `legacy/alexa/reporte_core.py` (única fuente de verdad).

Uso (desde la raíz talento/):
    python legacy/doc/reporte_web.py

Salida: legacy/reports/biometrico/reporte_biometrico.html
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

from alexa.reporte_core import write_html_report

sys.stdout.reconfigure(encoding="utf-8")


def main() -> None:
    print("Reporte Biométrico Web")
    print("─" * 40)
    try:
        write_html_report()
    except FileNotFoundError as exc:
        print(f"[ERROR] {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
