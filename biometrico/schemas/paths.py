import re
import unicodedata
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
RAW_BIOMETRICO_DIR = PROJECT_ROOT / "raw" / "biometrico"
REPORTS_BIOMETRICO_DIR = PROJECT_ROOT / "reports" / "biometrico"
EXCEL_REPORTS_DIR = REPORTS_BIOMETRICO_DIR / "excel"
WEB_REPORT_HTML = REPORTS_BIOMETRICO_DIR / "reporte_biometrico.html"
RAW_FERIADOS_PATH = PROJECT_ROOT / "raw" / "reglas" / "feriados_horarios.xlsx"


def slugify_filename(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_value = normalized.encode("ascii", "ignore").decode("ascii")
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "_", ascii_value).strip("_")
    return cleaned.lower() or "reporte"


def iter_biometrico_excels() -> list[Path]:
    return sorted(
        path
        for path in [*RAW_BIOMETRICO_DIR.glob("*.xlsx"), *RAW_BIOMETRICO_DIR.glob("*.xls")]
        if not path.name.startswith("~$")
    )


def get_latest_biometrico_excel() -> Path:
    files = iter_biometrico_excels()
    if not files:
        raise FileNotFoundError(
            f"No se encontraron archivos .xlsx en {RAW_BIOMETRICO_DIR}"
        )
    return max(files, key=lambda path: path.stat().st_mtime)


def build_excel_report_path(colaborador: str, mes_key: str) -> Path:
    filename = f"reporte_biometrico_{slugify_filename(colaborador)}_{mes_key}.xlsx"
    return EXCEL_REPORTS_DIR / filename
