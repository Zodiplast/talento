import re
import unicodedata
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]

RAW_BIOMETRICO_DIR = PROJECT_ROOT / "raw" / "biometrico"
RAW_BIOMETRICO_PARQUET_DIR = PROJECT_ROOT / "raw" / "biometrico_parquet"
RAW_REGLAS_DIR = PROJECT_ROOT / "raw" / "reglas"
RAW_FERIADOS_PATH = RAW_REGLAS_DIR / "feriados_horarios.xlsx"
COLLABORATORS_CONFIG_PATH = RAW_REGLAS_DIR / "colaboradores.yaml"

REPORTS_DIR = PROJECT_ROOT / "legacy" / "reports" / "biometrico"
EXCEL_REPORTS_DIR = REPORTS_DIR
WEB_REPORT_HTML = REPORTS_DIR / "reporte_biometrico.html"

DOC_DIR = PROJECT_ROOT / "legacy" / "doc"
REPORT_TEMPLATE_HTML = DOC_DIR / "reporte_template.html"


def stem_export_mes_biometrico(anio: int, mes: int) -> str:
    """YYYYMMDD (primer día del mes): mismo prefijo para .xlsx y .parquet mensuales."""
    return f"{anio:04d}{mes:02d}01"


def stem_export_memoria_biometrico(dt: datetime | None = None) -> str:
    """YYYYMMDDHHMMSS: volcado de memoria completa; mismo prefijo para xlsx y parquet."""
    return (dt or datetime.now()).strftime("%Y%m%d%H%M%S")


def slugify_filename(name: str) -> str:
    s = unicodedata.normalize("NFD", name).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower()
    return s or "colab"


def build_excel_report_path(colaborador: str, mes_key: str) -> Path:
    return EXCEL_REPORTS_DIR / f"reporte_biometrico_{slugify_filename(colaborador)}_{mes_key}.xlsx"
