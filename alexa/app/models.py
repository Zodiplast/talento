"""
models.py — Tipos de datos compartidos entre processor, database y main.

TypedDicts: contratos con la plantilla HTML (no cambian sin cambiar el template).
Dataclasses: contratos internos entre processor.py y database.py.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, TypedDict


# ── Tipos para la plantilla HTML ──────────────────────────────────────────────

class DiaData(TypedDict):
    fecha: str                      # "2026-01-15"
    dia: int                        # 15
    dow: str                        # "Lun"
    estado: str                     # "completo" | "incompleto" | "libre"
    motivo_feriado: Optional[str]
    entrada: Optional[str]          # "08:05"
    sal_des: Optional[str]
    ent_des: Optional[str]
    sal_alm: Optional[str]
    ent_alm: Optional[str]
    salida: Optional[str]
    t_desayuno: Optional[float]
    t_almuerzo: Optional[float]
    t_jornada: Optional[float]
    t_efectivo: Optional[float]
    jornada_objetivo: Optional[int]
    exceso_desayuno: Optional[float]
    exceso_almuerzo: Optional[float]
    exceso_descanso: Optional[float]
    horas_extra: Optional[float]
    resaltar_descanso: bool


class ResumenData(TypedDict):
    completos: int
    incompletos: int
    libres: int
    feriados: int
    total_efectivo: int
    total_jornada: int
    prom_efectivo: int
    prom_jornada: int
    total_exceso_descanso: int
    total_horas_extra: int


class MesPayload(TypedDict):
    dias: list[DiaData]
    resumen: ResumenData
    excel_file: Optional[str]       # "excel/reporte_biometrico_jose_2026-01.xlsx"


class ConfigData(TypedDict):
    desayuno_min: int
    almuerzo_min: int
    descanso_max_min: int
    jornada_lunes_viernes_min: int
    jornada_sabado_min: int


class MesInfo(TypedDict):
    value: str   # "2026-01"
    label: str   # "Enero 2026"


class ReporteJSON(TypedDict):
    config: ConfigData
    colaboradores: list[str]
    meses: list[MesInfo]
    datos: dict[str, dict[str, MesPayload]]  # datos[nombre][mes_key]


# ── Contratos internos processor ↔ database ───────────────────────────────────

@dataclass
class ColaboradorResult:
    """Resultado procesado de un colaborador para un mes dado."""
    colab_key: str               # "id:00042" o "nom:JOSE GARCIA"
    display_name: str            # "Jose Garcia"
    mes_key: str                 # "2026-01"
    mes_label: str               # "Enero 2026"
    dias: list[DiaData] = field(default_factory=list)
    resumen: Optional[ResumenData] = None


@dataclass
class ProcessedMonth:
    """Resultado completo de procesar un archivo Excel del biométrico."""
    mes_key: str                 # "2026-01"
    mes_label: str               # "Enero 2026"
    original_file: str           # "ENERO-2026.xlsx"
    colaboradores: list[ColaboradorResult] = field(default_factory=list)

    @property
    def total_dias(self) -> int:
        return sum(len(c.dias) for c in self.colaboradores)
