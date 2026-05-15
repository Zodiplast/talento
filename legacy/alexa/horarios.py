"""
horarios.py — Configuración parametrizable de turnos y descansos.

Copia alineada con `legacy/doc/horarios.py`; usado por `alexa.reporte_core` y la webapp en `legacy/`.
Editá `legacy/doc/horarios.py` o este archivo manteniéndolos alineados.
"""

from dataclasses import dataclass

DESAYUNO_MIN = 15
ALMUERZO_MIN = 30
DESCANSO_MAX_MIN = DESAYUNO_MIN + ALMUERZO_MIN

JORNADA_LUNES_VIERNES_MIN = 8 * 60
JORNADA_SABADO_MIN = 4 * 60
JORNADA_BASE_MIN = JORNADA_LUNES_VIERNES_MIN

TOLERANCIA_TURNO_MIN = 90


@dataclass(frozen=True)
class Turno:
    nombre: str
    entrada: str
    salida: str

    @property
    def horas_contractuales(self) -> float:
        h_in, m_in = map(int, self.entrada.split(":"))
        h_out, m_out = map(int, self.salida.split(":"))
        return ((h_out * 60 + m_out) - (h_in * 60 + m_in)) / 60

    @property
    def horas_efectivas(self) -> float:
        return self.horas_contractuales - (DESAYUNO_MIN + ALMUERZO_MIN) / 60


TURNOS: dict[str, Turno] = {
    "madrugada": Turno("Madrugada", "05:30", "17:00"),
    "ordinario": Turno("Ordinario", "08:00", "18:00"),
    "inventario": Turno("Inventario", "09:00", "20:00"),
    "domingo": Turno("Domingo", "06:00", "14:00"),
}

DIAS_LIBRES_POR_SEMANA = 1
DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


def get_jornada_objetivo_min(weekday: int) -> int:
    if weekday <= 4:
        return JORNADA_LUNES_VIERNES_MIN
    if weekday == 5:
        return JORNADA_SABADO_MIN
    return 0


def listar_turnos() -> None:
    print("\nTURNOS CONFIGURADOS")
    print("─" * 60)
    print(f"  Desayuno: {DESAYUNO_MIN} min  |  Almuerzo: {ALMUERZO_MIN} min")
    print("─" * 60)
    for key, t in TURNOS.items():
        print(
            f"  [{key:<10}]  {t.nombre:<12}  "
            f"{t.entrada} → {t.salida}  |  "
            f"{t.horas_contractuales:.1f}h brutas  "
            f"{t.horas_efectivas:.1f}h efectivas"
        )
    print("─" * 60)
    print(f"  Día libre por semana  : {DIAS_LIBRES_POR_SEMANA}")
    print(f"  Descanso máx. total   : {DESCANSO_MAX_MIN} min")
    print(
        "  Jornada base (extras) : "
        f"L-V {JORNADA_LUNES_VIERNES_MIN // 60}h, "
        f"Sáb {JORNADA_SABADO_MIN // 60}h"
    )
    print(f"  Tolerancia turno      : ±{TOLERANCIA_TURNO_MIN} min")
    print()


if __name__ == "__main__":
    listar_turnos()
