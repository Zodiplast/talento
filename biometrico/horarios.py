"""
horarios.py — Configuración parametrizable de turnos y descansos.

Edita este archivo para ajustar cualquier horario sin tocar la lógica
del reporte. Todas las horas son strings "HH:MM".
"""

from dataclasses import dataclass

# ── Descansos ─────────────────────────────────────────────────────────────────
# Tiempo estándar de descanso aplicado a todos los turnos (en minutos).
DESAYUNO_MIN = 15
ALMUERZO_MIN = 30
DESCANSO_MAX_MIN = DESAYUNO_MIN + ALMUERZO_MIN  # 45 min — límite permitido en total

# ── Jornada base ──────────────────────────────────────────────────────────────
# Horas ordinarias diarias. Todo tiempo efectivo por encima = horas extra.
JORNADA_LUNES_VIERNES_MIN = 8 * 60  # 480 min
JORNADA_SABADO_MIN = 4 * 60  # 240 min
JORNADA_BASE_MIN = JORNADA_LUNES_VIERNES_MIN

# ── Tolerancia de clasificación de turno ──────────────────────────────────────
# Ventana (en minutos) alrededor de la hora de entrada del turno para
# asignar automáticamente el turno a un día de trabajo.
TOLERANCIA_TURNO_MIN = 90


# ── Definición de turnos ──────────────────────────────────────────────────────
@dataclass(frozen=True)
class Turno:
    nombre: str  # nombre descriptivo
    entrada: str  # hora de inicio esperada  "HH:MM"
    salida: str  # hora de fin esperada      "HH:MM"

    @property
    def horas_contractuales(self) -> float:
        """Horas brutas del turno (sin descontar descansos)."""
        h_in, m_in = map(int, self.entrada.split(":"))
        h_out, m_out = map(int, self.salida.split(":"))
        return ((h_out * 60 + m_out) - (h_in * 60 + m_in)) / 60

    @property
    def horas_efectivas(self) -> float:
        """Horas reales trabajadas (descontando desayuno y almuerzo)."""
        return self.horas_contractuales - (DESAYUNO_MIN + ALMUERZO_MIN) / 60


# ── Catálogo de turnos ────────────────────────────────────────────────────────
# Agrega, elimina o edita turnos aquí. La clave es el ID corto que usarás
# para asignarle un turno a cada colaborador.

TURNOS: dict[str, Turno] = {
    "madrugada": Turno("Madrugada", "05:30", "17:00"),
    "ordinario": Turno("Ordinario", "08:00", "18:00"),
    "inventario": Turno("Inventario", "09:00", "20:00"),
    "domingo": Turno("Domingo", "06:00", "14:00"),
}

# Días libres: cada colaborador tiene 1 día libre a la semana.
# Se gestiona en la asignación de horarios (ver asignacion.py).
DIAS_LIBRES_POR_SEMANA = 1

# Nombres de días (para legibilidad en reportes)
DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def get_jornada_objetivo_min(weekday: int) -> int:
    """Devuelve la jornada objetivo por día de la semana.

    weekday sigue la convención de datetime.date.weekday():
    0=Lunes, 5=Sábado, 6=Domingo.
    """
    if weekday <= 4:
        return JORNADA_LUNES_VIERNES_MIN
    if weekday == 5:
        return JORNADA_SABADO_MIN
    return 0


def listar_turnos() -> None:
    """Imprime en consola el resumen de todos los turnos configurados."""
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
