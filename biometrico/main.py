import sys
from datetime import datetime, date

import polars as pl

sys.stdout.reconfigure(encoding="utf-8")

# ── Configuración ─────────────────────────────────────────────────────────────
ARCHIVO = "raw/biometrico/ENERO-2026.xlsx"
USUARIO = "DARLA RAMON"
UMBRAL_DUPLICADO_MIN = 5   # marcaciones a menos de X min = duplicado

# ── 1. Carga ──────────────────────────────────────────────────────────────────
df = pl.read_excel(ARCHIVO, sheet_name="Sheet")

# Separar nombre y apellido (primera palabra = nombre, resto = apellido)
df = df.with_columns([
    pl.col("Nombre").str.split(" ").list.first().alias("nombre"),
    pl.col("Nombre").str.split(" ").list.slice(1).list.join(" ").alias("apellido"),
    pl.col("Tiempo").cast(pl.Date).alias("fecha"),
    pl.col("Tiempo").cast(pl.Time).alias("hora"),
])

# ── 2. Filtro usuario ─────────────────────────────────────────────────────────
darla = (
    df.filter(pl.col("Nombre") == USUARIO)
    .sort(["fecha", "Tiempo"])
)

# ── 3. Eliminar duplicados cercanos por día ───────────────────────────────────
# Convertir a Python para lógica por día
filas = darla.to_dicts()

limpias = []
for i, fila in enumerate(filas):
    if i == 0:
        limpias.append(fila)
        continue
    prev = limpias[-1]
    # Si mismo día y diferencia < umbral, descartar
    if fila["fecha"] == prev["fecha"]:
        diff_min = (fila["Tiempo"] - prev["Tiempo"]).total_seconds() / 60
        if diff_min < UMBRAL_DUPLICADO_MIN:
            continue
    limpias.append(fila)

# ── 4. Agrupar por día y asignar roles ───────────────────────────────────────
from collections import defaultdict

dias: dict[date, list[datetime]] = defaultdict(list)
for fila in limpias:
    dias[fila["fecha"]].append(fila["Tiempo"])

# Ordenar fechas
fechas_ordenadas = sorted(dias.keys())


def hhmm(dt: datetime | None) -> str:
    return dt.strftime("%H:%M:%S") if dt else "   --   "


def fmt_dur(minutos: float | None) -> str:
    if minutos is None:
        return "  --  "
    h = int(minutos) // 60
    m = int(minutos) % 60
    return f"{h}h {m:02d}m"


def diff_min(a: datetime | None, b: datetime | None) -> float | None:
    if a is None or b is None:
        return None
    return (b - a).total_seconds() / 60


resultados = []

for fecha in fechas_ordenadas:
    marcas = sorted(dias[fecha])
    n = len(marcas)

    entrada = salida = None
    sal_des = ent_des = sal_alm = ent_alm = None

    if n >= 1:
        entrada = marcas[0]
    if n >= 2:
        salida = marcas[-1]

    if n == 2:
        pass  # solo entrada y salida, sin breaks
    elif n == 3:
        # Probable: entrada, salida almuerzo (sin regreso), salida
        sal_alm = marcas[1]
    elif n == 4:
        # Entrada, salida_almuerzo, entrada_almuerzo, salida
        sal_alm = marcas[1]
        ent_alm = marcas[2]
    elif n == 5:
        # Entrada, sal_des?, sal_alm, ent_alm, salida  — ambiguo
        # Asumimos: el par de None más cercano entre sí es almuerzo
        sal_alm = marcas[2]
        ent_alm = marcas[3]
    elif n >= 6:
        # Entrada, sal_des, ent_des, sal_alm, ent_alm, salida
        sal_des = marcas[1]
        ent_des = marcas[2]
        sal_alm = marcas[3]
        ent_alm = marcas[4]

    t_desayuno = diff_min(sal_des, ent_des)
    t_almuerzo = diff_min(sal_alm, ent_alm)
    t_jornada = diff_min(entrada, salida)

    if t_jornada is not None:
        t_efectivo = t_jornada - (t_desayuno or 0) - (t_almuerzo or 0)
    else:
        t_efectivo = None

    resultados.append({
        "fecha":        fecha,
        "n":            n,
        "entrada":      entrada,
        "sal_des":      sal_des,
        "ent_des":      ent_des,
        "sal_alm":      sal_alm,
        "ent_alm":      ent_alm,
        "salida":       salida,
        "t_desayuno":   t_desayuno,
        "t_almuerzo":   t_almuerzo,
        "t_jornada":    t_jornada,
        "t_efectivo":   t_efectivo,
    })

# ── 5. Reporte en terminal ────────────────────────────────────────────────────
ANCHO = 110
SEP   = "─" * ANCHO

print()
print("═" * ANCHO)
print(f"  REPORTE BIOMÉTRICO  |  {USUARIO}  |  ENERO 2026".center(ANCHO))
print("═" * ANCHO)
print(
    f"  {'FECHA':<11}"
    f"{'ENTRADA':<10}"
    f"{'S.DESAYUNO':<12}"
    f"{'E.DESAYUNO':<12}"
    f"{'S.ALMUERZO':<12}"
    f"{'E.ALMUERZO':<12}"
    f"{'SALIDA':<10}"
    f"{'DESAYUNO':>9}"
    f"{'ALMUERZO':>9}"
    f"{'EFECTIVO':>9}"
    f"{'JORNADA':>9}"
    f"  #"
)
print(SEP)

total_efectivo = 0.0
total_jornada  = 0.0
dias_completos = 0

for r in resultados:
    fecha_str = r["fecha"].strftime("%a %d/%m/%y")
    incompleto = " ⚠" if (r["entrada"] is None or r["salida"] is None) else "  "

    print(
        f"  {fecha_str:<11}"
        f"{hhmm(r['entrada']):<10}"
        f"{hhmm(r['sal_des']):<12}"
        f"{hhmm(r['ent_des']):<12}"
        f"{hhmm(r['sal_alm']):<12}"
        f"{hhmm(r['ent_alm']):<12}"
        f"{hhmm(r['salida']):<10}"
        f"{fmt_dur(r['t_desayuno']):>9}"
        f"{fmt_dur(r['t_almuerzo']):>9}"
        f"{fmt_dur(r['t_efectivo']):>9}"
        f"{fmt_dur(r['t_jornada']):>9}"
        f"  {r['n']}{incompleto}"
    )

    if r["t_efectivo"] is not None:
        total_efectivo += r["t_efectivo"]
        dias_completos += 1
    if r["t_jornada"] is not None:
        total_jornada += r["t_jornada"]

print(SEP)

prom_efectivo = total_efectivo / dias_completos if dias_completos else None
prom_jornada  = total_jornada  / dias_completos if dias_completos else None

print()
print(f"  Empleado                   : {USUARIO}")
print(f"  Días con jornada completa  : {dias_completos}")
print(f"  Días con jornada incompleta: {sum(1 for r in resultados if r['entrada'] is None or r['salida'] is None)}")
print(f"  Total horas efectivas      : {fmt_dur(total_efectivo)}")
print(f"  Promedio horas efectivas   : {fmt_dur(prom_efectivo)}")
print(f"  Total horas jornada        : {fmt_dur(total_jornada)}")
print(f"  Promedio jornada           : {fmt_dur(prom_jornada)}")
print()
print("  Nota: # = marcaciones brutas del día (antes de deduplicar).")
print("  Columnas --:-- = sin registro para ese turno.")
print("  ⚠ = día con jornada incompleta (falta entrada o salida).")
print("═" * ANCHO)
print()
