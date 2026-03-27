"""
reporte_web.py — Genera reporte HTML interactivo del biométrico.

Uso (desde la carpeta talento/):
    python biometrico/reporte_web.py
    → genera biometrico/reporte_biometrico.html

Abre el HTML resultante en cualquier navegador.
La configuración de descansos y turnos está en biometrico/horarios.py.
"""

import calendar
import json
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

import polars as pl

sys.stdout.reconfigure(encoding="utf-8")

from horarios import (
    ALMUERZO_MIN, DESAYUNO_MIN,
    DESCANSO_MAX_MIN, JORNADA_BASE_MIN, TOLERANCIA_TURNO_MIN, TURNOS,
)

# ── Configuración ──────────────────────────────────────────────────────────────
UMBRAL_DUPLICADO_MIN = 5
RAW_DIR     = Path("raw/biometrico")
OUTPUT_HTML = Path("biometrico/reporte_biometrico.html")

MESES_ES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}
MESES_ES_INV = {v: k.capitalize() for k, v in MESES_ES.items()}

# ── Helpers de procesamiento ───────────────────────────────────────────────────

def parse_mes(nombre: str) -> tuple[int, int] | None:
    """Extrae (año, mes_num) del nombre 'ENERO-2026.xlsx'."""
    stem = Path(nombre).stem.upper()
    parts = stem.split("-")
    if len(parts) != 2:
        return None
    mes_str, anio_str = parts
    if mes_str not in MESES_ES:
        return None
    try:
        return int(anio_str), MESES_ES[mes_str]
    except ValueError:
        return None


def dedup(filas: list[dict]) -> list[dict]:
    """Elimina marcaciones duplicadas dentro del mismo día (< umbral minutos)."""
    out = []
    for i, f in enumerate(filas):
        if i == 0:
            out.append(f)
            continue
        prev = out[-1]
        if f["fecha"] == prev["fecha"]:
            diff = (f["Tiempo"] - prev["Tiempo"]).total_seconds() / 60
            if diff < UMBRAL_DUPLICADO_MIN:
                continue
        out.append(f)
    return out


def procesar_dia(marcas: list) -> dict:
    """
    Asigna roles a las marcaciones del día y calcula duraciones.
    Misma lógica que main.py.
    """
    n = len(marcas)
    entrada = salida = sal_des = ent_des = sal_alm = ent_alm = None

    if n >= 1: entrada = marcas[0]
    if n >= 2: salida  = marcas[-1]
    if n == 3:
        sal_alm = marcas[1]
    elif n == 4:
        sal_alm, ent_alm = marcas[1], marcas[2]
    elif n == 5:
        sal_alm, ent_alm = marcas[2], marcas[3]
    elif n >= 6:
        sal_des, ent_des = marcas[1], marcas[2]
        sal_alm, ent_alm = marcas[3], marcas[4]

    def dm(a, b):
        return (b - a).total_seconds() / 60 if (a and b) else None

    t_des = dm(sal_des, ent_des)
    t_alm = dm(sal_alm, ent_alm)
    t_jor = dm(entrada, salida)
    t_ef  = (t_jor - (t_des or 0) - (t_alm or 0)) if t_jor is not None else None

    def ft(dt): return dt.strftime("%H:%M") if dt else None
    def rr(v):  return round(v, 1) if v is not None else None

    t_descanso_total = (t_des or 0) + (t_alm or 0)
    exceso_descanso  = round(t_descanso_total - DESCANSO_MAX_MIN, 1) if t_jor is not None else None
    horas_extra      = round(max(0.0, (t_ef or 0) - JORNADA_BASE_MIN), 1) if t_ef is not None else None

    return {
        "entrada":         ft(entrada),  "sal_des":    ft(sal_des),
        "ent_des":         ft(ent_des),  "sal_alm":    ft(sal_alm),
        "ent_alm":         ft(ent_alm),  "salida":     ft(salida),
        "t_desayuno":      rr(t_des),    "t_almuerzo": rr(t_alm),
        "t_jornada":       rr(t_jor),    "t_efectivo": rr(t_ef),
        "exceso_descanso": exceso_descanso,
        "horas_extra":     horas_extra,
        "n_marcas":        n,
    }


def detectar_turno(entrada_str: str | None) -> str | None:
    """Devuelve la clave del turno más cercano a la hora de entrada, o None."""
    if not entrada_str:
        return None
    h, m = map(int, entrada_str.split(":"))
    entrada_min = h * 60 + m
    mejor, menor_diff = None, TOLERANCIA_TURNO_MIN + 1
    for key, turno in TURNOS.items():
        th, tm = map(int, turno.entrada.split(":"))
        diff = abs(entrada_min - (th * 60 + tm))
        if diff < menor_diff:
            menor_diff, mejor = diff, key
    return mejor if menor_diff <= TOLERANCIA_TURNO_MIN else None


# ── Carga y procesamiento de todos los archivos ────────────────────────────────

def cargar_datos() -> dict:
    archivos = sorted(RAW_DIR.glob("*.xlsx"))
    archivos = [a for a in archivos if not a.name.startswith("~$")]  # ignorar temporales Excel
    if not archivos:
        print(f"[ERROR] No hay archivos .xlsx en {RAW_DIR.resolve()}")
        sys.exit(1)

    meses_map = {parse_mes(a.name): a for a in archivos if parse_mes(a.name)}
    if not meses_map:
        print("[ERROR] Ningún archivo tiene el formato MESNAME-YYYY.xlsx")
        sys.exit(1)

    todos_colab: set[str] = set()
    datos: dict[str, dict[str, dict]] = defaultdict(dict)

    for (anio, mes_n), arch in sorted(meses_map.items()):
        mes_key = f"{anio}-{mes_n:02d}"
        print(f"  → Procesando {arch.name}  ({mes_key})")

        df = pl.read_excel(arch, sheet_name="Sheet")
        df = df.with_columns([
            pl.col("Tiempo").cast(pl.Date).alias("fecha"),
        ])

        nombres = sorted(df["Nombre"].unique().to_list())
        todos_colab.update(nombres)

        for nombre in nombres:
            sub = (
                df.filter(pl.col("Nombre") == nombre)
                .sort(["fecha", "Tiempo"])
                .to_dicts()
            )
            sub = dedup(sub)

            dias_marcas: dict[date, list] = defaultdict(list)
            for f in sub:
                dias_marcas[f["fecha"]].append(f["Tiempo"])

            _, total_dias = calendar.monthrange(anio, mes_n)
            dias_mes = []

            for d in range(1, total_dias + 1):
                fecha = date(anio, mes_n, d)
                dow = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"][fecha.weekday()]
                marcas = sorted(dias_marcas.get(fecha, []))

                if not marcas:
                    estado = "libre"
                    info = {
                        "entrada": None, "sal_des": None, "ent_des": None,
                        "sal_alm": None, "ent_alm": None, "salida":  None,
                        "t_desayuno": None, "t_almuerzo": None,
                        "t_jornada": None,  "t_efectivo": None,
                        "exceso_descanso": None, "horas_extra": None,
                        "n_marcas": 0,
                    }
                else:
                    info   = procesar_dia(marcas)
                    estado = "completo" if (info["entrada"] and info["salida"]) else "incompleto"

                dias_mes.append({
                    "fecha":  fecha.strftime("%Y-%m-%d"),
                    "dia":    d,
                    "dow":    dow,
                    "estado": estado,
                    "turno":  detectar_turno(info.get("entrada")) if estado != "libre" else None,
                    **info,
                })

            comp   = sum(1 for d in dias_mes if d["estado"] == "completo")
            incomp = sum(1 for d in dias_mes if d["estado"] == "incompleto")
            libre  = sum(1 for d in dias_mes if d["estado"] == "libre")
            t_ef   = sum(d["t_efectivo"] or 0 for d in dias_mes)
            t_jor  = sum(d["t_jornada"]  or 0 for d in dias_mes)

            datos[nombre][mes_key] = {
                "dias": dias_mes,
                "resumen": {
                    "completos":       comp,
                    "incompletos":     incomp,
                    "libres":          libre,
                    "total_efectivo":  round(t_ef),
                    "total_jornada":   round(t_jor),
                    "prom_efectivo":   round(t_ef / comp) if comp else 0,
                    "prom_jornada":    round(t_jor / comp) if comp else 0,
                },
            }

    meses_lista = [
        {"label": f"{MESES_ES_INV[mes_n]} {anio}", "value": f"{anio}-{mes_n:02d}"}
        for (anio, mes_n) in sorted(meses_map.keys())
    ]

    return {
        "colaboradores": sorted(todos_colab),
        "meses":         meses_lista,
        "datos":         dict(datos),
        "config": {
            "desayuno_min":       DESAYUNO_MIN,
            "almuerzo_min":       ALMUERZO_MIN,
            "descanso_max_min":   DESCANSO_MAX_MIN,
            "jornada_base_min":   JORNADA_BASE_MIN,
            "turnos": {k: {"nombre": t.nombre, "entrada": t.entrada, "salida": t.salida}
                       for k, t in TURNOS.items()},
        },
    }


# ── Plantilla HTML ─────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Reporte Biométrico — Zodiplast</title>
  <style>
    :root {
      --blue-950: #0f2045;
      --blue-900: #1e3a5f;
      --blue-700: #1d4ed8;
      --blue-600: #2563eb;
      --blue-100: #dbeafe;
      --blue-50:  #eff6ff;
      --green-100: #dcfce7;
      --green-800: #166534;
      --amber-100: #fef3c7;
      --amber-800: #92400e;
      --slate-50:  #f8fafc;
      --slate-100: #f1f5f9;
      --slate-200: #e2e8f0;
      --slate-300: #cbd5e1;
      --slate-400: #94a3b8;
      --slate-500: #64748b;
      --slate-600: #475569;
      --slate-700: #334155;
      --slate-800: #1e293b;
      --bg: #eef2f7;
      --radius: 10px;
      --shadow: 0 1px 3px rgba(0,0,0,.08), 0 1px 2px rgba(0,0,0,.05);
      --shadow-md: 0 4px 12px rgba(0,0,0,.1);
    }

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
      background: var(--bg);
      color: var(--slate-800);
      font-size: 13.5px;
      line-height: 1.5;
    }

    /* ── Header ── */
    header {
      background: linear-gradient(135deg, var(--blue-950) 0%, var(--blue-700) 100%);
      color: #fff;
      padding: 18px 32px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 24px;
      box-shadow: 0 2px 10px rgba(0,0,0,.25);
      position: sticky;
      top: 0;
      z-index: 100;
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 14px;
    }

    .brand-icon {
      width: 40px;
      height: 40px;
      background: rgba(255,255,255,.15);
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 20px;
    }

    .brand h1 {
      font-size: 18px;
      font-weight: 700;
      letter-spacing: -.3px;
    }

    .brand p {
      font-size: 11px;
      opacity: .65;
      margin-top: 1px;
    }

    .controls {
      display: flex;
      align-items: flex-end;
      gap: 16px;
      flex-wrap: wrap;
    }

    .ctrl {
      display: flex;
      flex-direction: column;
      gap: 5px;
    }

    .ctrl label {
      font-size: 10px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .6px;
      opacity: .65;
    }

    .ctrl select {
      background: rgba(255,255,255,.12);
      border: 1px solid rgba(255,255,255,.25);
      color: #fff;
      padding: 7px 12px;
      border-radius: 7px;
      font-size: 13px;
      min-width: 200px;
      cursor: pointer;
      outline: none;
      transition: background .15s, border-color .15s;
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='rgba(255,255,255,.7)' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 10px center;
      padding-right: 30px;
    }

    .ctrl select:hover, .ctrl select:focus {
      background-color: rgba(255,255,255,.22);
      border-color: rgba(255,255,255,.5);
    }

    .ctrl select option {
      background: var(--blue-900);
      color: #fff;
    }

    /* ── Main ── */
    main {
      max-width: 1440px;
      margin: 0 auto;
      padding: 28px 32px;
    }

    /* ── Stats ── */
    .stats {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
      gap: 14px;
      margin-bottom: 22px;
    }

    .card {
      background: #fff;
      border-radius: var(--radius);
      padding: 18px 22px;
      box-shadow: var(--shadow);
    }

    .card-label {
      font-size: 10.5px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .5px;
      color: var(--slate-400);
      margin-bottom: 6px;
    }

    .card-value {
      font-size: 30px;
      font-weight: 800;
      line-height: 1;
      letter-spacing: -1px;
    }

    .card-sub {
      font-size: 11px;
      color: var(--slate-400);
      margin-top: 4px;
    }

    .card.completo   .card-value { color: #16a34a; }
    .card.incompleto .card-value { color: #d97706; }
    .card.libre      .card-value { color: var(--slate-500); }
    .card.efectivo   .card-value { color: var(--blue-600); font-size: 24px; }
    .card.prom       .card-value { color: var(--slate-700); font-size: 22px; }

    /* ── Table card ── */
    .tcard {
      background: #fff;
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      overflow: hidden;
    }

    .tcard-head {
      padding: 16px 22px;
      border-bottom: 1px solid var(--slate-100);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      flex-wrap: wrap;
    }

    .tcard-head h2 {
      font-size: 15px;
      font-weight: 600;
      color: var(--slate-700);
    }

    .legend {
      display: flex;
      gap: 18px;
      flex-wrap: wrap;
    }

    .leg {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 11.5px;
      color: var(--slate-500);
    }

    .leg-dot {
      width: 9px;
      height: 9px;
      border-radius: 50%;
      flex-shrink: 0;
    }

    .dot-c { background: #22c55e; }
    .dot-i { background: #f59e0b; }
    .dot-l { background: var(--slate-300); }

    .table-wrap { overflow-x: auto; }

    table {
      width: 100%;
      border-collapse: collapse;
    }

    thead th {
      background: var(--blue-50);
      color: var(--blue-600);
      font-size: 10px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .5px;
      padding: 10px 14px;
      text-align: center;
      border-bottom: 2px solid var(--blue-100);
      white-space: nowrap;
      position: sticky;
      top: 0; /* makes header sticky inside scroll container - best effort */
    }

    thead th:first-child,
    thead th:nth-child(2) { text-align: left; }

    tbody tr {
      border-bottom: 1px solid var(--slate-100);
      transition: filter .1s;
    }

    tbody tr:last-child { border-bottom: none; }
    tbody tr:hover { filter: brightness(.97); }

    tbody td {
      padding: 9px 14px;
      text-align: center;
      white-space: nowrap;
    }

    tbody td:first-child { text-align: left; }
    tbody td:nth-child(2) { text-align: left; }

    tr.completo   { background: #fff; }
    tr.incompleto { background: #fffdf0; }
    tr.libre      { background: var(--slate-50); }
    tr.libre td   { color: var(--slate-400); }

    /* fecha cell */
    .fecha-cell {
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .dow {
      font-size: 9.5px;
      font-weight: 700;
      text-transform: uppercase;
      color: var(--slate-400);
      width: 28px;
    }

    .libre .dow { color: var(--slate-300); }

    .dnum {
      font-size: 14px;
      font-weight: 700;
      color: var(--slate-700);
      min-width: 18px;
    }

    .libre .dnum { color: var(--slate-400); }

    /* badges */
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 3px 9px;
      border-radius: 999px;
      font-size: 10px;
      font-weight: 700;
      letter-spacing: .3px;
      text-transform: uppercase;
    }

    .bc { background: #dcfce7; color: #15803d; }
    .bi { background: #fef3c7; color: #b45309; }
    .bl { background: var(--slate-100); color: var(--slate-500); }

    /* turno badges */
    .bt-madrugada  { background: #e0e7ff; color: #3730a3; }
    .bt-ordinario  { background: #dcfce7; color: #166534; }
    .bt-inventario { background: #fef9c3; color: #854d0e; }
    .bt-domingo    { background: #fce7f3; color: #9d174d; }
    .bt-unknown    { background: var(--slate-100); color: var(--slate-500); }

    /* exceso de descanso → fila magenta */
    tr.exceso-descanso { background: #fdf0fb !important; }
    tr.exceso-descanso:hover { filter: brightness(.96) !important; }
    .exceso-val {
      color: #a21caf;
      font-weight: 700;
    }
    .exceso-cell {
      color: #a21caf !important;
      font-weight: 700;
    }

    /* horas extra */
    .extra-val { color: #0e7490; font-weight: 700; }

    /* time & duration */
    .mono {
      font-family: ui-monospace, "SF Mono", Consolas, monospace;
      font-size: 12px;
    }

    .ef { color: var(--blue-600); font-weight: 700; }
    .nd { color: var(--slate-300); }
    .nm { color: var(--slate-400); font-size: 11px; }

    /* empty */
    .empty {
      text-align: center;
      padding: 72px 20px;
      color: var(--slate-400);
    }

    .empty-icon { font-size: 48px; margin-bottom: 14px; }
    .empty p    { font-size: 15px; }

    /* footer */
    footer {
      text-align: center;
      padding: 28px;
      color: var(--slate-400);
      font-size: 11px;
    }

    @media (max-width: 768px) {
      header { flex-direction: column; align-items: flex-start; padding: 14px 18px; }
      main   { padding: 16px 14px; }
    }
  </style>
</head>
<body>

<header>
  <div class="brand">
    <div class="brand-icon">⏱</div>
    <div>
      <h1>Reporte Biométrico</h1>
      <p>Zodiplast · Control de Asistencia</p>
    </div>
  </div>
  <div class="controls">
    <div class="ctrl">
      <label>Colaborador</label>
      <select id="selColab"><option value="">— Seleccionar —</option></select>
    </div>
    <div class="ctrl">
      <label>Mes</label>
      <select id="selMes"><option value="">— Seleccionar —</option></select>
    </div>
  </div>
</header>

<main>
  <div id="statsWrap" class="stats" style="display:none">
    <div class="card completo">
      <div class="card-label">Días completos</div>
      <div class="card-value" id="stComp">—</div>
    </div>
    <div class="card incompleto">
      <div class="card-label">Incompletos</div>
      <div class="card-value" id="stIncomp">—</div>
      <div class="card-sub">Sin entrada o salida</div>
    </div>
    <div class="card libre">
      <div class="card-label">Días libres / ausentes</div>
      <div class="card-value" id="stLibre">—</div>
    </div>
    <div class="card efectivo">
      <div class="card-label">Total horas efectivas</div>
      <div class="card-value" id="stTotEf">—</div>
      <div class="card-sub" id="stTotJor"></div>
    </div>
    <div class="card prom">
      <div class="card-label">Promedio efectivo / día</div>
      <div class="card-value" id="stPromEf">—</div>
      <div class="card-sub" id="stPromJor"></div>
    </div>
  </div>

  <div class="tcard">
    <div class="tcard-head">
      <h2 id="tTitle">Selecciona un colaborador y mes para ver el reporte</h2>
      <div class="legend">
        <div class="leg"><span class="leg-dot dot-c"></span> Completo</div>
        <div class="leg"><span class="leg-dot dot-i"></span> Incompleto</div>
        <div class="leg"><span class="leg-dot dot-l"></span> Libre / sin registro</div>
      </div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Fecha</th>
            <th>Estado</th>
            <th>Turno</th>
            <th>Entrada</th>
            <th>S. Desayuno</th>
            <th>E. Desayuno</th>
            <th>S. Almuerzo</th>
            <th>E. Almuerzo</th>
            <th>Salida</th>
            <th>Desayuno</th>
            <th>Almuerzo</th>
            <th>Efectivo</th>
            <th>Jornada</th>
            <th>H. Extra</th>
            <th>#</th>
          </tr>
        </thead>
        <tbody id="tbody">
          <tr>
            <td colspan="15">
              <div class="empty">
                <div class="empty-icon">📋</div>
                <p>Selecciona un colaborador y un mes para ver el reporte</p>
              </div>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</main>

<footer id="footer"></footer>

<script>
const DATA = __DATA_JSON__;

const selColab  = document.getElementById('selColab');
const selMes    = document.getElementById('selMes');
const tbody     = document.getElementById('tbody');
const tTitle    = document.getElementById('tTitle');
const statsWrap = document.getElementById('statsWrap');
const footer    = document.getElementById('footer');

// ── Config de descansos
const CFG = DATA.config;
footer.textContent =
  `Generado el ${new Date().toLocaleDateString('es', {year:'numeric',month:'long',day:'numeric'})}` +
  ` · Desayuno: ${CFG.desayuno_min} min · Almuerzo: ${CFG.almuerzo_min} min`;

// ── Poblar selects
DATA.colaboradores.forEach(c => {
  selColab.insertAdjacentHTML('beforeend', `<option value="${esc(c)}">${title(c)}</option>`);
});
DATA.meses.forEach(m => {
  selMes.insertAdjacentHTML('beforeend', `<option value="${esc(m.value)}">${esc(m.label)}</option>`);
});

// Auto-seleccionar primeros valores
if (DATA.meses.length)         selMes.value   = DATA.meses[0].value;
if (DATA.colaboradores.length) selColab.value = DATA.colaboradores[0];

selColab.addEventListener('change', render);
selMes.addEventListener('change', render);
render();

// ── Funciones
function render() {
  const colab = selColab.value;
  const mes   = selMes.value;

  if (!colab || !mes) {
    showEmpty('📋', 'Selecciona un colaborador y un mes');
    statsWrap.style.display = 'none';
    return;
  }

  const mesData = DATA.datos?.[colab]?.[mes];
  if (!mesData) {
    showEmpty('🔍', 'Sin datos para este colaborador en este mes');
    statsWrap.style.display = 'none';
    return;
  }

  const mesLabel = DATA.meses.find(m => m.value === mes)?.label ?? mes;
  tTitle.textContent = `${title(colab)}  ·  ${mesLabel}`;

  // Stats
  const r = mesData.resumen;
  document.getElementById('stComp').textContent   = r.completos;
  document.getElementById('stIncomp').textContent = r.incompletos;
  document.getElementById('stLibre').textContent  = r.libres;
  document.getElementById('stTotEf').textContent  = fmtH(r.total_efectivo);
  document.getElementById('stTotJor').textContent = `Jornada total: ${fmtH(r.total_jornada)}`;
  document.getElementById('stPromEf').textContent = fmtH(r.prom_efectivo);
  document.getElementById('stPromJor').textContent= `Jornada prom: ${fmtH(r.prom_jornada)}`;
  statsWrap.style.display = 'grid';

  // Filas
  tbody.innerHTML = mesData.dias.map(d => {
    const libre   = d.estado === 'libre';
    const exceso  = !libre && d.exceso_descanso !== null && d.exceso_descanso > 0;

    const badgeCls = {completo:'bc', incompleto:'bi', libre:'bl'}[d.estado];
    const badgeTxt = {completo:'✓ Completo', incompleto:'⚠ Incompleto', libre:'Libre'}[d.estado];

    // Turno badge
    const turnoInfo = DATA.config.turnos[d.turno] ?? null;
    const turnoCls  = d.turno ? `bt-${d.turno}` : 'bt-unknown';
    const turnoHtml = libre
      ? nd()
      : turnoInfo
        ? `<span class="badge ${turnoCls}">${turnoInfo.nombre}</span>`
        : `<span class="badge bt-unknown">?</span>`;

    const rowCls = [d.estado, exceso ? 'exceso-descanso' : ''].join(' ').trim();

    const t = (v) => v ? `<span class="mono">${v}</span>` : nd();
    const dur = (v, extra='') =>
      v !== null && v !== undefined
        ? `<span class="mono ${extra}">${fmtMin(v)}</span>`
        : nd();

    // Celdas de descanso: magenta si hay exceso
    const desayunoCls = exceso ? 'exceso-cell' : '';
    const almuerzoCls = exceso ? 'exceso-cell' : '';
    const desayunoCell = libre ? nd() : `<span class="mono ${desayunoCls}">${d.t_desayuno !== null ? fmtMin(d.t_desayuno) : '—'}</span>`;
    const almuerzoCell = libre ? nd() : `<span class="mono ${almuerzoCls}">${d.t_almuerzo !== null ? fmtMin(d.t_almuerzo) : '—'}</span>`;

    // Horas extra
    let extraCell = nd();
    if (!libre && d.horas_extra !== null) {
      extraCell = d.horas_extra > 0
        ? `<span class="mono extra-val">+${fmtMin(d.horas_extra)}</span>`
        : `<span class="nd">—</span>`;
    }

    return `<tr class="${rowCls}">
      <td>
        <div class="fecha-cell">
          <span class="dow">${d.dow}</span>
          <span class="dnum">${d.dia}</span>
        </div>
      </td>
      <td><span class="badge ${badgeCls}">${badgeTxt}</span></td>
      <td>${turnoHtml}</td>
      <td>${libre ? nd() : t(d.entrada)}</td>
      <td>${libre ? nd() : t(d.sal_des)}</td>
      <td>${libre ? nd() : t(d.ent_des)}</td>
      <td>${libre ? nd() : t(d.sal_alm)}</td>
      <td>${libre ? nd() : t(d.ent_alm)}</td>
      <td>${libre ? nd() : t(d.salida)}</td>
      <td>${desayunoCell}</td>
      <td>${almuerzoCell}</td>
      <td>${libre ? nd() : dur(d.t_efectivo, 'ef')}</td>
      <td>${libre ? nd() : dur(d.t_jornada)}</td>
      <td>${extraCell}</td>
      <td>${libre ? '' : `<span class="nm">${d.n_marcas}</span>`}</td>
    </tr>`;
  }).join('');
}

function showEmpty(icon, msg) {
  tbody.innerHTML = `<tr><td colspan="15"><div class="empty">
    <div class="empty-icon">${icon}</div><p>${msg}</p></div></td></tr>`;
}

function nd() { return '<span class="nd">—</span>'; }

function fmtMin(min) {
  if (min === null || min === undefined) return '—';
  const h = Math.floor(min / 60);
  const m = Math.round(min % 60);
  return h > 0 ? `${h}h ${String(m).padStart(2,'0')}m` : `${m}m`;
}

function fmtH(min) {
  if (!min && min !== 0) return '—';
  const h = Math.floor(min / 60);
  const m = Math.round(min % 60);
  return h > 0 ? `${h}h ${String(m).padStart(2,'0')}m` : `${m}m`;
}

function title(str) {
  return str.toLowerCase().replace(/(?:^|\s)\S/g, c => c.toUpperCase());
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>
"""


# ── Punto de entrada ───────────────────────────────────────────────────────────

def main() -> None:
    print("Reporte Biométrico Web")
    print("─" * 40)
    datos = cargar_datos()

    n_colab = len(datos["colaboradores"])
    n_meses = len(datos["meses"])
    print(f"  Colaboradores: {n_colab}")
    print(f"  Meses:         {n_meses}  ({', '.join(m['label'] for m in datos['meses'])})")

    # Inyectar datos en el HTML
    json_str = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    html = HTML.replace("__DATA_JSON__", json_str)

    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_HTML.write_text(html, encoding="utf-8")

    print(f"\n  ✓ Reporte generado: {OUTPUT_HTML.resolve()}")
    print("  Abre el archivo en tu navegador.\n")


if __name__ == "__main__":
    main()
