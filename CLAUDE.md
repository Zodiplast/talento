# CLAUDE.md

## Propósito

Extractor de asistencia del biométrico ZKTeco (192.168.100.251:4370).
Descarga el mes, guarda **Excel + Parquet**, opcionalmente **MotherDuck**, y genera el **reporte HTML** (CLI o **FastAPI** sin Postgres).

## Estructura

```
talento/
├── biometrico/
│   ├── sync_device.py        ← ZKTeco → raw/biometrico + raw/biometrico_parquet
│   ├── upload_motherduck.py  ← Parquet → MotherDuck (esquema talento)
│   ├── reporte_core.py       ← única lógica: Excel/Parquet → JSON plantilla
│   ├── horarios.py           ← turnos / jornadas (duplicado lógico de doc/horarios.py)
│   └── schemas/paths.py      ← rutas proyecto + slugify + reports/
├── doc/
│   ├── reporte_template.html ← UI (colaborador, mes, día, tabla)
│   ├── reporte_web.py        ← CLI fino → reporte_core.write_html_report
│   └── horarios.py           ← referencia / edición manual (alinear con biometrico/)
├── webapp/main.py            ← FastAPI: / , /api/reporte.json , /api/dia.json , /health
├── raw/biometrico/           ← Excel MES-YYYY.xlsx
├── raw/biometrico_parquet/   ← Parquet MES-YYYY.parquet
├── raw/reglas/
│   ├── feriados_horarios.xlsx
│   └── colaboradores.yaml    ← excluir inactivos del reporte
├── reports/biometrico/       ← HTML generado + Excel por colaborador
├── scripts/publish.ps1       ← pipeline ejemplo: sync → build → (MotherDuck)
├── scripts/deploy/nginx-talento.example.conf
├── Makefile
├── Dockerfile
└── alexa/                    ← legado; `alexa/main.py` delega a doc/reporte_web
```

## Uso rápido

```bash
uv pip install -r requirements.txt

# 1) Sync (en PC con acceso al ZK)
python biometrico/sync_device.py --mes mayo --anio 2026

# 2) HTML estático
python doc/reporte_web.py
# → reports/biometrico/reporte_biometrico.html

# 3) Servicio web (lee raw/ cada request)
#    Opcional: $env:AUTH_PASSWORD = "..."  (Basic auth)
uvicorn webapp.main:app --host 0.0.0.0 --port 8080 --reload
```

### MotherDuck (opcional, BI / nube)

```powershell
$env:MOTHERDUCK_TOKEN = "..."
python biometrico/upload_motherduck.py --mes mayo --anio 2026
```

Tabla `talento.marcaciones_raw` en la base `md:zodiplast`. El token **no** va en el navegador; un backend o herramienta con credenciales consulta MotherDuck (p. ej. DuckDB/Java en servidor).

### Colaboradores inactivos

Editá `raw/reglas/colaboradores.yaml` (`exclude_numeros`, `exclude_colab_keys`, `exclude_name_contains`) y volvé a correr `doc/reporte_web.py` o el webapp.

### Deploy (talento.zodiplast.com.ec)

1. **Docker**: `docker build -t talento-biometrico .` y montá volumen con `raw/` actualizado o copiá datos al build.
2. **Nginx**: ver `scripts/deploy/nginx-talento.example.conf` (TLS + proxy a `:8080`).
3. **Pipeline diario**: `scripts/publish.ps1 --mes mayo --anio 2026` (ajustá deploy al final; subdominio `talento.zodiplast.com.ec`).

## Flujo de datos

1. `sync_device` → Excel (reporte legacy) + Parquet (analítica / MotherDuck).
2. `reporte_core.build_report_payload` lee **Parquet con prioridad sobre Excel** por mes si ambos existen.
3. Plantilla `doc/reporte_template.html` consume el JSON (incluye `marcaciones` por día para el selector de día).
