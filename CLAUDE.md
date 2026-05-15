# CLAUDE.md

## Propósito

Extractor de asistencia del biométrico ZKTeco (192.168.100.251:4370).
Descarga el mes, guarda **Excel + Parquet**, opcionalmente **MotherDuck**, y genera el **reporte HTML** (CLI o **FastAPI** sin Postgres).

Roadmap del nuevo front (React/HTML/CSS): **`PLAN.md`**.

## Estructura

```
talento/
├── biometrico/              ← sync/extract hacia raw/ (según rama actual)
│   └── schemas/paths.py     ← rutas proyecto + slugify + reports/
├── alexa/                   ← reporte_core, horarios, upload_motherduck, menú legacy
├── doc/
│   ├── reporte_template.html ← UI (colaborador, mes, día, tabla)
│   ├── reporte_web.py        ← CLI → alexa.reporte_core.write_html_report
│   └── horarios.py           ← reexporta alexa.horarios
├── legacy/
│   ├── README.md
│   ├── webapp/               ← FastAPI: / , /api/reporte.json , /api/dia.json , /health
│   └── scripts/
│       ├── publish.ps1       ← pipeline ejemplo: sync → build → (MotherDuck)
│       └── deploy/nginx-talento.example.conf
├── raw/biometrico/
├── raw/biometrico_parquet/
├── raw/reglas/
├── reports/biometrico/
├── PLAN.md
├── Makefile
├── Dockerfile
└── requirements.txt
```

## Uso rápido

```bash
uv pip install -r requirements.txt

# 1) Sync (en PC con acceso al ZK) — ajustar al script vigente en biometrico/
python biometrico/extract_biometrico.py --mes-actual

# 2) HTML estático
python doc/reporte_web.py
# → reports/biometrico/reporte_biometrico.html

# 3) Servicio web (lee raw/ cada request)
#    Opcional: $env:AUTH_PASSWORD = "..."  (Basic auth)
#    PYTHONPATH debe incluir legacy/ para importar webapp
```

PowerShell:

```powershell
$env:PYTHONPATH = ".;legacy"
uvicorn webapp.main:app --host 0.0.0.0 --port 8080 --reload
```

Make (Unix): `make serve` (define `PYTHONPATH=.:legacy`).

### MotherDuck (opcional, BI / nube)

```powershell
$env:MOTHERDUCK_TOKEN = "..."
python alexa/upload_motherduck.py --mes mayo --anio 2026
```

Tabla `talento.marcaciones_raw` en la base `md:zodiplast`. El token **no** va en el navegador; un backend o herramienta con credenciales consulta MotherDuck (p. ej. DuckDB/Java en servidor).

### Colaboradores inactivos

Editá `raw/reglas/colaboradores.yaml` (`exclude_numeros`, `exclude_colab_keys`, `exclude_name_contains`) y volvé a correr `doc/reporte_web.py` o el webapp en `legacy/webapp`.

### Deploy (talento.zodiplast.com.ec)

1. **Docker**: `docker build -t talento-biometrico .` y montá volumen con `raw/` actualizado o copiá datos al build.
2. **Nginx**: ver `legacy/scripts/deploy/nginx-talento.example.conf` (TLS + proxy a `:8080`).
3. **Pipeline diario**: `.\legacy\scripts\publish.ps1` (ajustá deploy al final; subdominio `talento.zodiplast.com.ec`).

## Flujo de datos

1. Extract / sync → Excel + Parquet bajo `raw/`.
2. `alexa.reporte_core.build_report_payload` lee **Parquet con prioridad sobre Excel** por mes si ambos existen.
3. Plantilla `doc/reporte_template.html` consume el JSON (incluye `marcaciones` por día para el selector de día).
