# CLAUDE.md

## Propósito

Extractor de asistencia del biométrico ZKTeco (192.168.100.251:4370).
**Producto actual en la raíz:** `biometrico/` (ingesta → `raw/`).

Todo lo demás del flujo viejo (plantilla HTML, `alexa`, reportes, FastAPI de ejemplo) está en **`legacy/`** y **está marcado para borrarse** cuando tengas el stack nuevo (ver `legacy/AVISO_ELIMINACION_PRONTO.md`).

Roadmap: **`PLAN.md`**.

## Estructura

```
talento/
├── biometrico/              ← ingesta ZKTeco → raw/
├── raw/
├── legacy/                  ← stack anterior; ver AVISO_ELIMINACION_PRONTO.md
│   ├── doc/                 ← plantilla + CLI reporte (antes en raíz)
│   ├── alexa/
│   ├── reports/
│   ├── webapp/
│   └── scripts/
├── PLAN.md
├── Makefile
├── Dockerfile
└── requirements.txt
```

### Variables de entorno (hosting)

- Copiá `.env.example` → `.env` (no se versiona).
- **cPanel / UAPI:** `../hosting_web/.env` con `CPANEL_BASE_URL`, `CPANEL_USER`, `CPANEL_API_TOKEN`, `CPANEL_DOMAIN`.
- **Subdominio y ruta remota:** `UPLOAD_SUBDOMINIO=talento` → `https://talento.zodiplast.com.ec`, carpeta típica `public_html/talento` (ajustá si tu cPanel usa otra ruta).

## Uso rápido

```bash
uv pip install -r requirements.txt

python biometrico/extract_biometrico.py --mes-actual
```

### Reporte HTML legacy (opcional)

```bash
python legacy/doc/reporte_web.py
# → legacy/reports/biometrico/reporte_biometrico.html
```

### FastAPI legacy

```powershell
$env:PYTHONPATH = ".;legacy"
uvicorn webapp.main:app --host 0.0.0.0 --port 8080 --reload
```

### MotherDuck (legacy)

```powershell
$env:PYTHONPATH = ".;legacy"
python legacy/alexa/upload_motherduck.py --mes mayo --anio 2026
```

### Pipeline PowerShell

`.\legacy\scripts\publish.ps1`

### Docker

`docker build -t talento-biometrico .` — copia `legacy/` completo; `PYTHONPATH=/app:/app/legacy`.
