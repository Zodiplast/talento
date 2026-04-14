# GEMINI.md - talento

FastAPI app para el reporte biométrico de asistencias.

## Commands

```bash
# Iniciar entorno de desarrollo
bash dev.sh

# Levantar con Docker
docker-compose up -d

# Instalar dependencias localmente
pip install -r requirements.txt
```

## Endpoints Principales

- `GET /` — Reporte HTML completo (datos desde PostgreSQL)
- `POST /upload` — Sube Excel biométrico, procesa y persiste en BD
- `GET /excel/{name}` — Descarga Excel generado por colaborador
- `GET /health` — Health check

## Architecture

- **`app/main.py`** — Punto de entrada FastAPI y definición de rutas.
- **`app/processor.py`** — Lógica de procesamiento de archivos Excel biométricos.
- **`app/database.py`** — Conexión y operaciones con la base de datos PostgreSQL.
- **`app/models.py`** — Modelos de SQLAlchemy para la base de datos.
- **`biometrico/`** — Directorio para almacenar archivos relacionados con el biométrico.
- **`migrations/`** — Migraciones de base de datos (si aplica).

## Environment Variables

- `DATABASE_URL` — URL de conexión a PostgreSQL.
- `AUTH_PASSWORD` — Password para Basic Auth (opcional).

## Conventions

- **Auth:** Si `AUTH_PASSWORD` está definido, todos los endpoints (excepto `/health`) requieren HTTP Basic Auth.
- **Procesamiento:** El Excel biométrico se procesa en `processor.py` antes de persistirse.
