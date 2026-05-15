"""
main.py — FastAPI app para el reporte biométrico de asistencias.

Rutas:
    GET  /               — Reporte HTML completo (datos desde PostgreSQL)
    POST /upload         — Sube Excel biométrico, procesa y persiste en BD
    GET  /excel/{name}   — Descarga Excel generado por colaborador
    GET  /health         — Health check

Auth:
    Si AUTH_PASSWORD está definido en el entorno, todos los endpoints
    (excepto /health) requieren HTTP Basic Auth con ese password.
    Usuario puede ser cualquier valor (solo se valida el password).
"""

from __future__ import annotations

import json
import os
import secrets
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

import app.database as db
from app.processor import generar_excel_colaborador, process_excel_file

# ── Rutas de archivos ─────────────────────────────────────────────────────────

_ROOT = Path(__file__).resolve().parents[3]
_TEMPLATE_PATH = _ROOT / "legacy" / "doc" / "reporte_template.html"
_EXCEL_DIR = _ROOT / "legacy" / "reports" / "biometrico" / "excel"
_FERIADOS_PATH = _ROOT / "raw" / "reglas" / "feriados_horarios.xlsx"

# Leer template una sola vez al arrancar (no cambia en producción)
_TEMPLATE: str = _TEMPLATE_PATH.read_text(encoding="utf-8")

# ── Auth ──────────────────────────────────────────────────────────────────────

_AUTH_PASSWORD: str | None = os.environ.get("AUTH_PASSWORD")


def _check_auth(request: Request) -> bool:
    """Verifica HTTP Basic Auth. Retorna True si está autorizado."""
    if _AUTH_PASSWORD is None:
        return True  # auth deshabilitada

    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Basic "):
        return False

    import base64
    try:
        decoded = base64.b64decode(auth_header[6:]).decode("utf-8")
        _, _, provided = decoded.partition(":")
    except Exception:
        return False

    return secrets.compare_digest(provided, _AUTH_PASSWORD)


_UNAUTHORIZED = JSONResponse(
    status_code=401,
    content={"detail": "Autenticación requerida"},
    headers={"WWW-Authenticate": 'Basic realm="Biométrico Zodiplast"'},
)


# ── Lifespan ──────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(application: FastAPI):
    # Arranque
    db.init_pool()
    db.run_migration()
    _EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    yield
    # Apagado — el pool de psycopg2 se cierra solo al terminar el proceso


# ── App ───────────────────────────────────────────────────────────────────────

app = FastAPI(
    title="Reporte Biométrico — Zodiplast",
    description="Portal de asistencias para Talento y Cultura",
    version="1.0.0",
    docs_url=None,   # sin Swagger en producción
    redoc_url=None,
    lifespan=lifespan,
)


# ── Rutas ─────────────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    """Health check — no requiere autenticación."""
    ok = db.health_check()
    return {"status": "ok" if ok else "error", "db": "ok" if ok else "error"}


@app.get("/", response_class=HTMLResponse)
async def reporte(request: Request):
    """Sirve el reporte HTML completo con todos los datos desde la BD."""
    if not _check_auth(request):
        return _UNAUTHORIZED

    try:
        data = db.load_reporte_json()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error cargando datos: {exc}") from exc

    json_str = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = _TEMPLATE.replace("__DATA_JSON__", json_str)
    return HTMLResponse(content=html)


@app.post("/upload")
async def upload(request: Request, file: UploadFile = File(...)):
    """Sube un archivo Excel biométrico y lo persiste en la BD (carga incremental).

    El nombre del archivo debe contener mes y año, p.ej.:
      - ENERO-2026.xlsx
      - BIO-MARZO-2026.xlsx
    """
    if not _check_auth(request):
        return _UNAUTHORIZED

    filename = file.filename or "upload.xlsx"
    if not filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Solo se aceptan archivos .xlsx o .xls")

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="Archivo vacío")

    # Procesar Excel → estructura en memoria
    try:
        processed = process_excel_file(
            file_bytes=file_bytes,
            filename=filename,
            feriados_path=_FERIADOS_PATH if _FERIADOS_PATH.exists() else None,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error procesando Excel: {exc}") from exc

    # Persistir en BD
    try:
        db.upsert_month(processed)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error guardando en BD: {exc}") from exc

    # Generar Excel por colaborador (post-commit)
    for colab in processed.colaboradores:
        try:
            excel_filename = generar_excel_colaborador(
                colaborador=colab.display_name,
                mes_key=colab.mes_key,
                mes_label=colab.mes_label,
                dias=colab.dias,
                resumen=colab.resumen,
                output_dir=_EXCEL_DIR,
            )
            db.update_excel_filename(colab.colab_key, colab.mes_key, excel_filename)
        except Exception as exc:
            # No es crítico — los datos están en BD. El Excel se puede regenerar.
            print(f"  [WARN] No se pudo generar Excel para {colab.display_name}: {exc}")

    return {
        "ok": True,
        "mes_key": processed.mes_key,
        "mes_label": processed.mes_label,
        "colaboradores": len(processed.colaboradores),
        "dias_procesados": processed.total_dias,
    }


@app.get("/excel/{filename}")
async def download_excel(filename: str, request: Request):
    """Descarga el Excel generado para un colaborador y mes."""
    if not _check_auth(request):
        return _UNAUTHORIZED

    # Prevenir path traversal
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="Nombre de archivo inválido")

    path = _EXCEL_DIR / filename
    if not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    return FileResponse(
        path=path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )
