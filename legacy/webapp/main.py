"""
FastAPI: sirve el reporte HTML con datos actuales desde raw/ (Excel o Parquet).

Variables de entorno:
  AUTH_PASSWORD  — si está definido, exige HTTP Basic (cualquier usuario, ese password)
"""

from __future__ import annotations

import json
import os
import secrets
import sys
from pathlib import Path


def _repo_root() -> Path:
    """Raíz del repo (carpetas `alexa/` y `doc/`), aunque webapp viva en `legacy/webapp/`."""
    p = Path(__file__).resolve().parent
    for _ in range(8):
        if (p / "alexa").is_dir() and (p / "doc").is_dir():
            return p
        p = p.parent
    raise RuntimeError("No se encontró la raíz del repo (faltan alexa/ y doc/).")


_ROOT = _repo_root()
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
_legacy = _ROOT / "legacy"
if _legacy.is_dir() and str(_legacy) not in sys.path:
    sys.path.insert(0, str(_legacy))

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse

from alexa.reporte_core import build_report_payload
from biometrico.schemas.paths import REPORT_TEMPLATE_HTML

AUTH_PASSWORD = os.environ.get("AUTH_PASSWORD")

app = FastAPI(title="Talento biométrico", version="1.0")


def _check_auth(request: Request) -> bool:
    if not AUTH_PASSWORD:
        return True
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Basic "):
        return False
    import base64

    try:
        raw = base64.b64decode(auth[6:].encode()).decode("utf-8", errors="replace")
        _, pw = raw.split(":", 1)
    except Exception:
        return False
    return secrets.compare_digest(pw, AUTH_PASSWORD)


def _unauthorized() -> JSONResponse:
    return JSONResponse(
        status_code=401,
        content={"detail": "Autenticación requerida"},
        headers={"WWW-Authenticate": 'Basic realm="talento"'},
    )


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/api/reporte.json")
async def reporte_json(request: Request):
    if not _check_auth(request):
        return _unauthorized()
    try:
        data = build_report_payload()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return JSONResponse(content=data)


@app.get("/api/dia.json")
async def api_dia(
    request: Request,
    colaborador: str,
    mes: str,
    fecha: str,
):
    """Un día concreto (fecha YYYY-MM-DD) para un colaborador y mes_key (YYYY-MM)."""
    if not _check_auth(request):
        return _unauthorized()
    try:
        data = build_report_payload()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    mes_data = (data.get("datos") or {}).get(colaborador, {}).get(mes)
    if not mes_data:
        raise HTTPException(status_code=404, detail="Colaborador o mes no encontrado")
    for d in mes_data.get("dias") or []:
        if d.get("fecha") == fecha:
            return JSONResponse(
                {
                    "colaborador": colaborador,
                    "mes": mes,
                    "fecha": fecha,
                    "dia": d,
                    "resumen_mes": mes_data.get("resumen"),
                }
            )
    raise HTTPException(status_code=404, detail="Fecha no encontrada en ese mes")


@app.get("/", response_class=HTMLResponse)
async def reporte_html(request: Request):
    if not _check_auth(request):
        return HTMLResponse(
            content="<h1>401</h1><p>Autenticación requerida.</p>",
            status_code=401,
            headers={"WWW-Authenticate": 'Basic realm="talento"'},
        )
    try:
        data = build_report_payload()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    template = REPORT_TEMPLATE_HTML.read_text(encoding="utf-8")
    json_str = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = template.replace("__DATA_JSON__", json_str)
    return HTMLResponse(content=html)


def main():
    import uvicorn

    uvicorn.run(
        "webapp.main:app",
        host=os.environ.get("HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "8080")),
        reload=os.environ.get("UVICORN_RELOAD", "").lower() in ("1", "true", "yes"),
    )


if __name__ == "__main__":
    main()
