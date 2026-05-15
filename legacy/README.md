# Legacy (referencia y stack anterior)

Aquí vive lo que **ya no es el foco principal** del repo pero sirve como contexto, copiar ideas, o seguir usando hasta migrar.

| Ruta | Qué es |
|------|--------|
| `legacy/webapp/` | FastAPI que sirve el reporte HTML y JSON (`/api/reporte.json`, `/api/dia.json`). Sigue siendo válido; se movió aquí para dejar la raíz lista para el nuevo front (React/HTML/CSS). |
| `legacy/scripts/` | `publish.ps1` (sync → HTML → MotherDuck opcional) y ejemplo de Nginx. |

**Importante:** la lógica de negocio del reporte **no** está en `legacy/`; sigue en `alexa/reporte_core.py`, `alexa/horarios.py`, `doc/reporte_web.py` y `biometrico/`. Renombrar o mover `alexa/` es un paso aparte (ver `PLAN.md` en la raíz del repo).
