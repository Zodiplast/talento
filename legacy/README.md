# Legacy — **se eliminará pronto**

> **Lee primero:** [`AVISO_ELIMINACION_PRONTO.md`](AVISO_ELIMINACION_PRONTO.md) — aquí está todo lo del stack anterior; **no es la base del proyecto nuevo** (solo referencia hasta migrar o borrar).

| Ruta | Contenido |
|------|-----------|
| `legacy/doc/` | Plantilla HTML, `reporte_web.py`, `horarios.py` (CLI / plantilla del reporte viejo). |
| `legacy/alexa/` | Paquete `alexa`: `reporte_core`, MotherDuck, sync de respaldo, menú, `app/` Postgres antiguo. |
| `legacy/reports/` | HTML y Excel generados por el pipeline antiguo. |
| `legacy/webapp/` | FastAPI que leía `raw/` y rellenaba la plantilla. |
| `legacy/scripts/` | `publish.ps1`, ejemplo Nginx. |

Para `import alexa` o `import doc` hace falta **`PYTHONPATH` con la raíz del repo y `legacy/`** (véase `Makefile`, `Dockerfile`, `legacy/scripts/publish.ps1`).
