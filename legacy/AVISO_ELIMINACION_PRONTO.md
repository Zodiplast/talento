# AVISO: carpeta `legacy/` — basura histórica, se borrará

**Todo lo que está dentro de `legacy/` es código y datos del enfoque anterior** (reporte HTML monolítico, FastAPI de ejemplo, paquete `alexa`, plantillas, scripts viejos, reportes generados).

**No dependas de nada aquí para el producto nuevo.** La ingesta actual vive en `biometrico/` y `raw/`.

Cuando el nuevo front (React/HTML/CSS) y el pipeline nuevo estén listos, **se puede eliminar la carpeta `legacy/` entera** (o archivarla fuera del repo). Antes de borrar, buscá referencias a `legacy/` en el código que sí vaya a producción.
