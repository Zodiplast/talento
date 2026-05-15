# Plan Talento — orden y nuevo front

Documento vivo: ir marcando checkboxes y añadiendo fases cuando aparezcan decisiones nuevas.

## Objetivo

- Mantener la **descarga y normalización** de datos del biométrico (Excel / Parquet, reglas, MotherDuck opcional).
- Sustituir progresivamente la **presentación** (HTML monolítico + plantilla + FastAPI de referencia) por una app **HTML + CSS + React** (o stack equivalente) alineada a tu diseño.

## Estado actual (2026-05)

- [x] Carpeta `legacy/` con `webapp` (FastAPI) y `scripts` (pipeline / nginx de ejemplo).
- [x] Rutas de ejecución actualizadas (`Makefile`, `Dockerfile`, `PYTHONPATH` para `uvicorn webapp.main:app`).
- [ ] Nuevo proyecto front (carpeta propia, p. ej. `frontend/` o repo aparte).
- [ ] Contrato de API / JSON estable entre back y front (hoy `build_report_payload` define la forma del JSON).
- [ ] Renombrar paquete `alexa/` → algo tipo `talento_core` o mover módulos a `biometrico/` (solo cuando no rompa scripts ni hábitos).

## Fases sugeridas

### 1. Inventario y contrato de datos

- Documentar qué campos consume `doc/reporte_template.html` del JSON (o extraer tipos TypeScript desde ejemplos).
- Decidir fuente de verdad en runtime: JSON estático generado en build vs API en vivo.

### 2. Nuevo front

- Crear app (Vite + React + TS recomendado; ajustar si preferís otra base).
- Replicar vistas del diseño usando los mismos datos de prueba (Parquet/JSON de ejemplo en repo o fixture).

### 3. Integración

- Opción A: seguir generando HTML estático con `doc/reporte_web.py` mientras el front madura.
- Opción B: exponer solo API (puede seguir siendo FastAPI en `legacy/webapp` o un binario mínimo en la raíz) y que el front sea SPA desplegada aparte o servida por el mismo contenedor.

### 4. Limpieza final (cuando el nuevo front esté en producción)

- Valorar eliminar o reducir `doc/reporte_template.html` si ya no se usa.
- Valorar quitar FastAPI del contenedor si el despliegue solo sirve archivos estáticos.
- Revisar duplicados bajo `alexa/reports/`, `alexa/raw/` vs `reports/`, `raw/` (mucho artefacto histórico; no versionar datos masivos si no hace falta).

## Cómo correr lo legacy hoy

Desde la raíz del repo `talento/`:

```powershell
$env:PYTHONPATH = ".;legacy"
uvicorn webapp.main:app --host 0.0.0.0 --port 8080 --reload
```

O con Make (entorno tipo Unix): `make serve`.

Pipeline PowerShell: `.\legacy\scripts\publish.ps1`

## Notas

- Añadir aquí decisiones (hosting, auth, CI, subdominio) en cuanto estén cerradas.
