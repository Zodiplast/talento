# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Menú interactivo (reporte CLI, servidor, ver turnos)
python main.py

# Desarrollo con 1Password + hot-reload
bash dev.sh

# Producción con Docker
docker-compose up -d

# Generar reporte HTML directamente (sin servidor)
python biometrico/reporte_web.py

# Instalar dependencias
pip install -r requirements.txt
```

Variables de entorno requeridas (`DATABASE_URL`, `AUTH_PASSWORD` opcional) — ver `.env.example`.

## Arquitectura

El proyecto tiene **dos modos de operación independientes** sobre la misma lógica de negocio:

### Modo CLI (`biometrico/reporte_web.py`)
Lee los Excel de `raw/biometrico/`, procesa todo en memoria y genera `reports/biometrico/reporte_biometrico.html` inyectando el JSON en el placeholder `__DATA_JSON__` de `biometrico/reporte_template.html`. No requiere base de datos.

### Modo servidor (`app/main.py` → FastAPI)
Recibe un Excel via `POST /upload`, lo procesa con `app/processor.py` y persiste en PostgreSQL (`talento.*` schema). `GET /` sirve el mismo HTML pero con datos de la BD. El template se carga una sola vez al arrancar (`lifespan`).

### Flujo de procesamiento Excel (compartido)
`parse_mes(filename)` → `_leer_excel()` / `process_excel_file()` → `dedup()` (5 min threshold) → `procesar_dia()` (asigna roles a marcas: entrada, sal_des, ent_des, sal_alm, ent_alm, salida) → cálculo de jornada, descansos y horas extra → `generar_excel_colaborador()` (2 hojas: Resumen + Detalle).

La lógica de `procesar_dia()` asigna roles por **posición** de las marcas (n=1 solo entrada, n=4 almuerzo completo, n=6 desayuno+almuerzo completo, etc.).

## Configuración de negocio

Todo lo que no es lógica de código está en **`biometrico/horarios.py`**: tiempos de descanso, jornadas por día de semana, catálogo de turnos. Editar ahí, no en el procesador.

Feriados vienen de `raw/reglas/feriados_horarios.xlsx` (hoja `feriados`, columnas `Fecha` y `Motivo del Feriado`).

## Base de datos

Schema `talento.*` en la instancia PostgreSQL compartida (para no colisionar con otros servicios). Las migraciones (`migrations/001_init.sql`) son idempotentes y se ejecutan automáticamente al arrancar vía `db.run_migration()`.

Tablas: `colaboradores`, `meses`, `asistencia_dias`, `resumen_mes`. La carga es atómica por mes: UPSERT de encabezados + DELETE+bulk INSERT de días.

La clave de colaborador (`colab_key`) se construye como `id:<numero>` si tiene número de empleado, o `nom:<nombre_normalizado>` si no.

## Convenciones

- **Nombre de archivo Excel**: debe contener mes en español y año — `ENERO-2026.xlsx` o `BIO-MARZO-2026.xlsx`. `parse_mes()` busca el primer token válido como mes.
- **Auth**: HTTP Basic Auth opcional. Si `AUTH_PASSWORD` está definido, aplica a todos los endpoints excepto `/health`. El usuario puede ser cualquier valor.
- **Archivos `.xls` antiguos**: el procesador normaliza columnas con encoding corrupto (`_normalizar_columnas`) y parsea `Tiempo` como string `DD/MM/YYYY H:MM:SS`.
- **Imports**: `biometrico/` es un paquete Python (tiene `__init__.py`). Usar imports absolutos (`from biometrico.horarios import ...`), no relativos.
