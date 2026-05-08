-- migrations/001_init.sql
-- Schema talento.* para el reporte biométrico de asistencias.
-- Usa el prefijo talento. para no colisionar con tablas de la Intranet
-- en la misma instancia PostgreSQL compartida (base cortana).
-- Todos los DDL usan IF NOT EXISTS → seguro re-ejecutar en cada arranque.

CREATE SCHEMA IF NOT EXISTS talento;

-- ── Colaboradores ─────────────────────────────────────────────────────────────
-- colab_key es la identidad estable: "id:00042" o "nom:JOSE LUIS GARCIA"
-- display_name es el nombre legible en Title Case (deduplicado entre meses)
CREATE TABLE IF NOT EXISTS talento.colaboradores (
    id           SERIAL PRIMARY KEY,
    colab_key    TEXT NOT NULL UNIQUE,
    display_name TEXT NOT NULL,
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- ── Meses procesados ──────────────────────────────────────────────────────────
-- Una fila por mes cargado. Re-subir el mismo mes reemplaza los datos
-- de ese mes sin tocar los demás (carga incremental).
CREATE TABLE IF NOT EXISTS talento.meses (
    id            SERIAL PRIMARY KEY,
    mes_key       TEXT NOT NULL UNIQUE,  -- "2026-01"
    mes_label     TEXT NOT NULL,         -- "Enero 2026"
    original_file TEXT,
    uploaded_at   TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- ── Marcaciones diarias ───────────────────────────────────────────────────────
-- Una fila por (colaborador, fecha). Los tiempos se guardan como "HH:MM"
-- (string) para coincidir exactamente con la salida del procesador.
CREATE TABLE IF NOT EXISTS talento.asistencia_dias (
    id                SERIAL PRIMARY KEY,
    colaborador_id    INTEGER NOT NULL REFERENCES talento.colaboradores(id),
    mes_key           TEXT    NOT NULL REFERENCES talento.meses(mes_key),
    fecha             DATE    NOT NULL,
    dia               SMALLINT NOT NULL,
    dow               TEXT    NOT NULL,   -- "Lun","Mar","Mié","Jue","Vie","Sáb","Dom"
    estado            TEXT    NOT NULL,   -- "completo","incompleto","libre"
    motivo_feriado    TEXT,

    -- Marcaciones (HH:MM o NULL)
    entrada           TEXT,
    sal_des           TEXT,
    ent_des           TEXT,
    sal_alm           TEXT,
    ent_alm           TEXT,
    salida            TEXT,

    -- Duraciones en minutos (nullable)
    t_desayuno        NUMERIC(7,1),
    t_almuerzo        NUMERIC(7,1),
    t_jornada         NUMERIC(7,1),
    t_efectivo        NUMERIC(7,1),
    jornada_objetivo  SMALLINT,
    exceso_desayuno   NUMERIC(7,1),
    exceso_almuerzo   NUMERIC(7,1),
    exceso_descanso   NUMERIC(7,1),
    horas_extra       NUMERIC(7,1),
    resaltar_descanso BOOLEAN NOT NULL DEFAULT false,

    UNIQUE (colaborador_id, fecha)
);

CREATE INDEX IF NOT EXISTS idx_asistencia_colab_mes
    ON talento.asistencia_dias (colaborador_id, mes_key);

-- ── Resumen mensual por colaborador ──────────────────────────────────────────
-- Precalculado en cada carga para evitar re-agregar en cada page load.
-- excel_filename: solo el nombre del archivo (sin ruta), p.ej.
--   "reporte_biometrico_jose_garcia_2026-01.xlsx"
CREATE TABLE IF NOT EXISTS talento.resumen_mes (
    id                    SERIAL PRIMARY KEY,
    colaborador_id        INTEGER NOT NULL REFERENCES talento.colaboradores(id),
    mes_key               TEXT    NOT NULL REFERENCES talento.meses(mes_key),
    completos             SMALLINT NOT NULL DEFAULT 0,
    incompletos           SMALLINT NOT NULL DEFAULT 0,
    libres                SMALLINT NOT NULL DEFAULT 0,
    feriados              SMALLINT NOT NULL DEFAULT 0,
    total_efectivo        INTEGER  NOT NULL DEFAULT 0,
    total_jornada         INTEGER  NOT NULL DEFAULT 0,
    prom_efectivo         INTEGER  NOT NULL DEFAULT 0,
    prom_jornada          INTEGER  NOT NULL DEFAULT 0,
    total_exceso_descanso INTEGER  NOT NULL DEFAULT 0,
    total_horas_extra     INTEGER  NOT NULL DEFAULT 0,
    excel_filename        TEXT,

    UNIQUE (colaborador_id, mes_key)
);
