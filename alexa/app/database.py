"""
database.py — Conexión a PostgreSQL y operaciones de base de datos.

Usa psycopg2 con un pool de conexiones. No hay ORM.

Funciones públicas:
    init_pool(database_url)       — inicializa el pool (llamado en lifespan)
    run_migration()               — ejecuta migrations/001_init.sql
    upsert_month(processed)       — persiste un ProcessedMonth en la BD
    load_reporte_json() -> dict   — lee todos los datos y arma el JSON para la plantilla
    health_check() -> bool        — SELECT 1 para verificar conectividad
"""

from __future__ import annotations

import os
from contextlib import contextmanager
from pathlib import Path
from typing import Generator

import psycopg2
import psycopg2.pool
import psycopg2.extras

from app.models import (
    ConfigData,
    MesInfo,
    MesPayload,
    ProcessedMonth,
    ReporteJSON,
    ResumenData,
)

# ── Estado global del pool ────────────────────────────────────────────────────

_pool: psycopg2.pool.ThreadedConnectionPool | None = None

MIGRATIONS_PATH = Path(__file__).resolve().parents[1] / "migrations" / "001_init.sql"


def init_pool(database_url: str | None = None) -> None:
    """Inicializa el pool de conexiones. Llamar una sola vez al arrancar."""
    global _pool
    url = database_url or os.environ["DATABASE_URL"]
    _pool = psycopg2.pool.ThreadedConnectionPool(minconn=1, maxconn=5, dsn=url)


@contextmanager
def get_conn() -> Generator:
    """Context manager que presta una conexión del pool y la devuelve al salir."""
    if _pool is None:
        raise RuntimeError("Pool no inicializado. Llama init_pool() primero.")
    conn = _pool.getconn()
    try:
        yield conn
    except Exception:
        conn.rollback()
        raise
    finally:
        _pool.putconn(conn)


# ── Migración ─────────────────────────────────────────────────────────────────

def run_migration() -> None:
    """Ejecuta el SQL de migraciones (idempotente — usa IF NOT EXISTS)."""
    sql = MIGRATIONS_PATH.read_text(encoding="utf-8")
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql)
        conn.commit()


# ── Health check ──────────────────────────────────────────────────────────────

def health_check() -> bool:
    try:
        with get_conn() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
        return True
    except Exception:
        return False


# ── Upsert de un mes completo ─────────────────────────────────────────────────

def upsert_month(processed: ProcessedMonth) -> None:
    """Persiste un ProcessedMonth en la BD de forma incremental.

    - Hace UPSERT de talento.meses (un mes = un registro).
    - Para cada colaborador:
        1. UPSERT talento.colaboradores
        2. DELETE + bulk INSERT de talento.asistencia_dias para ese mes
        3. UPSERT talento.resumen_mes
    Todo en una sola transacción.
    """
    with get_conn() as conn:
        with conn.cursor() as cur:
            # 1. UPSERT mes
            cur.execute(
                """
                INSERT INTO talento.meses (mes_key, mes_label, original_file, uploaded_at)
                VALUES (%s, %s, %s, now())
                ON CONFLICT (mes_key) DO UPDATE
                    SET mes_label     = EXCLUDED.mes_label,
                        original_file = EXCLUDED.original_file,
                        uploaded_at   = now()
                """,
                (processed.mes_key, processed.mes_label, processed.original_file),
            )

            for colab in processed.colaboradores:
                # 2. UPSERT colaborador → obtener id
                cur.execute(
                    """
                    INSERT INTO talento.colaboradores (colab_key, display_name, updated_at)
                    VALUES (%s, %s, now())
                    ON CONFLICT (colab_key) DO UPDATE
                        SET display_name = EXCLUDED.display_name,
                            updated_at   = now()
                    RETURNING id
                    """,
                    (colab.colab_key, colab.display_name),
                )
                colab_id: int = cur.fetchone()[0]

                # 3. DELETE días de ese mes para ese colaborador (reemplaza el mes completo)
                cur.execute(
                    "DELETE FROM talento.asistencia_dias WHERE colaborador_id = %s AND mes_key = %s",
                    (colab_id, colab.mes_key),
                )

                # 4. Bulk INSERT días
                if colab.dias:
                    psycopg2.extras.execute_values(
                        cur,
                        """
                        INSERT INTO talento.asistencia_dias (
                            colaborador_id, mes_key, fecha, dia, dow, estado, motivo_feriado,
                            entrada, sal_des, ent_des, sal_alm, ent_alm, salida,
                            t_desayuno, t_almuerzo, t_jornada, t_efectivo, jornada_objetivo,
                            exceso_desayuno, exceso_almuerzo, exceso_descanso, horas_extra,
                            resaltar_descanso
                        ) VALUES %s
                        """,
                        [
                            (
                                colab_id, colab.mes_key, d["fecha"], d["dia"], d["dow"],
                                d["estado"], d.get("motivo_feriado"),
                                d["entrada"], d["sal_des"], d["ent_des"],
                                d["sal_alm"], d["ent_alm"], d["salida"],
                                d["t_desayuno"], d["t_almuerzo"], d["t_jornada"],
                                d["t_efectivo"], d["jornada_objetivo"],
                                d["exceso_desayuno"], d["exceso_almuerzo"],
                                d["exceso_descanso"], d["horas_extra"],
                                d["resaltar_descanso"],
                            )
                            for d in colab.dias
                        ],
                    )

                # 5. UPSERT resumen_mes (sin excel_filename aún — se actualiza después)
                r = colab.resumen
                cur.execute(
                    """
                    INSERT INTO talento.resumen_mes (
                        colaborador_id, mes_key,
                        completos, incompletos, libres, feriados,
                        total_efectivo, total_jornada, prom_efectivo, prom_jornada,
                        total_exceso_descanso, total_horas_extra
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (colaborador_id, mes_key) DO UPDATE
                        SET completos             = EXCLUDED.completos,
                            incompletos           = EXCLUDED.incompletos,
                            libres                = EXCLUDED.libres,
                            feriados              = EXCLUDED.feriados,
                            total_efectivo        = EXCLUDED.total_efectivo,
                            total_jornada         = EXCLUDED.total_jornada,
                            prom_efectivo         = EXCLUDED.prom_efectivo,
                            prom_jornada          = EXCLUDED.prom_jornada,
                            total_exceso_descanso = EXCLUDED.total_exceso_descanso,
                            total_horas_extra     = EXCLUDED.total_horas_extra
                    """,
                    (
                        colab_id, colab.mes_key,
                        r["completos"], r["incompletos"], r["libres"], r["feriados"],
                        r["total_efectivo"], r["total_jornada"],
                        r["prom_efectivo"], r["prom_jornada"],
                        r["total_exceso_descanso"], r["total_horas_extra"],
                    ),
                )

        conn.commit()


def update_excel_filename(colab_key: str, mes_key: str, excel_filename: str) -> None:
    """Actualiza el nombre del Excel generado en talento.resumen_mes."""
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE talento.resumen_mes rm
                SET excel_filename = %s
                FROM talento.colaboradores c
                WHERE rm.colaborador_id = c.id
                  AND c.colab_key = %s
                  AND rm.mes_key = %s
                """,
                (excel_filename, colab_key, mes_key),
            )
        conn.commit()


# ── Construcción del JSON para la plantilla ───────────────────────────────────

def load_reporte_json() -> ReporteJSON:
    """Lee todos los datos de la BD y arma el dict que espera la plantilla HTML."""
    # Importamos aquí para evitar importación circular
    from app.processor import MESES_ES_INV  # noqa: F401 (solo para referencia)

    # Importar constantes de horarios
    import sys
    from pathlib import Path as _Path
    _bio = _Path(__file__).resolve().parents[1] / "biometrico"
    if str(_bio) not in sys.path:
        sys.path.insert(0, str(_bio))
    from horarios import (
        ALMUERZO_MIN, DESAYUNO_MIN, DESCANSO_MAX_MIN,
        JORNADA_LUNES_VIERNES_MIN, JORNADA_SABADO_MIN,
    )

    with get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:

            # Colaboradores
            cur.execute("SELECT id, colab_key, display_name FROM talento.colaboradores ORDER BY display_name")
            colaboradores_rows = cur.fetchall()
            colab_by_id: dict[int, dict] = {r["id"]: r for r in colaboradores_rows}

            # Meses
            cur.execute("SELECT mes_key, mes_label FROM talento.meses ORDER BY mes_key")
            meses_rows = cur.fetchall()

            # Resúmenes
            cur.execute(
                "SELECT colaborador_id, mes_key, completos, incompletos, libres, feriados, "
                "       total_efectivo, total_jornada, prom_efectivo, prom_jornada, "
                "       total_exceso_descanso, total_horas_extra, excel_filename "
                "FROM talento.resumen_mes"
            )
            resumenes: dict[tuple[int, str], dict] = {
                (r["colaborador_id"], r["mes_key"]): r
                for r in cur.fetchall()
            }

            # Días (todos de una vez, luego agrupamos en Python)
            cur.execute(
                "SELECT colaborador_id, mes_key, fecha, dia, dow, estado, motivo_feriado, "
                "       entrada, sal_des, ent_des, sal_alm, ent_alm, salida, "
                "       t_desayuno, t_almuerzo, t_jornada, t_efectivo, jornada_objetivo, "
                "       exceso_desayuno, exceso_almuerzo, exceso_descanso, horas_extra, "
                "       resaltar_descanso "
                "FROM talento.asistencia_dias "
                "ORDER BY colaborador_id, fecha"
            )
            dias_rows = cur.fetchall()

    # Agrupar días por (colaborador_id, mes_key)
    from collections import defaultdict
    dias_by_colab_mes: dict[tuple[int, str], list] = defaultdict(list)
    for row in dias_rows:
        key = (row["colaborador_id"], row["mes_key"])
        dias_by_colab_mes[key].append(row)

    # Construir datos[display_name][mes_key] = MesPayload
    datos: dict[str, dict[str, MesPayload]] = {}

    for colab_id, colab_info in colab_by_id.items():
        display_name = colab_info["display_name"]
        datos[display_name] = {}

        for mes_row in meses_rows:
            mes_key = mes_row["mes_key"]
            dias_raw = dias_by_colab_mes.get((colab_id, mes_key), [])
            if not dias_raw and (colab_id, mes_key) not in resumenes:
                continue  # este colaborador no tiene datos para este mes

            resumen_raw = resumenes.get((colab_id, mes_key), {})
            excel_filename = resumen_raw.get("excel_filename")

            dias = [
                {
                    "fecha": str(d["fecha"]),
                    "dia": d["dia"],
                    "dow": d["dow"],
                    "estado": d["estado"],
                    "motivo_feriado": d["motivo_feriado"],
                    "entrada": d["entrada"],
                    "sal_des": d["sal_des"],
                    "ent_des": d["ent_des"],
                    "sal_alm": d["sal_alm"],
                    "ent_alm": d["ent_alm"],
                    "salida": d["salida"],
                    "t_desayuno": float(d["t_desayuno"]) if d["t_desayuno"] is not None else None,
                    "t_almuerzo": float(d["t_almuerzo"]) if d["t_almuerzo"] is not None else None,
                    "t_jornada": float(d["t_jornada"]) if d["t_jornada"] is not None else None,
                    "t_efectivo": float(d["t_efectivo"]) if d["t_efectivo"] is not None else None,
                    "jornada_objetivo": d["jornada_objetivo"],
                    "exceso_desayuno": float(d["exceso_desayuno"]) if d["exceso_desayuno"] is not None else None,
                    "exceso_almuerzo": float(d["exceso_almuerzo"]) if d["exceso_almuerzo"] is not None else None,
                    "exceso_descanso": float(d["exceso_descanso"]) if d["exceso_descanso"] is not None else None,
                    "horas_extra": float(d["horas_extra"]) if d["horas_extra"] is not None else None,
                    "resaltar_descanso": bool(d["resaltar_descanso"]),
                }
                for d in dias_raw
            ]

            resumen = ResumenData(
                completos=resumen_raw.get("completos", 0),
                incompletos=resumen_raw.get("incompletos", 0),
                libres=resumen_raw.get("libres", 0),
                feriados=resumen_raw.get("feriados", 0),
                total_efectivo=resumen_raw.get("total_efectivo", 0),
                total_jornada=resumen_raw.get("total_jornada", 0),
                prom_efectivo=resumen_raw.get("prom_efectivo", 0),
                prom_jornada=resumen_raw.get("prom_jornada", 0),
                total_exceso_descanso=resumen_raw.get("total_exceso_descanso", 0),
                total_horas_extra=resumen_raw.get("total_horas_extra", 0),
            )

            datos[display_name][mes_key] = MesPayload(
                dias=dias,
                resumen=resumen,
                excel_file=f"excel/{excel_filename}" if excel_filename else None,
            )

    # Eliminar colaboradores sin datos en ningún mes
    datos = {k: v for k, v in datos.items() if v}

    config = ConfigData(
        desayuno_min=DESAYUNO_MIN,
        almuerzo_min=ALMUERZO_MIN,
        descanso_max_min=DESCANSO_MAX_MIN,
        jornada_lunes_viernes_min=JORNADA_LUNES_VIERNES_MIN,
        jornada_sabado_min=JORNADA_SABADO_MIN,
    )

    return ReporteJSON(
        config=config,
        colaboradores=sorted(datos.keys()),
        meses=[MesInfo(value=m["mes_key"], label=m["mes_label"]) for m in meses_rows],
        datos=datos,
    )
