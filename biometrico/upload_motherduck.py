"""
Sube marcaciones raw (Parquet del sync) al esquema `talento` en MotherDuck.

El biométrico es IP local: ejecutá `sync_device.py` en la PC en la LAN; este
script empuja el Parquet ya generado por internet hacia MotherDuck.

Variables de entorno:
  MOTHERDUCK_TOKEN    — token de servicio (docs MotherDuck)
  MOTHERDUCK_DATABASE — nombre de la base (default: zodiplast)

Uso (desde la raíz de talento/):
  uv pip install -r requirements.txt
  $env:MOTHERDUCK_TOKEN = "..."   # PowerShell
  python biometrico/upload_motherduck.py --mes mayo --anio 2026
  python biometrico/upload_motherduck.py --parquet raw/biometrico_parquet/MAYO-2026.parquet

La web no debe incrustar el token: un backend lee MotherDuck con credenciales
en servidor (Python duckdb, Java JDBC, etc.).
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from biometrico.schemas.paths import RAW_BIOMETRICO_PARQUET_DIR

try:
    import duckdb
    import polars as pl
except ImportError as e:
    sys.exit(f"[ERROR] Dependencia faltante: {e}\n  Ejecuta: pip install -r requirements.txt")

SCHEMA = "talento"
TABLE = "marcaciones_raw"

MESES_ES_NUM: dict[str, int] = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}


def _connect_motherduck(database: str) -> duckdb.DuckDBPyConnection:
    token = os.environ.get("MOTHERDUCK_TOKEN", "").strip()
    if not token:
        sys.exit(
            "[ERROR] Falta MOTHERDUCK_TOKEN en el entorno.\n"
            "  PowerShell: $env:MOTHERDUCK_TOKEN='tu_token'"
        )
    return duckdb.connect(
        f"md:{database}",
        config={"motherduck_token": token},
    )


def ensure_table(con: duckdb.DuckDBPyConnection) -> None:
    con.execute(f"CREATE SCHEMA IF NOT EXISTS {SCHEMA}")
    con.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {SCHEMA}.{TABLE} (
            numero VARCHAR,
            nombre VARCHAR,
            tiempo TIMESTAMP,
            anio INTEGER,
            mes INTEGER,
            synced_at TIMESTAMP
        )
        """
    )


def upload_parquet(con: duckdb.DuckDBPyConnection, parquet_path: Path) -> None:
    if not parquet_path.is_file():
        sys.exit(f"[ERROR] No existe el archivo: {parquet_path}")

    df = pl.read_parquet(parquet_path)
    required = ("numero", "nombre", "tiempo", "anio", "mes", "synced_at")
    for col in required:
        if col not in df.columns:
            sys.exit(f"[ERROR] El Parquet no tiene columna '{col}'. Regenerá con sync_device.py")

    anios = df["anio"].unique().to_list()
    meses = df["mes"].unique().to_list()
    if len(anios) != 1 or len(meses) != 1:
        sys.exit("[ERROR] Se espera un solo mes/año por archivo (como genera sync_device).")

    anio, mes = int(anios[0]), int(meses[0])
    ensure_table(con)
    con.register("src_marc", df)
    try:
        con.execute(
            f"DELETE FROM {SCHEMA}.{TABLE} WHERE anio = ? AND mes = ?",
            [anio, mes],
        )
        con.execute(f"INSERT INTO {SCHEMA}.{TABLE} SELECT * FROM src_marc")
    finally:
        con.unregister("src_marc")

    n = con.execute(
        f"SELECT COUNT(*) FROM {SCHEMA}.{TABLE} WHERE anio = ? AND mes = ?",
        [anio, mes],
    ).fetchone()[0]
    print(f"  OK MotherDuck → {SCHEMA}.{TABLE}  ({anio}-{mes:02d}): {n} filas para ese mes.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Sube Parquet de marcaciones a MotherDuck (esquema talento).")
    parser.add_argument("--parquet", type=Path, default=None, help="Ruta al .parquet")
    parser.add_argument("--mes", default=None, help="Mes en español (ej. mayo), si no pasás --parquet")
    parser.add_argument("--anio", type=int, default=None, help="Año, si no pasás --parquet")
    parser.add_argument(
        "--database",
        default=os.environ.get("MOTHERDUCK_DATABASE", "zodiplast"),
        help="Base MotherDuck (default: zodiplast o MOTHERDUCK_DATABASE)",
    )
    args = parser.parse_args()

    if args.parquet:
        pq = args.parquet if args.parquet.is_absolute() else _ROOT / args.parquet
    else:
        if not args.mes or args.anio is None:
            sys.exit("[ERROR] Pasá --parquet o bien --mes y --anio")
        mes_key = args.mes.strip().upper()
        if mes_key not in MESES_ES_NUM:
            sys.exit(f"[ERROR] Mes no válido: {args.mes}")
        from biometrico.sync_device import MESES_ES

        mes_num = MESES_ES_NUM[mes_key]
        pq = RAW_BIOMETRICO_PARQUET_DIR / f"{MESES_ES[mes_num]}-{args.anio}.parquet"

    db = (args.database or "").strip() or "zodiplast"
    print(f"Conectando a MotherDuck md:{db} ...")
    con = _connect_motherduck(db)
    try:
        upload_parquet(con, pq)
    finally:
        con.close()


if __name__ == "__main__":
    main()
