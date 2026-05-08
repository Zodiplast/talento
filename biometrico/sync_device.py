"""
sync_device.py — Descarga asistencia del biométrico ZKTeco y guarda en raw/biometrico/.

Genera:
  - Excel compatible con el procesador (raw/biometrico/MES-YYYY.xlsx)
  - Parquet con columnas analíticas (raw/biometrico_parquet/MES-YYYY.parquet)

Uso (desde la raíz del proyecto):
    python biometrico/sync_device.py              # mes actual
    python biometrico/sync_device.py --mes mayo --anio 2026
    python biometrico/sync_device.py --ip 192.168.100.251 --port 4370
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from biometrico.schemas.paths import RAW_BIOMETRICO_DIR, RAW_BIOMETRICO_PARQUET_DIR

try:
    import openpyxl
    import polars as pl
    from zk import ZK
except ImportError as e:
    sys.exit(f"[ERROR] Dependencia faltante: {e}\n  Ejecuta: pip install -r requirements.txt")

MESES_ES: dict[int, str] = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
MESES_ES_NUM: dict[str, int] = {v: k for k, v in MESES_ES.items()}


def sync(ip: str, port: int, mes: int, anio: int, password: int = 0) -> Path:
    print(f"Conectando a {ip}:{port} ...")
    zk = ZK(ip, port=port, timeout=15, password=password, force_udp=False, ommit_ping=False)
    conn = zk.connect()
    print(f"  Firmware : {conn.get_firmware_version()}")
    print(f"  Serial   : {conn.get_serialnumber()}")

    print("Descargando usuarios ...")
    users = conn.get_users()
    user_map: dict[str, tuple[str, str]] = {}
    for u in users:
        uid_str = str(u.user_id)
        nombre = (u.name or "").strip() or f"NN-{u.user_id}"
        user_map[uid_str] = (nombre, uid_str)

    print(f"  {len(users)} usuarios encontrados.")

    print("Descargando asistencia ...")
    attendances = conn.get_attendance()
    conn.disconnect()
    print(f"  {len(attendances)} registros en total.")

    # Filtrar por mes/año solicitado
    filtrados = [
        a for a in attendances
        if a.timestamp and a.timestamp.year == anio and a.timestamp.month == mes
    ]
    print(f"  {len(filtrados)} registros para {MESES_ES[mes]}-{anio}.")

    if not filtrados:
        sys.exit(
            f"[ERROR] No hay registros para {MESES_ES[mes]}-{anio}. "
            "Verifica el mes/año o revisa los datos del dispositivo."
        )

    # Ordenar por usuario y timestamp
    filtrados.sort(key=lambda a: (str(a.user_id), a.timestamp))

    # Construir Excel con el formato que espera el procesador
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet"
    ws.append(["Número", "Nombre", "Tiempo"])

    numeros: list[str] = []
    nombres: list[str] = []
    tiempos: list[datetime] = []
    for a in filtrados:
        uid_str = str(a.user_id)
        nombre, numero = user_map.get(uid_str, (f"NN-{a.user_id}", uid_str))
        ws.append([numero, nombre, a.timestamp])
        numeros.append(str(numero))
        nombres.append(nombre)
        tiempos.append(a.timestamp)

    RAW_BIOMETRICO_DIR.mkdir(parents=True, exist_ok=True)
    output = RAW_BIOMETRICO_DIR / f"{MESES_ES[mes]}-{anio}.xlsx"
    wb.save(output)
    print(f"\n  Guardado en: {output}")

    # Naive UTC (mismo estilo que columnas del Excel) para evitar mezcla tz en Polars
    synced_at = datetime.now(timezone.utc).replace(tzinfo=None)
    df = pl.DataFrame(
        {
            "numero": numeros,
            "nombre": nombres,
            "tiempo": tiempos,
            "anio": [anio] * len(filtrados),
            "mes": [mes] * len(filtrados),
            "synced_at": [synced_at] * len(filtrados),
        }
    )
    RAW_BIOMETRICO_PARQUET_DIR.mkdir(parents=True, exist_ok=True)
    pq_out = RAW_BIOMETRICO_PARQUET_DIR / f"{MESES_ES[mes]}-{anio}.parquet"
    df.write_parquet(pq_out)
    print(f"  Parquet : {pq_out}")

    return output


def _parse_args() -> argparse.Namespace:
    hoy = datetime.today()
    parser = argparse.ArgumentParser(description="Sincroniza biométrico ZKTeco → raw/biometrico/")
    parser.add_argument("--ip",     default="192.168.100.251", help="IP del dispositivo")
    parser.add_argument("--port",   default=4370, type=int,    help="Puerto TCP (default 4370)")
    parser.add_argument("--mes",    default=None,              help="Mes en español (ej. mayo)")
    parser.add_argument("--anio",   default=hoy.year, type=int, help="Año (default año actual)")
    parser.add_argument("--password", default=0, type=int,    help="Contraseña del dispositivo")
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()

    if args.mes:
        mes_upper = args.mes.upper()
        if mes_upper not in MESES_ES_NUM:
            meses_validos = ", ".join(m.capitalize() for m in MESES_ES_NUM)
            sys.exit(f"[ERROR] Mes '{args.mes}' no válido. Opciones: {meses_validos}")
        mes_num = MESES_ES_NUM[mes_upper]
    else:
        mes_num = datetime.today().month

    sync(
        ip=args.ip,
        port=args.port,
        mes=mes_num,
        anio=args.anio,
        password=args.password,
    )
