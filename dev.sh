#!/usr/bin/env bash
# talento — Levantar FastAPI con secrets inyectados desde 1Password
# Uso: ./dev.sh
set -euo pipefail

echo "🔑 Inyectando secrets desde 1Password..."

POSTGRES_PASSWORD=$(op read "op://Data Analytics/Postgres Admin/password")
export DATABASE_URL="postgresql://josue:${POSTGRES_PASSWORD}@localhost:5432/cortana"

exec op run --env-file .env -- uvicorn app.main:app --reload --host 0.0.0.0 --port 8010
