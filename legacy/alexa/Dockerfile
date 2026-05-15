FROM python:3.12-slim

WORKDIR /app

# psycopg2-binary incluye el driver compilado; no se necesita libpq-dev
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Código de la aplicación
COPY app/        ./app/
COPY biometrico/ ./biometrico/
COPY migrations/ ./migrations/

# reports/ y raw/ se montan como volúmenes en docker-compose
# (no se copian al imagen para no incluir datos sensibles)

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
