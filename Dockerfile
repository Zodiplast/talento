FROM python:3.12-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH=/app:/app/legacy

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY biometrico ./biometrico
COPY raw ./raw
COPY legacy ./legacy

# El contenedor sirve datos montados o copiados en ./raw/biometrico* y ./raw/reglas
EXPOSE 8080
CMD ["uvicorn", "webapp.main:app", "--host", "0.0.0.0", "--port", "8080"]
