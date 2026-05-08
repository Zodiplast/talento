# Uso: desde la raíz del repo talento/
.PHONY: sync build serve docker-build

sync:
	python biometrico/sync_device.py

build:
	python doc/reporte_web.py

serve:
	uvicorn webapp.main:app --host 0.0.0.0 --port 8080 --reload

docker-build:
	docker build -t talento-biometrico:latest .
