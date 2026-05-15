# Uso: desde la raíz del repo talento/
.PHONY: sync build serve docker-build

sync:
	python biometrico/extract_biometrico.py --mes-actual

build:
	python doc/reporte_web.py

serve:
	PYTHONPATH=.:legacy uvicorn webapp.main:app --host 0.0.0.0 --port 8080 --reload

docker-build:
	docker build -t talento-biometrico:latest .
