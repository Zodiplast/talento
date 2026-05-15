"""
main.py — Punto de entrada del proyecto Talento.

Uso:
    python main.py
"""

import sys
from pathlib import Path

# Asegura paths: `alexa` en legacy/alexa; `doc` en legacy/doc; repo root para `biometrico`.
ALEXA_DIR = Path(__file__).resolve().parent
LEGACY_DIR = ALEXA_DIR.parent
TALENTO_ROOT = LEGACY_DIR.parent
for p in (ALEXA_DIR, LEGACY_DIR, TALENTO_ROOT):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))


def _sep(ancho: int = 45) -> None:
    print("─" * ancho)


def generar_reporte() -> None:
    from doc.reporte_web import main

    main()


def iniciar_servidor() -> None:
    from dotenv import load_dotenv
    import uvicorn

    load_dotenv()
    print("\nIniciando servidor FastAPI en http://localhost:8000")
    print("Presiona Ctrl+C para detener.\n")
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)


def ver_turnos() -> None:
    from alexa.horarios import listar_turnos
    listar_turnos()


MENU = [
    ("Generar reporte biométrico HTML", generar_reporte),
    ("Iniciar servidor FastAPI",        iniciar_servidor),
    ("Ver configuración de turnos",     ver_turnos),
]


def menu() -> None:
    while True:
        print()
        print("Talento — Sistema de Reportes Biométricos")
        _sep()
        for i, (label, _) in enumerate(MENU, 1):
            print(f"  {i}. {label}")
        print("  0. Salir")
        _sep()

        opcion = input("Selecciona una opción: ").strip()

        if opcion == "0":
            print("Saliendo.")
            break

        if opcion.isdigit() and 1 <= int(opcion) <= len(MENU):
            _, fn = MENU[int(opcion) - 1]
            print()
            try:
                fn()
            except KeyboardInterrupt:
                print("\nInterrumpido.")
        else:
            print("Opción no válida, intenta de nuevo.")


if __name__ == "__main__":
    menu()
