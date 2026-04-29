"""
programar_monitor.py
--------------------
Controla la tarea programada del monitor de correos en el Programador de Windows.

Uso:
    python programar_monitor.py activar [--horas N]   # Crear/actualizar tarea (default: 6h)
    python programar_monitor.py desactivar            # Eliminar tarea
    python programar_monitor.py estado                # Ver si está activa y cuándo corre
"""

import sys
import subprocess
from pathlib import Path

NOMBRE_TAREA  = "ClasificadorDocs_Monitor"
PYTHON_EXE    = r"C:\Users\aux22.gg\AppData\Local\Programs\Python\Python312\python.exe"
SCRIPT_PATH   = Path(__file__).parent / "monitor_correos.py"
HORAS_DEFAULT = 6


def _ejecutar(cmd: list[str]) -> tuple[int, str]:
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r.returncode, (r.stdout + r.stderr).strip()


def activar(horas: int):
    """Crea o reemplaza la tarea programada."""
    minutos = horas * 60

    # Eliminar si ya existe para recrear con el nuevo intervalo
    _ejecutar(["schtasks", "/Delete", "/TN", NOMBRE_TAREA, "/F"])

    cmd = [
        "schtasks", "/Create",
        "/TN", NOMBRE_TAREA,
        "/TR", f'"{PYTHON_EXE}" "{SCRIPT_PATH}"',
        "/SC", "MINUTE",
        "/MO", str(minutos),
        "/ST", "00:00",          # primer disparo a medianoche del día de creación
        "/RL", "HIGHEST",        # permisos elevados (necesario para COM de Outlook)
        "/F",                    # forzar si ya existe
    ]
    codigo, salida = _ejecutar(cmd)
    if codigo == 0:
        print(f"[OK] Tarea '{NOMBRE_TAREA}' activada — corre cada {horas} hora(s).")
        print(f"     Script: {SCRIPT_PATH}")
    else:
        print(f"[ERROR] No se pudo crear la tarea:\n{salida}")
        sys.exit(1)


def desactivar():
    """Elimina la tarea programada."""
    codigo, salida = _ejecutar(["schtasks", "/Delete", "/TN", NOMBRE_TAREA, "/F"])
    if codigo == 0:
        print(f"[OK] Tarea '{NOMBRE_TAREA}' eliminada.")
    else:
        if "no existe" in salida.lower() or "not exist" in salida.lower() or "cannot find" in salida.lower():
            print(f"[INFO] La tarea '{NOMBRE_TAREA}' no estaba activa.")
        else:
            print(f"[ERROR] {salida}")
            sys.exit(1)


def estado():
    """Muestra el estado actual de la tarea."""
    codigo, salida = _ejecutar(["schtasks", "/Query", "/TN", NOMBRE_TAREA, "/FO", "LIST"])
    if codigo != 0:
        print(f"[INFO] La tarea '{NOMBRE_TAREA}' no está programada.")
    else:
        print(salida)


def _parsear_horas(args: list[str]) -> int:
    if "--horas" in args:
        idx = args.index("--horas")
        try:
            return int(args[idx + 1])
        except (IndexError, ValueError):
            print("[ERROR] Uso: --horas N  (donde N es un número entero)")
            sys.exit(1)
    return HORAS_DEFAULT


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("activar", "desactivar", "estado"):
        print(__doc__)
        sys.exit(0)

    accion = sys.argv[1]
    resto  = sys.argv[2:]

    if accion == "activar":
        horas = _parsear_horas(resto)
        activar(horas)
    elif accion == "desactivar":
        desactivar()
    elif accion == "estado":
        estado()
