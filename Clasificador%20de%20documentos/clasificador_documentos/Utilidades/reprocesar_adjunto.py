"""
Reprocesa correos específicos buscándolos por nombre de adjunto en la bandeja de entrada.
Ignora procesados.txt — fuerza el reprocesamiento aunque ya estén marcados.

Uso:
    python reprocesar_adjunto.py "Z64005.pdf" "Datos adjuntos sin título 00003.eml"
"""

import sys
import io
import pythoncom
import win32com.client

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from pathlib import Path
_BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_BASE))
sys.path.insert(0, str(_BASE.parent.parent / "Agente_Seguridad"))

from configuracion.ajustes import OUTLOOK_EMAIL
from Integraciones.graph_client import GraphClient
from Integraciones.monitor_correos import _procesar_mensaje, _cargar_rutas_subidas


def reprocesar(nombres_adjunto: list[str]):
    pythoncom.CoInitialize()
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        bandeja = None
        for store in namespace.Stores:
            try:
                if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                    bandeja = store.GetDefaultFolder(6)
                    break
            except Exception:
                continue

        if bandeja is None:
            print(f"[ERROR] No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")
            return

        cliente       = GraphClient()
        rutas_subidas = _cargar_rutas_subidas()
        nombres_lower = [n.lower() for n in nombres_adjunto]
        encontrados   = []

        print(f"[Búsqueda] Buscando {len(nombres_adjunto)} adjunto(s) en bandeja...")
        for item in bandeja.Items:
            try:
                if item.Class != 43:
                    continue
                for i in range(1, item.Attachments.Count + 1):
                    try:
                        nombre_adj = item.Attachments.Item(i).FileName.lower()
                    except Exception:
                        continue
                    if nombre_adj in nombres_lower:
                        encontrados.append((item, nombre_adj))
                        print(f"[OK] Encontrado: {nombre_adj} — {item.Subject}")
                        break
            except Exception:
                continue

        if not encontrados:
            print("[WARN] No se encontró ningún correo con esos adjuntos.")
            return

        print(f"\n[Reprocesando] {len(encontrados)} correo(s)...\n")
        for msg, nombre_adj in encontrados:
            print(f"--- {msg.Subject} ---")
            _procesar_mensaje(msg, cliente, rutas_subidas)

        print("\n[Listo] Reprocesamiento completado.")

    except KeyboardInterrupt:
        print("\nInterrumpido por el usuario.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python reprocesar_adjunto.py \"nombre_adjunto1.pdf\" \"nombre_adjunto2.pdf\"")
        sys.exit(1)
    reprocesar(sys.argv[1:])
