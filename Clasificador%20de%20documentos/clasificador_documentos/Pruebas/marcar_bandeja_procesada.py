"""
Utilidad de una sola vez: recorre TODA la bandeja de Outlook y registra
todos los EntryIDs en procesados.txt — sin subir nada ni tocar ningún correo.

Después de correr esto, el pipeline normal solo procesará correos nuevos.

Uso:
    python pruebas/marcar_bandeja_procesada.py
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from pathlib import Path
from configuracion.ajustes import OUTLOOK_EMAIL, EXTENSIONES_BLOQUEADAS

_PROCESADOS_FILE = Path(__file__).parent.parent / "procesados.txt"


def main():
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError("pywin32 no instalado. Ejecuta: pip install pywin32")

    ya_en_txt = set()
    if _PROCESADOS_FILE.exists():
        ya_en_txt = set(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines())
    print(f"  {len(ya_en_txt)} IDs ya registrados en procesados.txt")

    pythoncom.CoInitialize()
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        bandeja = None
        for store in namespace.Stores:
            try:
                if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                    bandeja = store.GetDefaultFolder(6)
                    print(f"  Cuenta: {store.DisplayName}")
                    break
            except Exception:
                continue

        if bandeja is None:
            raise RuntimeError(f"No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")

        mensajes = bandeja.Items
        total    = mensajes.Count
        print(f"  Total mensajes en bandeja: {total}")
        print("  Escaneando... (puede tardar varios minutos)")

        nuevos = []
        for i, msg in enumerate(mensajes, 1):
            if i % 100 == 0:
                print(f"  ... {i}/{total} escaneados, {len(nuevos)} nuevos a registrar")
            try:
                msg_id = msg.EntryID
                if msg_id and msg_id not in ya_en_txt:
                    nuevos.append(msg_id)
            except Exception:
                continue

    finally:
        pythoncom.CoUninitialize()

    if nuevos:
        with _PROCESADOS_FILE.open("a", encoding="utf-8") as f:
            for mid in nuevos:
                f.write(mid + "\n")
        print(f"\n  Registrados {len(nuevos)} IDs nuevos en procesados.txt")
    else:
        print("\n  No hay IDs nuevos — procesados.txt ya estaba completo.")

    total_final = len(ya_en_txt) + len(nuevos)
    print(f"  Total en procesados.txt ahora: {total_final}")
    print("\nListo. El pipeline ahora solo procesará correos nuevos.")


if __name__ == "__main__":
    main()
