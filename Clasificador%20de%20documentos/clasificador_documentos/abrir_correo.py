"""
Abre un correo en Outlook a partir de su EntryID guardado en registros_subidas.log.

Uso:
    python abrir_correo.py 00000000D9CA9CA1DA9553438EA51EE97642E572646E2000
"""

import sys
import win32com.client
import pywintypes

if len(sys.argv) < 2:
    print("Uso: python abrir_correo.py <EntryID>")
    print()
    print("El EntryID se encuentra en registros_subidas.log, columna 4:")
    print("  N | timestamp | remitente | ENTRY_ID | ruta_sharepoint")
    sys.exit(1)

entry_id = sys.argv[1].strip()

try:
    outlook = win32com.client.GetActiveObject("Outlook.Application")
except pywintypes.com_error:
    print("ERROR — Outlook no esta abierto. Abri Outlook e intenta de nuevo.")
    sys.exit(1)

try:
    ns  = outlook.GetNamespace("MAPI")
    msg = ns.GetItemFromID(entry_id)
    msg.Display()
    print(f"Abierto: {msg.Subject}")
except Exception as e:
    print(f"ERROR — No se pudo abrir el correo: {e}")
    print("El EntryID puede haber cambiado si se reconfiguró el perfil de Outlook.")
    sys.exit(1)
