"""
Prueba de etiquetado de asunto en Outlook.

Toma el correo mas reciente de documentacion@grupointeca.com,
le agrega [#TEST] al asunto y verifica que se pueda buscar.

Uso:
    python Pruebas/prueba_etiqueta_asunto.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client
import pywintypes

print("=" * 60)
print("  PRUEBA ETIQUETA ASUNTO")
print("=" * 60)

outlook = win32com.client.GetActiveObject("Outlook.Application")
ns = outlook.GetNamespace("MAPI")

inbox = None
for store in ns.Stores:
    try:
        if "documentacion@grupointeca.com" in store.DisplayName.lower():
            inbox = store.GetDefaultFolder(6)
            break
    except Exception:
        continue

if inbox is None:
    print("ERROR — no se encontro la bandeja de documentacion@grupointeca.com")
    sys.exit(1)

items = inbox.Items
items.Sort("[ReceivedTime]", True)
msg = items[1]

asunto_original = msg.Subject
print(f"\nCorreo seleccionado:")
print(f"  Asunto original : {asunto_original}")
print(f"  De              : {msg.SenderEmailAddress}")
print(f"  Recibido        : {msg.ReceivedTime.strftime('%Y-%m-%d %H:%M')}")

etiqueta = " [#TEST]"
if etiqueta in asunto_original:
    print(f"\nYa tiene la etiqueta. Asunto actual: {asunto_original}")
else:
    msg.Subject = asunto_original + etiqueta
    msg.Save()
    print(f"\nAsunto modificado a: {msg.Subject}")

print(f"""
Ahora en Outlook 2016 buscá:
    [#TEST]

Deberia aparecer ese correo exacto.

Para revertir: selecciona el correo, presiona F2 y borra [#TEST] del asunto.
""")
print("=" * 60)
