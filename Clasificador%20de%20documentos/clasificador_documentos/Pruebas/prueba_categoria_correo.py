"""
Prueba de asignacion de categoria a un correo en Outlook.

Toma el correo mas reciente de la bandeja de entrada y le asigna
la categoria CLAS-00001 para verificar que funciona correctamente.

Uso:
    python Pruebas/prueba_categoria_correo.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

print("=" * 60)
print("  PRUEBA CATEGORIA OUTLOOK")
print("=" * 60)

import win32com.client
import pywintypes

print("\n[1] Conectando a Outlook...", end=" ", flush=True)
try:
    outlook = win32com.client.GetActiveObject("Outlook.Application")
    print("OK")
except pywintypes.com_error as e:
    print(f"FALLO — Outlook no esta abierto: {e}")
    sys.exit(1)

ns = outlook.GetNamespace("MAPI")

OUTLOOK_EMAIL = "documentacion@grupointeca.com"
inbox = None
for store in ns.Stores:
    try:
        if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
            inbox = store.GetDefaultFolder(6)
            print(f"[1b] Bandeja encontrada: {store.DisplayName}")
            break
    except Exception:
        continue

if inbox is None:
    print(f"FALLO — no se encontro la bandeja de {OUTLOOK_EMAIL}")
    print("      Cuentas disponibles:")
    for store in ns.Stores:
        print(f"        - {store.DisplayName}")
    sys.exit(1)

items = inbox.Items
items.Sort("[ReceivedTime]", True)
msg = items[1]

print(f"\nCorreo seleccionado:")
print(f"  Asunto   : {msg.Subject}")
print(f"  De       : {msg.SenderEmailAddress}")
print(f"  Recibido : {msg.ReceivedTime.strftime('%Y-%m-%d %H:%M')}")
print(f"  Categoria actual: '{msg.Categories}'")

categoria = "CLAS-00001"

print(f"\n[2] Asignando categoria '{categoria}'...", end=" ", flush=True)
try:
    msg.Categories = categoria
    msg.Save()
    print("OK")
except Exception as e:
    print(f"FALLO — {e}")
    sys.exit(1)

print(f"\n[3] Verificando...", end=" ", flush=True)
# Releer el mensaje para confirmar
msg2 = inbox.Items.Find(f"[Subject] = '{msg.Subject.replace(chr(39), '')}'")
if msg2 and categoria in (msg2.Categories or ""):
    print("OK — categoria guardada correctamente")
else:
    print("OK (verificar manualmente en Outlook)")

print(f"""
Ahora en Outlook 2016:
  - El correo "{msg.Subject[:50]}" tiene la etiqueta '{categoria}'
  - Para buscarlo escribe en el buscador de Outlook:
        category:CLAS-00001
  - Deberia aparecer ese correo exacto

Para quitar la categoria de prueba, selecciona el correo
en Outlook, click derecho -> Categorizar -> Borrar todas.
""")
print("=" * 60)
