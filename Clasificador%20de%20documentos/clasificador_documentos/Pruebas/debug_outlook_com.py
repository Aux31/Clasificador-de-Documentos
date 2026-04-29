"""
Debug de conexión Outlook vía win32com.
Muestra paso a paso qué pasa al intentar acceder a la bandeja de entrada.

Uso:
    python pruebas/debug_outlook_com.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

print("=" * 60)
print("  DEBUG OUTLOOK COM")
print("=" * 60)

# 1. Importar win32com
print("\n[1] Importando win32com...", end=" ", flush=True)
try:
    import win32com.client
    import pywintypes
    print("OK")
except ImportError as e:
    print(f"FALLO — {e}")
    sys.exit(1)

# 2. Conectar a Outlook
print("[2] Conectando a Outlook (GetActiveObject)...", end=" ", flush=True)
try:
    outlook = win32com.client.GetActiveObject("Outlook.Application")
    print("OK — Outlook ya estaba abierto")
except pywintypes.com_error as e:
    print(f"FALLO — {e}")
    print("      Outlook no está abierto o no responde. Ábrelo e intenta de nuevo.")
    sys.exit(1)

# 3. Acceder al namespace MAPI
print("[3] Obteniendo namespace MAPI...", end=" ", flush=True)
try:
    ns = outlook.GetNamespace("MAPI")
    print("OK")
except Exception as e:
    print(f"FALLO — {e}")
    sys.exit(1)

# 4. Listar cuentas
print("[4] Cuentas configuradas:")
try:
    for i, cuenta in enumerate(ns.Accounts, 1):
        print(f"      {i}. {cuenta.SmtpAddress}")
except Exception as e:
    print(f"      ERROR listando cuentas: {e}")

# 5. Acceder a bandeja de entrada
print("[5] Accediendo a bandeja de entrada...", end=" ", flush=True)
try:
    inbox = ns.GetDefaultFolder(6)  # 6 = olFolderInbox
    print(f"OK — '{inbox.Name}'")
except Exception as e:
    print(f"FALLO — {e}")
    sys.exit(1)

# 6. Contar elementos
print("[6] Contando elementos...", end=" ", flush=True)
try:
    items = inbox.Items
    total = items.Count
    print(f"OK — {total} mensajes en bandeja")
except Exception as e:
    print(f"FALLO — {e}")
    sys.exit(1)

# 7. Leer los últimos 3 correos
print("[7] Últimos 3 correos:")
try:
    items.Sort("[ReceivedTime]", True)
    for i in range(1, min(4, total + 1)):
        msg = items[i]
        print(f"      {i}. [{msg.ReceivedTime.strftime('%Y-%m-%d %H:%M')}] "
              f"{msg.Subject[:60]} | De: {msg.SenderEmailAddress}")
except Exception as e:
    print(f"      ERROR leyendo mensajes: {e}")

# 8. Probar adjuntos del primero
print("[8] Adjuntos del correo más reciente:")
try:
    items.Sort("[ReceivedTime]", True)
    msg = items[1]
    if msg.Attachments.Count == 0:
        print("      (sin adjuntos)")
    else:
        for att in msg.Attachments:
            print(f"      - {att.FileName} ({att.Size // 1024} KB)")
except Exception as e:
    print(f"      ERROR leyendo adjuntos: {e}")

print("\n" + "=" * 60)
print("  COM OK — Outlook responde correctamente")
print("=" * 60 + "\n")
