"""
Limpia etiquetas de prueba de los asuntos de correos en Outlook.

Elimina: [#TEST], [#1], [#2], [REF-1], [REF-2] de los asuntos,
dejando solo [REFERENCE-N] si existe.

Uso:
    python Pruebas/limpiar_etiquetas_prueba.py
"""

import sys
import os
import re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client
import pywintypes

_ETIQUETAS_BASURA = re.compile(r'\s*\[#[^\]]+\]|\s*\[REF-\d+\]')

print("=" * 60)
print("  LIMPIAR ETIQUETAS DE PRUEBA")
print("=" * 60)

try:
    outlook = win32com.client.GetActiveObject("Outlook.Application")
except pywintypes.com_error:
    print("ERROR — Outlook no esta abierto.")
    sys.exit(1)

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

modificados = 0
revisados   = 0
for i in range(1, min(items.Count + 1, 51)):
    try:
        msg = items[i]
        if msg.Class != 43:
            continue
        asunto = msg.Subject or ""
        if not _ETIQUETAS_BASURA.search(asunto):
            continue
        revisados += 1
        asunto_limpio = _ETIQUETAS_BASURA.sub("", asunto).strip()
        print(f"\nAntes : {asunto}")
        print(f"Despues: {asunto_limpio}")
        msg.Subject = asunto_limpio
        msg.Save()
        modificados += 1
    except Exception as e:
        print(f"[WARN] Error en item {i}: {e}")
        continue

print(f"\n{modificados} correo(s) limpiados.")
print("=" * 60)
