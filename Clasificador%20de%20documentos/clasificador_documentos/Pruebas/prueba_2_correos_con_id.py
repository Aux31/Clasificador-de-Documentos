"""
Prueba de etiquetado + log con ID unico.

Toma los 2 correos mas recientes de documentacion@grupointeca.com,
les asigna #1 y #2 en el asunto, y escribe las entradas en registros_subidas.log.

Uso:
    python Pruebas/prueba_2_correos_con_id.py

Para revertir: en Outlook selecciona cada correo, presiona F2 y borra [#1] / [#2] del asunto.
Para limpiar el log: borra las ultimas entradas de prueba en registros_subidas.log.
Para reiniciar el contador: borra contador_correos.txt o ponle 0.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
import win32com.client
import pywintypes
from datetime import datetime
from pathlib import Path

_LOG_SUBIDAS      = Path(__file__).parent.parent / "registros_subidas.log"
_CONTADOR_CORREOS = Path(__file__).parent.parent / "contador_correos.txt"

_logger_subidas = logging.getLogger("subidas_sharepoint_prueba")
if not _logger_subidas.handlers:
    _h = logging.FileHandler(_LOG_SUBIDAS, encoding="utf-8")
    _h.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_h)
    _logger_subidas.setLevel(logging.INFO)


def _siguiente_consecutivo() -> int:
    if not _LOG_SUBIDAS.exists():
        return 1
    lines = _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines()
    return len([l for l in lines if l.strip()]) + 1


def _siguiente_id_correo() -> int:
    if not _CONTADOR_CORREOS.exists():
        _CONTADOR_CORREOS.write_text("0", encoding="utf-8")
    n = int(_CONTADOR_CORREOS.read_text(encoding="utf-8").strip() or "0") + 1
    _CONTADOR_CORREOS.write_text(str(n), encoding="utf-8")
    return n


def _log_subida(remitente: str, ruta_sharepoint: str, id_correo: int) -> str:
    ruta_completa = ruta_sharepoint.replace("\\", "/")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n  = _siguiente_consecutivo()
    entrada = f"{n} | REFERENCE:{id_correo} | {ts} | {remitente} | {ruta_completa}"
    _logger_subidas.info(entrada)
    return entrada


print("=" * 60)
print("  PRUEBA 2 CORREOS CON ID")
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

correos = []
for i in range(1, 4):
    try:
        msg = items[i]
        if msg.Class == 43:
            correos.append(msg)
        if len(correos) == 2:
            break
    except Exception:
        break

if len(correos) < 2:
    print(f"ERROR — se encontraron solo {len(correos)} correo(s) en la bandeja.")
    sys.exit(1)

print()
for idx, msg in enumerate(correos):
    id_correo = _siguiente_id_correo()
    etiqueta  = f" [REFERENCE:{id_correo}]"

    print(f"Correo #{id_correo}:")
    print(f"  De       : {msg.SenderEmailAddress}")
    print(f"  Recibido : {msg.ReceivedTime.strftime('%Y-%m-%d %H:%M')}")
    print(f"  Asunto   : {msg.Subject}")

    if etiqueta in (msg.Subject or ""):
        print(f"  -> Ya tiene {etiqueta.strip()}, no se modifica")
    else:
        msg.Subject = (msg.Subject or "") + etiqueta
        msg.Save()
        print(f"  -> Asunto ahora: {msg.Subject}")

    ruta_prueba = f"cmer-OC-00199999-TEST00{idx}->4. DOCUMENTACION->OTROS->PRUEBA_CORREO_{id_correo}.pdf"
    entrada = _log_subida(msg.SenderEmailAddress, ruta_prueba, id_correo)
    print(f"  -> Log   : {entrada}")
    print()

print("Ahora en Outlook 2016 busca:")
print("  [REFERENCE:1]   para encontrar el primer correo")
print("  [REFERENCE:2]   para encontrar el segundo correo")
print()
print("=" * 60)
