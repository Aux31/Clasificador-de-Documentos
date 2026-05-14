"""
Prueba: inserta 2 entradas de prueba en registros_subidas.log con el nuevo formato.

Formato nuevo: N | timestamp | remitente | #ID_CORREO | ruta_sharepoint

Uso:
    python Pruebas/prueba_log_consecutivos.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
from datetime import datetime
from pathlib import Path

_LOG_SUBIDAS      = Path(__file__).parent.parent / "registros_subidas.log"
_CONTADOR_CORREOS = Path(__file__).parent.parent / "contador_correos.txt"

_logger_subidas = logging.getLogger("subidas_sharepoint")
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


def _log_subida(remitente: str, ruta_sharepoint: str, id_correo: int):
    ruta_completa = ruta_sharepoint.replace("\\", "/")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n  = _siguiente_consecutivo()
    entrada = f"{n} | {ts} | {remitente} | #{id_correo} | {ruta_completa}"
    _logger_subidas.info(entrada)
    return entrada


print("=" * 60)
print("  PRUEBA LOG CONSECUTIVOS")
print("=" * 60)

# Correo 1 — 2 adjuntos (misma id_correo)
id_correo_1 = _siguiente_id_correo()
print(f"\nCorreo #{id_correo_1}: proveedor1@ejemplo.com")
e1 = _log_subida("proveedor1@ejemplo.com",
                 "cmer-OC-00199999-TEST000->4. DOCUMENTACION->4.05 BL-AWB-Porte definitivo->BL_PRUEBA.pdf",
                 id_correo_1)
e2 = _log_subida("proveedor1@ejemplo.com",
                 "cmer-OC-00199999-TEST000->4. DOCUMENTACION->4.02 Factura Definitiva->INVOICE_PRUEBA.pdf",
                 id_correo_1)
print(f"  Entrada 1: {e1}")
print(f"  Entrada 2: {e2}")

# Correo 2 — 1 adjunto (id_correo diferente)
id_correo_2 = _siguiente_id_correo()
print(f"\nCorreo #{id_correo_2}: proveedor2@ejemplo.com")
e3 = _log_subida("proveedor2@ejemplo.com",
                 "cmer-OC-00200000-TEST001->4. DOCUMENTACION->4.27 Packing list definitivo->PACKING_PRUEBA.pdf",
                 id_correo_2)
print(f"  Entrada 3: {e3}")

print(f"""
Entradas escritas en:
  {_LOG_SUBIDAS}

Para buscar el correo 1 en Outlook: busca  [#{id_correo_1}]
Para buscar el correo 2 en Outlook: busca  [#{id_correo_2}]

Para limpiar las entradas de prueba, elimina las ultimas 3 lineas del log.
""")
print("=" * 60)
