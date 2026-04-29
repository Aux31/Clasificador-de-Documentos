"""
Procesador de aprobaciones de sugerencias a proveedores.

Lee la carpeta pendientes_aprobacion/, procesa todos los .txt
que tengan DECISION: APROBAR o DECISION: RECHAZAR y termina.

  APROBAR  → envia el correo al proveedor via Graph API + notifica Teams
  RECHAZAR → descarta sin enviar

Uso:
    python watcher_aprobaciones.py
"""

import sys
import io
import warnings
import urllib3

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

from pathlib import Path

from Aprobaciones.cola_aprobacion import (
    listar_pendientes,
    leer_decision,
    leer_metadatos,
    leer_mensaje_editado,
    archivar,
)
from Integraciones.graph_client import GraphClient
from Integraciones.notificador  import notificar_sugerencia


def _procesar_archivo(ruta: Path, cliente: GraphClient):
    decision = leer_decision(ruta)
    if decision == "PENDIENTE":
        print(f"[PENDIENTE] {ruta.name} — sin decision, omitido.")
        return

    meta      = leer_metadatos(ruta)
    remitente = meta.get("remitente", "")
    asunto    = meta.get("asunto", "")
    mensaje   = leer_mensaje_editado(ruta)

    if decision == "APROBAR":
        print(f"[APROBAR] {ruta.name}")
        print(f"  Para   : {remitente}")
        print(f"  Asunto : {asunto}")

        if not remitente:
            print("  [ERROR] No se encontro destinatario — archivando sin enviar.")
            archivar(ruta, "RECHAZAR")
            return

        ok = cliente.enviar_correo(
            destinatario=remitente,
            asunto=asunto,
            cuerpo=mensaje,
        )
        if ok:
            print("  [OK] Correo enviado.")
            notificar_sugerencia(remitente, asunto, f"[ENVIADO]\n\n{mensaje}")
        else:
            print("  [ERROR] No se pudo enviar — revisar permisos Mail.Send en Azure AD.")
            return  # no archivar para poder reintentar

    elif decision == "RECHAZAR":
        print(f"[RECHAZAR] {ruta.name} — descartado sin enviar.")

    archivar(ruta, decision)


def main():
    print("=" * 60)
    print("APROBACIONES DE SUGERENCIAS")
    print("=" * 60)

    cliente    = GraphClient()
    pendientes = listar_pendientes()

    if not pendientes:
        print("  Sin sugerencias pendientes.")
        return

    print(f"  {len(pendientes)} archivo(s) encontrado(s)\n")
    for ruta in pendientes:
        try:
            _procesar_archivo(ruta, cliente)
        except Exception as e:
            print(f"[ERROR] {ruta.name}: {e}")

    print("\nListo.")


if __name__ == "__main__":
    main()
