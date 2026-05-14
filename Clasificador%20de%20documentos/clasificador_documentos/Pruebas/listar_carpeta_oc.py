"""
Lista el contenido de una subcarpeta de una OC en SharePoint.

Uso:
    cd clasificador_documentos
    python Pruebas/listar_carpeta_oc.py cmer-OC-00187133 OTROS
    python Pruebas/listar_carpeta_oc.py cmer-OC-00187133 "4. DOCUMENTACION"
"""

import sys
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def listar(gc, ruta_carpeta: str):
    ruta_enc = quote(ruta_carpeta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_enc}:/children?$select=name,file,folder&$top=999"
    items = []
    while url:
        data = gc._get(url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def main():
    oc        = sys.argv[1] if len(sys.argv) > 1 else "cmer-OC-00187133"
    subcarpeta = sys.argv[2] if len(sys.argv) > 2 else "OTROS"

    ruta = f"{SHAREPOINT_CARPETA_OCS}/{oc}/4. DOCUMENTACION/{subcarpeta}"
    print(f"\nListando: {ruta}\n")

    gc = GraphClient()
    try:
        items = listar(gc, ruta)
    except Exception as e:
        print(f"[ERROR] {e}")
        return

    if not items:
        print("(carpeta vacía o no existe)")
        return

    for item in items:
        tipo = "[DIR]" if "folder" in item else "[ARQ]"
        print(f"  {tipo} {item['name']!r}")

    print(f"\nTotal: {len(items)} item(s)")


if __name__ == "__main__":
    main()
