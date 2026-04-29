"""
limpiar_versiones.py
--------------------
Enumera TODOS los archivos del drive SharePoint con una sola pasada delta,
identifica archivos _vN que tienen su original en la misma carpeta,
muestra la lista y borra previa confirmación.

Uso:
    python c:\\Users\\aux22.gg\\Desktop\\PROYECTOS\\SHAREPOINT\\clasificador_documentos\\pruebas\\limpiar_versiones.py
"""

import re
import sys
import urllib3
from pathlib import Path
from urllib.parse import quote

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import requests
import msal
from configuracion.ajustes import (
    TENANT_ID, CLIENT_ID, CLIENT_SECRET,
    SHAREPOINT_DRIVE_ID,
)

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_RE_VERSION = re.compile(r"^(.+?)(_v\d+)(\.[^.]+)$", re.IGNORECASE)


def obtener_token() -> dict:
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"Token error: {res.get('error_description')}")
    return {"Authorization": f"Bearer {res['access_token']}"}


def enumerar_todos_los_archivos(headers: dict) -> dict[str, str]:
    """
    Usa el delta query para obtener un inventario completo del drive.
    Devuelve un dict: { ruta_relativa_lower -> item_id }
    La ruta es relativa a la raíz del drive, con '/' como separador.
    """
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root/delta"
    params = {
        "$select": "id,name,parentReference,file",
        "$top": 1000,
    }

    # Construimos un mapa de id -> parentPath para reconstruir rutas
    items_raw = []
    pagina = 1

    while url:
        print(f"  Descargando página {pagina}...", end=" ", flush=True)
        r = requests.get(url, headers=headers, params=params if pagina == 1 else None, verify=False)
        r.raise_for_status()
        data = r.json()
        batch = data.get("value", [])
        items_raw.extend(batch)
        print(f"{len(batch)} ítems")
        pagina += 1
        url = data.get("@odata.nextLink")

    print(f"\n  Total ítems descargados: {len(items_raw)}")

    # Filtrar solo archivos (no carpetas) y construir mapa ruta -> id
    archivo_map = {}
    for item in items_raw:
        if "file" not in item:
            continue  # es carpeta, saltar
        parent_path = item.get("parentReference", {}).get("path", "")
        # parent_path es algo como "/drives/{id}/root:/OC´s/cmer-OC-00196893/4. DOCUMENTACION"
        # Extraemos solo la parte después de "root:/"
        if "/root:/" in parent_path:
            carpeta = parent_path.split("/root:/", 1)[1]
        elif "/root:" in parent_path:
            carpeta = ""
        else:
            carpeta = parent_path

        nombre = item["name"]
        ruta = f"{carpeta}/{nombre}" if carpeta else nombre
        archivo_map[ruta] = item["id"]

    return archivo_map



def main():
    print("Conectando a SharePoint...")
    headers = obtener_token()

    print("Descargando inventario completo del drive (delta query)...")
    archivo_map = enumerar_todos_los_archivos(headers)

    if not archivo_map:
        print("No se encontraron archivos en el drive.")
        return

    print(f"\nAnalizando {len(archivo_map)} archivos en busca de duplicados _vN...")

    # Para búsqueda case-insensitive, construimos también un set de rutas en minúscula -> ruta_real
    rutas_lower = {r.lower(): r for r in archivo_map}

    candidatos = []
    for ruta, item_id in archivo_map.items():
        nombre = ruta.rsplit("/", 1)[-1] if "/" in ruta else ruta
        m = _RE_VERSION.match(nombre)
        if not m:
            continue
        stem, _v, ext = m.groups()
        nombre_original = f"{stem}{ext}"
        carpeta = ruta.rsplit("/", 1)[0] if "/" in ruta else ""
        ruta_original = f"{carpeta}/{nombre_original}" if carpeta else nombre_original

        # Verificar si el original existe en el drive (case-insensitive)
        if ruta_original.lower() in rutas_lower:
            ruta_original_real = rutas_lower[ruta_original.lower()]
            candidatos.append({
                "ruta_vn": ruta,
                "id_vn": item_id,
                "ruta_original": ruta_original_real,
                "nombre": nombre,
                "original": nombre_original,
            })

    if not candidatos:
        print("No se encontraron archivos _vN con original en la misma carpeta.")
        return

    # Ordenar por ruta para presentar agrupado
    candidatos.sort(key=lambda c: c["ruta_vn"])

    print(f"\nSe borrarán {len(candidatos)} archivos _vN (el original se conserva):\n")
    for i, c in enumerate(candidatos, 1):
        print(f"  {i:3}. BORRAR:   {c['ruta_vn']}")
        print(f"       CONSERVAR: {c['ruta_original']}")

    print(f"\nTotal de versiones _vN encontradas: {len(candidatos)}")
    print("(El borrado de versiones esta deshabilitado. Eliminalas manualmente desde SharePoint.)")


if __name__ == "__main__":
    main()
