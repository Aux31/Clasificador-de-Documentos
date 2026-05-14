# -*- coding: utf-8 -*-
"""
mover_carpeta_oc.py
-------------------
Mueve todos los archivos de una carpeta OC sin código de proveedor
(ej. cmer-OC-00206784) hacia la carpeta con código de proveedor
(ej. cmer-OC-00206784-EFM000C), incluyendo subcarpetas, y luego
elimina la carpeta origen que queda vacía.

Uso:
    python mover_carpeta_oc.py cmer-OC-00206784 cmer-OC-00206784-EFM000C
    python mover_carpeta_oc.py cmer-OC-00206784 cmer-OC-00206784-EFM000C --dry-run
"""

import sys
import argparse
import time
import msal
import requests
from urllib.parse import quote
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------------------------------------------------------------------------
# Config inline — mismas credenciales que el clasificador
# ---------------------------------------------------------------------------
import os
from pathlib import Path
from dotenv import load_dotenv

_env = Path(__file__).resolve().parent / "configuracion" / ".env"
load_dotenv(dotenv_path=_env, encoding="utf-8")

TENANT_ID       = os.getenv("GRAPH_TENANT_ID")
CLIENT_ID       = os.getenv("GRAPH_CLIENT_ID")
CLIENT_SECRET   = os.getenv("GRAPH_CLIENT_SECRET")
DRIVE_ID        = os.getenv("SHAREPOINT_DRIVE_ID")
CARPETA_OCS     = "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s"

_GRAPH = "https://graph.microsoft.com/v1.0"
_SCOPES = ["https://graph.microsoft.com/.default"]


# ---------------------------------------------------------------------------
# Auth + session
# ---------------------------------------------------------------------------

def _get_token():
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    r = app.acquire_token_for_client(scopes=_SCOPES)
    if "access_token" not in r:
        raise RuntimeError(f"No se pudo obtener token: {r.get('error_description')}")
    return r["access_token"]


def _session(token):
    s = requests.Session()
    s.headers.update({"Authorization": f"Bearer {token}"})
    retry = Retry(total=5, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=["GET", "POST", "DELETE", "PATCH"])
    s.mount("https://", HTTPAdapter(max_retries=retry))
    return s


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------

def _item_url(ruta: str) -> str:
    return f"{_GRAPH}/drives/{DRIVE_ID}/root:/{quote(ruta, safe='/')}:"


def listar_hijos(session, ruta: str) -> list[dict]:
    url = f"{_item_url(ruta)}/children?$select=id,name,file,folder,parentReference&$top=999"
    items = []
    while url:
        r = session.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def mover_item(session, item_id: str, nuevo_nombre: str, destino_id: str, dry_run: bool):
    if dry_run:
        print(f"  [DRY-RUN] Mover '{nuevo_nombre}' → id_destino={destino_id}")
        return
    url = f"{_GRAPH}/drives/{DRIVE_ID}/items/{item_id}"
    body = {
        "parentReference": {"driveId": DRIVE_ID, "id": destino_id},
        "name": nuevo_nombre,
    }
    r = session.patch(url, json=body, timeout=30)
    if not r.ok:
        raise RuntimeError(f"Error moviendo {nuevo_nombre}: {r.status_code} {r.text[:300]}")


def obtener_o_crear_carpeta(session, parent_id: str, nombre: str, dry_run: bool) -> str | None:
    url = f"{_GRAPH}/drives/{DRIVE_ID}/items/{parent_id}/children?$filter=name eq '{nombre}'&$select=id,name"
    r = session.get(url, timeout=30)
    r.raise_for_status()
    items = r.json().get("value", [])
    if items:
        return items[0]["id"]
    if dry_run:
        print(f"  [DRY-RUN] Crear subcarpeta '{nombre}' bajo id={parent_id}")
        return None
    body = {"name": nombre, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    r = session.post(f"{_GRAPH}/drives/{DRIVE_ID}/items/{parent_id}/children",
                     json=body, timeout=30)
    r.raise_for_status()
    return r.json()["id"]


def get_item_id(session, ruta: str) -> str | None:
    r = session.get(_item_url(ruta), timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()["id"]


def eliminar_item(session, item_id: str, nombre: str, dry_run: bool):
    if dry_run:
        print(f"  [DRY-RUN] Eliminar carpeta '{nombre}' (id={item_id})")
        return
    r = session.delete(f"{_GRAPH}/drives/{DRIVE_ID}/items/{item_id}", timeout=30)
    if r.status_code not in (200, 204):
        raise RuntimeError(f"Error eliminando {nombre}: {r.status_code} {r.text[:300]}")


# ---------------------------------------------------------------------------
# Lógica principal: mover recursivamente
# ---------------------------------------------------------------------------

def mover_recursivo(session, origen_id: str, origen_ruta: str,
                    destino_id: str, destino_ruta: str,
                    dry_run: bool, nivel: int = 0):
    sangria = "  " * nivel
    hijos = listar_hijos(session, origen_ruta)

    if not hijos:
        print(f"{sangria}(vacía)")
        return

    for item in hijos:
        nombre = item["name"]
        es_carpeta = "folder" in item

        if es_carpeta:
            print(f"{sangria}[carpeta] {nombre}")
            sub_destino_id = obtener_o_crear_carpeta(session, destino_id, nombre, dry_run)
            mover_recursivo(
                session,
                item["id"],
                f"{origen_ruta}/{nombre}",
                sub_destino_id or destino_id,
                f"{destino_ruta}/{nombre}",
                dry_run,
                nivel + 1,
            )
            eliminar_item(session, item["id"], nombre, dry_run)
        else:
            print(f"{sangria}[archivo] {nombre}")
            mover_item(session, item["id"], nombre, destino_id, dry_run)
            time.sleep(0.15)


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Mueve archivos entre carpetas OC en SharePoint")
    parser.add_argument("origen",  help="Carpeta sin código de proveedor, ej: cmer-OC-00206784")
    parser.add_argument("destino", help="Carpeta con código de proveedor, ej: cmer-OC-00206784-EFM000C")
    parser.add_argument("--dry-run", action="store_true",
                        help="Simula el proceso sin hacer cambios reales")
    args = parser.parse_args()

    ruta_origen  = f"{CARPETA_OCS}/{args.origen}"
    ruta_destino = f"{CARPETA_OCS}/{args.destino}"

    print(f"\nOrigen : {ruta_origen}")
    print(f"Destino: {ruta_destino}")
    if args.dry_run:
        print("MODO DRY-RUN — no se realizarán cambios\n")
    else:
        print()

    print("Autenticando con Graph API...")
    token   = _get_token()
    session = _session(token)

    print("Verificando carpetas...")
    origen_id = get_item_id(session, ruta_origen)
    if not origen_id:
        print(f"ERROR: No se encontró la carpeta origen: {ruta_origen}")
        sys.exit(1)

    destino_id = get_item_id(session, ruta_destino)
    if not destino_id:
        print(f"ERROR: No se encontró la carpeta destino: {ruta_destino}")
        sys.exit(1)

    print(f"\nMoviendo contenido de '{args.origen}' -> '{args.destino}'...\n")
    mover_recursivo(session, origen_id, ruta_origen, destino_id, ruta_destino,
                    dry_run=args.dry_run)

    print(f"\nEliminando carpeta origen '{args.origen}'...")
    eliminar_item(session, origen_id, args.origen, dry_run=args.dry_run)

    print("\nListo.")


if __name__ == "__main__":
    main()
