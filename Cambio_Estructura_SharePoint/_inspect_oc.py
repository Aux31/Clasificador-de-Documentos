# -*- coding: utf-8 -*-
"""
_inspect_oc.py — Lista recursivamente todas las carpetas de una OC en SharePoint
y las compara contra la estructura definida en migration-config.json.
"""

import json
import sys
import os
from pathlib import Path
from urllib.parse import quote

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).parent

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Clasificador de documentos", "clasificador_documentos"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Clasificador%20de%20documentos", "clasificador_documentos"))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
OC_TARGET   = "cmer-OC-00203761-AIP000"


def _get_item(gc, ruta):
    encoded = quote(ruta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{encoded}:"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def _list_folders(gc, item_id):
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/children?$select=id,name,folder&$top=999"
    items = []
    while url:
        data = gc._get(url)
        items.extend(h for h in data.get("value", []) if "folder" in h)
        url = data.get("@odata.nextLink")
    return items


def _list_recursive(gc, item_id, prefix=""):
    """Retorna lista de rutas relativas de todas las carpetas bajo item_id."""
    result = []
    folders = _list_folders(gc, item_id)
    for f in sorted(folders, key=lambda x: x["name"]):
        rel = f"{prefix}/{f['name']}" if prefix else f["name"]
        result.append(rel)
        result.extend(_list_recursive(gc, f["id"], rel))
    return result


def load_config():
    config_file = SCRIPT_DIR / "migration-config.json"
    with config_file.open(encoding="utf-8") as fh:
        return json.load(fh)


def main():
    config = load_config()
    gc = GraphClient()

    ruta_oc = f"{SHAREPOINT_CARPETA_OCS}/{OC_TARGET}"
    print(f"Buscando: {ruta_oc}")
    oc_item = _get_item(gc, ruta_oc)
    if not oc_item:
        print(f"[ERROR] No se encontro la carpeta: {ruta_oc}")
        sys.exit(1)

    print(f"Encontrada (id={oc_item['id']}). Listando carpetas recursivamente...\n")
    sp_folders = _list_recursive(gc, oc_item["id"])

    print("=" * 60)
    print(f"ESTRUCTURA ACTUAL EN SHAREPOINT ({len(sp_folders)} carpetas)")
    print("=" * 60)
    for f in sp_folders:
        print(f"  {f}")

    # ----------------------------------------------------------------
    # Comparacion contra config
    # ----------------------------------------------------------------
    new_struct  = config["new_structure"]
    cfg_folders = set(new_struct["folders"])           # carpetas raiz esperadas
    cfg_subfolders: dict = new_struct["subfolders"]    # subcarpetas por carpeta padre

    # Construir set completo de rutas esperadas segun config
    expected_paths: set[str] = set()
    for folder in new_struct["folders"]:
        expected_paths.add(folder)
        for sub in cfg_subfolders.get(folder, []):
            expected_paths.add(f"{folder}/{sub}")

    sp_set = set(sp_folders)

    # Solo comparar en primer y segundo nivel (ignorar niveles mas profundos)
    sp_level1 = {p for p in sp_set if "/" not in p}
    sp_level2 = {p for p in sp_set if p.count("/") == 1}
    sp_shallow = sp_level1 | sp_level2

    only_in_sp       = sp_shallow - expected_paths   # existen en SP pero no en config
    only_in_config   = expected_paths - sp_shallow   # estan en config pero no en SP

    print()
    print("=" * 60)
    print("COMPARACION CONTRA migration-config.json")
    print("=" * 60)

    print(f"\n[1] En SharePoint pero NO en config ({len(only_in_sp)}):")
    if only_in_sp:
        for p in sorted(only_in_sp):
            print(f"  SP_EXTRA: {p}")
    else:
        print("  (ninguna — todas las carpetas de SP estan en el config)")

    print(f"\n[2] En config pero NO en SharePoint ({len(only_in_config)}):")
    if only_in_config:
        for p in sorted(only_in_config):
            print(f"  CONFIG_FALTANTE: {p}")
    else:
        print("  (ninguna — config y SharePoint coinciden completamente)")

    print()
    # Resumen niveles mas profundos (si hay)
    sp_deep = {p for p in sp_set if p.count("/") > 1}
    if sp_deep:
        print(f"[NOTA] {len(sp_deep)} carpetas en nivel 3+ (no comparadas contra config):")
        for p in sorted(sp_deep):
            print(f"  DEEP: {p}")


if __name__ == "__main__":
    main()
