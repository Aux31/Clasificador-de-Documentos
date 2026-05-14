# -*- coding: utf-8 -*-
"""
clone_oc.py — Clona carpetas OC en SharePoint hacia carpetas de prueba.

Clonaciones:
  cmer-OC-00194453-IITC000  →  Prueba_1
  cmer-OC-00191063-TCHC001  →  Prueba_2

Copia toda la estructura de carpetas y archivos usando la Graph API.
Los archivos se copian descargando el contenido y subiéndolo al destino
(más confiable que la copia asíncrona de Graph para archivos cross-folder).

Uso:
  python clone_oc.py
  python clone_oc.py --source cmer-OC-00194453-IITC000 --dest Prueba_1
"""

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Clasificador de documentos", "clasificador_documentos"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Clasificador%20de%20documentos", "clasificador_documentos"))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient

_GRAPH_BASE  = "https://graph.microsoft.com/v1.0"
SCRIPT_DIR   = Path(__file__).parent
MAX_RETRIES  = 3
RETRY_DELAY  = 3  # segundos entre reintentos

CLONES = [
    {"source": "cmer-OC-00194453-IITC000", "dest": "Prueba_1"},
    {"source": "cmer-OC-00191063-TCHC001", "dest": "Prueba_2"},
]


# ---------------------------------------------------------------------------
# Helpers de red con reintentos
# ---------------------------------------------------------------------------

def _con_reintento(fn, *args, **kwargs):
    """Ejecuta fn(*args, **kwargs) hasta MAX_RETRIES veces ante excepciones."""
    ultimo_error = None
    for intento in range(1, MAX_RETRIES + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            ultimo_error = e
            if intento < MAX_RETRIES:
                time.sleep(RETRY_DELAY * intento)
    raise ultimo_error


# ---------------------------------------------------------------------------
# Helpers Graph API
# ---------------------------------------------------------------------------

def _obtener_item(gc: GraphClient, ruta: str) -> dict | None:
    encoded = quote(ruta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{encoded}:"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def _obtener_root_id(gc: GraphClient) -> str:
    item = _obtener_item(gc, SHAREPOINT_CARPETA_OCS)
    if not item:
        sys.exit(f"[ERROR] No se encontró la carpeta raíz: {SHAREPOINT_CARPETA_OCS}")
    return item["id"]


def _listar_todos(gc: GraphClient, item_id: str) -> list[dict]:
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/children?$select=id,name,folder,file,size&$top=999"
    items = []
    while url:
        data = _con_reintento(gc._get, url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _crear_carpeta(gc: GraphClient, parent_id: str, nombre: str) -> dict:
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{parent_id}/children"
    return _con_reintento(
        gc._post, url,
        {"name": nombre, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
    )


def _descargar_archivo(gc: GraphClient, file_id: str) -> bytes:
    """Descarga el contenido binario de un archivo."""
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{file_id}/content"
    gc._renovar_token_si_necesario()

    def _get_content():
        r = gc.session.get(url, headers=gc._headers, timeout=120, allow_redirects=True)
        r.raise_for_status()
        return r.content

    return _con_reintento(_get_content)


def _subir_archivo(gc: GraphClient, parent_id: str, nombre: str, contenido: bytes) -> dict:
    """Sube un archivo al destino. Usa upload session para archivos > 4 MB."""
    SIZE_LIMIT = 4 * 1024 * 1024  # 4 MB

    if len(contenido) <= SIZE_LIMIT:
        # Subida simple
        encoded = quote(nombre)
        url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{parent_id}:/{encoded}:/content"
        gc._renovar_token_si_necesario()

        def _put():
            headers = dict(gc._headers)
            headers["Content-Type"] = "application/octet-stream"
            r = gc.session.put(url, headers=headers, data=contenido, timeout=120)
            r.raise_for_status()
            return r.json()

        return _con_reintento(_put)

    else:
        # Upload session para archivos grandes
        session_url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{parent_id}:/{quote(nombre)}:/createUploadSession"
        gc._renovar_token_si_necesario()
        r = gc.session.post(session_url, headers=gc._headers,
                            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
                            timeout=30)
        r.raise_for_status()
        upload_url = r.json()["uploadUrl"]

        # Subir en fragmentos de 4 MB
        chunk_size = SIZE_LIMIT
        total = len(contenido)
        result = None
        for offset in range(0, total, chunk_size):
            chunk = contenido[offset:offset + chunk_size]
            end   = offset + len(chunk) - 1
            headers = {
                "Content-Range": f"bytes {offset}-{end}/{total}",
                "Content-Length": str(len(chunk)),
            }

            def _put_chunk(h=headers, c=chunk):
                r2 = gc.session.put(upload_url, headers=h, data=c, timeout=120)
                r2.raise_for_status()
                return r2.json() if r2.content else {}

            result = _con_reintento(_put_chunk)

        return result or {}


# ---------------------------------------------------------------------------
# Lógica de clonación recursiva
# ---------------------------------------------------------------------------

def _clonar_recursivo(gc: GraphClient, src_id: str, dest_id: str, ruta: str, log: list) -> None:
    hijos = _listar_todos(gc, src_id)

    for hijo in hijos:
        nombre    = hijo["name"]
        ruta_hijo = f"{ruta}/{nombre}"

        if "folder" in hijo:
            try:
                nueva = _crear_carpeta(gc, dest_id, nombre)
                print(f"  [DIR]  {ruta_hijo}")
                log.append({"path": ruta_hijo, "type": "folder", "result": "created"})
                _clonar_recursivo(gc, hijo["id"], nueva["id"], ruta_hijo, log)
            except Exception as e:
                print(f"  [ERR]  {ruta_hijo}: {e}")
                log.append({"path": ruta_hijo, "type": "folder", "result": "error", "detail": str(e)})

        elif "file" in hijo:
            try:
                contenido = _descargar_archivo(gc, hijo["id"])
                _subir_archivo(gc, dest_id, nombre, contenido)
                size_kb = len(contenido) // 1024
                print(f"  [FILE] {ruta_hijo}  ({size_kb} KB)")
                log.append({"path": ruta_hijo, "type": "file", "result": "copied", "size": len(contenido)})
            except Exception as e:
                print(f"  [ERR]  {ruta_hijo}: {e}")
                log.append({"path": ruta_hijo, "type": "file", "result": "error", "detail": str(e)})


def clonar_oc(gc: GraphClient, source: str, dest: str) -> dict:
    ruta_src  = f"{SHAREPOINT_CARPETA_OCS}/{source}"
    ruta_dest = f"{SHAREPOINT_CARPETA_OCS}/{dest}"

    print(f"\n{'='*60}")
    print(f"  Origen : {source}")
    print(f"  Destino: {dest}")

    src_item = _obtener_item(gc, ruta_src)
    if not src_item:
        msg = f"No se encontró la carpeta origen: {ruta_src}"
        print(f"  [ERROR] {msg}")
        return {"source": source, "dest": dest, "result": "error", "detail": msg, "operations": []}

    dest_item = _obtener_item(gc, ruta_dest)
    if dest_item:
        print(f"  [INFO]  La carpeta destino '{dest}' ya existe — se sobreescribirá el contenido.")
        dest_id = dest_item["id"]
    else:
        root_id   = _obtener_root_id(gc)
        dest_item = _crear_carpeta(gc, root_id, dest)
        dest_id   = dest_item["id"]
        print(f"  [DIR]  {dest} (creada)")

    log = []
    _clonar_recursivo(gc, src_item["id"], dest_id, dest, log)

    dirs  = sum(1 for op in log if op["type"] == "folder" and op["result"] == "created")
    files = sum(1 for op in log if op["type"] == "file"   and op["result"] == "copied")
    errs  = sum(1 for op in log if op["result"] == "error")

    print(f"\n  Resumen: {dirs} carpetas creadas, {files} archivos copiados, {errs} errores")
    return {
        "source": source, "dest": dest,
        "result": "ok" if errs == 0 else "partial",
        "folders_created": dirs, "files_copied": files, "errors": errs,
        "operations": log,
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Clona carpetas OC en SharePoint")
    parser.add_argument("--source", default=None)
    parser.add_argument("--dest",   default=None)
    args = parser.parse_args()

    gc    = GraphClient()
    pares = [{"source": args.source, "dest": args.dest}] if args.source and args.dest else CLONES

    resultados = []
    for par in pares:
        res = clonar_oc(gc, par["source"], par["dest"])
        resultados.append(res)

    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = SCRIPT_DIR / f"clone_log_{ts}.json"
    log_file.write_text(
        json.dumps({"generated_at": datetime.now(timezone.utc).isoformat(), "clones": resultados},
                   ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    print(f"\n[OK] Log guardado en {log_file}")


if __name__ == "__main__":
    main()
