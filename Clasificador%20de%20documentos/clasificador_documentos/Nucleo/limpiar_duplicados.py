"""
Detecta y elimina archivos duplicados en SharePoint por CONTENIDO (hash).

El archivo duplicado puede tener un nombre completamente distinto al original
porque el clasificador lo renombra diferente en cada ejecución. Este script
compara el hash que SharePoint ya calcula internamente (quickXorHash) — sin
necesidad de descargar los archivos.

Lógica:
  1. Lee registros_subidas.log para obtener las carpetas afectadas.
  2. Lista todos los archivos de cada carpeta con su hash.
  3. Agrupa por hash idéntico → mismo contenido = duplicado.
  4. Muestra los grupos con 2+ archivos.
  5. Con --borrar, elimina el más reciente de cada grupo (conserva el original).

Uso:
    python limpiar_duplicados.py              # solo muestra duplicados
    python limpiar_duplicados.py --borrar     # borra el más reciente de cada grupo
"""

import sys
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from configuracion.ajustes import TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_DRIVE_ID
import msal
import requests

_GRAPH_BASE  = "https://graph.microsoft.com/v1.0"
_LOG_SUBIDAS = Path(__file__).parent.parent / "registros_subidas.log"


def _obtener_token() -> str:
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError(f"No se pudo obtener token: {res.get('error_description')}")
    return res["access_token"]


def _ocs_del_log() -> set[str]:
    """Rutas base de cada OC que aparece en el log (hasta '4. DOCUMENTACION')."""
    if not _LOG_SUBIDAS.exists():
        print(f"[ERROR] No existe: {_LOG_SUBIDAS}")
        return set()
    ocs = set()
    for linea in _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines():
        partes = linea.strip().split(" | ")
        if len(partes) < 4:
            continue
        ruta = partes[3].strip().split(" | YA_EXISTIA")[0].strip()
        partes_ruta = ruta.split("/")
        # Buscar el segmento "4. DOCUMENTACION" y quedarse con la ruta hasta ahí
        try:
            idx = next(i for i, p in enumerate(partes_ruta) if p.startswith("4. DOCUMENTACION"))
            ocs.add("/".join(partes_ruta[:idx + 1]))
        except StopIteration:
            continue
    return ocs


def _subcarpetas_documentacion(session, headers, ruta_doc: str) -> list[str]:
    """Devuelve todas las subcarpetas dentro de '4. DOCUMENTACION/' de una OC."""
    url = (
        f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
        f"/root:/{quote(ruta_doc, safe='/')}:/children"
        f"?$select=name,folder&$top=200"
    )
    try:
        r = session.get(url, headers=headers, timeout=30)
        if r.status_code == 404:
            return []
        r.raise_for_status()
        return [
            f"{ruta_doc}/{item['name']}"
            for item in r.json().get("value", [])
            if "folder" in item
        ]
    except Exception as e:
        print(f"  [WARN] No se pudo listar subcarpetas de '{ruta_doc}': {e}")
        return []


def _hash_archivo(item: dict) -> str | None:
    """Devuelve el mejor hash disponible, o None si no hay ninguno."""
    hashes = item.get("file", {}).get("hashes", {})
    return (
        hashes.get("quickXorHash")
        or hashes.get("sha256Hash")
        or hashes.get("sha1Hash")
    )


def _clave_comparacion(item: dict) -> str | None:
    """
    Clave para detectar duplicados:
    - Si hay hash → lo usa (más confiable).
    - Si no → usa tamaño en bytes como fallback (puede dar falsos positivos
      en archivos del mismo tamaño pero distinto contenido, por eso se avisa).
    """
    h = _hash_archivo(item)
    if h:
        return f"hash:{h}"
    size = item.get("size")
    if size is not None:
        return f"size:{size}"
    return None


def _listar_archivos(session, headers, carpeta: str) -> list[dict]:
    """Lista archivos con nombre, id, tamaño, hashes y fecha de creación."""
    url = (
        f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
        f"/root:/{quote(carpeta, safe='/')}:/children"
        f"?$select=name,id,createdDateTime,size,file&$top=200"
    )
    try:
        r = session.get(url, headers=headers, timeout=30)
        if r.status_code == 404:
            return []
        r.raise_for_status()
        # Solo archivos (no subcarpetas)
        return [item for item in r.json().get("value", []) if "file" in item]
    except Exception as e:
        print(f"  [WARN] No se pudo listar '{carpeta}': {e}")
        return []



def ejecutar(borrar: bool = False):
    ocs = _ocs_del_log()
    if not ocs:
        print("No se encontraron OCs en el log.")
        return

    token   = _obtener_token()
    headers = {"Authorization": f"Bearer {token}"}
    session = requests.Session()

    # Expandir cada OC a todas sus subcarpetas de documentación
    carpetas: set[str] = set()
    for ruta_doc in sorted(ocs):
        subs = _subcarpetas_documentacion(session, headers, ruta_doc)
        carpetas.update(subs)

    print(f"OCs en el log: {len(ocs)} → subcarpetas a revisar: {len(carpetas)}\n")

    duplicados: list[tuple[str, str, str, str]] = []  # (carpeta, nombre_orig, nombre_dup, id_dup)

    for carpeta in sorted(carpetas):
        archivos = _listar_archivos(session, headers, carpeta)
        if len(archivos) < 2:
            continue

        # Agrupar por clave de contenido (hash o tamaño como fallback)
        por_clave: dict[str, list[dict]] = {}
        sin_clave = []
        for arch in archivos:
            clave = _clave_comparacion(arch)
            if clave:
                por_clave.setdefault(clave, []).append(arch)
            else:
                sin_clave.append(arch["name"])
        if sin_clave:
            print(f"  [AVISO] Sin hash ni tamaño (no comparables): {', '.join(sin_clave)}")

        for clave, grupo in por_clave.items():
            if len(grupo) < 2:
                continue
            es_fallback = clave.startswith("size:")
            # El más antiguo es el original
            grupo.sort(key=lambda x: x.get("createdDateTime", ""))
            original = grupo[0]
            carpeta_corta = "/".join(carpeta.split("/")[-2:])
            etiqueta = "[POSIBLE DUPLICADO — mismo tamaño]" if es_fallback else "[DUPLICADO]"
            print(f"  {etiqueta} .../{carpeta_corta}/")
            print(f"    Original  : {original['name']}  ({original.get('createdDateTime','?')[:10]})")
            for dup in grupo[1:]:
                print(f"    Duplicado : {dup['name']}  ({dup.get('createdDateTime','?')[:10]})")
                duplicados.append((carpeta, original["name"], dup["name"], dup["id"]))

    if not duplicados:
        print("No se encontraron archivos duplicados por contenido.")
        return

    print(f"\nTotal de duplicados encontrados: {len(duplicados)}")
    print("(El borrado de duplicados esta deshabilitado. Eliminalos manualmente desde SharePoint.)")
    print("\nListo.")


if __name__ == "__main__":
    ejecutar(borrar="--borrar" in sys.argv)
