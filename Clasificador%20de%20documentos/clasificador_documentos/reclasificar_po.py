"""
Reclasifica documentos subidos a una PO incorrecta y los mueve a la PO correcta.

Flujo:
  1. Lista todos los archivos bajo la PO origen (recursivo, todas las subcarpetas).
  2. Para cada archivo construye la ruta equivalente bajo la PO destino.
  3. Crea las carpetas necesarias en la PO destino.
  4. Copia el archivo a la ruta destino (Graph API move via PATCH).
  5. Imprime un resumen y genera un log en /Registros/.

Uso:
    python reclasificar_po.py                         # usa PO_ORIGEN / PO_DESTINO hardcoded abajo
    python reclasificar_po.py 4060220260 177192        # o como argumentos posicionales
    python reclasificar_po.py --dry-run                # simula sin mover nada
"""

import sys
import io
import time
import logging
from pathlib import Path
from urllib.parse import quote
from datetime import datetime

# --- stdout UTF-8 en Windows ------------------------------------------------
if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# --- sys.path ----------------------------------------------------------------
_BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(_BASE))
sys.path.insert(0, str(_BASE.parent.parent / "Agente_Seguridad"))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient
from Nucleo.clasificador import formatear_numero_oc

# ---------------------------------------------------------------------------
# Configuración por defecto — se sobreescribe por argumentos CLI
# ---------------------------------------------------------------------------
PO_ORIGEN  = "4060220260"
PO_DESTINO = "177192"

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------
_LOG_DIR = _BASE / "Registros"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"reclasificar_po_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(str(_LOG_FILE), encoding="utf-8"),
    ],
)
log = logging.getLogger("reclasificar_po")


# ---------------------------------------------------------------------------
# Helpers Graph API
# ---------------------------------------------------------------------------

def _listar_archivos_recursivo(gc: GraphClient, carpeta_ruta: str) -> list[dict]:
    """
    Devuelve todos los ítems (archivos) bajo carpeta_ruta de forma recursiva.
    Cada ítem es un dict con 'id', 'name', 'ruta_relativa' (relativa a drive root),
    'parentReference.path' y 'folder' (presente si es carpeta).
    """
    archivos = []
    cola = [carpeta_ruta]

    while cola:
        ruta_actual = cola.pop(0)
        ruta_enc = quote(ruta_actual, safe="/")
        url = (
            f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}"
            f"/root:/{ruta_enc}:/children"
            f"?$select=id,name,folder,parentReference&$top=500"
        )
        try:
            data = gc._get(url)
        except Exception as e:
            log.warning(f"  No se pudo listar '{ruta_actual}': {e}")
            continue

        for item in data.get("value", []):
            ruta_item = f"{ruta_actual}/{item['name']}"
            if "folder" in item:
                cola.append(ruta_item)
            else:
                archivos.append({
                    "id":   item["id"],
                    "name": item["name"],
                    "ruta": ruta_item,
                })

        # Paginación
        next_link = data.get("@odata.nextLink")
        while next_link:
            try:
                gc._renovar_token_si_necesario()
                r = gc.session.get(next_link, headers=gc._headers, timeout=30)
                r.raise_for_status()
                data = r.json()
                for item in data.get("value", []):
                    ruta_item = f"{ruta_actual}/{item['name']}"
                    if "folder" in item:
                        cola.append(ruta_item)
                    else:
                        archivos.append({
                            "id":   item["id"],
                            "name": item["name"],
                            "ruta": ruta_item,
                        })
                next_link = data.get("@odata.nextLink")
            except Exception as e:
                log.warning(f"  Error paginando: {e}")
                break

    return archivos


def _mover_item(gc: GraphClient, item_id: str, nombre: str, ruta_padre_destino: str, dry_run: bool) -> bool:
    """
    Mueve un ítem (por ID) a la carpeta ruta_padre_destino.
    Usa PATCH /drives/{id}/items/{item-id} con parentReference + name.
    """
    if dry_run:
        log.info(f"  [DRY-RUN] Movería: {nombre}  →  {ruta_padre_destino}/{nombre}")
        return True

    # Obtener el ID de la carpeta destino
    ruta_enc = quote(ruta_padre_destino, safe="/")
    url_carpeta = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_enc}:"
    try:
        gc._renovar_token_si_necesario()
        r = gc.session.get(url_carpeta, headers=gc._headers, timeout=30)
        r.raise_for_status()
        carpeta_id = r.json()["id"]
    except Exception as e:
        log.error(f"  No se pudo obtener ID de carpeta destino '{ruta_padre_destino}': {e}")
        return False

    url_patch = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}"
    payload = {
        "parentReference": {"id": carpeta_id},
        "name": nombre,
    }
    try:
        gc._renovar_token_si_necesario()
        r = gc.session.patch(url_patch, headers=gc._headers, json=payload, timeout=60)
        if r.status_code in (200, 201):
            return True
        log.error(f"  PATCH {r.status_code}: {r.text[:300]}")
        return False
    except Exception as e:
        log.error(f"  Error moviendo '{nombre}': {e}")
        return False


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def reclasificar(po_origen: str, po_destino: str, dry_run: bool = False):
    log.info("=" * 70)
    log.info(f"Reclasificación PO  {po_origen}  →  {po_destino}")
    log.info(f"Modo: {'DRY-RUN (simulación)' if dry_run else 'REAL — los archivos serán movidos'}")
    log.info("=" * 70)

    gc = GraphClient()

    # Construir rutas raíz de cada OC
    oc_origen  = gc.buscar_carpeta_oc(formatear_numero_oc(po_origen))
    oc_destino = gc.buscar_carpeta_oc(formatear_numero_oc(po_destino))

    raiz_origen  = f"{SHAREPOINT_CARPETA_OCS}/{oc_origen}"
    raiz_destino = f"{SHAREPOINT_CARPETA_OCS}/{oc_destino}"

    log.info(f"Carpeta origen : {raiz_origen}")
    log.info(f"Carpeta destino: {raiz_destino}")
    log.info("")

    # 1. Listar todos los archivos bajo la OC origen
    log.info("Listando archivos en la OC origen...")
    archivos = _listar_archivos_recursivo(gc, raiz_origen)
    log.info(f"  {len(archivos)} archivo(s) encontrado(s)")

    if not archivos:
        log.warning("No hay archivos que mover. Terminando.")
        return

    # 2. Procesar cada archivo
    movidos   = 0
    fallidos  = 0

    for arch in archivos:
        ruta_origen_completa = arch["ruta"]

        # Calcular ruta relativa dentro de la OC (quitando la raíz origen)
        ruta_relativa = ruta_origen_completa[len(raiz_origen):].lstrip("/")

        # Construir ruta destino (reemplazando la raíz)
        ruta_destino_completa = f"{raiz_destino}/{ruta_relativa}"
        carpeta_destino = str(Path(ruta_destino_completa).parent).replace("\\", "/")

        log.info(f"[{movidos + fallidos + 1}/{len(archivos)}] {arch['name']}")
        log.info(f"  Desde: {ruta_origen_completa}")
        log.info(f"  Hacia: {ruta_destino_completa}")

        # 3. Crear carpeta destino si no existe
        if not dry_run:
            try:
                gc.crear_carpeta_si_no_existe(ruta_destino_completa)
            except Exception as e:
                log.warning(f"  crear_carpeta_si_no_existe: {e}")

        # 4. Mover
        ok = _mover_item(gc, arch["id"], arch["name"], carpeta_destino, dry_run)
        if ok:
            movidos += 1
            log.info("  OK")
        else:
            fallidos += 1
            log.error("  FALLIDO")

        time.sleep(0.3)  # respetar throttling de Graph API

    # 5. Resumen
    log.info("")
    log.info("=" * 70)
    log.info(f"RESUMEN: {movidos} movidos  |  {fallidos} fallidos  |  {len(archivos)} total")
    log.info(f"Log guardado en: {_LOG_FILE}")
    log.info("=" * 70)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    dry_run    = "--dry-run" in sys.argv
    args_limpios = [a for a in sys.argv[1:] if not a.startswith("-")]

    po_origen  = args_limpios[0] if len(args_limpios) >= 1 else PO_ORIGEN
    po_destino = args_limpios[1] if len(args_limpios) >= 2 else PO_DESTINO

    reclasificar(po_origen, po_destino, dry_run=dry_run)
