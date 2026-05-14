"""
Mueve todos los archivos de la PO 10003696 (carpeta cmer-OC-10003696)
a la carpeta correcta de la PO 186697 (cmer-OC-00186697-AKSB000).

El script descubre automáticamente todos los archivos en la carpeta origen
(recorriendo subcarpetas) y los mueve a la misma estructura en destino.
"""

import sys
import io
import os
import time
import logging
from pathlib import Path
from datetime import datetime
from urllib.parse import quote
import requests

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

_BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_BASE))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID
from Integraciones.graph_client import GraphClient

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
CARPETA_OCS = "Centro de Documentación Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC´s"
CARPETA_ORIGEN  = f"{CARPETA_OCS}/cmer-OC-10003696/4. DOCUMENTACION"
CARPETA_DESTINO = f"{CARPETA_OCS}/cmer-OC-00186697-AKSB000/4. DOCUMENTACION"

# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------
_LOG_DIR = _BASE / "Registros"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"mover_po10003696_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(str(_LOG_FILE), encoding="utf-8"),
    ],
)
log = logging.getLogger("mover_po10003696")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_item(gc: GraphClient, ruta: str) -> dict | None:
    ruta_encoded = quote(ruta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_encoded}:"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    if r.status_code == 200:
        return r.json()
    if r.status_code == 404:
        return None
    r.raise_for_status()


def listar_hijos(gc: GraphClient, item_id: str) -> list[dict]:
    """Lista todos los hijos (archivos y carpetas) de un item por su id."""
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/children?$top=999"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])


def obtener_o_crear_carpeta(gc: GraphClient, ruta_carpeta: str) -> str | None:
    """Obtiene el id de la carpeta; la crea si no existe."""
    item = get_item(gc, ruta_carpeta)
    if item:
        return item["id"]
    # Crear carpeta usando el método existente del cliente
    gc.crear_carpeta_si_no_existe(ruta_carpeta)
    item = get_item(gc, ruta_carpeta)
    return item["id"] if item else None


def copiar_item(gc: GraphClient, item_id: str, dest_folder_id: str, nombre: str) -> str | None:
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/copy"
    payload = {
        "parentReference": {"driveId": SHAREPOINT_DRIVE_ID, "id": dest_folder_id},
        "name": nombre,
    }
    gc._renovar_token_si_necesario()
    r = gc.session.post(
        url,
        headers={**gc._headers, "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    if r.status_code == 202:
        return r.headers.get("Location")
    r.raise_for_status()


def esperar_copia(gc: GraphClient, monitor_url: str, timeout: int = 180) -> bool:
    inicio = time.time()
    while time.time() - inicio < timeout:
        r = gc.session.get(monitor_url, timeout=30)
        if r.status_code == 202:
            time.sleep(2)
            continue
        if r.status_code == 200:
            data = r.json()
            status = data.get("status", "")
            if status == "completed":
                return True
            if status == "failed":
                log.error(f"Copia falló: {data}")
                return False
            time.sleep(2)
        else:
            log.warning(f"Monitor respondió {r.status_code}")
            time.sleep(2)
    log.error("Timeout esperando copia")
    return False


def eliminar_item(gc: GraphClient, item_id: str):
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}"
    gc._renovar_token_si_necesario()
    r = gc.session.delete(url, headers=gc._headers, timeout=30)
    if r.status_code not in (200, 204):
        r.raise_for_status()


def descubrir_archivos(gc: GraphClient, item_id: str, ruta_relativa: str = "") -> list[tuple[str, str, str]]:
    """
    Recorre recursivamente la carpeta origen.
    Retorna lista de (item_id, ruta_relativa_subcarpeta, nombre_archivo).
    """
    resultados = []
    hijos = listar_hijos(gc, item_id)
    for hijo in hijos:
        nombre = hijo["name"]
        if "folder" in hijo:
            subruta = f"{ruta_relativa}/{nombre}" if ruta_relativa else nombre
            resultados.extend(descubrir_archivos(gc, hijo["id"], subruta))
        else:
            resultados.append((hijo["id"], ruta_relativa, nombre))
    return resultados


def main():
    dry_run = "--dry-run" in sys.argv
    log.info("=" * 70)
    log.info(f"MOVER PO 10003696 → 186697 — {'DRY-RUN' if dry_run else 'MODO REAL'}")
    log.info(f"  Origen:  {CARPETA_ORIGEN}")
    log.info(f"  Destino: {CARPETA_DESTINO}")
    log.info("=" * 70)

    gc = GraphClient()

    # Verificar carpeta origen
    origen_item = get_item(gc, CARPETA_ORIGEN)
    if origen_item is None:
        log.error(f"Carpeta origen no encontrada: {CARPETA_ORIGEN}")
        sys.exit(1)

    log.info("Descubriendo archivos en carpeta origen...")
    archivos = descubrir_archivos(gc, origen_item["id"])
    log.info(f"  Encontrados: {len(archivos)} archivos")

    if not archivos:
        log.info("No hay archivos que mover.")
        return

    movidos = 0
    errores = 0

    for item_id, subcarpeta, nombre in archivos:
        ruta_dest_carpeta = f"{CARPETA_DESTINO}/{subcarpeta}" if subcarpeta else CARPETA_DESTINO
        display_origen = f".../{subcarpeta}/{nombre}" if subcarpeta else f".../{nombre}"
        display_dest  = f"...AKSB000/4. DOCUMENTACION/{subcarpeta}/{nombre}" if subcarpeta else f"...AKSB000/4. DOCUMENTACION/{nombre}"

        log.info(f"\n  Archivo: {nombre}")
        log.info(f"  Origen:  {display_origen}")
        log.info(f"  Destino: {display_dest}")

        if dry_run:
            log.info(f"  [DRY-RUN] Se copiaría y eliminaría (item_id={item_id})")
            continue

        # Obtener/crear carpeta destino
        dest_folder_id = obtener_o_crear_carpeta(gc, ruta_dest_carpeta)
        if dest_folder_id is None:
            log.error(f"  [ERROR] No se pudo obtener/crear carpeta destino: {ruta_dest_carpeta}")
            errores += 1
            continue

        # Copiar
        log.info(f"  Copiando...")
        monitor_url = copiar_item(gc, item_id, dest_folder_id, nombre)
        if monitor_url is None:
            log.error(f"  [ERROR] No se obtuvo URL de monitoreo")
            errores += 1
            continue

        if not esperar_copia(gc, monitor_url):
            log.error(f"  [ERROR] Fallo al copiar {nombre}")
            errores += 1
            continue

        log.info(f"  Copia completada.")

        # Eliminar origen
        log.info(f"  Eliminando de carpeta origen...")
        eliminar_item(gc, item_id)
        log.info(f"  [OK] Movido correctamente.")
        movidos += 1

    log.info("\n" + "=" * 70)
    log.info(f"Movidos: {movidos} | Errores: {errores}")
    log.info(f"Log: {_LOG_FILE}")
    log.info("=" * 70)


if __name__ == "__main__":
    main()
