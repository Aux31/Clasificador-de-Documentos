"""
Mueve los 16 archivos de PO 201096 de la carpeta incorrecta (sin código de proveedor)
a la carpeta correcta (con código de proveedor -EFM000C).

Acciones:
  1. Para cada archivo: GET item-id en carpeta incorrecta → DELETE
  2. Re-sube el fragmento desde la carpeta temporal (o lo copia via PATCH move en Graph API)

Como los fragmentos ya no existen en disco (fueron temporales), usamos PATCH /copy en Graph API
para mover dentro de SharePoint sin necesidad de re-subir desde local.
"""

import sys
import io
import os
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
CARPETA_INCORRECTA = f"{CARPETA_OCS}/cmer-OC-00201096/4. DOCUMENTACION"
CARPETA_CORRECTA   = f"{CARPETA_OCS}/cmer-OC-00201096-EFM000C/4. DOCUMENTACION"

# Los 16 archivos subidos incorrectamente (subcarpeta → nombre)
ARCHIVOS = [
    ("4.05 BL-AWB-Porte definitivo",                  "12193 Set_01_BL_MEDUJX202719.pdf"),
    ("4.02 Factura Definitiva",                        "12193 Set_02_INVOICE_EXP2619193.pdf"),
    ("4.27 Packing list definitivo",                   "12193 Set_03_PACKING LIST_01052026.pdf"),
    ("OTROS",                                          "12193 Set_04_OTROS_01052026.pdf"),
    ("OTROS",                                          "12193 Set_05_QUALITY CERTIFICATE_01052026.pdf"),
    ("OTROS",                                          "12193 Set_06_FUMIGATION CERTIFICATE_01052026.pdf"),
    ("OTROS",                                          "12193 Set_07_QUALITY CERTIFICATE_01052026.pdf"),
    ("OTROS",                                          "12193 Set_08_QUALITY CERTIFICATE_260414010125.pdf"),
    ("OTROS",                                          "12193 Set_09_QUALITY CERTIFICATE_12193.pdf"),
    ("OTROS",                                          "12193 Set_10_QUALITY CERTIFICATE_12193.pdf"),
    ("OTROS",                                          "12193 Set_11_QUALITY CERTIFICATE_2605080101228.pdf"),
    ("OTROS",                                          "12193 Set_12_QUALITY CERTIFICATE_12193.pdf"),
    ("OTROS",                                          "12193 Set_13_QUALITY CERTIFICATE_12193.pdf"),
    ("OTROS",                                          "12193 Set_14_WEIGHT CERTIFICATE_5385422030AGRIGO.pdf"),
    ("4.12 Aprob Borr Cert fito origen",               "12193 Set_15_FITOSANITARIO_ECTRA6384016.pdf"),
    ("4.10 Certificado Origen definitivo (COO)",        "12193 Set_16_CO_1822674.pdf"),
]

# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------
_LOG_DIR = _BASE / "Registros"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"mover_po201096_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(str(_LOG_FILE), encoding="utf-8"),
    ],
)
log = logging.getLogger("mover_po201096")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_item_id(gc: GraphClient, ruta: str) -> str | None:
    """Obtiene el item-id de un archivo en SharePoint por su ruta."""
    ruta_encoded = quote(ruta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_encoded}:"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    if r.status_code == 200:
        return r.json()["id"]
    if r.status_code == 404:
        return None
    r.raise_for_status()


def get_folder_id(gc: GraphClient, ruta_carpeta: str) -> str | None:
    """Obtiene el item-id de una carpeta en SharePoint."""
    ruta_encoded = quote(ruta_carpeta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_encoded}:"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    if r.status_code == 200:
        return r.json()["id"]
    return None


def copiar_item(gc: GraphClient, item_id: str, dest_folder_id: str, nuevo_nombre: str) -> str | None:
    """
    Copia un item a otra carpeta usando Graph API /copy.
    Retorna la URL de monitoreo de la operación asíncrona.
    """
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/copy"
    payload = {
        "parentReference": {
            "driveId": SHAREPOINT_DRIVE_ID,
            "id": dest_folder_id,
        },
        "name": nuevo_nombre,
    }
    gc._renovar_token_si_necesario()
    r = gc.session.post(url, headers={**gc._headers, "Content-Type": "application/json"},
                        json=payload, timeout=60)
    if r.status_code == 202:
        return r.headers.get("Location")
    r.raise_for_status()


def esperar_copia(gc: GraphClient, monitor_url: str, timeout: int = 120) -> bool:
    """Espera a que la operación asíncrona de copia termine."""
    import time
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
    """Elimina un item de SharePoint por su item-id."""
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}"
    gc._renovar_token_si_necesario()
    r = gc.session.delete(url, headers=gc._headers, timeout=30)
    if r.status_code not in (200, 204):
        r.raise_for_status()


def main():
    dry_run = "--dry-run" in sys.argv
    log.info("=" * 70)
    log.info(f"MOVER PO 201096 — {'DRY-RUN' if dry_run else 'MODO REAL'}")
    log.info(f"  Origen:  {CARPETA_INCORRECTA}")
    log.info(f"  Destino: {CARPETA_CORRECTA}")
    log.info("=" * 70)

    gc = GraphClient()

    movidos = 0
    errores = 0

    for subcarpeta, nombre in ARCHIVOS:
        ruta_origen  = f"{CARPETA_INCORRECTA}/{subcarpeta}/{nombre}"
        ruta_destino_carpeta = f"{CARPETA_CORRECTA}/{subcarpeta}"

        log.info(f"\n  Archivo: {nombre}")
        log.info(f"  Origen:  .../{subcarpeta}/{nombre}")
        log.info(f"  Destino: ...EFM000C/4. DOCUMENTACION/{subcarpeta}/{nombre}")

        # 1. Obtener item-id del archivo en la carpeta incorrecta
        item_id = get_item_id(gc, ruta_origen)
        if item_id is None:
            log.warning(f"  [SKIP] No encontrado en carpeta incorrecta — puede ya haber sido movido.")
            continue

        if dry_run:
            log.info(f"  [DRY-RUN] Se copiaría y eliminaría (item_id={item_id})")
            continue

        # 2. Obtener o crear carpeta destino
        dest_folder_id = get_folder_id(gc, ruta_destino_carpeta)
        if dest_folder_id is None:
            log.info(f"  Creando carpeta destino: {ruta_destino_carpeta}")
            gc.crear_carpeta_si_no_existe(f"{ruta_destino_carpeta}/placeholder")
            dest_folder_id = get_folder_id(gc, ruta_destino_carpeta)
            if dest_folder_id is None:
                log.error(f"  [ERROR] No se pudo crear/obtener carpeta destino")
                errores += 1
                continue

        # 3. Copiar a carpeta correcta
        log.info(f"  Copiando...")
        monitor_url = copiar_item(gc, item_id, dest_folder_id, nombre)
        if monitor_url is None:
            log.error(f"  [ERROR] No se obtuvo URL de monitoreo para la copia")
            errores += 1
            continue

        ok = esperar_copia(gc, monitor_url)
        if not ok:
            log.error(f"  [ERROR] Fallo al copiar {nombre}")
            errores += 1
            continue

        log.info(f"  Copia completada.")

        # 4. Eliminar de carpeta incorrecta
        log.info(f"  Eliminando de carpeta incorrecta...")
        eliminar_item(gc, item_id)
        log.info(f"  [OK] Movido correctamente.")
        movidos += 1

    log.info("\n" + "=" * 70)
    log.info(f"Movidos: {movidos} | Errores: {errores}")
    log.info(f"Log: {_LOG_FILE}")
    log.info("=" * 70)


if __name__ == "__main__":
    main()
