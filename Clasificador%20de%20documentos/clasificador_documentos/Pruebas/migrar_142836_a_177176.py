"""
Migra los archivos subidos por error a cmer-OC-00142836 que en realidad
pertenecen a la PO 177176.

Detectados en registros_subidas.log líneas 910-913 (2026-05-02 14:04:xx,
lstone@coopelev.com). El proveedor envió adjuntos de dos OCs en un mismo
correo; el clasificador usó la PO del asunto (142836) para todos.

Uso:
    cd clasificador_documentos
    python Pruebas/migrar_142836_a_177176.py
"""

import sys
import time
from pathlib import Path
from urllib.parse import quote

# Ajustar sys.path para importar desde la raíz del proyecto
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Archivos a mover: (ruta_origen_relativa, ruta_destino_relativa)
# Las rutas son relativas a la raíz del drive de SharePoint.
# ---------------------------------------------------------------------------
OC_ORIGEN  = "cmer-OC-00142836"
OC_DESTINO = "cmer-OC-00177176"

# Archivos con "177176" en el nombre subidos a la carpeta 00142836
MIGRACIONES = [
    # (subcarpeta_origen,                              nombre_archivo,                                  subcarpeta_destino)
    ("OTROS",                                          "adj_a3tqjb6r_10_OTROS_177176.pdf",              "OTROS"),
    ("4.10 Certificado Origen definitivo (COO)",       "adj_a3tqjb6r_11_WEIGHT CERTIFICATE_177176.pdf", "OTROS"),
]


def _item_id_por_ruta(gc: GraphClient, ruta: str) -> str | None:
    """Devuelve el id del item de Graph API dado una ruta relativa al drive."""
    ruta_enc = quote(ruta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_enc}:"
    try:
        data = gc._get(url)
        return data.get("id")
    except Exception as e:
        print(f"  [ERROR] No se encontró el archivo en SharePoint: {ruta}\n         {e}")
        return None


def _parent_id_por_ruta(gc: GraphClient, ruta_carpeta: str) -> str | None:
    """Devuelve el id de la carpeta destino, creándola si no existe."""
    ruta_enc = quote(ruta_carpeta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_enc}:"
    try:
        data = gc._get(url)
        return data.get("id")
    except Exception:
        # La carpeta no existe — intentar crearla
        print(f"  [INFO] Carpeta destino no existe, creándola: {ruta_carpeta}")
        gc.crear_carpeta_si_no_existe(ruta_carpeta + "/placeholder")
        time.sleep(1)
        try:
            data = gc._get(url)
            return data.get("id")
        except Exception as e:
            print(f"  [ERROR] No se pudo crear/localizar la carpeta destino: {e}")
            return None


def mover_archivo(gc: GraphClient, item_id: str, parent_id_destino: str, nombre: str):
    """Mueve un item a otra carpeta vía PATCH (Graph API move)."""
    gc._renovar_token_si_necesario()
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}"
    payload = {
        "parentReference": {"id": parent_id_destino},
        "name": nombre,
    }
    r = gc.session.patch(url, headers=gc._headers, json=payload, timeout=30)
    if not r.ok:
        raise RuntimeError(f"PATCH {r.status_code}: {r.text[:300]}")


def main():
    print("=" * 60)
    print("Migración OC-00142836 → OC-00177176")
    print("=" * 60)

    gc = GraphClient()

    # Buscar el nombre real de la carpeta OC destino en SharePoint
    # (puede tener sufijo como cmer-OC-00177176-PROVEEDOR)
    carpeta_destino_real = gc.buscar_carpeta_oc("cmer-OC-00177176")
    print(f"Carpeta destino: {carpeta_destino_real}\n")

    exitos = 0
    errores = 0

    for subcarpeta_origen, nombre_archivo, subcarpeta_destino in MIGRACIONES:
        ruta_origen = (
            f"{SHAREPOINT_CARPETA_OCS}/{OC_ORIGEN}"
            f"/4. DOCUMENTACION/{subcarpeta_origen}/{nombre_archivo}"
        )
        ruta_carpeta_destino = (
            f"{SHAREPOINT_CARPETA_OCS}/{carpeta_destino_real}"
            f"/4. DOCUMENTACION/{subcarpeta_destino}"
        )

        print(f"Archivo : {nombre_archivo}")
        print(f"  Origen : .../{OC_ORIGEN}/4. DOCUMENTACION/{subcarpeta_origen}/")
        print(f"  Destino: .../{carpeta_destino_real}/4. DOCUMENTACION/{subcarpeta_destino}/")

        item_id = _item_id_por_ruta(gc, ruta_origen)
        if not item_id:
            errores += 1
            print()
            continue

        parent_id = _parent_id_por_ruta(gc, ruta_carpeta_destino)
        if not parent_id:
            errores += 1
            print()
            continue

        try:
            mover_archivo(gc, item_id, parent_id, nombre_archivo)
            print(f"  [OK] Movido correctamente.")
            exitos += 1
        except Exception as e:
            print(f"  [ERROR] {e}")
            errores += 1

        print()
        time.sleep(0.5)

    print("=" * 60)
    print(f"Resultado: {exitos} movidos, {errores} errores.")
    print("=" * 60)


if __name__ == "__main__":
    main()
