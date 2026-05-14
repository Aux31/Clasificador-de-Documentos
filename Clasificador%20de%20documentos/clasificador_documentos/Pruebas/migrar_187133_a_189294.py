"""
Migra los archivos subidos por error a cmer-OC-00187133 que pertenecen
a la OC-00189294-LESUI000.

Detectados en registros_subidas.log líneas 1253, 1255-1257 (2026-05-09 13:29-13:31,
aux23.gg@grupointeca.com).

Uso:
    cd clasificador_documentos
    python Pruebas/migrar_187133_a_189294.py
    python Pruebas/migrar_187133_a_189294.py --dry-run
"""

import sys
import time
import argparse
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"

OC_ORIGEN  = "cmer-OC-00187133"
OC_DESTINO = "cmer-OC-00189294"  # buscar_carpeta_oc resolverá el sufijo -LESUI000

# (subcarpeta_origen, nombre_archivo, subcarpeta_destino)
MIGRACIONES = [
    ("4.03 Borr BL oAWB o Porte",    "177WJMJMJ008506X(BL draft).pdf",                    "4.03 Borr BL oAWB o Porte"),
    ("4.27 Packing list definitivo",  "PACKING LIST 177WJMJMJ008506X.pdf",                 "4.27 Packing list definitivo"),
    ("OTROS",                         "REQUEST\xa0FOR\xa0PAYMENT\xa0V.2.0\xa02025.12.09.docx", "OTROS"),
    ("OTROS",                         "WEIGHT\xa0CERTIFICATE.DOCX",                          "OTROS"),
]


def _item_id_por_ruta(gc: GraphClient, ruta: str) -> str | None:
    # Intenta la ruta tal cual; si 404, prueba con la extensión en minúsculas
    for ruta_intento in _variantes_ruta(ruta):
        ruta_enc = quote(ruta_intento, safe="/")
        url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_enc}:"
        try:
            data = gc._get(url)
            item_id = data.get("id")
            if item_id:
                if ruta_intento != ruta:
                    print(f"  [INFO] Encontrado con nombre alternativo: {ruta_intento.split('/')[-1]}")
                return item_id
        except Exception:
            continue
    print(f"  [ERROR] No encontrado en SharePoint (ninguna variante): {ruta}")
    return None


def _variantes_ruta(ruta: str) -> list[str]:
    """Devuelve variantes de la ruta con distinta capitalización de extensión."""
    variantes = [ruta]
    p = Path(ruta)
    ext = p.suffix
    if ext and ext != ext.lower():
        variantes.append(str(p.with_suffix(ext.lower())).replace("\\", "/"))
    if ext and ext != ext.upper():
        variantes.append(str(p.with_suffix(ext.upper())).replace("\\", "/"))
    return variantes


def _parent_id_por_ruta(gc: GraphClient, ruta_carpeta: str) -> str | None:
    ruta_enc = quote(ruta_carpeta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{ruta_enc}:"
    try:
        data = gc._get(url)
        return data.get("id")
    except Exception:
        print(f"  [INFO] Carpeta destino no existe, creándola: {ruta_carpeta}")
        gc.crear_carpeta_si_no_existe(ruta_carpeta + "/placeholder")
        time.sleep(1)
        try:
            data = gc._get(url)
            return data.get("id")
        except Exception as e:
            print(f"  [ERROR] No se pudo crear la carpeta destino: {e}")
            return None


def mover_archivo(gc: GraphClient, item_id: str, parent_id: str, nombre: str):
    gc._renovar_token_si_necesario()
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}"
    payload = {"parentReference": {"id": parent_id}, "name": nombre}
    r = gc.session.patch(url, headers=gc._headers, json=payload, timeout=30)
    if not r.ok:
        raise RuntimeError(f"PATCH {r.status_code}: {r.text[:300]}")


def main():
    parser = argparse.ArgumentParser(description="Migra archivos de OC-00187133 a OC-00189294-LESUI000")
    parser.add_argument("--dry-run", action="store_true", help="Simula sin hacer cambios")
    args = parser.parse_args()

    print("=" * 60)
    print("Migración OC-00187133 → OC-00189294-LESUI000")
    if args.dry_run:
        print("MODO DRY-RUN — no se realizarán cambios")
    print("=" * 60)

    gc = GraphClient()

    carpeta_destino_real = gc.buscar_carpeta_oc("cmer-OC-00189294")
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

        if args.dry_run:
            print(f"  [DRY-RUN] Se movería correctamente.")
            exitos += 1
            print()
            continue

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
        time.sleep(0.3)

    print("=" * 60)
    print(f"Resultado: {exitos} movidos, {errores} errores.")
    print("=" * 60)


if __name__ == "__main__":
    main()
