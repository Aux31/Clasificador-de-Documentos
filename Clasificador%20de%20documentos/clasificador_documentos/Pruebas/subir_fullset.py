# Corta, clasifica y sube a SharePoint un PDF Full Set.
# Uso: python pruebas/subir_fullset.py "C:/ruta/al/archivo.pdf" PO_NUMBER
#      python pruebas/subir_fullset.py "C:/Users/aux22.gg/Downloads/3310 FULL SET DOCUMENT.PDF" 200157

import sys
import os
import tempfile
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from clasificador  import procesar_adjunto, formatear_numero_oc
from graph_client  import GraphClient


def main():
    if len(sys.argv) < 3:
        print("Uso: python pruebas/subir_fullset.py \"ruta/al/fullset.pdf\" NUMERO_PO")
        print("Ej:  python pruebas/subir_fullset.py \"C:/Users/aux22.gg/Downloads/3310 FULL SET DOCUMENT.PDF\" 200157")
        sys.exit(1)

    ruta_pdf  = sys.argv[1]
    numero_po = sys.argv[2]

    if not Path(ruta_pdf).exists():
        print(f"Archivo no encontrado: {ruta_pdf}")
        sys.exit(1)

    nombre = Path(ruta_pdf).name
    ext    = Path(ruta_pdf).suffix

    print(f"\n{'='*60}")
    print(f"SUBIR FULL SET A SHAREPOINT")
    print(f"Archivo : {nombre}")
    print(f"PO      : {numero_po}")
    print(f"{'='*60}\n")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="adj_")
    tmp.close()
    shutil.copy2(ruta_pdf, tmp.name)

    cliente = GraphClient()

    try:
        oc_base = formatear_numero_oc(numero_po)
        oc_real = cliente.buscar_carpeta_oc(oc_base)

        print(f"  [OC] Carpeta: {oc_real}\n")

        destinos = procesar_adjunto(
            nombre_archivo=nombre,
            numero_po_asunto=numero_po,
            ruta_local=tmp.name,
        )

        if not destinos:
            print("[ERROR] No se generaron destinos — verificar PO y contenido del PDF.")
            return

        subidos = 0
        for info in destinos:
            ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)
            print(f"  Tipo : {info['tipo']} ({info['certeza']}%)")
            print(f"  -> {ruta_real}")
            ruta_archivo = info.get("ruta_local") or tmp.name
            cliente.crear_carpeta_si_no_existe(ruta_real)
            cliente.subir_archivo(ruta_archivo, ruta_real)
            print(f"  [OK] Subido\n")
            subidos += 1

        print(f"{'='*60}")
        print(f"Subidos: {subidos} archivo(s)")
        print(f"{'='*60}\n")

    finally:
        try:
            os.remove(tmp.name)
        except Exception:
            pass


if __name__ == "__main__":
    main()
