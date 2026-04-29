"""
Prueba del separador de Full Set — corta el PDF y sube cada fragmento a SharePoint.

Uso:
    python pruebas/prueba_separador.py "C:/ruta/al/documento.pdf" OC-200159
    python pruebas/prueba_separador.py "C:/ruta/al/documento.pdf" 200159
"""

import sys
import os
import io
import shutil
import tempfile
from pathlib import Path

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from separador_fullset import separar_fullset
from clasificador import procesar_adjunto, formatear_numero_oc
from graph_client import GraphClient


def main():
    if len(sys.argv) < 3:
        print("Uso: python pruebas/prueba_separador.py <ruta_pdf> <numero_oc>")
        print("Ejemplo: python pruebas/prueba_separador.py \"C:/docs/fullset.pdf\" OC-200159")
        sys.exit(1)

    ruta_pdf = Path(sys.argv[1])
    numero_oc_raw = sys.argv[2]

    if not ruta_pdf.exists():
        print(f"[ERROR] No se encontró el archivo: {ruta_pdf}")
        sys.exit(1)

    if ruta_pdf.suffix.lower() != ".pdf":
        print(f"[ERROR] El archivo debe ser un PDF. Recibido: {ruta_pdf.suffix}")
        sys.exit(1)

    # Normalizar número de OC — extraer solo los dígitos
    numero_po = numero_oc_raw.replace("OC-", "").replace("OC", "").lstrip("0")
    if not numero_po.isdigit():
        print(f"[ERROR] Número de OC inválido: {numero_oc_raw}")
        sys.exit(1)

    print()
    print("=" * 60)
    print(f"  SEPARADOR DE FULL SET")
    print(f"  Archivo : {ruta_pdf.name}")
    print(f"  OC      : {formatear_numero_oc(numero_po)}")
    print("=" * 60)

    # Contar páginas
    try:
        import fitz
        doc = fitz.open(str(ruta_pdf))
        num_paginas = len(doc)
        doc.close()
        print(f"  Páginas totales: {num_paginas}")
    except Exception as e:
        print(f"  [WARN] No se pudo contar páginas: {e}")
        num_paginas = None

    print()

    # Carpeta temporal para los fragmentos
    carpeta_tmp = tempfile.mkdtemp(prefix="fullset_")

    try:
        # 1. Cortar el PDF
        fragmentos = separar_fullset(
            ruta_pdf=str(ruta_pdf),
            nombre_archivo=ruta_pdf.name,
            carpeta_temp=carpeta_tmp,
        )

        print()
        print("=" * 60)

        if fragmentos is None:
            print("  El PDF NO es un Full Set — documento único.")
            print("  No se generaron fragmentos.")
            return

        if fragmentos == []:
            print("  Claude detectó Full Set pero la separación FALLÓ.")
            print("  Revisar logs para más detalles.")
            sys.exit(2)

        print(f"  Full Set confirmado: {len(fragmentos)} fragmento(s)\n")

        # 2. Clasificar cada fragmento y subir a SharePoint
        cliente = GraphClient()

        for i, frag in enumerate(fragmentos, 1):
            print(f"  [{i}/{len(fragmentos)}] {frag['nombre']}")
            print(f"    Tipo sugerido : {frag['tipo_sugerido']}")
            print(f"    Referencia    : {frag.get('referencia') or '(sin referencia)'}")

            destinos = procesar_adjunto(
                nombre_archivo=frag["nombre"],
                numero_po_asunto=numero_po,
                ruta_local=frag["ruta"],
                _es_fragmento=True,
            )

            if not destinos:
                print(f"    [SKIP] No se pudo clasificar el fragmento")
                print()
                continue

            for info in destinos:
                oc_base = formatear_numero_oc(info["numero_po"])
                oc_real = cliente.buscar_carpeta_oc(oc_base)
                ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)
                ruta_real = cliente.resolver_ruta_versionada(ruta_real)

                cliente.crear_carpeta_si_no_existe(ruta_real)
                cliente.subir_archivo(frag["ruta"], ruta_real)

                print(f"    Tipo clasificado : {info['tipo']} ({info.get('certeza', 0)}%)")
                print(f"    Subido a         : {ruta_real}")

            print()

        print("=" * 60)
        print("  Proceso completado.")
        print("=" * 60)

    finally:
        shutil.rmtree(carpeta_tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
