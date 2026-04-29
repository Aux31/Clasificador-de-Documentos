"""
Genera CATALOGO.txt a partir de los PDFs de referencia en la carpeta CATALOGO.

Uso:
    python herramientas/generar_catalogo_txt.py

El archivo resultante (CATALOGO/CATALOGO.txt) contiene el texto extraído de todos
los documentos de referencia, agrupado por tipo, listo para ser enviado a Claude API.

Formato del archivo:
    === BL ===
    texto ejemplo 1
    ---
    texto ejemplo 2

    === CO ===
    texto ejemplo 1
    ...
"""

import sys
from pathlib import Path

# Apuntar al raíz del proyecto para importar extractor_texto
sys.path.insert(0, str(Path(__file__).parent.parent))

from extractor_texto import extraer_texto

CATALOGO_RUTA  = Path(__file__).parent.parent.parent.parent / "CATALOGO"
SALIDA         = CATALOGO_RUTA / "CATALOGO.txt"
MAX_CHARS      = 1500  # caracteres máximos por documento
EXTENSIONES_OK = {".pdf", ".docx", ".doc", ".xlsx", ".xls"}


def main():
    if not CATALOGO_RUTA.exists():
        print(f"ERROR: No se encontró la carpeta {CATALOGO_RUTA}")
        sys.exit(1)

    carpetas = sorted(
        c for c in CATALOGO_RUTA.iterdir()
        if c.is_dir()
    )

    if not carpetas:
        print("No se encontraron subcarpetas en el catálogo.")
        sys.exit(1)

    bloques = []

    for carpeta in carpetas:
        tipo = carpeta.name.upper()
        archivos = sorted(
            f for f in carpeta.iterdir()
            if f.is_file() and f.suffix.lower() in EXTENSIONES_OK
        )

        if not archivos:
            print(f"  [{tipo}] Sin archivos — omitido")
            continue

        print(f"  [{tipo}] Procesando {len(archivos)} archivo(s)...")
        textos = []
        for archivo in archivos:
            print(f"    > {archivo.name}")
            texto = extraer_texto(str(archivo))
            if texto.strip():
                textos.append(texto[:MAX_CHARS].strip())
            else:
                print(f"      (sin texto extraído)")

        if not textos:
            print(f"  [{tipo}] Sin texto extraíble — omitido")
            continue

        bloque = f"=== {tipo} ===\n" + "\n---\n".join(textos)
        bloques.append(bloque)

    contenido = "\n\n".join(bloques)
    SALIDA.write_text(contenido, encoding="utf-8")
    print(f"\nArchivo generado: {SALIDA}")
    print(f"Tamaño: {len(contenido):,} caracteres")


if __name__ == "__main__":
    main()
