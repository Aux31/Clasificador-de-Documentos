# Prueba del separador de Full Set con el PDF de ejemplo.
# Uso: python pruebas/prueba_fullset.py
#      python pruebas/prueba_fullset.py "C:/ruta/al/archivo.pdf"

import sys
import os
from pathlib import Path

# Agregar el directorio padre al path para importar los módulos del proyecto
sys.path.insert(0, str(Path(__file__).parent.parent))

from separador_fullset import separar_fullset

# PDF de prueba por defecto
PDF_PRUEBA = r"C:/Users/aux22.gg/Downloads/3310 FULL SET DOCUMENT.PDF"


def main():
    ruta_pdf = sys.argv[1] if len(sys.argv) > 1 else PDF_PRUEBA

    if not Path(ruta_pdf).exists():
        print(f"Archivo no encontrado: {ruta_pdf}")
        sys.exit(1)

    nombre = Path(ruta_pdf).name
    print(f"\n{'='*60}")
    print(f"PRUEBA SEPARADOR FULL SET")
    print(f"Archivo: {nombre}")
    print(f"{'='*60}\n")

    # Carpeta de salida: misma carpeta que el PDF de prueba
    carpeta_salida = str(Path(ruta_pdf).parent)

    fragmentos = separar_fullset(
        ruta_pdf=ruta_pdf,
        nombre_archivo=nombre,
        carpeta_temp=carpeta_salida,
    )

    print(f"\n{'='*60}")
    if fragmentos is None:
        print("RESULTADO: Documento único — no es un Full Set")
    elif not fragmentos:
        print("RESULTADO: Full Set detectado pero la separación falló")
    else:
        print(f"RESULTADO: {len(fragmentos)} fragmento(s) generados")
        print()
        for i, f in enumerate(fragmentos, 1):
            print(f"  {i}. {f['nombre']}")
            print(f"     Tipo sugerido : {f['tipo_sugerido']}")
            print(f"     Referencia    : {f['referencia'] or '(sin referencia)'}")
            print(f"     Descripción   : {f['descripcion']}")
            print(f"     Ruta          : {f['ruta']}")
            print()
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
