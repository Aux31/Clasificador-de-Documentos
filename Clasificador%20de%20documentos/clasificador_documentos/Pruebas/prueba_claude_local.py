"""
Prueba principal del clasificador Claude.
Dado un archivo local, muestra:
  - Tipo y certeza según Claude API
  - Comparación con fallback por keywords (nombre + contenido)
  - Ruta SharePoint simulada (si hay PO en el nombre o asunto)

Uso:
    python pruebas/prueba_claude_local.py "C:\\Temp\\BL PO 196893.pdf"
    python pruebas/prueba_claude_local.py "C:\\Temp\\factura.pdf" "asunto del correo PO 196893"
    python pruebas/prueba_claude_local.py "C:\\Temp\\doc.pdf" "PO 196893" --solo-claude
"""

import sys
import os

# Asegurar que el directorio padre esté en el path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pathlib import Path
from clasificador_claude import clasificar_con_claude, es_fallo
from clasificador import (
    _clasificar_fallback_nombre,
    extraer_numero_po,
    construir_ruta,
    formatear_numero_oc,
)


def _separador(titulo: str = "", ancho: int = 60):
    if titulo:
        print(f"\n{'─' * 3} {titulo} {'─' * max(0, ancho - len(titulo) - 5)}")
    else:
        print("─" * ancho)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = [a for a in sys.argv[1:] if a.startswith("--")]
    solo_claude = "--solo-claude" in flags

    if not args:
        print("Uso: python pruebas/prueba_claude_local.py <ruta_archivo> [asunto_correo] [--solo-claude]")
        sys.exit(1)

    ruta = args[0]
    asunto = args[1] if len(args) > 1 else ""

    if not Path(ruta).exists():
        print(f"[ERROR] Archivo no encontrado: {ruta}")
        sys.exit(1)

    nombre = Path(ruta).name

    print(f"\n{'═' * 60}")
    print(f"  PRUEBA CLASIFICADOR CLAUDE")
    print(f"{'═' * 60}")
    print(f"  Archivo : {nombre}")
    print(f"  Ruta    : {ruta}")
    print(f"  Asunto  : {asunto or '(no proporcionado)'}")

    # --- Claude API (envío directo del archivo) ---
    _separador("Claude API")
    tipo_claude, certeza, justificacion, inconsistencias, *_ = clasificar_con_claude(
        nombre_archivo=nombre,
        ruta_local=ruta,
        asunto_correo=asunto,
    )

    if es_fallo(tipo_claude):
        print(f"  Estado  : FALLO ({tipo_claude})")
        print(f"  ⚠  Claude no pudo clasificar — se usaría el fallback en producción")
    else:
        barra = "█" * (certeza // 10) + "░" * (10 - certeza // 10)
        print(f"  Tipo    : {tipo_claude}")
        print(f"  Certeza : {certeza}%  [{barra}]")
        nivel = "Alta" if certeza >= 80 else "Media" if certeza >= 50 else "Baja"
        print(f"  Nivel   : {nivel}")
        if justificacion:
            print(f"  Motivo  : {justificacion}")

    if inconsistencias:
        _separador("Inconsistencias detectadas")
        for idx, inc in enumerate(inconsistencias, 1):
            sev    = inc.get("severidad", "?").upper()
            campo  = inc.get("campo", "—")
            desc   = inc.get("descripcion", "—")
            print(f"  {idx}. [{sev}] {campo}: {desc}")
    else:
        print(f"  Inconsistencias: ninguna")

    if solo_claude:
        print()
        return

    # --- Fallback por nombre ---
    _separador("Fallback — keywords en nombre del archivo")
    tipos_nombre = _clasificar_fallback_nombre(nombre)
    print(f"  Resultado: {', '.join(tipos_nombre)}")

    # --- Comparación ---
    _separador("Comparación")
    if not es_fallo(tipo_claude):
        fb_nombre = tipos_nombre[0] if tipos_nombre else "OTROS"
        print(f"  Claude vs. fallback nombre: {'✓ coinciden' if tipo_claude == fb_nombre else f'✗ difieren ({tipo_claude} vs {fb_nombre})'}")

    # --- Ruta SharePoint simulada ---
    _separador("Ruta SharePoint simulada")
    po = extraer_numero_po(nombre) or (extraer_numero_po(asunto) if asunto else None)
    if po and not es_fallo(tipo_claude):
        ruta_sp = construir_ruta(po, tipo_claude, nombre)
        oc = formatear_numero_oc(po)
        print(f"  PO      : {po}")
        print(f"  OC      : {oc}")
        print(f"  Ruta    : {ruta_sp}")
    elif not po:
        print("  ⚠  No se encontró número de PO en el nombre del archivo ni en el asunto.")
        print("     Para ver la ruta completa, incluye el PO en el nombre o asunto.")
    else:
        print("  ⚠  No se puede calcular ruta porque Claude no clasificó correctamente.")

    print(f"\n{'═' * 60}\n")


if __name__ == "__main__":
    main()
