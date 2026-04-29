"""
Prueba del pipeline completo sin subir archivos a SharePoint.
Lee correos reales de Outlook, clasifica adjuntos con Claude API y muestra
el resultado con tipo, certeza y método de clasificación.

Uso:
    python pruebas/prueba_pipeline_completo.py          # últimos 6 correos
    python pruebas/prueba_pipeline_completo.py 10       # últimos 10 correos
    python pruebas/prueba_pipeline_completo.py 5 skip   # omite AV check
"""

import sys
import os
import io
import time
import tempfile
import shutil
import zipfile
import warnings
import urllib3

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

# Asegurar que el directorio padre esté en el path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from configuracion.ajustes import EXTENSIONES_BLOQUEADAS
from clasificador import procesar_adjunto, extraer_po_y_bl_de_asunto
from graph_client import GraphClient


def _expandir_zip(ruta_zip: str) -> tuple[list[dict], str]:
    """Descomprime ZIP en carpeta temporal."""
    carpeta_tmp = tempfile.mkdtemp(prefix="zip_adj_")
    adjuntos = []
    try:
        with zipfile.ZipFile(ruta_zip, "r") as zf:
            for entry in zf.infolist():
                if entry.is_dir():
                    continue
                nombre = os.path.basename(entry.filename)
                ext = os.path.splitext(nombre)[1].lower()
                if ext in EXTENSIONES_BLOQUEADAS or not nombre or nombre.startswith("._"):
                    continue
                ruta_extraida = zf.extract(entry, carpeta_tmp)
                adjuntos.append({"nombre": nombre, "ruta_local": ruta_extraida})
    except Exception as e:
        print(f"  [WARN] No se pudo descomprimir ZIP: {e}")
        shutil.rmtree(carpeta_tmp, ignore_errors=True)
        return [], ""
    return adjuntos, carpeta_tmp


def main():
    args = sys.argv[1:]
    numeros = [a for a in args if a.isdigit()]
    skip_av = "skip" in args
    max_correos = int(numeros[0]) if numeros else 6

    print(f"\n{'═' * 70}")
    print(f"  PRUEBA PIPELINE COMPLETO — sin subir a SharePoint")
    print(f"  Correos a revisar: {max_correos} | AV check: {'OMITIDO' if skip_av else 'activo'}")
    print(f"{'═' * 70}")

    # Verificar AV (opcional)
    if not skip_av:
        from agente_seguridad import verificar_servicio_av
        ok_av, err_av = verificar_servicio_av()
        if not ok_av:
            print(f"\n[SEG] ABORTADO — Bitdefender no está activo ({err_av}).")
            print("      Usa 'skip' como argumento para omitir el check AV en pruebas.")
            sys.exit(1)
        print(f"  AV: OK\n")
    else:
        print(f"  AV: omitido (modo prueba)\n")

    cliente = GraphClient()
    correos = cliente.obtener_correos_nuevos(procesados=set())[:max_correos]

    if not correos:
        print("  Sin correos disponibles.")
        return

    total_adjuntos   = 0
    total_claude     = 0
    total_fallback   = 0
    total_sin_po     = 0

    for i, correo in enumerate(correos, 1):
        asunto    = correo["asunto"]
        remitente = correo["remitente"]
        po_asunto, bl_asunto = extraer_po_y_bl_de_asunto(asunto)

        print(f"\n[{i}/{len(correos)}] {asunto}")
        print(f"         De: {remitente}")
        print(f"         PO asunto: {po_asunto or 'no detectado'} | BL asunto: {bl_asunto or 'no detectado'}")

        # Expandir ZIPs
        adjuntos_expandidos = []
        carpetas_tmp = []
        for adj in correo.get("adjuntos", []):
            if adj["nombre"].lower().endswith(".zip"):
                internos, tmp = _expandir_zip(adj["ruta_local"])
                if internos:
                    adjuntos_expandidos.extend(internos)
                    carpetas_tmp.append(tmp)
                    print(f"         ZIP: {len(internos)} archivo(s) extraídos de {adj['nombre']}")
            else:
                adjuntos_expandidos.append(adj)

        for adj in adjuntos_expandidos:
            nombre = adj["nombre"]
            total_adjuntos += 1

            # Seguridad básica (solo extensión, sin AV completo en prueba)
            if not skip_av:
                from agente_seguridad import ejecutar as seg_ejecutar
                seg = seg_ejecutar(
                    ruta_local=adj.get("ruta_local", nombre),
                    nombre_orig=nombre,
                )
                if seg["resultado"] == "rechazado":
                    print(f"         ✗ {nombre} — rechazado por seguridad ({seg.get('motivo','?')})")
                    continue

            destinos = procesar_adjunto(
                nombre,
                numero_po_asunto=po_asunto,
                numero_bl=bl_asunto,
                ruta_local=adj.get("ruta_local"),
                asunto_correo=asunto,
            )

            if not destinos:
                total_sin_po += 1
                print(f"         ✗ {nombre} — sin PO")
                continue

            metodo_adjunto = destinos[0].get("metodo_clasificacion", "?") if destinos else "?"
            if metodo_adjunto == "claude":
                total_claude += 1
            else:
                total_fallback += 1

            for info in destinos:
                metodo  = info.get("metodo_clasificacion", "?")
                certeza = info.get("certeza", 0)
                tipo    = info["tipo"]
                icono   = "✓" if metodo == "claude" else "↩"

                certeza_str = f"{certeza}%" if certeza else "-"
                print(f"         {icono} {nombre}")
                print(f"           Tipo: {tipo:25s} Método: {metodo:20s} Certeza: {certeza_str}")
                print(f"           Ruta: {info['ruta_sharepoint']}")

        # Limpiar temporales
        for tmp in carpetas_tmp:
            shutil.rmtree(tmp, ignore_errors=True)

    # --- Resumen ---
    print(f"\n{'═' * 70}")
    print(f"  RESUMEN")
    print(f"  Correos procesados : {len(correos)}")
    print(f"  Adjuntos totales   : {total_adjuntos}")
    print(f"  → Claude API       : {total_claude}")
    print(f"  → Fallback keywords: {total_fallback}")
    print(f"  → Sin PO (saltados): {total_sin_po}")
    print(f"{'═' * 70}\n")


if __name__ == "__main__":
    main()
