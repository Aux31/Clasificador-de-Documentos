"""
Prueba del pipeline completo filtrando solo correos de hoy.
No sube nada a SharePoint — solo clasifica y muestra resultados.

Uso:
    python pruebas/prueba_hoy.py          # correos de hoy, AV activo
    python pruebas/prueba_hoy.py skip     # omite AV check
"""

import sys
import os
import io
import shutil
import zipfile
import tempfile
import warnings
import urllib3
from datetime import datetime, date

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from configuracion.ajustes import EXTENSIONES_BLOQUEADAS
from clasificador import procesar_adjunto, extraer_po_y_bl_de_asunto
from graph_client import GraphClient


def _expandir_zip(ruta_zip: str) -> tuple[list[dict], str]:
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
    skip_av = "skip" in args
    hoy = date.today()

    print(f"\n{'═' * 70}")
    print(f"  PRUEBA PIPELINE — solo correos de hoy ({hoy.strftime('%d/%m/%Y')})")
    print(f"  AV check: {'OMITIDO' if skip_av else 'activo'}")
    print(f"{'═' * 70}")

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

    # Obtener solo los 20 más recientes (ordenados más reciente primero)
    cliente = GraphClient()
    cliente._max_scan = 20
    cliente._max_adj  = 20
    todos = cliente.obtener_correos_nuevos(procesados=set())

    correos_hoy = [c for c in todos if c.get("fecha_recibido") == hoy]

    if not correos_hoy:
        print(f"  Sin correos de hoy ({hoy.strftime('%d/%m/%Y')}) en la bandeja.")
        return

    print(f"  {len(correos_hoy)} correo(s) de hoy encontrados.\n")

    total_adjuntos  = 0
    total_claude    = 0
    total_fallback  = 0
    total_sin_po    = 0
    documentos_log  = []  # [(nombre_archivo, tipo, certeza, ruta_sharepoint)]

    for i, correo in enumerate(correos_hoy, 1):
        asunto    = correo["asunto"]
        remitente = correo["remitente"]
        po_asunto, bl_asunto = extraer_po_y_bl_de_asunto(asunto)

        print(f"[{i}/{len(correos_hoy)}] {asunto}")
        print(f"         De: {remitente}")
        print(f"         PO: {po_asunto or 'no detectado'} | BL: {bl_asunto or 'no detectado'}")

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

            for info in destinos:
                metodo  = info.get("metodo_clasificacion", "?")
                certeza = info.get("certeza", 0)
                tipo    = info["tipo"]
                icono   = "✓" if metodo == "claude" else "↩"
                if metodo == "claude":
                    total_claude += 1
                else:
                    total_fallback += 1

                certeza_str = f"{certeza}%" if certeza else "-"
                print(f"         {icono} {nombre}")
                print(f"           Tipo   : {tipo:25s} Certeza: {certeza_str}")
                print(f"           Método : {metodo}")
                print(f"           Ruta   : {info['ruta_sharepoint']}")
                documentos_log.append((nombre, tipo, certeza_str, info["ruta_sharepoint"]))

        for tmp in carpetas_tmp:
            shutil.rmtree(tmp, ignore_errors=True)

        print()

    print(f"{'═' * 70}")
    print(f"  RESUMEN")
    print(f"  Correos de hoy     : {len(correos_hoy)}")
    print(f"  Adjuntos totales   : {total_adjuntos}")
    print(f"  → Claude API       : {total_claude}")
    print(f"  → Fallback keywords: {total_fallback}")
    print(f"  → Sin PO (saltados): {total_sin_po}")

    if documentos_log:
        print(f"\n  DOCUMENTOS QUE SE SUBIRÍAN ({len(documentos_log)}):")
        print(f"  {'#':<4} {'Archivo':<40} {'Tipo':<22} {'Certeza':<8} Destino SharePoint")
        print(f"  {'-'*4} {'-'*40} {'-'*22} {'-'*8} {'-'*50}")
        for n, (nombre, tipo, certeza, ruta) in enumerate(documentos_log, 1):
            # Mostrar solo las últimas 3 partes de la ruta para que quepa
            partes = ruta.replace("\\", "/").split("/")
            ruta_corta = "/".join(partes[-3:]) if len(partes) >= 3 else ruta
            print(f"  {n:<4} {nombre[:40]:<40} {tipo[:22]:<22} {certeza:<8} {ruta_corta}")

    print(f"{'═' * 70}\n")


if __name__ == "__main__":
    main()
