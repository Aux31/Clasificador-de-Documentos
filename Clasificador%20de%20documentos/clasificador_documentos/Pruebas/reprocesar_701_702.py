"""
Reprocesa las subidas incorrectas 701 y 702, forzando PO 177176.

  701 - MEDUWN969219_Freight (1).pdf  →  fue a cmer-OC-00142836  (142836 = N° de factura, no PO)
  702 - Invoice scanner (doc04...)    →  fue a cmer-OC-4068220260 (falso positivo: "oc" en "doc")

Los archivos ya fueron borrados de SharePoint.
El script busca en Outlook los correos originales por nombre de adjunto,
fuerza PO = 177176 y sube al lugar correcto.
"""

import sys
import io
import os
import tempfile
import shutil
import pythoncom
import win32com.client

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from pathlib import Path

_BASE = Path(__file__).resolve().parent.parent
os.chdir(str(_BASE))          # asegura cwd = clasificador_documentos sin importar desde dónde se corra
sys.path.insert(0, str(_BASE))
sys.path.insert(0, str(_BASE.parent.parent / "Agente_Seguridad"))

from configuracion.ajustes          import OUTLOOK_EMAIL, EXTENSIONES_BLOQUEADAS
from Nucleo.clasificador            import procesar_adjunto, formatear_numero_oc
from Integraciones.graph_client     import GraphClient
from Integraciones.monitor_correos  import _es_imagen_inline, _save_con_timeout
from agente_seguridad               import ejecutar as seguridad_adjunto, verificar_servicio_av
from Nucleo.recopilador_documentos  import _expandir_zip, _expandir_rar

# ---------------------------------------------------------------------------
# Configuración fija para este reprocesamiento
# ---------------------------------------------------------------------------

PO_FORZADA = "177176"

# Adjuntos a buscar (nombre original en Outlook, sin importar mayúsculas)
ADJUNTOS_OBJETIVO = {
    "meduwn969219_freight (1).pdf",   # registro 701
    "doc04068220260423124813.pdf",    # registro 702
}


def _procesar_adjunto_forzado(nombre: str, ruta_tmp: str, cliente: GraphClient) -> bool:
    """
    Llama a procesar_adjunto forzando PO_FORZADA.
    Retorna True si se subió al menos un destino.
    """
    destinos = procesar_adjunto(
        nombre,
        numero_po_asunto=PO_FORZADA,   # <-- fuerza PO, ignora nombre del archivo y asunto
        ruta_local=ruta_tmp,
        asunto_correo=f"PO {PO_FORZADA}",
    )
    if not destinos:
        print(f"  [SKIP] procesar_adjunto no retornó destinos para: {nombre}")
        return False

    subido = False
    for info in destinos:
        oc_base   = formatear_numero_oc(info["numero_po"])
        oc_real   = cliente.buscar_carpeta_oc(oc_base)
        ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)

        cliente.crear_carpeta_si_no_existe(ruta_real)
        ruta_archivo = info.get("ruta_local") or ruta_tmp
        cliente.subir_archivo(ruta_archivo, ruta_real)
        print(f"  [OK] PO {info['numero_po']} | {info['tipo']} | {ruta_real}")
        subido = True

    return subido


def reprocesar():
    pythoncom.CoInitialize()
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        bandeja = None
        for store in namespace.Stores:
            try:
                if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                    bandeja = store.GetDefaultFolder(6)
                    print(f"[Outlook] Conectado: {store.DisplayName}")
                    break
            except Exception:
                continue

        if bandeja is None:
            print(f"[ERROR] No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")
            return

        ok_av, err_av = verificar_servicio_av()
        if not ok_av:
            print(f"[SEG] ABORTADO — Bitdefender no activo: {err_av}")
            return

        cliente = GraphClient()

        pendientes_lower = set(ADJUNTOS_OBJETIVO)
        encontrados: list[tuple] = []   # (msg, adj, nombre_original)

        print(f"\n[Búsqueda] Buscando adjuntos en bandeja de entrada...")
        for item in bandeja.Items:
            try:
                if item.Class != 43:
                    continue
                for i in range(1, item.Attachments.Count + 1):
                    try:
                        adj    = item.Attachments.Item(i)
                        nombre = adj.FileName or ""
                    except Exception:
                        continue
                    if nombre.lower() in pendientes_lower:
                        encontrados.append((item, adj, nombre))
                        print(f"  [Encontrado] {nombre}  —  Asunto: {item.Subject}")
            except Exception:
                continue

        if not encontrados:
            print("\n[WARN] No se encontraron los adjuntos en la bandeja.")
            print("       Verifica que los correos originales del 2026-04-23 sigan en Bandeja de Entrada.")
            return

        print(f"\n[Reprocesando] {len(encontrados)} adjunto(s) con PO forzada = {PO_FORZADA}\n")

        for msg, adj, nombre in encontrados:
            ext = Path(nombre).suffix.lower()
            if ext in EXTENSIONES_BLOQUEADAS:
                print(f"  [SKIP] Extensión bloqueada: {nombre}")
                continue
            if _es_imagen_inline(adj):
                print(f"  [SKIP] Imagen inline: {nombre}")
                continue

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="adj_")
            tmp.close()
            dir_extraccion = None

            print(f"--- {nombre} ---")
            try:
                _save_con_timeout(adj, tmp.name)

                seg = seguridad_adjunto(ruta_local=tmp.name, nombre_orig=nombre)
                if seg["resultado"] == "rechazado":
                    print(f"  [SEG] Rechazado por antivirus: {nombre}")
                    continue

                if ext == ".zip":
                    internos, dir_extraccion = _expandir_zip(tmp.name)
                    adjuntos_a_procesar = internos or []
                elif ext == ".rar":
                    internos, dir_extraccion = _expandir_rar(tmp.name)
                    adjuntos_a_procesar = internos or []
                else:
                    adjuntos_a_procesar = [{"nombre": nombre, "ruta_local": tmp.name}]

                for adj_interno in adjuntos_a_procesar:
                    _procesar_adjunto_forzado(
                        adj_interno["nombre"],
                        adj_interno["ruta_local"],
                        cliente,
                    )

            except Exception as e:
                print(f"  [ERROR] {type(e).__name__}: {e}")
            finally:
                try:
                    os.remove(tmp.name)
                except Exception:
                    pass
                if dir_extraccion and Path(dir_extraccion).exists():
                    try:
                        shutil.rmtree(dir_extraccion)
                    except Exception:
                        pass

        print("\n[Listo] Reprocesamiento completado.")

    except KeyboardInterrupt:
        print("\nInterrumpido.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    print("=" * 60)
    print("REPROCESAMIENTO 701-702  —  PO forzada: 177176")
    print("=" * 60)
    reprocesar()
