"""
Reprocesa correos pendientes de un rango de fechas dado.

A diferencia de 'python recopilador_documentos.py reprocesar' (que reprocesa
toda la bandeja), este script solo toca los correos cuyo EntryID NO esté en
procesados.txt dentro del período indicado.

Útil para recuperar correos que fallaron por error de API sin volver a subir
los que ya estaban correctamente clasificados.

Uso:
    python reprocesar_pendientes.py [--desde YYYY-MM-DD] [--hasta YYYY-MM-DD] [--max N]

    Sin argumentos → reprocesa correos de hoy que no estén en procesados.txt.

Ejemplos:
    python reprocesar_pendientes.py
    python reprocesar_pendientes.py --desde 2026-03-25
    python reprocesar_pendientes.py --desde 2026-03-24 --hasta 2026-03-26
    python reprocesar_pendientes.py --max 10
"""

import sys
import io
import os
import re
import shutil
import argparse
import tempfile
import warnings
import pythoncom
import win32com.client
import urllib3

from datetime import datetime, date, timedelta
from pathlib import Path

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# Bootstrap: asegura que el directorio padre esté en sys.path
# ---------------------------------------------------------------------------
_DIR_PROYECTO = Path(__file__).parent.parent
sys.path.insert(0, str(_DIR_PROYECTO))

from configuracion.ajustes import OUTLOOK_EMAIL, EXTENSIONES_BLOQUEADAS
from logger_errores import log_error, log_advertencia, log_evento
from agente_seguridad import ejecutar as seguridad_adjunto, verificar_servicio_av
from clasificador import procesar_adjunto, extraer_po_y_bl_de_asunto, formatear_numero_oc
from graph_client import GraphClient
from recopilador_documentos import _expandir_zip, _expandir_rar
from clasificador_claude import generar_respuesta_proveedor_consolidada
from generador_reporte import generar_reporte_word
from cola_aprobacion import encolar_sugerencia

# ---------------------------------------------------------------------------
# Reutiliza helpers de monitor_correos
# ---------------------------------------------------------------------------
_PROCESADOS_FILE = _DIR_PROYECTO / "procesados.txt"
_RE_IMAGEN_INLINE = re.compile(r"^image\d+\.", re.IGNORECASE)
_EXTS_IMAGEN      = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
_TAMANO_MAX_FIRMA = 100 * 1024  # 100 KB


def _cargar_procesados() -> set:
    if _PROCESADOS_FILE.exists():
        return set(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines())
    return set()


def _marcar_procesado(correo_id: str):
    with _PROCESADOS_FILE.open("a", encoding="utf-8") as f:
        f.write(correo_id + "\n")


def _es_imagen_inline(adj) -> bool:
    nombre = adj.FileName or ""
    ext    = Path(nombre).suffix.lower()
    if _RE_IMAGEN_INLINE.match(nombre):
        return True
    if ext in _EXTS_IMAGEN:
        try:
            if adj.Size <= _TAMANO_MAX_FIRMA:
                return True
        except Exception:
            pass
    try:
        content_id = adj.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
        )
        if content_id:
            return True
    except Exception:
        pass
    try:
        if adj.Type == 6:
            return True
    except Exception:
        pass
    return False


def _save_adjunto(adj, ruta: str) -> bool:
    try:
        datos = adj.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x37010102"
        )
        with open(ruta, "wb") as f:
            f.write(datos)
        return True
    except Exception:
        pass
    try:
        adj.SaveAsFile(ruta)
        return True
    except Exception as e:
        raise e


def _procesar_mensaje(msg, cliente: GraphClient, procesados: set) -> bool:
    """
    Procesa un MailItem. Retorna True si al menos un adjunto fue subido
    correctamente (lo que habilita marcar el correo en procesados.txt).
    """
    asunto    = msg.Subject or ""
    remitente = msg.SenderEmailAddress or ""
    msg_id    = msg.EntryID

    if msg.Attachments.Count == 0:
        return False

    ok_av, err_av = verificar_servicio_av()
    if not ok_av:
        print(f"[SEG] ABORTADO — Bitdefender no activo ({err_av})")
        return False

    po_asunto, bl_asunto = extraer_po_y_bl_de_asunto(asunto)
    archivos_subidos = []
    docs_con_inconsistencias = []

    for i in range(1, msg.Attachments.Count + 1):
        try:
            adj    = msg.Attachments.Item(i)
            nombre = adj.FileName
        except Exception as e:
            print(f"[ERROR] No se pudo acceder al adjunto {i}: {e}")
            continue

        ext = Path(nombre).suffix.lower()
        if ext in EXTENSIONES_BLOQUEADAS:
            continue
        if _es_imagen_inline(adj):
            continue

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="adj_")
        tmp.close()
        dir_extraccion = None

        try:
            _save_adjunto(adj, tmp.name)
        except Exception as e_save:
            print(f"[ERROR] No se pudo guardar adjunto {nombre}: {e_save}")
            try:
                os.remove(tmp.name)
            except Exception:
                pass
            continue

        try:
            seg = seguridad_adjunto(ruta_local=tmp.name, nombre_orig=nombre)
            if seg["resultado"] == "rechazado":
                continue

            # Descomprimir ZIP/RAR antes de clasificar
            if ext == ".zip":
                internos, dir_extraccion = _expandir_zip(tmp.name)
                if not internos:
                    print(f"[WARN] ZIP vacío o no se pudo descomprimir: {nombre}")
                    continue
                adjuntos_a_procesar = internos
            elif ext == ".rar":
                internos, dir_extraccion = _expandir_rar(tmp.name)
                if not internos:
                    print(f"[WARN] RAR vacío o no se pudo descomprimir: {nombre}")
                    continue
                adjuntos_a_procesar = internos
            else:
                adjuntos_a_procesar = [{"nombre": nombre, "ruta_local": tmp.name}]

            for adj_interno in adjuntos_a_procesar:
                destinos = procesar_adjunto(
                    adj_interno["nombre"],
                    numero_po_asunto=po_asunto,
                    numero_bl=bl_asunto,
                    ruta_local=adj_interno["ruta_local"],
                )
                if not destinos:
                    continue

                for info in destinos:
                    oc_base   = formatear_numero_oc(info["numero_po"])
                    oc_real   = cliente.buscar_carpeta_oc(oc_base)
                    ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)

                    cliente.crear_carpeta_si_no_existe(ruta_real)
                    ruta_archivo_subir = info.get("ruta_local") or adj_interno["ruta_local"]
                    cliente.subir_archivo(ruta_archivo_subir, ruta_real)
                    print(f"[OK] PO {info['numero_po']} | {info['tipo']} | {info['nombre_archivo']}")
                    archivos_subidos.append(info["tipo"])

                    if info.get("inconsistencias"):
                        docs_con_inconsistencias.append({
                            "nombre_archivo": info["nombre_archivo"],
                            "tipo":           info["tipo"],
                            "inconsistencias": info["inconsistencias"],
                        })

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

    if docs_con_inconsistencias:
        print(f"[INCONSISTENCIAS] {len(docs_con_inconsistencias)} documento(s) con problemas — {asunto}")
        sugerencia = generar_respuesta_proveedor_consolidada(
            remitente=remitente,
            asunto=asunto,
            documentos=docs_con_inconsistencias,
        )
        try:
            generar_reporte_word(remitente, asunto, docs_con_inconsistencias, sugerencia)
        except Exception as e:
            print(f"[WARN] No se pudo generar reporte Word: {e}")
        ruta_pendiente = encolar_sugerencia(remitente, asunto, sugerencia or "")
        print(f"[PENDIENTE] Sugerencia guardada: {ruta_pendiente.name}")

    return bool(archivos_subidos)


def main():
    parser = argparse.ArgumentParser(description="Reprocesa correos pendientes por rango de fechas")
    parser.add_argument("--desde", default=None, help="Fecha inicio YYYY-MM-DD (default: hoy)")
    parser.add_argument("--hasta", default=None, help="Fecha fin YYYY-MM-DD (default: hoy)")
    parser.add_argument("--max",   default=None, type=int, help="Máximo de correos a procesar")
    parser.add_argument("--po",    default=None, help="Filtrar por número(s) de PO en el asunto, separados por coma (ej: 191065,194172)")
    args = parser.parse_args()

    filtro_po = [po.strip() for po in args.po.split(",")] if args.po else []

    hoy       = date.today().strftime("%Y-%m-%d")
    # Si se filtra por PO sin fecha explícita, buscar en los últimos 30 días
    default_desde = (date.today() - timedelta(days=7)).strftime("%Y-%m-%d") if filtro_po and not args.desde else hoy
    desde_str = args.desde or default_desde
    hasta_str = args.hasta or hoy

    try:
        desde_dt = datetime.strptime(desde_str, "%Y-%m-%d")
        hasta_dt = datetime.strptime(hasta_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    except ValueError as e:
        print(f"[ERROR] Formato de fecha inválido: {e}")
        sys.exit(1)

    print("=" * 60)
    print(f"REPROCESAR PENDIENTES")
    print(f"  Rango : {desde_str} → {hasta_str}")
    if filtro_po:
        print(f"  Filtro PO: {', '.join(filtro_po)}")
    print(f"  Cuenta: {OUTLOOK_EMAIL}")
    print("=" * 60)

    procesados = _cargar_procesados()

    pythoncom.CoInitialize()
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        bandeja = None
        for store in namespace.Stores:
            try:
                if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                    bandeja = store.GetDefaultFolder(6)
                    print(f"[Outlook] Conectado a: {store.DisplayName}")
                    break
            except Exception:
                continue

        if bandeja is None:
            print(f"[ERROR] No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")
            sys.exit(1)

        cliente = GraphClient()

        # Recopilar correos en el rango que no estén en procesados.txt
        items_bandeja = bandeja.Items
        items_bandeja.Sort("[ReceivedTime]", True)  # más recientes primero

        pendientes = []
        for item in items_bandeja:
            try:
                if item.Class != 43:
                    continue
                recibido = item.ReceivedTime
                # ReceivedTime es pywintypes.datetime — convertir a datetime estándar
                # para evitar comparaciones cruzadas de tipos que fallan silenciosamente
                try:
                    recibido_dt = datetime(
                        recibido.year, recibido.month, recibido.day,
                        recibido.hour, recibido.minute, recibido.second
                    )
                except Exception:
                    continue

                if recibido_dt < desde_dt:
                    break  # bandeja ordenada desc; ya pasamos el rango
                if recibido_dt > hasta_dt:
                    continue

                if item.EntryID in procesados:
                    continue

                if filtro_po:
                    asunto_item = item.Subject or ""
                    if not any(po in asunto_item for po in filtro_po):
                        continue

                pendientes.append(item)
            except Exception:
                continue

        if not pendientes:
            print("\nNo hay correos pendientes en el rango indicado.")
            return

        total = len(pendientes)
        if args.max:
            pendientes = pendientes[: args.max]
            print(f"  {total} pendiente(s) encontrado(s) — procesando {len(pendientes)} (--max {args.max})\n")
        else:
            print(f"  {total} pendiente(s) encontrado(s)\n")

        completados = 0
        for i, msg in enumerate(pendientes, 1):
            try:
                asunto    = msg.Subject or "(sin asunto)"
                remitente = msg.SenderEmailAddress or "?"
                try:
                    rt = msg.ReceivedTime
                    recibido_str = f"{rt.year}-{rt.month:02d}-{rt.day:02d} {rt.hour:02d}:{rt.minute:02d}"
                except Exception:
                    recibido_str = "?"
                print(f"\n[{i}/{len(pendientes)}] {asunto}")
                print(f"  De: {remitente} | Recibido: {recibido_str}")

                exito = _procesar_mensaje(msg, cliente, procesados)
                if exito:
                    _marcar_procesado(msg.EntryID)
                    procesados.add(msg.EntryID)
                    completados += 1
                else:
                    pass

            except KeyboardInterrupt:
                print(f"\n[!] Cancelado — {completados} correo(s) completados.")
                break
            except Exception as e:
                print(f"  [ERROR] {type(e).__name__}: {e}")
                continue

        print(f"\nCompletado: {completados}/{len(pendientes)} correo(s) procesados.")

    except KeyboardInterrupt:
        print("\nDetenido por el usuario.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()