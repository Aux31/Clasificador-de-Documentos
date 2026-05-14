"""
monitor_sin_claude.py
---------------------
Versión del monitor optimizada para correr en VM el fin de semana.

Diferencias respecto a monitor_correos.py:
  - Clasifica primero por keywords del nombre del archivo (sin costo)
  - Solo llama a Claude API si la certeza por keywords es <= 90%
    (nombre ambiguo como "DOCUMENTOS DESPACHO.pdf")
  - Si el nombre es claro (ej. "BL PO 196893.pdf" → 95%) no gasta tokens
  - Las inconsistencias se loguean en registros/inconsistencias_YYYY-MM-DD.log
    (texto plano, no genera Word .docx)
  - MODO_SIN_AV=true en .env → no necesita Bitdefender corriendo
  - Se queda en bucle esperando correos nuevos (Ctrl+C para detener)

Uso:
    python monitor_sin_claude.py

Detener con Ctrl+C.
"""

import sys
import io
import os
import re
import shutil
import time
import logging
import warnings
import tempfile
import pythoncom
import win32com.client
import urllib3
from datetime import datetime
from pathlib import Path

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

from configuracion.ajustes import OUTLOOK_EMAIL, MODO_MOCK, EXTENSIONES_BLOQUEADAS, MODO_SIN_AV, SEGURIDAD_SERVICIO_AV, SEGURIDAD_ESPERA_AV_SEGUNDOS, SEGURIDAD_TAMANO_MAX_MB, SEGURIDAD_EXTENSIONES_PERMITIDAS

from Utilidades.logger_errores import log_error, log_advertencia, log_evento
from Utilidades.verificador_cadenas import verificar_cadena
from Nucleo.clasificador import (
    extraer_po_y_bl_de_asunto,
    _clasificar_fallback_nombre,
    _clasificar_fallback_contenido,
    extraer_numero_po,
    construir_ruta,
    renombrar_para_tipo,
    formatear_numero_oc,
    MAPA_CARPETAS,
)
from Nucleo.clasificador_claude import clasificar_con_claude, es_fallo
from Nucleo.extractor_texto import extraer_texto
from Integraciones.graph_client import GraphClient
from agente_seguridad import ejecutar as seguridad_adjunto, verificar_servicio_av
from Integraciones.notificador import notificar_sugerencia, notificar_error
from configuracion.ajustes import CLAUDE_CERTEZA_MINIMA, CLAUDE_MAX_CHARS_TEXTO

# Umbrales de certeza
_CERTEZA_NOMBRE_DIRECTO = 95   # tipo aparece literalmente en el nombre del archivo
_CERTEZA_NOMBRE_ALIAS   = 80   # tipo detectado por alias en el nombre
_CERTEZA_CONTENIDO      = 90   # tipo detectado por keywords en el contenido del documento
_UMBRAL_SIN_CLAUDE      = 90   # certeza >= este valor → no se llama a Claude


def _clasificar_con_certeza(nombre_archivo: str, ruta_local: str = "") -> tuple[list[str], int, str]:
    """
    Clasifica por keywords en nombre y contenido. Devuelve (tipos, certeza, metodo).

    Flujo:
      1. Keywords en nombre directo (tipo literalmente en el nombre) → 95%
      2. Keywords en nombre por alias (BILL OF LADING → BL, etc.)   → 80%
      3. Keywords en contenido del documento (extrae texto)          → 90%
      4. OTROS                                                       →  0%
    """
    import re as _re
    nombre_upper = nombre_archivo.upper()

    # 1. Nombre directo
    tipos_directos = [t for t in MAPA_CARPETAS if t != "CO" and t.upper() in nombre_upper]
    if "CO" not in tipos_directos and _re.search(r'(?<![A-Z])CO(?![A-Z])', nombre_upper):
        tipos_directos.append("CO")
    if tipos_directos:
        return tipos_directos, _CERTEZA_NOMBRE_DIRECTO, "keywords_nombre"

    # 2. Nombre por alias
    tipos_alias = _clasificar_fallback_nombre(nombre_archivo)
    if tipos_alias != ["OTROS"]:
        return tipos_alias, _CERTEZA_NOMBRE_ALIAS, "keywords_nombre_alias"

    # 3. Contenido del documento
    if ruta_local:
        try:
            texto = extraer_texto(ruta_local)
            if texto.strip():
                tipo_contenido = _clasificar_fallback_contenido(texto[:CLAUDE_MAX_CHARS_TEXTO].upper())
                if tipo_contenido != "OTROS":
                    return [tipo_contenido], _CERTEZA_CONTENIDO, "keywords_contenido"
        except Exception as e:
            print(f"  [WARN] Error extrayendo texto para keywords: {e}")

    return ["OTROS"], 0, "sin_clasificacion"

# ---------------------------------------------------------------------------
# Loggers
# ---------------------------------------------------------------------------
_LOG_SUBIDAS       = Path(__file__).parent.parent / "registros_subidas.log"
_CONTADOR_CORREOS  = Path(__file__).parent.parent / "contador_correos.txt"
_DIR_REGISTROS  = Path(__file__).parent / "registros"
_DIR_REGISTROS.mkdir(exist_ok=True)

_logger_subidas = logging.getLogger("subidas_sharepoint")
if not _logger_subidas.handlers:
    _h = logging.FileHandler(_LOG_SUBIDAS, encoding="utf-8")
    _h.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_h)
    _logger_subidas.setLevel(logging.INFO)


def _siguiente_consecutivo() -> int:
    if not _LOG_SUBIDAS.exists():
        return 1
    lines = _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines()
    return len([l for l in lines if l.strip()]) + 1


def _siguiente_id_correo() -> int:
    """Devuelve el próximo ID único de correo e incrementa el contador."""
    if not _CONTADOR_CORREOS.exists():
        _CONTADOR_CORREOS.write_text("0", encoding="utf-8")
    n = int(_CONTADOR_CORREOS.read_text(encoding="utf-8").strip() or "0") + 1
    _CONTADOR_CORREOS.write_text(str(n), encoding="utf-8")
    return n


def _log_subida(remitente: str, ruta_sharepoint: str, id_correo: int = 0, asunto: str = ""):
    ruta_completa = ruta_sharepoint.replace("\\", "/")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n  = _siguiente_consecutivo()
    asunto_limpio = (asunto or "").replace("\n", " ").replace("\r", "").strip()
    _logger_subidas.info(f"{n} | GINT{id_correo:05d}Z | {ts} | {remitente} | {asunto_limpio} | {ruta_completa}")


def _log_inconsistencias_txt(remitente: str, asunto: str, docs_con_inconsistencias: list[dict]):
    """Escribe las inconsistencias en un log de texto plano (sin Word, sin Claude)."""
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    ruta_log  = _DIR_REGISTROS / f"inconsistencias_{fecha_hoy}.log"
    ts        = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lineas = [
        "",
        "=" * 70,
        f"Timestamp : {ts}",
        f"Remitente : {remitente}",
        f"Asunto    : {asunto}",
        f"Docs      : {len(docs_con_inconsistencias)}",
        "-" * 70,
    ]
    for doc_data in docs_con_inconsistencias:
        lineas.append(f"  Archivo : {doc_data['nombre_archivo']}")
        lineas.append(f"  Tipo    : {doc_data['tipo']}")
        for inc in doc_data.get("inconsistencias", []):
            sev   = inc.get("severidad", "?").upper()
            campo = inc.get("campo", "")
            desc  = inc.get("descripcion", "")
            lineas.append(f"    [{sev}] {campo}: {desc}")
        lineas.append("")
    lineas.append("=" * 70)
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write("\n".join(lineas) + "\n")
    log_evento(f"Inconsistencias registradas en: {ruta_log}")

# ---------------------------------------------------------------------------
# Helpers COM
# ---------------------------------------------------------------------------

_RE_IMAGEN_INLINE = re.compile(r"^image\d+\.", re.IGNORECASE)
_EXTS_IMAGEN      = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
_TAMANO_MAX_FIRMA = 100 * 1024  # 100 KB


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


def _save_con_timeout(adj, ruta: str) -> bool:
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

# ---------------------------------------------------------------------------
# Registro de correos ya procesados
# ---------------------------------------------------------------------------
_PROCESADOS_FILE = Path(__file__).parent / "procesados.txt"


def _cargar_procesados() -> set:
    if _PROCESADOS_FILE.exists():
        return set(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines())
    return set()


def _marcar_procesado(entry_id: str):
    with open(_PROCESADOS_FILE, "a", encoding="utf-8") as f:
        f.write(entry_id + "\n")

# ---------------------------------------------------------------------------
# Procesador de correo — sin Claude, clasificación solo por keywords
# ---------------------------------------------------------------------------

def _procesar_mensaje(msg, cliente: GraphClient):
    """Procesa un único objeto MailItem de Outlook. Sin Claude API."""
    try:
        asunto    = msg.Subject or ""
        remitente = msg.SenderEmailAddress or ""
        msg_id    = msg.EntryID
        id_correo = _siguiente_id_correo()

        if msg.Attachments.Count == 0:
            print(f"\n[SKIP] Sin adjuntos: {asunto}")
            log_evento("BASURA", "monitor_sin_claude", remitente=remitente, asunto=asunto,
                       detalle="correo sin adjuntos")
            return

        print(f"\n{'=' * 60}")
        print(f"[NUEVO] {asunto}")
        print(f"  De: {remitente}")

        # Verificar AV (se omite si MODO_SIN_AV=true)
        if MODO_SIN_AV:
            ok_av, err_av = True, None
        else:
            ok_av, err_av = verificar_servicio_av(SEGURIDAD_SERVICIO_AV)
        if not ok_av:
            print(f"[SEG] ABORTADO — Bitdefender no activo ({err_av})")
            return

        po_asunto, bl_asunto = extraer_po_y_bl_de_asunto(asunto)
        if po_asunto:
            print(f"  Asunto -> PO: {po_asunto} | BL: {bl_asunto or 'no detectado'}")

        archivos_subidos         = []
        ultimo_info              = None
        docs_con_inconsistencias = []

        # Verificar cadena de correos
        incs_cadena = verificar_cadena(msg, asunto, po_asunto, bl_asunto)
        if incs_cadena:
            docs_con_inconsistencias.append({
                "nombre_archivo": "(cadena de correos)",
                "tipo":           "CADENA",
                "inconsistencias": incs_cadena,
            })
            log_evento("CADENA", "monitor_sin_claude", remitente=remitente, asunto=asunto,
                       detalle=f"{len(incs_cadena)} inconsistencia(s) en cadena de correos")

        n_adjuntos_reales = 0
        for i in range(1, msg.Attachments.Count + 1):
            try:
                adj    = msg.Attachments.Item(i)
                nombre = adj.FileName
            except Exception as e:
                print(f"  [SKIP] No se pudo acceder al adjunto {i}: {e}")
                continue
            ext = Path(nombre).suffix.lower()

            if ext in EXTENSIONES_BLOQUEADAS:
                print(f"  [SKIP] Extensión bloqueada: {nombre}")
                continue

            if _es_imagen_inline(adj):
                print(f"  [SKIP] Imagen inline ignorada: {nombre}")
                continue

            print(f"  Procesando: {nombre}")
            n_adjuntos_reales += 1

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="adj_")
            tmp.close()
            try:
                _save_con_timeout(adj, tmp.name)
            except Exception as e_save:
                print(f"  [ERROR] No se pudo guardar adjunto: {nombre} — {e_save}")
                log_error("monitor_sin_claude", "MON-001", nombre,
                          f"No se pudo guardar adjunto: {type(e_save).__name__}: {e_save}",
                          remitente=remitente, asunto=asunto)
                try:
                    os.remove(tmp.name)
                except Exception:
                    pass
                continue

            dir_extraccion = None
            try:
                seg = seguridad_adjunto(
                    ruta_local      = tmp.name,
                    nombre_orig     = nombre,
                    extensiones_ok  = SEGURIDAD_EXTENSIONES_PERMITIDAS,
                    tamano_max_mb   = SEGURIDAD_TAMANO_MAX_MB,
                    espera_av       = SEGURIDAD_ESPERA_AV_SEGUNDOS,
                    nombre_proyecto = "sharepoint",
                )
                if seg["resultado"] == "rechazado":
                    continue

                # --- Clasificación: keywords primero, Claude solo si certeza <= 90 ---
                numero_po = po_asunto or extraer_numero_po(nombre)
                if not numero_po:
                    print(f"  [SKIP] No se encontró número de PO en: {nombre}")
                    log_advertencia("monitor_sin_claude", "CLAS-001", nombre,
                                    "No se encontró número de PO", remitente=remitente, asunto=asunto)
                    continue

                tipos_kw, certeza_kw, metodo_kw = _clasificar_con_certeza(nombre, tmp.name)
                inconsistencias_doc = []

                if certeza_kw > _UMBRAL_SIN_CLAUDE:
                    # Keywords (nombre o contenido) lo resolvió con suficiente certeza
                    tipos   = tipos_kw
                    certeza = certeza_kw
                    metodo  = metodo_kw
                    print(f"  [{metodo_kw.upper()}] {tipos} ({certeza_kw}%) — sin Claude")
                else:
                    # Ambiguo incluso tras leer el contenido — llamar a Claude
                    print(f"  [KEYWORDS] Certeza baja ({certeza_kw}%) — consultando Claude...")
                    tipo_claude, certeza_claude, justif, incs, *_ = clasificar_con_claude(
                        nombre_archivo=nombre,
                        ruta_local=tmp.name,
                        asunto_correo=asunto,
                    )
                    if not es_fallo(tipo_claude) and certeza_claude >= CLAUDE_CERTEZA_MINIMA:
                        tipos   = [tipo_claude] if tipo_claude != "PL + INV" else ["PACKING LIST", "INVOICE"]
                        certeza = certeza_claude
                        metodo  = "claude"
                        inconsistencias_doc = incs or []
                        print(f"  [CLAUDE] {tipo_claude} ({certeza_claude}%) — {justif or 'sin justificacion'}")
                    elif certeza_kw > 0:
                        # Claude falló pero keywords tenía algo
                        tipos   = tipos_kw
                        certeza = certeza_kw
                        metodo  = f"{metodo_kw}_fallback"
                        print(f"  [FALLBACK] Claude falló — usando keywords: {tipos}")
                        log_advertencia("monitor_sin_claude", "CLAS-002", nombre,
                                        f"Claude falló ({tipo_claude}), usando keywords", remitente=remitente, asunto=asunto)
                    else:
                        tipos   = ["OTROS"]
                        certeza = 0
                        metodo  = "sin_clasificacion"
                        print(f"  [SKIP] No se pudo clasificar: {nombre}")
                        log_advertencia("monitor_sin_claude", "CLAS-003", nombre,
                                        "Keywords y Claude sin resultado", remitente=remitente, asunto=asunto)

                destinos = [
                    {
                        "nombre_archivo":       nombre,
                        "nombre_destino":       renombrar_para_tipo(nombre, tipo, tipos),
                        "numero_po":            numero_po,
                        "numero_bl":            bl_asunto,
                        "tipo":                 tipo,
                        "ruta_sharepoint":      construir_ruta(
                                                    numero_po, tipo,
                                                    renombrar_para_tipo(nombre, tipo, tipos),
                                                ),
                        "certeza":              certeza,
                        "metodo_clasificacion": metodo,
                        "inconsistencias":      inconsistencias_doc,
                        "ruta_local":           tmp.name,
                    }
                    for tipo in tipos
                ]

                for info in destinos:
                    oc_base   = formatear_numero_oc(info["numero_po"])
                    oc_real   = cliente.buscar_carpeta_oc(oc_base)
                    ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)

                    print(f"    PO: {info['numero_po']} | BL: {info['numero_bl'] or '-'} | Tipo: {info['tipo']} ({info['certeza']}% / {info['metodo_clasificacion']})")
                    print(f"    -> {ruta_real}")
                    cliente.crear_carpeta_si_no_existe(ruta_real)
                    cliente.subir_archivo(tmp.name, ruta_real)
                    _log_subida(remitente, ruta_real, id_correo=id_correo, asunto=asunto)
                    archivos_subidos.append(info["tipo"])
                    ultimo_info = {**info, "ruta_sharepoint": ruta_real}

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

        if n_adjuntos_reales == 0 and msg.Attachments.Count > 0:
            log_evento("BASURA", "monitor_sin_claude", remitente=remitente, asunto=asunto,
                       detalle=f"todos los adjuntos ({msg.Attachments.Count}) son imágenes inline o extensión bloqueada")

        if archivos_subidos and ultimo_info:
            print(f"  [OK] {len(archivos_subidos)} archivo(s) subido(s) para PO {ultimo_info['numero_po']}")

        # Inconsistencias → solo log de texto, sin Word, sin Claude
        if docs_con_inconsistencias:
            print(f"\n  [INCONSISTENCIAS] {len(docs_con_inconsistencias)} problema(s) detectado(s) — registrado en log")
            _log_inconsistencias_txt(remitente, asunto, docs_con_inconsistencias)
            notificar_sugerencia(remitente, asunto, "(sin sugerencia — modo sin Claude)")

        # Insertar ID único en el cuerpo del correo para búsqueda en Outlook
        if archivos_subidos:
            try:
                etiqueta_id = f"GINT{id_correo:05d}Z"
                try:
                    html_actual = msg.HTMLBody or ""
                    if etiqueta_id not in html_actual:
                        bloque_html = f'<div style="color:#ffffff;font-size:1px;mso-hide:all">{etiqueta_id}</div>'
                        if "</body>" in html_actual.lower():
                            _idx = html_actual.lower().rfind("</body>")
                            html_actual = html_actual[:_idx] + bloque_html + html_actual[_idx:]
                        else:
                            html_actual = html_actual + bloque_html
                        msg.HTMLBody = html_actual
                except Exception:
                    pass
                try:
                    body_actual = msg.Body or ""
                    if etiqueta_id not in body_actual:
                        msg.Body = body_actual + f"\n\n{etiqueta_id}"
                except Exception:
                    pass
                msg.Save()
            except Exception as e:
                log_advertencia("monitor_sin_claude", "MON-007", detalle=f"No se pudo etiquetar cuerpo: {e}")

        # Marcar como procesado
        _marcar_procesado(msg_id)

    except Exception as e:
        print(f"  [ERROR] Error procesando mensaje: {e}")
        log_error("monitor_sin_claude", "MON-002", "-",
                  f"Error procesando mensaje: {type(e).__name__}: {e}",
                  remitente=remitente, asunto=asunto)
        notificar_error(f"Error procesando mensaje de {remitente}",
                        detalle=f"{type(e).__name__}: {e}",
                        paso="monitor_sin_claude._procesar_mensaje",
                        como_resolver="Revisar el log de errores y verificar el adjunto manualmente.")


# ---------------------------------------------------------------------------
# COM Event sink
# ---------------------------------------------------------------------------

class _BandejaEventos:
    _cliente_compartido: GraphClient = None

    def OnItemAdd(self, item):
        try:
            if item.Class != 43:
                return
            _procesar_mensaje(item, self._cliente_compartido)
        except Exception as e:
            print(f"[ERROR] OnItemAdd: {e}")
            log_error("monitor_sin_claude", "MON-003", "-",
                      f"OnItemAdd: {type(e).__name__}: {e}")
            notificar_error("Error en listener de correos (OnItemAdd)",
                            detalle=f"{type(e).__name__}: {e}",
                            paso="monitor_sin_claude.OnItemAdd",
                            como_resolver="Reiniciar el monitor.")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main():
    modo = "MOCK" if MODO_MOCK else "REAL"
    print(f"Monitor iniciado [{modo}]")
    print(f"Cuenta          : {OUTLOOK_EMAIL}")
    print(f"Clasificacion   : keywords primero, Claude solo si certeza <= {_UMBRAL_SIN_CLAUDE}%")
    print(f"Reportes        : log de texto (sin Word)")
    print("Esperando correos nuevos... (Ctrl+C para detener)\n")

    pythoncom.CoInitialize()
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        bandeja = None
        for store in namespace.Stores:
            try:
                if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                    bandeja = store.GetDefaultFolder(6)
                    print(f"[Outlook] Suscrito a: {store.DisplayName}")
                    break
            except Exception:
                continue

        if bandeja is None:
            print(f"[ERROR] No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")
            log_error("monitor_sin_claude", "MON-004", "-",
                      f"Cuenta '{OUTLOOK_EMAIL}' no encontrada en Outlook")
            notificar_error(f"Cuenta '{OUTLOOK_EMAIL}' no encontrada",
                            paso="monitor_sin_claude.main",
                            como_resolver="Abrir Outlook e iniciar sesión.")
            sys.exit(1)

        cliente = GraphClient()

        # Escaneo inicial: procesa correos no vistos desde la última ejecución
        procesados    = _cargar_procesados()
        items_bandeja = bandeja.Items
        items_bandeja.Sort("[ReceivedTime]", True)
        print(f"[Monitor] Escaneo inicial ({len(procesados)} ya procesados)...")
        n_inicial = 0
        iter_bandeja = iter(items_bandeja)
        while True:
            try:
                item = next(iter_bandeja)
            except StopIteration:
                print(f"[Monitor] Escaneo inicial completo ({n_inicial} nuevos).")
                break
            except Exception as e:
                print(f"[WARN] Error al avanzar en bandeja: {e}")
                break
            try:
                if item.Class != 43:
                    continue
                if item.EntryID in procesados:
                    print(f"[Monitor] Encontrado correo ya procesado — escaneo completo ({n_inicial} nuevos).")
                    break
                _procesar_mensaje(item, cliente)
                n_inicial += 1
            except Exception as e:
                print(f"[WARN] Error procesando correo en escaneo inicial: {e}")
                continue

        # Suscribirse a eventos de correo nuevo (via Items, no via la carpeta)
        _BandejaEventos._cliente_compartido = cliente
        items_eventos = bandeja.Items
        win32com.client.DispatchWithEvents(items_eventos, _BandejaEventos)
        print("[Monitor] Escuchando correos nuevos en tiempo real...\n")

        # Bucle principal — mantiene COM vivo y procesa mensajes de Windows
        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(1)

    except KeyboardInterrupt:
        print("\nMonitor detenido por el usuario.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
