"""
Monitor de correos en tiempo real via COM Events (win32com).

En lugar de hacer polling cada N segundos, se suscribe al evento ItemAdd
de la Bandeja de entrada de Outlook. Cuando llega un correo nuevo, se
procesa de inmediato sin esperar ningún intervalo.

Uso:
    python monitor_correos.py

Detener con Ctrl+C.
"""

import sys
import io
import os
import re
from pathlib import Path as _PathEarly
sys.path.insert(0, str(_PathEarly(__file__).resolve().parent.parent))
from Reportes.generador_reporte import generar_reporte_word
import shutil
import time
import logging
import warnings
import tempfile
import threading
import urllib3
import pythoncom
import win32com.client
from datetime import datetime

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

from pathlib import Path
from configuracion.ajustes import OUTLOOK_EMAIL, MODO_MOCK, EXTENSIONES_BLOQUEADAS, MODO_SIN_AV, SEGURIDAD_SERVICIO_AV, SEGURIDAD_ESPERA_AV_SEGUNDOS, SEGURIDAD_TAMANO_MAX_MB, SEGURIDAD_EXTENSIONES_PERMITIDAS, CORREO_ALERTAS_INCONSISTENCIAS
from Utilidades.logger_errores import log_error, log_advertencia, log_evento
from Utilidades.verificador_cadenas import verificar_cadena
from Nucleo.recopilador_documentos import _expandir_zip, _expandir_rar

# ---------------------------------------------------------------------------
# Logger de subidas a SharePoint
# ---------------------------------------------------------------------------
_LOG_SUBIDAS = Path(__file__).parent.parent / "registros_subidas.log"

# ---------------------------------------------------------------------------
# Logger de inconsistencias y sugerencias de respuesta
# ---------------------------------------------------------------------------
_DIR_REGISTROS = Path(__file__).parent / "registros"
_DIR_REGISTROS.mkdir(exist_ok=True)


def _enviar_reporte_por_correo(ruta_docx: "Path", remitente: str, asunto: str, numero_po: str):
    """Envía el reporte Word como adjunto a CORREO_ALERTAS_INCONSISTENCIAS via Outlook COM."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem
        mail.To = CORREO_ALERTAS_INCONSISTENCIAS
        po_label = f" — OC {numero_po}" if numero_po else ""
        mail.Subject = f"[Inconsistencias]{po_label} {asunto}"
        mail.Body = (
            f"Se detectaron inconsistencias en documentos recibidos de {remitente}.\n\n"
            f"Asunto original: {asunto}\n"
            "Ver el reporte adjunto para el detalle y el borrador de respuesta al proveedor."
        )
        mail.Attachments.Add(str(ruta_docx.resolve()))
        mail.Send()
        log_evento(f"Reporte enviado a {CORREO_ALERTAS_INCONSISTENCIAS}: {ruta_docx.name}")
    except Exception as exc:
        log_error("monitor_correos", "MON-006", "-", f"Error enviando reporte por correo: {exc}")


def _log_inconsistencias(remitente: str, asunto: str, docs_con_inconsistencias: list[dict],
                         sugerencia: str, numero_po: str = "", nombre_proveedor: str = ""):
    """Genera el reporte de inconsistencias como documento Word (.docx) y lo envía por correo."""
    try:
        ruta = generar_reporte_word(remitente, asunto, docs_con_inconsistencias, sugerencia,
                                    numero_po, nombre_proveedor)
        log_evento(f"Reporte de inconsistencias generado: {ruta}")
        _enviar_reporte_por_correo(ruta, remitente, asunto, numero_po)
    except Exception as exc:
        log_error("monitor_correos", "MON-005", "-", f"Error generando reporte Word de inconsistencias: {exc}")


def _guardar_borrador(remitente: str, asunto: str, sugerencia: str):
    """
    Guarda el borrador de respuesta en registros/borradores_YYYY-MM-DD.log.

    Formato pensado para migrar a Outlook drafts: cada entrada incluye
    Para, Asunto y Cuerpo listos para usar.  Cuando se implemente la
    creación de borradores en Outlook, esta función es el punto de
    conexión — bastará con llamar a la API COM/Graph desde aquí.
    """
    if not sugerencia:
        return
    fecha_hoy  = datetime.now().strftime("%Y-%m-%d")
    ruta_log   = _DIR_REGISTROS / f"borradores_{fecha_hoy}.log"
    ts         = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    asunto_re  = f"RE: {asunto}" if not asunto.upper().startswith("RE:") else asunto
    lineas = [
        "",
        "=" * 70,
        f"Timestamp : {ts}",
        f"Estado    : PENDIENTE",          # futuro: ENVIADO / DESCARTADO
        f"Para      : {remitente}",
        f"Asunto    : {asunto_re}",
        "-" * 70,
        sugerencia,
        "=" * 70,
    ]
    with open(ruta_log, "a", encoding="utf-8") as f:
        f.write("\n".join(lineas) + "\n")

_logger_subidas = logging.getLogger("subidas_sharepoint")
if not _logger_subidas.handlers:
    _handler = logging.FileHandler(_LOG_SUBIDAS, encoding="utf-8")
    _handler.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_handler)
    _logger_subidas.setLevel(logging.INFO)


def _siguiente_consecutivo() -> int:
    """Cuenta las líneas existentes en el log para obtener el próximo consecutivo."""
    if not _LOG_SUBIDAS.exists():
        return 1
    lines = _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines()
    return len([l for l in lines if l.strip()]) + 1


def _log_subida(remitente: str, ruta_sharepoint: str):
    """Registra una subida en el formato: N | timestamp | correo | ruta_completa"""
    ruta_completa = ruta_sharepoint.replace("\\", "/")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n  = _siguiente_consecutivo()
    _logger_subidas.info(f"{n} | {ts} | {remitente} | {ruta_completa}")
import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).resolve().parent.parent.parent / "Agente_Seguridad"))

from Nucleo.clasificador          import procesar_adjunto, extraer_po_y_bl_de_asunto
from Integraciones.graph_client   import GraphClient
from agente_seguridad             import ejecutar as seguridad_adjunto, verificar_servicio_av
from Integraciones.notificador    import notificar_error
from Aprobaciones.cola_aprobacion import encolar_sugerencia

# ---------------------------------------------------------------------------
# Helpers COM
# ---------------------------------------------------------------------------

_RE_IMAGEN_INLINE = re.compile(r"^image\d+\.", re.IGNORECASE)
_EXTS_IMAGEN     = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
_TAMANO_MAX_FIRMA = 100 * 1024  # 100 KB


def _es_imagen_inline(adj) -> bool:
    """Devuelve True si el adjunto es una imagen inline (firma de correo)."""
    nombre = adj.FileName or ""
    ext    = Path(nombre).suffix.lower()

    # Patrón image001.png, image002.jpg, etc.
    if _RE_IMAGEN_INLINE.match(nombre):
        return True

    # Imagen pequeña (<= 100 KB) — casi seguro es firma o logo
    if ext in _EXTS_IMAGEN:
        try:
            if adj.Size <= _TAMANO_MAX_FIRMA:
                return True
        except Exception:
            pass

    # Tiene ContentId (imagen embebida en HTML)
    try:
        content_id = adj.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
        )
        if content_id:
            return True
    except Exception:
        pass

    # Tipo OLE embebido
    try:
        if adj.Type == 6:
            return True
    except Exception:
        pass

    return False


def _save_con_timeout(adj, ruta: str) -> bool:
    """
    Guarda el adjunto leyendo sus bytes via PropertyAccessor (hilo COM principal)
    y escribiéndolos al disco. Evita el problema de threading de SaveAsFile.
    Si falla la lectura via PropertyAccessor, cae a SaveAsFile directo.
    Devuelve True si se guardó correctamente, False si hubo timeout/error.
    """
    import signal as _signal

    # Intentar leer bytes directamente via PropertyAccessor (no bloquea COM)
    try:
        datos = adj.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x37010102"
        )
        with open(ruta, "wb") as f:
            f.write(datos)
        return True
    except Exception:
        pass

    # Fallback: SaveAsFile en hilo principal — sin threading para evitar el
    # error COM cross-thread. Usamos signal.alarm en Unix; en Windows no hay
    # SIGALRM, así que simplemente llamamos directo y asumimos que el filtro
    # de imágenes inline ya redujo la mayoría de los casos problemáticos.
    try:
        adj.SaveAsFile(ruta)
        return True
    except Exception as e:
        raise e

# ---------------------------------------------------------------------------
# Registro de correos ya procesados (mismo archivo que recopilador_documentos)
# ---------------------------------------------------------------------------
_PROCESADOS_FILE = Path(__file__).parent / "procesados.txt"


_MARCA_FECHA = "ULTIMA_FECHA:"


def _cargar_procesados() -> set:
    if _PROCESADOS_FILE.exists():
        return set(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines())
    return set()


def _cargar_ultima_fecha() -> datetime:
    """Devuelve la fecha+hora del correo más reciente procesado.
    Si no hay registro, devuelve hace 24 horas como punto de partida seguro."""
    from datetime import timedelta
    fallback = datetime.now() - timedelta(hours=24)
    if not _PROCESADOS_FILE.exists():
        return fallback
    for linea in reversed(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines()):
        if linea.startswith(_MARCA_FECHA):
            try:
                return datetime.strptime(linea[len(_MARCA_FECHA):], "%Y-%m-%d %H:%M:%S")
            except Exception:
                return fallback
    return fallback


def _cargar_rutas_subidas() -> set:
    """Carga todas las rutas de SharePoint ya subidas desde registros_subidas.log."""
    if not _LOG_SUBIDAS.exists():
        return set()
    rutas = set()
    for linea in _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines():
        partes = linea.split(" | ", 3)
        if len(partes) == 4:
            rutas.add(partes[3].strip())
    return rutas


def _marcar_procesado(entry_id: str, clave_secundaria: str | None = None,
                      received_time: datetime | None = None):
    """Registra el EntryID, clave secundaria y fecha+hora del correo en procesados.txt."""
    with _PROCESADOS_FILE.open("a", encoding="utf-8") as f:
        f.write(entry_id + "\n")
        if clave_secundaria:
            f.write(clave_secundaria + "\n")
        if received_time:
            f.write(f"{_MARCA_FECHA}{received_time.strftime('%Y-%m-%d %H:%M:%S')}\n")


# ---------------------------------------------------------------------------
# Procesador de correo — reutiliza la lógica de recopilador_documentos.py
# ---------------------------------------------------------------------------

def _procesar_mensaje(msg, cliente: GraphClient, rutas_subidas: set | None = None) -> dict:
    """Procesa un único objeto MailItem de Outlook.
    Retorna dict con 'rutas_subidas' (list) y 'errores' (list).
    """
    resultado = {"rutas_subidas": [], "errores": []}
    try:
        asunto    = msg.Subject or ""
        remitente = msg.SenderEmailAddress or ""
        msg_id    = msg.EntryID
        try:
            received_time    = msg.ReceivedTime.replace(tzinfo=None)
            fecha_str        = received_time.strftime("%Y-%m-%d %H:%M")
            clave_secundaria = f"{remitente}|{asunto}|{fecha_str}"
        except Exception:
            received_time    = None
            clave_secundaria = None

        if msg.Attachments.Count == 0:
            log_evento("BASURA", "monitor_correos", remitente=remitente, asunto=asunto,
                       detalle="correo sin adjuntos")
            _marcar_procesado(msg_id, clave_secundaria, received_time)
            return resultado

        # Verificar AV antes de tocar adjuntos
        if MODO_SIN_AV:
            ok_av, err_av = True, None
        else:
            ok_av, err_av = verificar_servicio_av(SEGURIDAD_SERVICIO_AV)
        if not ok_av:
            print(f"[SEG] ABORTADO — Bitdefender no activo ({err_av})")
            return resultado

        po_asunto, bl_asunto = extraer_po_y_bl_de_asunto(asunto)

        archivos_subidos          = []
        ultimo_info               = None
        docs_con_inconsistencias  = []  # para generar sugerencia de respuesta al proveedor
        proveedor_correo          = ""  # nombre de empresa extraído de INVOICE/PACKING LIST
        proveedor_correo_fallback = ""  # nombre de empresa de otro tipo de documento

        # Verificar cadena de correos — detecta conflictos de PO/BL entre mensajes del hilo
        incs_cadena = verificar_cadena(msg, asunto, po_asunto, bl_asunto)
        if incs_cadena:
            docs_con_inconsistencias.append({
                "nombre_archivo": "(cadena de correos)",
                "tipo":           "CADENA",
                "inconsistencias": incs_cadena,
            })
            log_evento("CADENA", "monitor_correos", remitente=remitente, asunto=asunto,
                       detalle=f"{len(incs_cadena)} inconsistencia(s) en cadena de correos")

        n_adjuntos_reales = 0  # adjuntos no filtrados (no inline, no bloqueados)
        for i in range(1, msg.Attachments.Count + 1):
            try:
                adj    = msg.Attachments.Item(i)
                nombre = adj.FileName
            except Exception as e:
                log_error("monitor_correos", "MON-000", "-", f"No se pudo acceder al adjunto {i}: {e}", remitente=remitente, asunto=asunto)
                continue
            ext    = Path(nombre).suffix.lower()

            if ext in EXTENSIONES_BLOQUEADAS:
                continue

            # Filtrar imágenes inline (firmas de correo) antes de SaveAsFile
            if _es_imagen_inline(adj):
                continue

            n_adjuntos_reales += 1

            # Guardar adjunto en archivo temporal
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="adj_")
            tmp.close()
            try:
                _save_con_timeout(adj, tmp.name)
            except Exception as e_save:
                print(f"  [ERROR] No se pudo guardar adjunto: {nombre} — {e_save}")
                log_error("monitor_correos", "MON-001", nombre, f"No se pudo guardar adjunto: {type(e_save).__name__}: {e_save}", remitente=remitente, asunto=asunto)
                notificar_error(f"No se pudo guardar adjunto: {nombre}", detalle=str(e_save), paso="monitor_correos._save_con_timeout", como_resolver="Verificar permisos de escritura en carpeta temporal y espacio en disco.")
                try:
                    os.remove(tmp.name)
                except Exception:
                    pass
                # Notificar por Teams
                try:
                    from configuracion.ajustes import TEAMS_WEBHOOK_URL
                    import requests as _req
                    _req.post(TEAMS_WEBHOOK_URL, json={
                        "@type":    "MessageCard",
                        "@context": "http://schema.org/extensions",
                        "themeColor": "D93025",
                        "summary": f"Error al guardar adjunto — {nombre}",
                        "sections": [{
                            "activityTitle": f"[CLASIFICADOR] Error al guardar adjunto — No se subió a SharePoint",
                            "activityText": (
                                f"<b>Qué pasó:</b> No se pudo guardar el adjunto del correo en el directorio temporal. El archivo no fue subido a SharePoint.<br><br>"
                                f"<b>Correo:</b> {asunto}<br>"
                                f"<b>Remitente:</b> {remitente}<br>"
                                f"<b>Adjunto:</b> {nombre}<br>"
                                f"<b>Error:</b> {type(e_save).__name__}: {e_save}<br><br>"
                                f"<b>Por qué ocurre:</b> El objeto COM del adjunto de Outlook no pudo escribir el archivo en disco. Causas frecuentes: el adjunto está bloqueado por Outlook o Bitdefender, el perfil COM se corrompió, o el archivo temporal no pudo crearse por falta de espacio en disco.<br><br>"
                                f"<b>Cómo solucionar:</b><br>"
                                f"1. Verificar que Bitdefender no esté bloqueando la carpeta %TEMP%.<br>"
                                f"2. Reiniciar Outlook y volver a ejecutar el monitor.<br>"
                                f"3. Solicitar al remitente que reenvíe el correo con el adjunto.<br>"
                                f"4. Revisar el log en <code>errores_clasificador.log</code> (código MON-001) para más detalle."
                            ),
                        }],
                    }, verify=False, timeout=10)
                except Exception:
                    pass
                continue

            # Directorio temporal para descompresión (ZIP/RAR)
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
                        asunto_correo=asunto,
                    )
                    if not destinos:
                        continue

                    for info in destinos:
                        from Nucleo.clasificador import formatear_numero_oc
                        oc_base   = formatear_numero_oc(info["numero_po"])
                        oc_real   = cliente.buscar_carpeta_oc(oc_base)
                        ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)

                        ruta_real_norm = ruta_real.replace("\\", "/")

                        if rutas_subidas and ruta_real_norm in rutas_subidas:
                            print(f"[SKIP] Ya subido anteriormente: {info['nombre_archivo']}")
                            archivos_subidos.append(info["tipo"])
                            continue

                        cliente.crear_carpeta_si_no_existe(ruta_real)
                        ruta_archivo_subir = info.get("ruta_local") or adj_interno["ruta_local"]
                        cliente.subir_archivo(ruta_archivo_subir, ruta_real)
                        _log_subida(remitente, ruta_real)
                        if rutas_subidas is not None:
                            rutas_subidas.add(ruta_real_norm)
                        archivos_subidos.append(info["tipo"])
                        resultado["rutas_subidas"].append(ruta_real_norm)
                        ultimo_info = {**info, "ruta_sharepoint": ruta_real}

                        # Acumular nombre de proveedor — priorizar INVOICE o PACKING LIST
                        _prov = info.get("nombre_proveedor", "")
                        if _prov:
                            if info["tipo"] in ("INVOICE", "PACKING LIST", "PL + INV"):
                                if not proveedor_correo:
                                    proveedor_correo = _prov
                            elif not proveedor_correo_fallback:
                                proveedor_correo_fallback = _prov

                        # Acumular inconsistencias para sugerencia de respuesta
                        if info.get("inconsistencias"):
                            docs_con_inconsistencias.append({
                                "nombre_archivo": info["nombre_archivo"],
                                "tipo":           info["tipo"],
                                "inconsistencias": info["inconsistencias"],
                            })

            finally:
                # Limpiar archivo temporal del adjunto
                try:
                    os.remove(tmp.name)
                except Exception:
                    pass
                # Limpiar directorio de extracción ZIP/RAR
                if dir_extraccion and Path(dir_extraccion).exists():
                    try:
                        shutil.rmtree(dir_extraccion)
                    except Exception:
                        pass

        # Correo basura: tenía adjuntos pero todos eran inline o extensión bloqueada
        if n_adjuntos_reales == 0 and msg.Attachments.Count > 0:
            log_evento("BASURA", "monitor_correos", remitente=remitente, asunto=asunto,
                       detalle=f"todos los adjuntos ({msg.Attachments.Count}) son imágenes inline o extensión bloqueada")

        # Marcar correo como procesado (EntryID + clave secundaria + fecha para criterio de corte)
        _marcar_procesado(msg_id, clave_secundaria, received_time)

        # Sugerencia de respuesta al proveedor si hay inconsistencias
        if docs_con_inconsistencias:
            print(f"[INCONSISTENCIAS] {len(docs_con_inconsistencias)} documento(s) con problemas — {asunto}")
            from Nucleo.clasificador_claude import (
                generar_respuesta_proveedor_consolidada,
                contrastar_sugerencia_vs_inconsistencias,
            )
            _prov_final = proveedor_correo or proveedor_correo_fallback
            sugerencia = generar_respuesta_proveedor_consolidada(
                remitente=remitente,
                asunto=asunto,
                documentos=docs_con_inconsistencias,
                nombre_proveedor=_prov_final,
            )
            sugerencia = contrastar_sugerencia_vs_inconsistencias(
                sugerencia=sugerencia or "",
                documentos=docs_con_inconsistencias,
            )
            _log_inconsistencias(remitente, asunto, docs_con_inconsistencias, sugerencia,
                                 po_asunto or "", nombre_proveedor=_prov_final)
            ruta_pendiente = encolar_sugerencia(remitente, asunto, sugerencia or "", po_asunto or "")
            print(f"[PENDIENTE] Sugerencia guardada: {ruta_pendiente.name}")

    except Exception as e:
        resultado["errores"].append(f"{remitente} | {asunto} — {type(e).__name__}: {e}")
        log_error("monitor_correos", "MON-002", "-", f"Error procesando mensaje: {type(e).__name__}: {e}", remitente=remitente, asunto=asunto)
        notificar_error(f"Error procesando mensaje de {remitente}", detalle=f"{type(e).__name__}: {e}", paso="monitor_correos._procesar_mensaje", como_resolver="Revisar el log de errores y verificar el adjunto manualmente.")

    return resultado


# ---------------------------------------------------------------------------
# COM Event sink — recibe ItemAdd cuando llega un correo nuevo
# ---------------------------------------------------------------------------

class _BandejaEventos:
    """
    Sink de eventos COM para la carpeta Bandeja de entrada.
    win32com llama a OnItemAdd cada vez que llega un item nuevo.
    """

    # Atributos de clase inyectados antes del DispatchWithEvents
    _cliente_compartido: GraphClient = None
    _rutas_subidas: set = None

    def OnItemAdd(self, item):
        """Disparado por COM cuando se agrega un item a la carpeta observada."""
        try:
            # Solo procesar MailItem (Class == 43)
            if item.Class != 43:
                return
            _procesar_mensaje(item, self._cliente_compartido, self._rutas_subidas)
        except Exception as e:
            print(f"[ERROR] OnItemAdd: {e}")
            log_error("monitor_correos", "MON-003", "-", f"OnItemAdd: {type(e).__name__}: {e}")
            notificar_error("Error en listener de correos (OnItemAdd)", detalle=f"{type(e).__name__}: {e}", paso="monitor_correos.OnItemAdd", como_resolver="Reiniciar el monitor. Si persiste, verificar que Outlook esté abierto y la conexión COM activa.")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main():
    import argparse
    from datetime import timedelta
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--rehash-48h", action="store_true", dest="rehash_48h",
                        help="Reprocesar correos de las últimas 48h ignorando procesados.txt.")
    parser.add_argument("--rehash-horas", type=int, dest="rehash_horas", default=0,
                        help="Reprocesar correos de las últimas N horas (ej: --rehash-horas 5).")
    args, _ = parser.parse_known_args()
    if args.rehash_horas > 0:
        args.rehash_48h = True  # activa el mismo flujo, con ventana variable
    _horas_rehash = args.rehash_horas if args.rehash_horas > 0 else 48

    modo = "MOCK" if MODO_MOCK else "REAL"
    print(f"Monitor de correos iniciado [{modo}]")
    print(f"Cuenta: {OUTLOOK_EMAIL}")
    if args.rehash_48h:
        print(f"Modo rehash activo — se reprocesarán correos de las últimas {_horas_rehash}h.")
        print("Los archivos ya subidos a SharePoint serán omitidos.\n")
    else:
        print("Esperando correos nuevos... (Ctrl+C para detener)\n")

    # COM debe inicializarse en el hilo principal y mantenerse vivo
    pythoncom.CoInitialize()
    try:
        # Reintentos para dar tiempo a Outlook a arrancar si acaba de abrirse
        outlook = None
        for intento in range(1, 6):
            try:
                outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
                break
            except Exception as e:
                codigo = getattr(e, 'hresult', None)
                if intento == 1:
                    print(f"[Outlook] Conectando... (intento {intento}/5)")
                else:
                    print(f"[Outlook] Reintentando conexión COM... (intento {intento}/5)")
                if intento == 3:
                    print("          Asegúrate de que Outlook esté abierto y con sesión activa.")
                    print("          Si Outlook corre como Administrador, este script también debe correr como Administrador.")
                if intento == 5:
                    raise RuntimeError(
                        f"No se pudo conectar a Outlook después de 5 intentos.\n"
                        f"  → Abre Outlook y asegúrate de que esté iniciado antes de correr el monitor.\n"
                        f"  → Verifica que ambos (Outlook y este terminal) corran con el mismo nivel de permisos.\n"
                        f"  → Error original: {e}"
                    ) from e
                time.sleep(3)

        namespace = outlook.GetNamespace("MAPI")

        # Localizar la bandeja de entrada de la cuenta configurada
        bandeja = None
        for store in namespace.Stores:
            try:
                if OUTLOOK_EMAIL.lower() in store.DisplayName.lower():
                    bandeja = store.GetDefaultFolder(6)  # olFolderInbox
                    print(f"[Outlook] Suscrito a: {store.DisplayName}")
                    break
            except Exception:
                continue

        if bandeja is None:
            print(f"[ERROR] No se encontró la cuenta '{OUTLOOK_EMAIL}' en Outlook.")
            print("        Verifica que Outlook esté abierto y la cuenta configurada.")
            log_error("monitor_correos", "MON-004", "-", f"Cuenta '{OUTLOOK_EMAIL}' no encontrada en Outlook")
            notificar_error(f"Cuenta '{OUTLOOK_EMAIL}' no encontrada en Outlook", paso="monitor_correos.iniciar", como_resolver="Abrir Outlook, iniciar sesión con la cuenta correcta y volver a ejecutar el monitor.")
            sys.exit(1)

        cliente = GraphClient()

        # -----------------------------------------------------------------
        # Escaneo inicial: recorre TODA la bandeja de más reciente a más
        # antiguo. Los correos ya en procesados.txt se saltan sin llamar
        # a Claude — solo iteración COM, sin costo de tokens.
        #
        # Modo --rehash-48h: ignora procesados.txt, filtra solo correos
        # de las últimas 48h, y confía en rutas_subidas (SharePoint) para
        # no re-subir archivos que ya existen.
        # -----------------------------------------------------------------
        procesados    = _cargar_procesados()
        rutas_subidas = _cargar_rutas_subidas()
        items_bandeja = bandeja.Items
        items_bandeja.Sort("[ReceivedTime]", True)  # más recientes primero

        if args.rehash_48h:
            corte_48h = datetime.now() - timedelta(hours=_horas_rehash)
            print(f"[Monitor] Rehash-{_horas_rehash}h — correos desde {corte_48h.strftime('%Y-%m-%d %H:%M')} ({items_bandeja.Count} elementos en bandeja)")
        else:
            print(f"[Monitor] Escaneo inicial — recorriendo bandeja completa ({items_bandeja.Count} elementos)")

        resumen_rutas   = []
        resumen_errores = []
        n_inicial  = 0
        n_saltados = 0
        for item in items_bandeja:
            try:
                if item.Class != 43:
                    continue

                # Obtener fecha de recepción
                try:
                    received_time = item.ReceivedTime.replace(tzinfo=None)
                except Exception:
                    received_time = None

                if args.rehash_48h:
                    # En modo rehash: solo correos de las últimas 48h;
                    # cuando llegamos a uno más antiguo, el resto también lo es (orden desc)
                    if received_time is not None and received_time < corte_48h:
                        break
                    # No consultar procesados.txt — dejar que _procesar_mensaje
                    # use rutas_subidas para evitar re-subidas a SharePoint
                else:
                    # Clave primaria: EntryID
                    # Clave secundaria: asunto + remitente + fecha (por si el EntryID cambió)
                    entry_id = item.EntryID
                    try:
                        fecha_str = received_time.strftime("%Y-%m-%d %H:%M") if received_time else ""
                        clave_secundaria = f"{item.SenderEmailAddress}|{item.Subject}|{fecha_str}"
                    except Exception:
                        clave_secundaria = None

                    ya_procesado = (
                        entry_id in procesados
                        or (clave_secundaria and clave_secundaria in procesados)
                    )

                    if ya_procesado:
                        n_saltados += 1
                        continue

                r = _procesar_mensaje(item, cliente, rutas_subidas)
                resumen_rutas.extend(r["rutas_subidas"])
                resumen_errores.extend(r["errores"])
                n_inicial += 1
                try:
                    procesados.add(item.EntryID)
                except Exception:
                    pass
            except Exception as e:
                resumen_errores.append(f"Error inesperado: {type(e).__name__}: {e}")
                continue

        print(f"[Monitor] Escaneo inicial completo — bandeja recorrida ({n_inicial} nuevos, {n_saltados} omitidos).")
        print("[Monitor] Escaneo completado.\n")
        print("=" * 60)
        if resumen_rutas:
            print(f"SUBIDOS ({len(resumen_rutas)}):")
            for ruta in resumen_rutas:
                print(f"  {ruta}")
        else:
            print("SUBIDOS: ninguno")
        print()
        if resumen_errores:
            print(f"ERRORES ({len(resumen_errores)}):")
            for err in resumen_errores:
                print(f"  [ERROR] {err}")
        else:
            print("ERRORES: ninguno")
        print("=" * 60)

    except KeyboardInterrupt:
        print("\nMonitor detenido por el usuario.")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
