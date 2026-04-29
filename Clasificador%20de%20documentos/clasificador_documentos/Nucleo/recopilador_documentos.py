"""
Script principal del clasificador automático de documentos.
- Sin argumentos: corre una sola vez y termina.
- Con argumento 'loop': corre en modo continuo cada INTERVALO_SEGUNDOS.
  Detener con Ctrl+C.
"""

import sys
import os
import io
import re
import time
import logging
import warnings
import urllib3
import zipfile
import shutil
from datetime import datetime
from pathlib import Path

# Asegurar que clasificador_documentos/ esté en el path independientemente de dónde se corra
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# Logger de subidas a SharePoint
# ---------------------------------------------------------------------------
_LOG_SUBIDAS = Path(__file__).parent.parent / "registros_subidas.log"
_LOG_SUBIDAS.parent.mkdir(parents=True, exist_ok=True)

_logger_subidas = logging.getLogger("subidas_sharepoint")
if not _logger_subidas.handlers:
    _handler = logging.FileHandler(_LOG_SUBIDAS, encoding="utf-8")
    _handler.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_handler)
    _logger_subidas.setLevel(logging.INFO)



# ---------------------------------------------------------------------------
# Logger de inconsistencias — separado de eventos operativos
# ---------------------------------------------------------------------------
_LOG_INCONSISTENCIAS = Path(__file__).parent / "registros_inconsistencias.log"

_logger_inconsistencias = logging.getLogger("inconsistencias_clasificador")
if not _logger_inconsistencias.handlers:
    _handler_inc = logging.FileHandler(_LOG_INCONSISTENCIAS, encoding="utf-8")
    _handler_inc.setFormatter(logging.Formatter("%(message)s"))
    _logger_inconsistencias.addHandler(_handler_inc)
    _logger_inconsistencias.setLevel(logging.INFO)


def _log_inconsistencia(remitente: str, asunto: str, adjunto: str, tipo: str,
                        resumen: str, ruta_txt: str):
    """Registra una línea en registros_inconsistencias.log con resumen y ruta del reporte."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger_inconsistencias.info(
        f"{ts} | {remitente} | {asunto} | {adjunto} | {tipo} | {resumen} | reporte: {ruta_txt}"
    )


def _siguiente_consecutivo() -> int:
    """Lee el último número registrado en el log y devuelve el siguiente."""
    if not _LOG_SUBIDAS.exists():
        return 1
    for linea in reversed(_LOG_SUBIDAS.read_text(encoding="utf-8").splitlines()):
        linea = linea.strip()
        if not linea:
            continue
        try:
            return int(linea.split("|", 1)[0].strip()) + 1
        except ValueError:
            continue
    return 1


def _extraer_empresa(nombre_remitente: str, email: str) -> str:
    """
    Intenta obtener el nombre de la empresa desde el nombre del remitente o el email.
    - Si el nombre tiene formato 'Persona - Empresa' o 'Persona / Empresa', extrae la parte de empresa.
    - Si el nombre parece ser solo un nombre personal (1-3 palabras sin indicadores de empresa), lo devuelve tal cual.
    - Si el nombre está vacío, extrae el dominio del email como referencia.
    """
    import re as _re
    nombre = (nombre_remitente or "").strip()

    # Separadores comunes entre nombre y empresa
    for sep in (" - ", " / ", " | ", " @ ", " :: "):
        if sep in nombre:
            return nombre.split(sep, 1)[1].strip()

    if nombre:
        return nombre

    # Fallback: dominio del email
    if "@" in email:
        dominio = email.split("@", 1)[1]
        return dominio.split(".")[0].upper()

    return ""


def _log_subida(remitente: str, ruta_sharepoint: str, ya_existia: bool = False,
                nombre_remitente: str = "", proveedor: str = ""):
    """Registra una subida en el formato: N | timestamp | correo | proveedor | ruta_completa [| YA_EXISTIA]"""
    ruta_completa = ruta_sharepoint.replace("\\", "/")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    n  = _siguiente_consecutivo()
    # Prioridad: nombre extraído del documento > SenderName de Outlook > dominio del email
    empresa = proveedor or _extraer_empresa(nombre_remitente, remitente)
    sufijo = " | YA_EXISTIA — omitida subida duplicada" if ya_existia else ""
    _logger_subidas.info(f"{n} | {ts} | {remitente} | {empresa} | {ruta_completa}{sufijo}")

import sys as _sys
from pathlib import Path as _Path
_sys.path.insert(0, str(_Path(__file__).resolve().parent.parent / "Agente_Seguridad"))

from configuracion.ajustes import (
    INTERVALO_SEGUNDOS, MODO_MOCK, EXTENSIONES_BLOQUEADAS,
    MODO_SIN_AV, SEGURIDAD_SERVICIO_AV, SEGURIDAD_ESPERA_AV_SEGUNDOS,
    SEGURIDAD_TAMANO_MAX_MB, SEGURIDAD_EXTENSIONES_PERMITIDAS,
)
from Nucleo.clasificador        import procesar_adjunto, extraer_po_y_bl_de_asunto
from Integraciones.graph_client import GraphClient
from agente_seguridad           import ejecutar as _seg_ejecutar, verificar_servicio_av
from Utilidades.logger_errores  import log_error, log_advertencia, log_evento
from Nucleo.clasificador_claude import generar_respuesta_proveedor, generar_respuesta_proveedor_consolidada, detectar_inconsistencias_cruzadas
from Utilidades.verificador_cadenas import verificar_cadena


_DIR_INCONSISTENCIAS = Path(__file__).parent / "inconsistencias"


def _generar_txt_inconsistencias(
    nombre_archivo: str,
    tipo: str,
    inconsistencias: list[dict],
    remitente: str,
    asunto: str,
) -> Path:
    """
    Genera un .txt con el listado de inconsistencias y una sugerencia de respuesta al proveedor.
    Lo guarda en la carpeta local 'inconsistencias/' y retorna la ruta del archivo creado.
    """
    _DIR_INCONSISTENCIAS.mkdir(parents=True, exist_ok=True)
    ts              = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem            = Path(nombre_archivo).stem
    proveedor_safe  = re.sub(r'[\\/*?:"<>|]', "", remitente).strip().replace(" ", "_")[:60]
    ruta            = _DIR_INCONSISTENCIAS / f"{proveedor_safe}_{stem}_{ts}_inconsistencias.txt"

    altas  = [i for i in inconsistencias if i.get("severidad") == "alta"]
    medias = [i for i in inconsistencias if i.get("severidad") == "media"]
    bajas  = [i for i in inconsistencias if i.get("severidad") == "baja"]

    lineas = [
        f"Este es un correo automatizado, favor no responder a esta direccion.",
        f"",
        f"{'=' * 60}",
        f"REPORTE DE INCONSISTENCIAS",
        f"{'=' * 60}",
        f"Archivo  : {nombre_archivo}",
        f"Tipo     : {tipo}",
        f"Correo   : {asunto}",
        f"De       : {remitente}",
        f"Generado : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"",
        f"RESUMEN: {len(inconsistencias)} inconsistencia(s) — "
        f"altas={len(altas)}, medias={len(medias)}, bajas={len(bajas)}",
        f"{'=' * 60}",
        f"",
    ]

    for seccion, items in [("ALTA", altas), ("MEDIA", medias), ("BAJA", bajas)]:
        if not items:
            continue
        lineas.append(f"--- Severidad {seccion} ---")
        for i, inc in enumerate(items, 1):
            campo = inc.get("campo", "")
            desc  = inc.get("descripcion", "")
            lineas.append(f"  {i}. [{campo}] {desc}")
        lineas.append("")

    print(f"  [INCONSISTENCIAS] Generando sugerencia de respuesta al proveedor...")
    respuesta = generar_respuesta_proveedor(nombre_archivo, tipo, inconsistencias)

    lineas += [
        f"{'=' * 60}",
        f"SUGERENCIA DE RESPUESTA AL PROVEEDOR",
        f"{'=' * 60}",
        f"",
        respuesta if respuesta else "(No se pudo generar la sugerencia — verificar conexión con Claude API)",
        f"",
    ]

    ruta.write_text("\n".join(lineas), encoding="utf-8")
    print(f"  [INCONSISTENCIAS] Reporte guardado: {ruta.name}")
    return ruta


def _generar_txt_consolidado(
    remitente: str,
    asunto: str,
    documentos: list[dict],
    msg_obj=None,
    numero_po: str = "",
    nombre_proveedor: str = "",
) -> Path:
    """
    Genera un .txt consolidado con todas las inconsistencias de un correo completo
    y una sugerencia de respuesta única al proveedor.
    Guarda en 'inconsistencias/' y retorna la ruta del archivo creado.

    Args:
        remitente: correo del remitente
        asunto: asunto del correo
        documentos: lista de dicts con inconsistencias
        msg_obj: objeto de mensaje de Outlook (para extraer ReceivedTime)
        numero_po: número de orden de compra para usar en el nombre del archivo
    """
    _DIR_INCONSISTENCIAS.mkdir(parents=True, exist_ok=True)

    # Formato de fecha: JUE 26/MAR, MIE 25/MAR, etc.
    dias = ["LUN", "MAR", "MIE", "JUE", "VIE", "SAB", "DOM"]
    meses = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

    # Usar fecha de recepción del correo si está disponible, si no usar fecha actual
    if msg_obj:
        try:
            fecha_recibida = msg_obj.ReceivedTime
            dia_sem = dias[fecha_recibida.weekday()]
            mes_abr = meses[fecha_recibida.month - 1]
            fecha_formateada = f"{dia_sem} {fecha_recibida.day} {mes_abr}"
        except Exception:
            # Fallback a fecha actual si hay error al leer ReceivedTime
            ahora = datetime.now()
            dia_sem = dias[ahora.weekday()]
            mes_abr = meses[ahora.month - 1]
            fecha_formateada = f"{dia_sem} {ahora.day} {mes_abr}"
    else:
        # Sin msg_obj, usar fecha actual
        ahora = datetime.now()
        dia_sem = dias[ahora.weekday()]
        mes_abr = meses[ahora.month - 1]
        fecha_formateada = f"{dia_sem} {ahora.day} {mes_abr}"

    # Asunto truncado (máximo 120 caracteres) para nombre de archivo
    asunto_safe    = re.sub(r'[\\/*?:"<>|]', "", asunto).strip()[:120]
    proveedor_safe = re.sub(r'[\\/*?:"<>|]', "", remitente).strip()
    po_safe        = re.sub(r'[\\/*?:"<>|]', "", numero_po).strip() if numero_po else ""
    if po_safe:
        # Versionar: OC-XXXXXXXX_v1_inconsistencias.txt, OC-XXXXXXXX_v2_inconsistencias.txt, ...
        existing = list(_DIR_INCONSISTENCIAS.glob(f"OC-{po_safe}_v*.txt"))
        nums_v = []
        for f in existing:
            m = re.search(r'_v(\d+)\.txt$', f.name)
            if m:
                nums_v.append(int(m.group(1)))
        version = (max(nums_v) + 1) if nums_v else 1
        ruta = _DIR_INCONSISTENCIAS / f"OC-{po_safe}_v{version}.txt"
    else:
        ruta = _DIR_INCONSISTENCIAS / f"{fecha_formateada}_{proveedor_safe}_{asunto_safe}_inconsistencias.txt"

    total = sum(len(d["inconsistencias"]) for d in documentos)
    altas  = sum(1 for d in documentos for i in d["inconsistencias"] if i.get("severidad") == "alta")
    medias = sum(1 for d in documentos for i in d["inconsistencias"] if i.get("severidad") == "media")
    bajas  = sum(1 for d in documentos for i in d["inconsistencias"] if i.get("severidad") == "baja")

    lineas = [
        "Este es un correo automatizado, favor no responder a esta direccion.",
        "",
        "=" * 60,
        "REPORTE CONSOLIDADO DE INCONSISTENCIAS",
        "=" * 60,
        f"Correo   : {asunto}",
        f"De       : {remitente}",
        f"Generado : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Documentos con inconsistencias: {len(documentos)}",
        f"",
        f"RESUMEN TOTAL: {total} inconsistencia(s) — altas={altas}, medias={medias}, bajas={bajas}",
        "=" * 60,
        "",
    ]

    for doc in documentos:
        inc_doc = doc["inconsistencias"]
        lineas.append(f"Documento: {doc['nombre_archivo']} (tipo: {doc['tipo']})")
        lineas.append("-" * 40)
        for i, inc in enumerate(inc_doc, 1):
            sev   = inc.get("severidad", "").upper()
            campo = inc.get("campo", "")
            desc  = inc.get("descripcion", "")
            lineas.append(f"  {i}. [{sev}] {campo}: {desc}")
        lineas.append("")

    print(f"  [CONSOLIDADO] Generando sugerencia de respuesta al proveedor...")
    respuesta = generar_respuesta_proveedor_consolidada(remitente, asunto, documentos, nombre_proveedor)

    lineas += [
        "=" * 60,
        "SUGERENCIA DE RESPUESTA AL PROVEEDOR",
        "=" * 60,
        "",
        respuesta if respuesta else "(No se pudo generar la sugerencia — verificar conexión con Claude API)",
        "",
    ]

    ruta.write_text("\n".join(lineas), encoding="utf-8")
    print(f"  [CONSOLIDADO] Reporte guardado: {ruta.name}")
    return ruta


def seguridad_adjunto(ruta_local: str, nombre_orig: str) -> dict:
    return _seg_ejecutar(
        ruta_local      = ruta_local,
        nombre_orig     = nombre_orig,
        extensiones_ok  = SEGURIDAD_EXTENSIONES_PERMITIDAS,
        tamano_max_mb   = SEGURIDAD_TAMANO_MAX_MB,
        espera_av       = SEGURIDAD_ESPERA_AV_SEGUNDOS,
        nombre_proyecto = "sharepoint",
    )

_PROCESADOS_FILE = Path(__file__).parent / "procesados.txt"
_DIR_TEMPORAL    = Path(__file__).parent / "temporal"


def _cargar_procesados() -> set:
    if _PROCESADOS_FILE.exists():
        return set(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines())
    return set()


def _marcar_procesado(correo_id: str):
    with _PROCESADOS_FILE.open("a", encoding="utf-8") as f:
        f.write(correo_id + "\n")


def _expandir_rar(ruta_rar: str) -> tuple[list[dict], str]:
    """
    Descomprime un RAR en una subcarpeta dentro de _DIR_TEMPORAL.
    Requiere: rarfile + unrar.exe en PATH (o WinRAR instalado).
    Retorna (lista de adjuntos virtuales, ruta carpeta temporal).
    """
    import uuid
    carpeta_tmp = str(_DIR_TEMPORAL / f"rar_{uuid.uuid4().hex[:8]}")
    Path(carpeta_tmp).mkdir(parents=True, exist_ok=True)
    adjuntos = []
    try:
        import rarfile
        rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"
        with rarfile.RarFile(ruta_rar, "r") as rf:
            for entry in rf.infolist():
                if entry.is_dir():
                    continue
                nombre = Path(entry.filename).name
                ext = Path(nombre).suffix.lower()
                if ext in EXTENSIONES_BLOQUEADAS or not nombre or nombre.startswith("._"):
                    continue
                ruta_extraida = rf.extract(entry, carpeta_tmp)
                adjuntos.append({"nombre": nombre, "ruta_local": ruta_extraida})
    except ImportError:
        print("  [WARN] rarfile no instalado — instalar con: pip install rarfile")
        log_advertencia("recopilador", "RECP-001", ruta_rar, "rarfile no instalado")
        shutil.rmtree(carpeta_tmp, ignore_errors=True)
        return [], ""
    except Exception as e:
        print(f"  [WARN] No se pudo descomprimir RAR: {e}")
        log_advertencia("recopilador", "RECP-002", ruta_rar, f"No se pudo descomprimir RAR: {e}")
        shutil.rmtree(carpeta_tmp, ignore_errors=True)
        return [], ""
    return adjuntos, carpeta_tmp


def _expandir_zip(ruta_zip: str) -> tuple[list[dict], str]:
    """
    Descomprime un ZIP en una subcarpeta dentro de _DIR_TEMPORAL.
    Retorna (lista de adjuntos virtuales, ruta carpeta temporal).
    Ignora extensiones bloqueadas y entradas que sean carpetas.
    """
    import uuid
    carpeta_tmp = str(_DIR_TEMPORAL / f"zip_{uuid.uuid4().hex[:8]}")
    Path(carpeta_tmp).mkdir(parents=True, exist_ok=True)
    adjuntos = []
    try:
        with zipfile.ZipFile(ruta_zip, "r") as zf:
            for entry in zf.infolist():
                if entry.is_dir():
                    continue
                nombre = Path(entry.filename).name
                ext = Path(nombre).suffix.lower()
                if ext in EXTENSIONES_BLOQUEADAS or not nombre or nombre.startswith("._") or "__MACOSX" in entry.filename:
                    continue
                ruta_extraida = zf.extract(entry, carpeta_tmp)
                adjuntos.append({"nombre": nombre, "ruta_local": ruta_extraida})
    except Exception as e:
        print(f"  [WARN] No se pudo descomprimir ZIP: {e}")
        log_advertencia("recopilador", "RECP-003", ruta_zip, f"No se pudo descomprimir ZIP: {e}")
        shutil.rmtree(carpeta_tmp, ignore_errors=True)
        return [], ""
    return adjuntos, carpeta_tmp


def ejecutar_ciclo(max_correos: int = None, reprocesar: bool = False):
    """Un ciclo completo: leer correos -> clasificar -> subir.

    reprocesar=True: ignora procesados.txt y vuelve a intentar todos los correos.
    Además, si un archivo se clasifica correctamente y estaba en OTROS, lo borra de ahí.
    """
    print("\n" + "-" * 50)
    if reprocesar:
        print("Iniciando ciclo de REPROCESAMIENTO (bandeja completa, ignorando procesados.txt)...")
    else:
        print("Iniciando ciclo de clasificación...")

    # Crear carpeta temporal limpia para este ciclo
    if _DIR_TEMPORAL.exists():
        shutil.rmtree(_DIR_TEMPORAL)
    _DIR_TEMPORAL.mkdir(parents=True)

    try:
        return _ejecutar_ciclo_interno(max_correos=max_correos, reprocesar=reprocesar)
    finally:
        # Limpiar carpeta temporal al terminar el ciclo (éxito o error)
        shutil.rmtree(_DIR_TEMPORAL, ignore_errors=True)


def _ejecutar_ciclo_interno(max_correos: int = None, reprocesar: bool = False):
    """Lógica interna del ciclo — llamada desde ejecutar_ciclo tras crear la carpeta temporal."""
    # Verificar AV una sola vez por ciclo — si no está activo, se aborta
    if MODO_SIN_AV:
        ok_av, err_av = True, None
    else:
        ok_av, err_av = verificar_servicio_av(SEGURIDAD_SERVICIO_AV)
    if not ok_av:
        print(f"[SEG] ABORTADO — Bitdefender no está activo ({err_av}). No se procesarán adjuntos.")
        log_evento("ABORTADO", "recopilador", detalle=f"Bitdefender no está activo ({err_av})")
        log_error("recopilador", "SEG-001", "-", f"Bitdefender no está activo ({err_av})")
        return

    ya_procesados = set() if reprocesar else _cargar_procesados()

    cliente = GraphClient()
    todos   = cliente.obtener_correos_nuevos(procesados=ya_procesados, dir_temporal=_DIR_TEMPORAL, max_correos=max_correos)
    pendientes = [c for c in todos if c["id"] not in ya_procesados]
    correos = pendientes[:max_correos] if max_correos else pendientes

    if not correos:
        print("Sin correos por procesar.")
        return

    print(f"  {len(pendientes)} pendiente(s) de {len(todos)} total — procesando {len(correos)}.")

    correos_completados = []  # IDs de correos completamente procesados en esta sesión

    try:
        for correo in correos:
            asunto           = correo["asunto"]
            remitente        = correo["remitente"]
            nombre_remitente = correo.get("nombre_remitente", "")

            po_asunto, bl_asunto = extraer_po_y_bl_de_asunto(asunto)
            print(f"\nCorreo: {asunto} | De: {remitente}")
            if po_asunto:
                print(f"  Asunto -> PO: {po_asunto} | BL: {bl_asunto or 'no detectado'}")
            else:
                cuerpo = correo.get("cuerpo", "")
                if cuerpo:
                    po_cuerpo, bl_cuerpo = extraer_po_y_bl_de_asunto(cuerpo)
                    if po_cuerpo:
                        po_asunto = po_cuerpo
                        bl_asunto = bl_asunto or bl_cuerpo
                        print(f"  Cuerpo  -> PO: {po_asunto} | BL: {bl_asunto or 'no detectado'}")

            archivos_subidos = []
            ultimo_info      = None
            docs_con_inconsistencias = []  # acumula por correo para reporte consolidado
            docs_para_cruzado        = []  # acumula textos para validación cruzada
            proveedor_correo         = ""   # nombre de empresa detectado en los documentos del correo
            proveedor_correo_fallback = ""  # proveedor de documento que no es factura/packing

            # Verificar cadena de correos — detecta conflictos de PO/BL entre mensajes del hilo
            msg_com     = correo.get("_msg_obj")
            incs_cadena = verificar_cadena(msg_com, asunto, po_asunto, bl_asunto)
            if incs_cadena:
                log_evento("CADENA", "recopilador", archivo="-", remitente=remitente, asunto=asunto,
                           detalle=f"{len(incs_cadena)} inconsistencia(s) en cadena de correos")
                docs_con_inconsistencias.append({
                    "nombre_archivo": "cadena_correos",
                    "tipo":           "CADENA",
                    "inconsistencias": incs_cadena,
                })

            # Expandir ZIPs antes de procesar — genera lista plana de adjuntos
            adjuntos_expandidos = []
            for adjunto in correo.get("adjuntos", []):
                if adjunto["nombre"].lower().endswith(".zip"):
                    print(f"  Descomprimiendo: {adjunto['nombre']}")
                    internos, _ = _expandir_zip(adjunto["ruta_local"])
                    if internos:
                        adjuntos_expandidos.extend(internos)
                        print(f"    {len(internos)} archivo(s) extraídos del ZIP")
                elif adjunto["nombre"].lower().endswith(".rar"):
                    print(f"  Descomprimiendo: {adjunto['nombre']}")
                    internos, _ = _expandir_rar(adjunto["ruta_local"])
                    if internos:
                        adjuntos_expandidos.extend(internos)
                        print(f"    {len(internos)} archivo(s) extraídos del RAR")
                else:
                    adjuntos_expandidos.append(adjunto)

            # Correo basura: tenía adjuntos pero todos fueron filtrados (inline, bloqueados)
            if not adjuntos_expandidos and correo.get("adjuntos"):
                log_evento("BASURA", "recopilador", archivo="-", remitente=remitente, asunto=asunto,
                           detalle="todos los adjuntos están bloqueados o son imágenes inline")

            for adjunto in adjuntos_expandidos:
                nombre = adjunto["nombre"]
                print(f"  Procesando: {nombre}")

                seg = seguridad_adjunto(
                    ruta_local=adjunto.get("ruta_local", nombre),
                    nombre_orig=nombre,
                )
                if seg["resultado"] == "rechazado":
                    detalle_seg = f"{seg.get('codigo_error', '')} — {seg.get('descripcion', 'seguridad')}"
                    log_evento("ABORTADO", "recopilador", archivo=nombre, remitente=remitente, asunto=asunto, detalle=detalle_seg)
                    log_error("recopilador", seg.get("codigo_error", "SEG-XXX"), nombre, seg.get("descripcion", ""), remitente=remitente, asunto=asunto)
                    continue

                destinos = procesar_adjunto(nombre, numero_po_asunto=po_asunto, numero_bl=bl_asunto,
                                            ruta_local=adjunto.get("ruta_local"), asunto_correo=asunto)
                if not destinos:
                    log_evento("SKIP", "recopilador", archivo=nombre, remitente=remitente, asunto=asunto, detalle="sin clasificacion")
                    continue

                for info in destinos:
                    # Resolver nombre real de la carpeta OC en SharePoint
                    from clasificador import formatear_numero_oc
                    oc_base  = formatear_numero_oc(info["numero_po"])
                    oc_real  = cliente.buscar_carpeta_oc(oc_base)
                    ruta_real = info["ruta_sharepoint"].replace(oc_base, oc_real, 1)

                    certeza_str = f"{info['certeza']}%" if info.get("certeza") else "-"
                    print(f"    PO: {info['numero_po']} | BL: {info['numero_bl'] or '-'} "
                          f"| Tipo: {info['tipo']} | {info.get('metodo_clasificacion','?')} ({certeza_str}) "
                          f"| Ruta: {ruta_real}")

                    # Si el archivo viene de OTROS (o modo reprocesar) y ahora clasificó en carpeta
                    # correcta, no versionar — simplemente reemplazar. Borrar de OTROS al final.
                    from configuracion.ajustes import SHAREPOINT_CARPETA_OCS
                    viene_de_otros = info["tipo"] != "OTROS" and cliente.archivo_existe(
                        f"{SHAREPOINT_CARPETA_OCS}/{oc_real}/4. DOCUMENTACION/OTROS/{info['nombre_archivo']}"
                    )
                    # En modo normal, si el archivo ya existe en destino omitir subida
                    # (evita duplicados _v2 cuando el correo se reprocesa por cambio de EntryID)
                    ya_subido = (
                        not reprocesar
                        and not viene_de_otros
                        and cliente.archivo_existe(ruta_real)
                    )
                    if ya_subido:
                        print(f"    [YA EXISTE] {info['nombre_archivo']} ya está en SharePoint — omitiendo subida duplicada")
                    elif not viene_de_otros:
                        ruta_real = cliente.resolver_ruta_versionada(ruta_real)

                    # Registrar en eventos si se usó fallback
                    metodo = info.get("metodo_clasificacion", "")
                    if metodo == "fallback_nombre":
                        log_evento("FALLBACK", "recopilador", archivo=nombre, remitente=remitente, asunto=asunto,
                                   detalle=f"certeza baja o error API — clasificado por nombre como {info['tipo']}")
                    elif metodo == "fallback_contenido":
                        log_evento("FALLBACK", "recopilador", archivo=nombre, remitente=remitente, asunto=asunto,
                                   detalle=f"certeza baja o error API — clasificado por contenido como {info['tipo']}")

                    # Acumular nombre de proveedor — priorizar INVOICE o PACKING LIST
                    _prov = info.get("nombre_proveedor", "")
                    if _prov:
                        if info["tipo"] in ("INVOICE", "PACKING LIST", "PL + INV"):
                            if not proveedor_correo:
                                proveedor_correo = _prov
                        elif not proveedor_correo_fallback:
                            proveedor_correo_fallback = _prov

                    # Acumular texto para validación cruzada (solo si Claude clasificó y hay texto)
                    texto_doc = info.get("texto_extraido", "")
                    if texto_doc and info.get("metodo_clasificacion") == "claude":
                        docs_para_cruzado.append({
                            "tipo":           info["tipo"],
                            "nombre_archivo": info["nombre_archivo"],
                            "texto":          texto_doc,
                        })

                    # Acumular inconsistencias — el reporte se genera al final del correo
                    if info.get("inconsistencias"):
                        docs_con_inconsistencias.append({
                            "nombre_archivo": info["nombre_archivo"],
                            "tipo":           info["tipo"],
                            "inconsistencias": info["inconsistencias"],
                        })

                    _prov_log = proveedor_correo or proveedor_correo_fallback
                    if ya_subido:
                        _log_subida(remitente, ruta_real, ya_existia=True,
                                    nombre_remitente=nombre_remitente, proveedor=_prov_log)
                    else:
                        cliente.crear_carpeta_si_no_existe(ruta_real)
                        ruta_a_subir = info.get("ruta_local") or adjunto.get("ruta_local", nombre)
                        cliente.subir_archivo(ruta_a_subir, ruta_real)
                        _log_subida(remitente, ruta_real,
                                    nombre_remitente=nombre_remitente, proveedor=_prov_log)
                    archivos_subidos.append(info["tipo"])
                    ultimo_info = {**info, "ruta_sharepoint": ruta_real}


            # Validación cruzada entre documentos del mismo correo
            if len(docs_para_cruzado) >= 2:
                print(f"\n  [CRUZADO] Validando consistencia entre {len(docs_para_cruzado)} documento(s)...")
                incs_cruzadas = detectar_inconsistencias_cruzadas(docs_para_cruzado)
                if incs_cruzadas:
                    docs_con_inconsistencias.append({
                        "nombre_archivo": "— validación cruzada entre documentos —",
                        "tipo":           "CRUZADO",
                        "inconsistencias": incs_cruzadas,
                    })

            # Generar reporte consolidado por correo si hubo al menos una inconsistencia
            if docs_con_inconsistencias:
                ruta_cons = _generar_txt_consolidado(
                    remitente=remitente,
                    asunto=asunto,
                    documentos=docs_con_inconsistencias,
                    msg_obj=msg_com,
                    numero_po=po_asunto or "",
                    nombre_proveedor=proveedor_correo or proveedor_correo_fallback,
                )
                total_incs = sum(len(d['inconsistencias']) for d in docs_con_inconsistencias)
                _log_inconsistencia(
                    remitente=remitente,
                    asunto=asunto,
                    adjunto="(consolidado)",
                    tipo=f"{len(docs_con_inconsistencias)} documento(s)",
                    resumen=f"reporte consolidado — {total_incs} inconsistencia(s) total",
                    ruta_txt=str(ruta_cons),
                )

            if archivos_subidos and ultimo_info:
                _marcar_procesado(correo["id"])
                correos_completados.append(correo["id"])

    except KeyboardInterrupt:
        print(f"\n[!] Cancelado por el usuario — {len(correos_completados)} correo(s) ya marcados como procesados.")
        raise

    print("\nCiclo completado.")


if __name__ == "__main__":
    modo = "MOCK" if MODO_MOCK else "REAL"
    print(f"Clasificador de Documentos iniciado [{modo}]")

    # Uso: python recopilador_documentos.py [loop|reprocesar] [N]
    #   loop        → modo continuo cada INTERVALO_SEGUNDOS
    #   reprocesar  → ignora procesados.txt, sube todo y limpia OTROS
    #   N           → límite de correos a procesar
    args = sys.argv[1:]
    es_loop        = "loop" in args
    es_reprocesar  = "reprocesar" in args
    nums           = [a for a in args if a.isdigit()]
    max_correos    = int(nums[0]) if nums else None

    if es_reprocesar:
        print("Modo REPROCESAR — se ignorará procesados.txt y se limpiará OTROS en SharePoint.")
        print("Este proceso puede tardar varios minutos según el volumen de correos.\n")
        ejecutar_ciclo(max_correos, reprocesar=True)
    elif es_loop:
        print(f"Modo continuo — intervalo: {INTERVALO_SEGUNDOS}s | Detener con Ctrl+C")
        while True:
            try:
                ejecutar_ciclo(max_correos)
            except Exception as e:
                print(f"[ERROR] Ciclo fallido: {e}")
                log_error("recopilador", "RECP-004", "-", f"Ciclo fallido: {type(e).__name__}: {e}")
            try:
                time.sleep(INTERVALO_SEGUNDOS)
            except KeyboardInterrupt:
                print("\nDetenido por el usuario.")
                sys.exit(0)
    else:
        ejecutar_ciclo(max_correos)

        