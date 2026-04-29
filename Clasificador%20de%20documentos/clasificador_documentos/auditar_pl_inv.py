# -*- coding: utf-8 -*-
"""
auditar_pl_inv.py — Audita todos los correos de la bandeja buscando Excel
que sean PL+INV (packing list + invoice en el mismo archivo), verifica si
están correctamente subidos en SharePoint y sube los que faltan.

Estrategia de detección:
  - Solo archivos .xlsx / .xls
  - Descarga el Excel y lee sus nombres de hoja
  - Se considera PL+INV si hay ≥1 hoja con keywords de PL y ≥1 con keywords de INV

Para cada PL+INV encontrado sube el mismo archivo a DOS carpetas:
  - 4.27 Packing list definitivo  (con prefijo PL_)
  - 4.02 Factura Definitiva       (con prefijo INV_)

Uso:
    cd clasificador_documentos
    python auditar_pl_inv.py

Salida: imprime por consola y genera auditoria_pl_inv_YYYY-MM-DD.log
"""

import sys
import io
import re
import shutil
import tempfile
import logging
import pythoncom
import win32com.client
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent))

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from configuracion.ajustes import (
    OUTLOOK_EMAIL,
    SHAREPOINT_DRIVE_ID,
    SHAREPOINT_CARPETA_OCS,
    MAPA_CARPETAS,
    MAPA_CARPETAS_BORRADOR,
)
from Nucleo.clasificador import formatear_numero_oc, extraer_numero_po
from Integraciones.graph_client import GraphClient

# ---------------------------------------------------------------------------
# Keywords para identificar hojas PL e INV dentro de un Excel
# ---------------------------------------------------------------------------
_KW_PL = re.compile(
    r"pack|packing|p\.?l\.?|lista\s*empaque|embalaje",
    re.IGNORECASE,
)
_KW_INV = re.compile(
    r"inv|invoice|factura|commercial",
    re.IGNORECASE,
)

# Solo nos interesan Excel
_EXTS_EXCEL = {".xlsx", ".xls"}

# Días hacia atrás a escanear (None = toda la bandeja)
DIAS_ATRAS = None  # cambia a un número para limitar, ej: 180

# ---------------------------------------------------------------------------
# Logging del reporte
# ---------------------------------------------------------------------------
_FECHA_HOY = datetime.now().strftime("%Y-%m-%d")
_LOG_PATH = Path(__file__).parent / f"auditoria_pl_inv_{_FECHA_HOY}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(_LOG_PATH, encoding="utf-8"),
    ],
)
log = logging.getLogger("auditoria")

# Logger de subidas — mismo formato que el monitor principal
_LOG_SUBIDAS = Path(__file__).parent / "registros_subidas.log"
_logger_subidas = logging.getLogger("subidas_sharepoint")
if not _logger_subidas.handlers:
    _h = logging.FileHandler(_LOG_SUBIDAS, encoding="utf-8")
    _h.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_h)
    _logger_subidas.setLevel(logging.INFO)


def _log_subida(remitente: str, ruta_sharepoint: str):
    ruta_completa = ruta_sharepoint.replace("\\", "/")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines() if _LOG_SUBIDAS.exists() else []
    n = len([l for l in lines if l.strip()]) + 1
    _logger_subidas.info(f"{n} | {ts} | {remitente} | {ruta_completa}")


# ---------------------------------------------------------------------------
# Detección por hojas del Excel
# ---------------------------------------------------------------------------

_tmp_files: list[str] = []   # para limpiar al final


def _hojas_son_pl_inv(ruta_excel: str) -> tuple[bool, list[str]]:
    """Retorna (es_pl_inv, lista_hojas)."""
    hojas = []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
        hojas = list(wb.sheetnames)
        wb.close()
    except Exception:
        try:
            import xlrd
            wb = xlrd.open_workbook(ruta_excel)
            hojas = list(wb.sheet_names())
        except Exception:
            return False, []

    tiene_pl  = any(_KW_PL.search(h)  for h in hojas)
    tiene_inv = any(_KW_INV.search(h) for h in hojas)
    return (tiene_pl and tiene_inv), hojas


def _escanear_adjuntos_del_mensaje(msg) -> list[dict]:
    """
    Descarga todos los adjuntos Excel, inspecciona sus hojas.
    Retorna lista de {nombre, ruta_tmp, hojas} para los que son PL+INV.
    """
    resultados = []
    try:
        count = msg.Attachments.Count
    except Exception:
        return resultados

    for i in range(1, count + 1):
        try:
            adj = msg.Attachments.Item(i)
            nombre = adj.FileName or ""
            ext = Path(nombre).suffix.lower()
            if ext not in _EXTS_EXCEL:
                continue

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext, prefix="audit_adj_")
            tmp.close()
            _tmp_files.append(tmp.name)

            try:
                adj.SaveAsFile(tmp.name)
            except Exception as e:
                log.info(f"  [WARN] No se pudo descargar '{nombre}': {e}")
                continue

            es_pl_inv, hojas = _hojas_son_pl_inv(tmp.name)
            hojas_str = ", ".join(hojas) if hojas else "(sin hojas)"

            if es_pl_inv:
                resultados.append({"nombre": nombre, "ruta_tmp": tmp.name, "hojas": hojas})
                log.info(f"  [PL+INV detectado] '{nombre}'  hojas: {hojas_str}")
            else:
                log.info(f"  [ignorado] '{nombre}'  hojas: {hojas_str}")

        except Exception:
            pass

    return resultados


# ---------------------------------------------------------------------------
# Verificación en SharePoint
# ---------------------------------------------------------------------------

def _listar_archivos_carpeta(cliente: GraphClient, ruta_carpeta: str) -> list[str]:
    try:
        ruta_enc = quote(ruta_carpeta, safe="/")
        url = (
            f"https://graph.microsoft.com/v1.0/drives/{SHAREPOINT_DRIVE_ID}"
            f"/root:/{ruta_enc}:/children?$select=name&$top=200"
        )
        data = cliente._get(url)
        return [item["name"] for item in data.get("value", [])]
    except Exception:
        return []


def _coincide_nombre(lista: list[str], stem: str) -> bool:
    for f in lista:
        fn = Path(f).stem.upper()
        if fn == stem:
            return True
        if fn in (f"PL_{stem}", f"INV_{stem}"):
            return True
        if stem in fn or fn in stem:
            return True
    return False


def _verificar_en_sharepoint(cliente: GraphClient, numero_oc: str, nombre_adj: str) -> dict:
    oc_real = cliente.buscar_carpeta_oc(formatear_numero_oc(numero_oc))
    base = f"{SHAREPOINT_CARPETA_OCS}/{oc_real}/4. DOCUMENTACION"

    carpeta_pl   = f"{base}/{MAPA_CARPETAS['PACKING LIST']}"
    carpeta_inv  = f"{base}/{MAPA_CARPETAS['INVOICE']}"
    carpeta_borr = f"{base}/{MAPA_CARPETAS_BORRADOR.get('INVOICE', '')}"

    archivos_pl   = _listar_archivos_carpeta(cliente, carpeta_pl)
    archivos_inv  = _listar_archivos_carpeta(cliente, carpeta_inv)
    archivos_borr = (
        _listar_archivos_carpeta(cliente, carpeta_borr)
        if MAPA_CARPETAS_BORRADOR.get("INVOICE") else []
    )

    stem = Path(nombre_adj).stem.upper()
    pl_ok  = _coincide_nombre(archivos_pl, stem)
    inv_ok = _coincide_nombre(archivos_inv, stem) or _coincide_nombre(archivos_borr, stem)

    return {
        "oc_real":     oc_real,
        "carpeta_pl":  carpeta_pl,
        "carpeta_inv": carpeta_inv,
        "archivos_pl": archivos_pl,
        "archivos_inv": archivos_inv + archivos_borr,
        "pl_ok":       pl_ok,
        "inv_ok":      inv_ok,
    }


# ---------------------------------------------------------------------------
# Subida de faltantes
# ---------------------------------------------------------------------------

def _subir_faltantes(cliente: GraphClient, nombre_adj: str,
                     ruta_tmp: str, remitente: str, v: dict):
    """
    Sube el archivo a las carpetas que faltan (PL y/o INV).
    Usa las rutas de carpeta que ya vienen de _verificar_en_sharepoint,
    donde oc_real es el nombre exacto de la carpeta en SharePoint.
    """
    ext  = Path(nombre_adj).suffix
    stem = Path(nombre_adj).stem
    subidos = []

    # --- Carpeta PL ---
    if not v["pl_ok"]:
        nombre_pl  = f"PL_{stem}{ext}"
        ruta_sp_pl = f"{v['carpeta_pl']}/{nombre_pl}"
        try:
            cliente.crear_carpeta_si_no_existe(ruta_sp_pl)
            cliente.subir_archivo(ruta_tmp, ruta_sp_pl)
            _log_subida(remitente, ruta_sp_pl)
            log.info(f"  → [SUBIDO PL]  {ruta_sp_pl}")
            subidos.append("PL")
        except Exception as e:
            log.info(f"  → [ERROR subiendo PL] {e}")

    # --- Carpeta INV ---
    if not v["inv_ok"]:
        nombre_inv  = f"INV_{stem}{ext}"
        ruta_sp_inv = f"{v['carpeta_inv']}/{nombre_inv}"
        try:
            cliente.crear_carpeta_si_no_existe(ruta_sp_inv)
            cliente.subir_archivo(ruta_tmp, ruta_sp_inv)
            _log_subida(remitente, ruta_sp_inv)
            log.info(f"  → [SUBIDO INV] {ruta_sp_inv}")
            subidos.append("INV")
        except Exception as e:
            log.info(f"  → [ERROR subiendo INV] {e}")

    return subidos


# ---------------------------------------------------------------------------
# Búsqueda de bandeja Outlook
# ---------------------------------------------------------------------------

def _buscar_bandeja(namespace):
    email_lower = OUTLOOK_EMAIL.lower()
    for store in namespace.Stores:
        try:
            if email_lower in store.DisplayName.lower():
                return store.GetDefaultFolder(6)
        except Exception:
            continue
    try:
        for cuenta in namespace.Accounts:
            try:
                if cuenta.SmtpAddress.lower() == email_lower:
                    return cuenta.DeliveryStore.GetDefaultFolder(6)
            except Exception:
                continue
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Auditoría principal
# ---------------------------------------------------------------------------

def auditar():
    pythoncom.CoInitialize()
    resultados_ok     = []
    resultados_subidos = []
    resultados_error   = []
    resultados_sin_oc  = []
    total_revisados    = 0

    try:
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            log.error("Outlook no está abierto. Ábrelo antes de ejecutar este script.")
            return

        namespace = outlook.GetNamespace("MAPI")
        bandeja = _buscar_bandeja(namespace)
        if bandeja is None:
            log.error(f"No se encontró la bandeja para {OUTLOOK_EMAIL}")
            return

        log.info(f"{'='*70}")
        log.info(f"AUDITORÍA + SUBIDA PL+INV — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        log.info(f"Cuenta: {OUTLOOK_EMAIL}")
        log.info(f"{'='*70}\n")

        mensajes = bandeja.Items
        mensajes.Sort("[ReceivedTime]", True)

        if DIAS_ATRAS:
            fecha_limite = (datetime.now() - timedelta(days=DIAS_ATRAS)).strftime("%m/%d/%Y")
            mensajes = mensajes.Restrict(f"[ReceivedTime] >= '{fecha_limite}'")
            mensajes.Sort("[ReceivedTime]", True)
            log.info(f"Período: últimos {DIAS_ATRAS} días\n")
        else:
            log.info("Período: TODA LA BANDEJA\n")

        cliente = GraphClient()
        total_mensajes = mensajes.Count
        log.info(f"Total mensajes en bandeja: {total_mensajes}\n")

        for idx in range(1, total_mensajes + 1):
            try:
                msg = mensajes[idx]
            except Exception:
                continue

            asunto = remitente = fecha = ""
            try:
                asunto    = msg.Subject or ""
                remitente = msg.SenderEmailAddress or ""
                fecha     = msg.ReceivedTime.strftime("%Y-%m-%d")
            except Exception:
                pass

            try:
                adjuntos = _escanear_adjuntos_del_mensaje(msg)
            except Exception:
                continue

            if not adjuntos:
                continue

            for item in adjuntos:
                nombre_adj = item["nombre"]
                ruta_tmp   = item["ruta_tmp"]
                hojas      = item["hojas"]
                total_revisados += 1

                log.info(f"\n[DETECTADO]  {fecha}  {nombre_adj}  hojas: {', '.join(hojas)}")
                log.info(f"  Remitente: {remitente}  |  Asunto: {asunto[:70]}")

                numero_oc = extraer_numero_po(asunto) or extraer_numero_po(nombre_adj)

                if not numero_oc:
                    resultados_sin_oc.append({"fecha": fecha, "adjunto": nombre_adj, "asunto": asunto[:80]})
                    log.info("  → [SIN OC] No se pudo extraer número de OC — omitido")
                    continue

                try:
                    v = _verificar_en_sharepoint(cliente, numero_oc, nombre_adj)
                except Exception as e:
                    resultados_error.append({"fecha": fecha, "adjunto": nombre_adj, "numero_oc": numero_oc, "error": str(e)})
                    log.info(f"  → [ERROR verificando SharePoint] {e}")
                    continue

                pl_ok  = v["pl_ok"]
                inv_ok = v["inv_ok"]

                if pl_ok and inv_ok:
                    resultados_ok.append({"fecha": fecha, "adjunto": nombre_adj, "numero_oc": numero_oc})
                    log.info(f"  → [OK]  OC {numero_oc}  PL✓  INV✓  (ya estaba subido)")
                    continue

                # Hay algo que falta — subir
                pl_str  = "PL✓" if pl_ok  else "PL✗"
                inv_str = "INV✓" if inv_ok else "INV✗"
                log.info(f"  → [FALTA]  OC {numero_oc}  {pl_str}  {inv_str}  — subiendo...")

                subidos = _subir_faltantes(cliente, nombre_adj, ruta_tmp, remitente, v)

                if subidos:
                    resultados_subidos.append({
                        "fecha":     fecha,
                        "adjunto":   nombre_adj,
                        "numero_oc": numero_oc,
                        "subidos":   subidos,
                    })
                else:
                    resultados_error.append({
                        "fecha":     fecha,
                        "adjunto":   nombre_adj,
                        "numero_oc": numero_oc,
                        "error":     "no se pudo subir ninguna carpeta",
                    })

    finally:
        pythoncom.CoUninitialize()
        for f in _tmp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass

    # ---------------------------------------------------------------------------
    # Resumen final
    # ---------------------------------------------------------------------------
    log.info(f"\n{'='*70}")
    log.info("RESUMEN")
    log.info(f"{'='*70}")
    log.info(f"Total Excel PL+INV encontrados   : {total_revisados}")
    log.info(f"  Ya estaban correctos            : {len(resultados_ok)}")
    log.info(f"  Subidos ahora                   : {len(resultados_subidos)}")
    log.info(f"  Con error                       : {len(resultados_error)}")
    log.info(f"  Sin número de OC                : {len(resultados_sin_oc)}")

    if resultados_subidos:
        log.info(f"\n{'─'*70}")
        log.info("SUBIDOS EN ESTA EJECUCIÓN:")
        log.info(f"{'─'*70}")
        for r in resultados_subidos:
            carpetas = " + ".join(r["subidos"])
            log.info(f"  OC {r['numero_oc']:>10}  |  {r['fecha']}  |  {r['adjunto'][:45]}  → {carpetas}")

    if resultados_error:
        log.info(f"\n{'─'*70}")
        log.info("ERRORES (revisar manualmente):")
        log.info(f"{'─'*70}")
        for r in resultados_error:
            log.info(f"  OC {r.get('numero_oc','?'):>10}  |  {r['fecha']}  |  {r['adjunto'][:45]}  — {r.get('error','')}")

    if resultados_sin_oc:
        log.info(f"\n{'─'*70}")
        log.info("SIN OC IDENTIFICADA (revisar manualmente):")
        log.info(f"{'─'*70}")
        for r in resultados_sin_oc:
            log.info(f"  {r['fecha']}  {r['adjunto'][:50]:<50}  asunto: {r['asunto'][:60]}")

    log.info(f"\nReporte guardado en: {_LOG_PATH}")


if __name__ == "__main__":
    auditar()
