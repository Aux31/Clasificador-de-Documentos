"""
Orquestador de clasificación de documentos.
- Clasificador principal : Claude API (clasificador_claude.py)
- Fallback               : keywords en nombre del archivo y contenido del documento
- Extracción de texto    : extractor_texto.py
- Catálogo de tipos      : catalogo_tipos.py

Exporta la misma interfaz pública que antes para no romper
recopilador_documentos.py ni monitor_correos.py.
"""

import re
import logging
from pathlib import Path

from Utilidades.logger_errores import log_error, log_advertencia

_logger_clas = logging.getLogger("clasificador")
from configuracion.ajustes import (
    MAPA_CARPETAS,
    MAPA_CARPETAS_BORRADOR,
    PREFIJOS_TIPO,
    SHAREPOINT_CARPETA_OCS,
    CLAUDE_CERTEZA_MINIMA,
    CLAUDE_CERTEZA_POR_TIPO,
)
from Nucleo.clasificador_claude import clasificar_con_claude, es_fallo
from Nucleo.separador_fullset import separar_fullset
from Nucleo.validador_campos import validar_campos

# ---------------------------------------------------------------------------
# Expresiones regulares para PO y BL
# ---------------------------------------------------------------------------

# OC seguido de guion/espacios y dígitos: 'OC-00194386', 'OC-196893'
# PO seguido de separadores y dígitos:    'PO 196893', 'PO#196893', 'PO NO.: 197341'
_PATRON_PO = re.compile(
    r'OC[-\s]?0*(\d{4,10})|PO[\s#\-_.]*(?:NO\.?[\s:]*)?(?:OC[-\s]?0*)?(\d{4,10})|#(\d{4,10})',
    re.IGNORECASE,
)

# BL MAEU9876543, BL# 265008831, BL-XXXXX
_PATRON_BL = re.compile(r'BL[\s#\-_]+([A-Z0-9]{4,20})', re.IGNORECASE)


# ---------------------------------------------------------------------------
# Extracción de PO / BL del asunto o nombre del archivo
# ---------------------------------------------------------------------------

def extraer_numero_po(texto: str) -> str | None:
    """
    Extrae el número de PO de un nombre de archivo o del asunto del correo.
    Soporta formatos: 'PO 196893', 'PO#196893', 'PO-196893', 'OC-00194386'.

    Returns:
        Número de PO como string, o None si no se encuentra.
    """
    stem = Path(texto).stem if "." in texto else texto
    match = _PATRON_PO.search(stem)
    return (match.group(1) or match.group(2) or match.group(3)) if match else None


def extraer_numero_bl(asunto: str) -> str | None:
    """
    Extrae el número de BL del asunto del correo.
    Soporta: 'BL XXXXX', 'BL# XXXXX', 'BL#XXXXX'.

    Returns:
        Número de BL como string, o None si no se encuentra.
    """
    match = _PATRON_BL.search(asunto)
    return match.group(1) if match else None


def extraer_po_y_bl_de_asunto(asunto: str) -> tuple[str | None, str | None]:
    """
    Extrae PO y BL del asunto del correo en un solo paso.

    Returns:
        Tupla (numero_po, numero_bl), cualquiera puede ser None.
    """
    match = _PATRON_PO.search(asunto)
    numero_po = (match.group(1) or match.group(2) or match.group(3)) if match else None
    return numero_po, extraer_numero_bl(asunto)


# ---------------------------------------------------------------------------
# Fallback: clasificación por keywords en el nombre del archivo
# ---------------------------------------------------------------------------

# Alias para detectar tipos por el nombre del archivo
_BL_ALIAS   = {"BILLOFLADING", "BILL OF LADING", "BILL_OF_LADING", "BILL LADING",
                "OCEAN BILL", "SEA WAYBILL", "SEAWAYBILL", "CONOCIMIENTO DE EMBARQUE",
                "ORIGINAL BL", "COPY BL", "B/L"}
_PL_ALIAS   = {"PACKLIST", "PACKING SLIP", "PACKING DETAILS", "LISTA DE EMPAQUE", "DETALLE DE EMPAQUE"}
_CO_ALIAS   = {"CERTIFICATE OF ORIGIN", "COO", "CERTIFICADO DE ORIGEN", "ORIGIN CERTIFICATE", "EUR1", "EUR.1", "EUR 1"}
_WC_ALIAS   = {"WEIGHT CERTIFICATE", "CERTIFICADO DE PESO"}
_QC_ALIAS   = {"QUALITY CERTIFICATE", "CERTIFICATE OF QUALITY",
                "INSPECTION CERTIFICATE", "TEST REPORT", "QC REPORT"}
_INV_ALIAS  = {"COMMERCIAL INVOICE", "TAX INVOICE", "FACTURA COMERCIAL"}
_FITO_ALIAS = {"PHYTOSANITARY", "PHYTO CERTIFICATE", "PLANT HEALTH", "CERTIFICADO FITOSANITARIO",
                "FITOSANITARIO"}
_ZOO_ALIAS  = {"ZOOSANITARIO", "ZOOSANITARY", "CERTIFICADO VETERINARIO",
                "VETERINARY CERTIFICATE", "ANIMAL HEALTH"}
_EXO_ALIAS  = {"EXONERACION", "EXONERACIÓN", "HACIENDA", "DGT"}
_MAR_ALIAS  = {"MARCHAMO", "PRECINTO", "SEAL PHOTO", "FOTO MARCHAMO", "FOTO PRECINTO"}

# Palabras clave para clasificación por CONTENIDO del documento
_PALABRAS_CLAVE_CONTENIDO: list[tuple[str, list[str]]] = [
    ("MBL",               ["MASTER BILL OF LADING", "MASTERBILLOFLADING", "MASTER BL",
                            "MASTER B/L", "MBL NO", "MBL NUMBER"]),
    ("HBL",               ["HOUSE BILL OF LADING", "HOUSEBILLOFLADING", "HOUSE BL",
                            "HOUSE B/L", "HBL NO", "HBL NUMBER"]),
    ("BL",                ["BILL OF LADING", "BILLOFLADING", "BILL LADING", "OCEAN BILL",
                            "B/L", "SEA WAYBILL", "SEAWAYBILL", "ORIGINAL BL", "ORIGINAL B/L",
                            "COPY BL", "CONOCIMIENTO DE EMBARQUE"]),
    ("AWB",               ["AIR WAYBILL", "AIRWAYBILL", "AIRWAY BILL", "AIR CARGO WAYBILL",
                            "AWB NO", "AWB NUMBER", "GUIA AEREA", "GUÍA AÉREA"]),
    ("INVOICE",           ["COMMERCIAL INVOICE", "TAX INVOICE", "INVOICE NO", "INVOICE NUMBER",
                            "FACTURA COMERCIAL", "FACTURA", "BILLING"]),
    ("WEIGHT CERTIFICATE",  ["WEIGHT CERTIFICATE", "CERTIFICADO DE PESO"]),
    ("QUALITY CERTIFICATE", ["QUALITY CERTIFICATE", "CERTIFICATE OF QUALITY",
                              "INSPECTION CERTIFICATE", "TEST REPORT", "QC REPORT", "QUALITY REPORT"]),
    ("PACKING LIST",      ["PACKING LIST", "PACKLIST", "PACKING DETAILS", "PACKING SLIP",
                            "LISTA DE EMPAQUE", "DETALLE DE EMPAQUE"]),
    ("CO",                ["CERTIFICATE OF ORIGIN", "CERTIFICATEOF ORIGIN", "CERTIFICATEOFORIGIN",
                            "COO", "CERTIFICADO DE ORIGEN", "ORIGIN CERTIFICATE", "COUNTRY OF ORIGIN",
                            "EUR.1", "EUR1", "EUR 1", "FORM A", "GSP CERTIFICATE"]),
    ("FITOSANITARIO",     ["CERTIFICADO FITOSANITARIO", "PHYTOSANITARY CERTIFICATE", "PHYTO CERTIFICATE",
                            "PLANT HEALTH CERTIFICATE", "FITOSANITARIO"]),
    ("ZOOSANITARIO",      ["CERTIFICADO ZOOSANITARIO", "ZOOSANITARY CERTIFICATE",
                            "CERTIFICADO VETERINARIO", "VETERINARY CERTIFICATE",
                            "ANIMAL HEALTH CERTIFICATE"]),
    ("PRINTER",           ["PRINTER DE AGENCIA", "DOCUMENTO DE AGENCIA", "AGENCIA DE ADUANA",
                            "CUSTOMS AGENCY", "AGENCIA ADUANAL"]),
    ("EXONERACION",       ["EXONERACION DE HACIENDA", "EXONERACIÓN DE HACIENDA", "EXONERACION DGT",
                            "DIRECCIÓN GENERAL DE TRIBUTACIÓN", "DIRECCION GENERAL DE TRIBUTACION",
                            "TAX EXEMPTION"]),
    ("MARCHAMO",          ["MARCHAMO", "PRECINTO", "SEAL NUMBER", "CONTAINER SEAL",
                            "SEAL PHOTO", "FOTO MARCHAMO", "FOTO PRECINTO"]),
]


def _clasificar_fallback_nombre(nombre_archivo: str) -> list[str]:
    """
    Detecta tipos de documento por keywords en el nombre del archivo.
    Permite múltiples tipos (ej. 'PL + INV' → Packing List + Factura).

    Returns:
        Lista de tipos encontrados, o ['OTROS'] si no coincide ninguno.
    """
    nombre_upper = nombre_archivo.upper()

    # Tipos que deben buscarse con límites de palabra para evitar falsos positivos
    # (ej: "BL" dentro de "VARIABLES", "AWB" dentro de "DRAWBACK", etc.)
    _TIPOS_PALABRA_COMPLETA = {"BL", "MBL", "HBL", "AWB", "CO"}

    tipos = []
    for t in MAPA_CARPETAS:
        t_upper = t.upper()
        if t_upper in _TIPOS_PALABRA_COMPLETA:
            if re.search(rf'(?<![A-Z]){re.escape(t_upper)}(?![A-Z])', nombre_upper):
                tipos.append(t)
        else:
            if t_upper in nombre_upper:
                tipos.append(t)

    # Alias adicionales (nombres alternativos no cubiertos por el nombre exacto del tipo)
    if "BL" not in tipos and "MBL" not in tipos and "HBL" not in tipos:
        if any(a in nombre_upper for a in _BL_ALIAS):
            tipos.append("BL")
    if "PACKING LIST" not in tipos and any(a in nombre_upper for a in _PL_ALIAS):
        tipos.append("PACKING LIST")
    if "CO" not in tipos and any(a in nombre_upper for a in _CO_ALIAS):
        tipos.append("CO")
    if "WEIGHT CERTIFICATE" not in tipos and any(a in nombre_upper for a in _WC_ALIAS):
        tipos.append("WEIGHT CERTIFICATE")
    if "QUALITY CERTIFICATE" not in tipos and any(a in nombre_upper for a in _QC_ALIAS):
        tipos.append("QUALITY CERTIFICATE")
    if "INVOICE" not in tipos and any(a in nombre_upper for a in _INV_ALIAS):
        tipos.append("INVOICE")
    if "FITOSANITARIO" not in tipos and any(a in nombre_upper for a in _FITO_ALIAS):
        tipos.append("FITOSANITARIO")
    if "ZOOSANITARIO" not in tipos and any(a in nombre_upper for a in _ZOO_ALIAS):
        tipos.append("ZOOSANITARIO")
    if "EXONERACION" not in tipos and any(a in nombre_upper for a in _EXO_ALIAS):
        tipos.append("EXONERACION")
    if "MARCHAMO" not in tipos and any(a in nombre_upper for a in _MAR_ALIAS):
        tipos.append("MARCHAMO")

    # Combinar PL + INV en tipo especial si ambos están presentes
    if "PL + INV" in tipos:
        tipos.remove("PL + INV")
        if "PACKING LIST" not in tipos:
            tipos.append("PACKING LIST")
        if "INVOICE" not in tipos:
            tipos.append("INVOICE")

    return tipos if tipos else ["OTROS"]


def _clasificar_fallback_contenido(texto_upper: str) -> str:
    """
    Detecta el tipo de documento por keywords en el contenido del texto.

    Args:
        texto_upper: Texto del documento ya convertido a mayúsculas.

    Returns:
        Tipo detectado, o 'OTROS' si no coincide ninguno.
    """
    for tipo, palabras in _PALABRAS_CLAVE_CONTENIDO:
        if any(p in texto_upper for p in palabras):
            return tipo
    return "OTROS"


# ---------------------------------------------------------------------------
# Funciones auxiliares de rutas y nombres
# ---------------------------------------------------------------------------

def renombrar_para_tipo(nombre_archivo: str, tipo: str, tipos_totales: list[str]) -> str:
    """
    Renombra el archivo cuando viene de un combinado (ej: 'PL + INV').
    Solo actúa si hay más de un destino y el tipo tiene prefijo definido.
    """
    if len(tipos_totales) <= 1 or tipo not in PREFIJOS_TIPO:
        return nombre_archivo
    p = Path(nombre_archivo)
    prefijo = PREFIJOS_TIPO[tipo]
    stem = p.stem
    if " OF " in stem.upper():
        resto = stem[stem.upper().index(" OF "):]
    else:
        resto = " " + stem
    return f"{prefijo}{resto}{p.suffix}"


def formatear_numero_oc(numero_po: str) -> str:
    """Convierte '196893' → 'cmer-OC-00196893' (siempre 8 dígitos)."""
    return f"cmer-OC-{numero_po.zfill(8)}"


_PATRON_BORRADOR_NOMBRE = re.compile(
    r'\b(draft|borrador|borr|preliminary|non.negotiable|specimen|copy)\b',
    re.IGNORECASE,
)


def es_borrador_por_nombre(nombre_archivo: str) -> bool:
    """Retorna True si el nombre del archivo indica que es un borrador."""
    return bool(_PATRON_BORRADOR_NOMBRE.search(Path(nombre_archivo).stem))


def construir_ruta(numero_po: str, tipo: str, nombre_archivo: str,
                   es_borrador: bool = False) -> str:
    """
    Construye la ruta relativa completa en SharePoint.
    Si es_borrador=True y el tipo tiene carpeta de borrador, usa esa carpeta.

    Returns:
        Ej: "OC´s/cmer-OC-00196893/4. DOCUMENTACION/4.05 BL-AWB-Porte definitivo/BL PO 196893.pdf"
    """
    if es_borrador and tipo in MAPA_CARPETAS_BORRADOR:
        carpeta = MAPA_CARPETAS_BORRADOR[tipo]
    else:
        carpeta = MAPA_CARPETAS.get(tipo, "OTROS")
    oc = formatear_numero_oc(numero_po)
    return f"{SHAREPOINT_CARPETA_OCS}/{oc}/4. DOCUMENTACION/{carpeta}/{nombre_archivo}"


# ---------------------------------------------------------------------------
# Funciones públicas de compatibilidad (usadas en tests o código externo)
# ---------------------------------------------------------------------------

def clasificar_tipos(nombre_archivo: str) -> list[str]:
    """Fallback por nombre — retorna lista de tipos."""
    return _clasificar_fallback_nombre(nombre_archivo)


def clasificar_tipo(nombre_archivo: str) -> str:
    """Retorna el primer tipo detectado por nombre (compatibilidad)."""
    return _clasificar_fallback_nombre(nombre_archivo)[0]


# ---------------------------------------------------------------------------
# Función principal del pipeline
# ---------------------------------------------------------------------------

def procesar_adjunto(
    nombre_archivo: str,
    numero_po_asunto: str | None = None,
    numero_bl: str | None = None,
    ruta_local: str | None = None,
    asunto_correo: str = "",
    **kwargs,
) -> list[dict] | None:
    """
    Procesa un adjunto y retorna la información de clasificación completa.

    Flujo de clasificación:
      1. Extrae texto del documento (si hay ruta_local).
      2. Clasifica con Claude API.
      3. Si Claude falla o retorna certeza < CLAUDE_CERTEZA_MINIMA:
         → fallback por keywords en nombre del archivo.
         → si el resultado sigue siendo OTROS y hay texto:
            → fallback por keywords en contenido.

    Args:
        nombre_archivo   : Nombre del archivo adjunto.
        numero_po_asunto : Número de PO extraído del asunto (tiene prioridad).
        numero_bl        : Número de BL extraído del asunto.
        ruta_local       : Ruta al archivo en disco (para extracción de texto).
        asunto_correo    : Asunto del correo (contexto adicional para Claude).

    Returns:
        Lista de dicts (uno por destino) con:
          - nombre_archivo, nombre_destino, numero_po, numero_bl, tipo,
            ruta_sharepoint, certeza, metodo_clasificacion
        Retorna None si no se puede extraer el número de PO.
    """
    # --- 1. PO es obligatorio (tres intentos antes de llegar a Claude) ---
    numero_po = numero_po_asunto or extraer_numero_po(nombre_archivo)
    if not numero_po and ruta_local:
        # Segundo intento: buscar el PO dentro del contenido del archivo (texto digital u OCR)
        try:
            from Nucleo.extractor_texto import extraer_texto as _extraer
            _texto_po = _extraer(ruta_local)
            if _texto_po:
                numero_po = extraer_numero_po(_texto_po)
                if numero_po:
                    print(f"  [PO-CONTENIDO] PO {numero_po} extraído del contenido de: {nombre_archivo}")
        except Exception:
            pass
    # Si aún no hay PO, no abortamos todavía — Claude puede extraerlo al leer el documento

    # --- 1b. Detección de Full Set (PDF con múltiples documentos) — solo si ya tenemos PO ---
    _es_fragmento   = kwargs.get("_es_fragmento", False)
    _tipo_sugerido  = kwargs.get("_tipo_sugerido", None)  # tipo detectado por fullset
    if numero_po and ruta_local and nombre_archivo.lower().endswith(".pdf"):
        fragmentos = separar_fullset(
            ruta_pdf=ruta_local,
            nombre_archivo=nombre_archivo,
            carpeta_temp=str(Path(ruta_local).parent),
            _es_fragmento=_es_fragmento,
        )
        if fragmentos is not None:
            if not fragmentos:
                _logger_clas.debug(f"  [FULLSET] Separación falló — procesando como documento único")
            else:
                _logger_clas.debug(f"  [FULLSET] Procesando {len(fragmentos)} fragmento(s) individualmente")
                resultados_fullset = []
                for fragmento in fragmentos:
                    resultados_fragmento = procesar_adjunto(
                        nombre_archivo=fragmento["nombre"],
                        numero_po_asunto=numero_po,
                        numero_bl=numero_bl,
                        ruta_local=fragmento["ruta"],
                        asunto_correo=asunto_correo,
                        _es_fragmento=True,
                        _tipo_sugerido=fragmento.get("tipo_sugerido"),
                    )
                    if resultados_fragmento:
                        resultados_fullset.extend(resultados_fragmento)
                # Inyectar inconsistencia grave en el primer resultado:
                # el proveedor no debe enviar todos los documentos concatenados en un solo PDF.
                if resultados_fullset:
                    tipos_encontrados = [f.get("tipo_sugerido", f.get("tipo", "")) for f in fragmentos]
                    inc_fullset = {
                        "campo": "Documentos enviados concatenados en un solo PDF",
                        "descripcion": (
                            f"El archivo '{nombre_archivo}' contiene {len(fragmentos)} documentos "
                            f"concatenados en un solo PDF "
                            f"({', '.join(t for t in tipos_encontrados if t)}). "
                            f"Cada documento debe enviarse como un archivo independiente."
                        ),
                        "severidad": "alta",
                    }
                    primer_resultado = resultados_fullset[0]
                    primer_resultado["inconsistencias"] = (
                        [inc_fullset] + list(primer_resultado.get("inconsistencias") or [])
                    )
                return resultados_fullset if resultados_fullset else None

    # --- 2. Clasificar con Claude API (envía el archivo directamente) ---
    tipo_claude, certeza, justificacion_claude, inconsistencias_claude, texto_extraido_claude, nombre_proveedor_claude, es_borrador_claude, numero_po_claude = clasificar_con_claude(
        nombre_archivo=nombre_archivo,
        ruta_local=ruta_local or "",
        asunto_correo=asunto_correo,
    )

    # Tercer intento: PO extraído por Claude al leer el documento (útil para PDFs escaneados sin asunto)
    if not numero_po and numero_po_claude:
        numero_po = numero_po_claude
        print(f"  [PO-CLAUDE] PO {numero_po} extraído por Claude del contenido de: {nombre_archivo}")

    if not numero_po:
        _logger_clas.debug(f"  [SKIP] No se encontró número de PO en: {nombre_archivo}")
        log_advertencia("clasificador", "CLAS-001", nombre_archivo, "No se encontró número de PO (asunto, nombre, contenido ni Claude)", asunto=asunto_correo)
        return None

    metodo = "claude"
    tipos  = []

    # Siempre mostrar lo que respondió Claude
    if es_fallo(tipo_claude):
        _logger_clas.debug(f"  [CLAUDE] ERROR en API: {tipo_claude}")
    else:
        _logger_clas.debug(f"  [CLAUDE] Respondió: {tipo_claude} ({certeza}%) — {justificacion_claude or 'sin justificación'}")

    certeza_minima_efectiva = CLAUDE_CERTEZA_POR_TIPO.get(tipo_claude, CLAUDE_CERTEZA_MINIMA)
    if not es_fallo(tipo_claude) and certeza >= certeza_minima_efectiva:
        # Claude clasificó con confianza suficiente
        if tipo_claude == "PL + INV":
            if _es_fragmento:
                # El separador fullset ya identificó el tipo real de este fragmento.
                # Si dijo INVOICE, conservar INVOICE; si dijo PACKING LIST (o nada), conservar PACKING LIST.
                if _tipo_sugerido == "INVOICE":
                    tipos = ["INVOICE"]
                    _logger_clas.debug(f"  [CLAUDE] Aceptado: PL + INV ({certeza}%) — fragmento INVOICE (fullset), se sube solo como INVOICE")
                else:
                    tipos = ["PACKING LIST"]
                    _logger_clas.debug(f"  [CLAUDE] Aceptado: PL + INV ({certeza}%) — fragmento full set, se sube solo como PACKING LIST")
            else:
                tipos = ["PACKING LIST", "INVOICE"]
        else:
            tipos = [tipo_claude]
        if not _es_fragmento or tipo_claude != "PL + INV":
            _logger_clas.debug(f"  [CLAUDE] Aceptado: {tipo_claude} ({certeza}%)")
    else:
        # --- 4. Fallback por contenido del documento ---
        if es_fallo(tipo_claude):
            log_advertencia("clasificador", "CLAS-002", nombre_archivo, f"Claude falló ({tipo_claude}), usando fallback por contenido", asunto=asunto_correo)
        else:
            log_advertencia("clasificador", "CLAS-003", nombre_archivo, f"Certeza Claude baja ({certeza}%), usando fallback por contenido", asunto=asunto_correo)

        texto_documento = ""
        if ruta_local:
            try:
                from Nucleo.extractor_texto import extraer_texto
                texto_documento = extraer_texto(ruta_local) or ""
            except Exception as e:
                _logger_clas.debug(f"  [FALLBACK] No se pudo extraer texto: {e}")

        if texto_documento:
            tipo_contenido = _clasificar_fallback_contenido(texto_documento.upper())
            if tipo_contenido != "OTROS":
                _logger_clas.debug(f"  [FALLBACK-CONTENIDO] Clasificado por contenido: {tipo_contenido}")
                tipos  = [tipo_contenido]
                metodo = "fallback_contenido"
                certeza = 0
            elif _tipo_sugerido and _tipo_sugerido != "OTROS":
                _logger_clas.debug(f"  [FALLBACK-FULLSET] Usando tipo del fullset: {_tipo_sugerido}")
                tipos  = [_tipo_sugerido]
                metodo = "fallback_fullset"
                certeza = 0
            else:
                tipos  = ["OTROS"]
                metodo = "fallback_contenido"
                certeza = 0
        else:
            if _tipo_sugerido and _tipo_sugerido != "OTROS":
                _logger_clas.debug(f"  [FALLBACK-FULLSET] Sin texto — usando tipo del fullset: {_tipo_sugerido}")
                tipos  = [_tipo_sugerido]
                metodo = "fallback_fullset"
                certeza = 0
            else:
                tipos  = ["OTROS"]
                metodo = "fallback_contenido"
                certeza = 0

    # Las inconsistencias, texto y proveedor solo están disponibles cuando Claude clasificó directamente
    inconsistencias  = inconsistencias_claude       if metodo == "claude" else []
    texto_extraido   = texto_extraido_claude        if metodo == "claude" else ""
    nombre_proveedor = nombre_proveedor_claude      if metodo == "claude" else ""

    # Fusionar con validaciones deterministas del SA (validador_campos.py).
    # Se ejecutan sobre todos los tipos con texto disponible, independientemente
    # del método de clasificación, para no perder validaciones SA en fallbacks.
    _texto_para_validar = texto_extraido or ""
    if not _texto_para_validar and ruta_local:
        try:
            from Nucleo.extractor_texto import extraer_texto as _extraer_v
            _texto_para_validar = _extraer_v(ruta_local) or ""
        except Exception:
            pass
    for _tipo_val in tipos:
        _incs_deterministas = validar_campos(_tipo_val, _texto_para_validar)
        if _incs_deterministas:
            inconsistencias = list(inconsistencias) + _incs_deterministas
    # Borrador: Claude lo detectó en el contenido O está en el nombre del archivo
    es_borrador = es_borrador_claude or es_borrador_por_nombre(nombre_archivo)

    # Regla de seguridad para CO: solo marcar definitivo si Claude confirmó el sello
    # Y el texto extraído lo respalda. Ante cualquier duda → borrador.
    if "CO" in tipos and not es_borrador:
        if texto_extraido:
            _PALABRAS_SELLO = ("sello", "stamp", "seal", "official", "authorized", "certified",
                               "certifi", "câmara", "camara", "chamber", "customs", "aduana",
                               "ministerio", "ministry", "original")
            texto_lower = texto_extraido.lower()
            tiene_sello = any(p in texto_lower for p in _PALABRAS_SELLO)
            if not tiene_sello:
                es_borrador = True
        # Si no hay texto extraíble (imagen pura), confiar solo en Claude;
        # el prompt ya instruye a Claude a asumir borrador por defecto para CO.

    # --- 6. Construir lista de destinos ---
    return [
        {
            "nombre_archivo":       nombre_archivo,
            "nombre_destino":       renombrar_para_tipo(nombre_archivo, tipo, tipos),
            "numero_po":            numero_po,
            "numero_bl":            numero_bl,
            "tipo":                 tipo,
            "ruta_sharepoint":      construir_ruta(
                                        numero_po,
                                        tipo,
                                        renombrar_para_tipo(nombre_archivo, tipo, tipos),
                                        es_borrador=es_borrador,
                                    ),
            "certeza":              certeza,
            "metodo_clasificacion": metodo,
            "inconsistencias":      inconsistencias,
            "ruta_local":           ruta_local,
            "texto_extraido":       texto_extraido,
            "nombre_proveedor":     nombre_proveedor,
            "es_borrador":          es_borrador,
        }
        for tipo in tipos
    ]
