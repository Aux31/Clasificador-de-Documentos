# -*- coding: utf-8 -*-
import os
from dotenv import load_dotenv
from pathlib import Path

# Cargar .env desde la misma carpeta
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path, encoding="utf-8")

# Graph API
TENANT_ID     = os.getenv("GRAPH_TENANT_ID")
CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID")
CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")

# Correo
OUTLOOK_EMAIL    = os.getenv("OUTLOOK_EMAIL")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")

# SharePoint
SHAREPOINT_SITE_ID      = os.getenv("SHAREPOINT_SITE_ID")
SHAREPOINT_DRIVE_ID     = os.getenv("SHAREPOINT_DRIVE_ID")
SHAREPOINT_CARPETA_RAIZ = os.getenv("SHAREPOINT_CARPETA_RAIZ", "Documentos PO")

# Teams — dos canales separados
TEAMS_WEBHOOK_PRUEBAS   = os.getenv("TEAMS_WEBHOOK_PRUEBAS",   os.getenv("TEAMS_WEBHOOK_URL", ""))
TEAMS_WEBHOOK_PROBLEMAS = os.getenv("TEAMS_WEBHOOK_PROBLEMAS", os.getenv("TEAMS_WEBHOOK_URL", ""))

# Comportamiento
MODO_MOCK          = os.getenv("MODO_MOCK", "false").lower() == "true"
INTERVALO_SEGUNDOS = int(os.getenv("INTERVALO_SEGUNDOS", "60"))
CARPETA_PROCESADOS = os.getenv("CARPETA_PROCESADOS", "Procesados")

# Extensiones bloqueadas — ejecutables, scripts, imágenes inline y correos anidados
EXTENSIONES_BLOQUEADAS = {
    ".exe", ".scr",
    ".js", ".vbs", ".bat", ".cmd", ".ps1",
    ".gif", ".bmp", ".webp", ".ico", ".svg",
    ".msg",   # correo anidado como adjunto — no es un documento de carga
}

# Extensiones de imagen que se procesan con OCR si vienen como adjunto
EXTENSIONES_IMAGEN_OCR = {".jpg", ".jpeg", ".png", ".tiff", ".tif"}

# Seguridad — agente_seguridad.py
# MODO_SIN_AV=true desactiva los checks de Bitdefender (para pruebas en VM sin AV)
MODO_SIN_AV                    = os.getenv("MODO_SIN_AV", "false").lower() == "true"
SEGURIDAD_SERVICIO_AV          = os.getenv("SEGURIDAD_SERVICIO_AV", "EPSecurityService")
SEGURIDAD_ESPERA_AV_SEGUNDOS   = int(os.getenv("SEGURIDAD_ESPERA_AV_SEGUNDOS", "3"))
SEGURIDAD_TAMANO_MAX_MB        = float(os.getenv("SEGURIDAD_TAMANO_MAX_MB", "50"))
SEGURIDAD_EXTENSIONES_PERMITIDAS = set(os.getenv(
    "SEGURIDAD_EXTENSIONES_PERMITIDAS",
    ".pdf,.xlsx,.xls,.docx,.doc,.pptx,.ppt,.zip,.rar,.txt,.csv,.xml,.json,.jpg,.jpeg,.png,.bmp,.tiff,.tif,.webp"
).split(","))

# Carpeta raíz de OC's en SharePoint
SHAREPOINT_CARPETA_OCS = "Centro de Documentaci\u00f3n Grupo Inteca Digitalizado/Proveedores/Transacciones con proveedores internacionales/OC\u00b4s"

# Estructura completa de carpetas por OC en SharePoint.
# Estas carpetas existen al mismo nivel dentro de cada OC.
# Las marcadas como activas son las que el bot usa actualmente para subir archivos.
# Las demás están registradas como referencia para implementaciones futuras.
ESTRUCTURA_CARPETAS_OC = [
    "1. CONTROL PROCESO",        # referencia — no activa
    "1. CONTROL PROCESO/1.01 LOG de eventos",          # referencia — no activa
    "1. CONTROL PROCESO/1.02 Flujo de trabajo",        # referencia — no activa
    "1. CONTROL PROCESO/1.03 Control de file original",# referencia — no activa
    "2. CONTRATO OC MADRE",      # referencia — no activa
    "2. CONTRATO OC MADRE/2.01 PO sin firmar",                         # referencia — no activa
    "2. CONTRATO OC MADRE/2.02 PO firmada por PROV",                   # referencia — no activa
    "2. CONTRATO OC MADRE/2.03 Shipping agreement . Instructions",     # referencia — no activa
    "2. CONTRATO OC MADRE/2.04 Factura proforma",                      # referencia — no activa
    "2. CONTRATO OC MADRE/2.05 OC hija",                               # referencia — no activa
    "3. FLETE INTERNACIONAL",    # referencia — no activa
    "3. FLETE INTERNACIONAL/3.01 Reservacion",                         # referencia — no activa
    "3. FLETE INTERNACIONAL/3.02 Orden de compra flete internacional", # referencia — no activa
    "3. FLETE INTERNACIONAL/3.03 Factura",                             # referencia — no activa
    "4. DOCUMENTACION",          # activa — el bot sube todos los documentos aquí
    "5. DESALMACENAJE",          # referencia — no activa
    "5. DESALMACENAJE/5.01 Precosteo",             # referencia — no activa
    "5. DESALMACENAJE/5.02 DUA de impuestos",      # referencia — no activa
    "5. DESALMACENAJE/5.03 OC de impuestos",       # referencia — no activa
    "5. DESALMACENAJE/5.04 Cheque pago de DUA",    # referencia — no activa
    "5. DESALMACENAJE/5.05 Declaracion de valor",  # referencia — no activa
    "6.0 ALMACEN FISCAL",        # referencia — no activa
    "6.0 ALMACEN FISCAL/6.01 OC compra AF",                      # referencia — no activa
    "6.0 ALMACEN FISCAL/6.02 Notificacion de video",             # referencia — no activa
    "6.0 ALMACEN FISCAL/6.03 TIRS",                              # referencia — no activa
    "6.0 ALMACEN FISCAL/6.04 Patio devolucion de contenedores",  # referencia — no activa
    "6.0 ALMACEN FISCAL/6.05 Video de descarga",                 # referencia — no activa
    "6.0 ALMACEN FISCAL/6.06 Reporte danos en AF",               # referencia — no activa
    "6.0 ALMACEN FISCAL/6.07 Movimiento de ingreso",             # referencia — no activa
    "6.0 ALMACEN FISCAL/6.08 Movimiento de salida",              # referencia — no activa
    "7. TRANSPORTE TERRESTRE",   # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.01 OC Transportista terrestre",   # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.02 Datos de choferes",            # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.03 Cita Transporte para recoger", # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.04 Pago de cita",                 # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.05 DUA transito",                 # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.06 Cita para devolver",           # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.07 Coordinacion de predio",       # referencia — no activa
    "7. TRANSPORTE TERRESTRE/7.08 Video de descarga",            # referencia — no activa
    "8. PAGOS",                  # referencia — no activa
    "8. PAGOS/8.01 Adelanto 1",  # referencia — no activa
    "8. PAGOS/8.02 Adelanto 2",  # referencia — no activa
    "8. PAGOS/8.03 Pago final",  # referencia — no activa
    "6. COSTEO",                 # referencia — no activa
    "6. COSTEO/6.4 Check List",  # referencia — no activa
    "6. COSTEO/6.1 Costeo 1",    # referencia — no activa
    "6. COSTEO/6.2 Costeo 2",    # referencia — no activa
    "6. COSTEO/6.3 Costeo 3",    # referencia — no activa
    "9. RECLAMOS",               # referencia — no activa
    "9. RECLAMOS/9.1.1 Comunicaciones",        # referencia — no activa
    "9. RECLAMOS/9.1.2 NC Reclamo",            # referencia — no activa
    "9. RECLAMOS/9.1.3 Reclamo inicial",       # referencia — no activa
    "9.2 RECLAMOS",              # referencia — no activa
    "9.2 RECLAMOS/9.2.1 Comunicaciones",        # referencia — no activa
    "9.2 RECLAMOS/9.2.2 NC Reclamo",            # referencia — no activa
    "9.2 RECLAMOS/9.2.3 Reclamo inicial",       # referencia — no activa
    "9.2 RECLAMOS/9.2.4 Resolucion de reclamo", # referencia — no activa
    "9.3 RECLAMOS",              # referencia — no activa
    "9.3 RECLAMOS/9.3.1 Comunicaciones",        # referencia — no activa
    "9.3 RECLAMOS/9.3.2 NC Reclamo",            # referencia — no activa
    "9.3 RECLAMOS/9.3.3 Reclamo inicial",       # referencia — no activa
    "9.3 RECLAMOS/9.3.4 Resolucion de reclamo", # referencia — no activa
]

# Mapeo tipo → carpeta de BORRADOR en SharePoint (solo tipos que tienen carpeta de borrador)
MAPA_CARPETAS_BORRADOR = {
    "INVOICE":   "4.01 Borrador de factura",
    "PL + INV":  "4.01 Borrador de factura",
    "BL":        "4.03 Borr BL oAWB o Porte",
    "MBL":       "4.03 Borr BL oAWB o Porte",
    "HBL":       "4.03 Borr BL oAWB o Porte",
    "AWB":       "4.03 Borr BL oAWB o Porte",
    "CO":        "4.07 Borr COO",
    "FITOSANITARIO": "4.011 Borr Cert Fito Origen",
    "ZOOSANITARIO":  "4.16 Borr Zoos Origen",
    "PRINTER":       "4.21 Borrador docs Agencia",
}

# Mapeo tipo → carpeta exacta en SharePoint (orden = prioridad de clasificación)
MAPA_CARPETAS = {
    "FOB LETTER":          "OTROS",
    "PL + INV":            "4.27 Packing list definitivo",
    "WEIGHT CERTIFICATE":  "OTROS",
    "QUALITY CERTIFICATE": "OTROS",
    "PACKING LIST":        "4.27 Packing list definitivo",
    "INVOICE":             "4.02 Factura Definitiva",
    "BL":                  "4.05 BL-AWB-Porte definitivo",
    "MBL":                 "4.05 BL-AWB-Porte definitivo",
    "HBL":                 "4.05 BL-AWB-Porte definitivo",
    "AWB":                 "4.05 BL-AWB-Porte definitivo",
    "CO":                  "4.10 Certificado Origen definitivo (COO)",
    "FITOSANITARIO":       "4.12 Aprob Borr Cert fito origen",
    "ZOOSANITARIO":        "4.18 Certifi Zoos Origen",
    "PRINTER":             "4.24 Aceptacion documentos de agencia",
    "EXONERACION":         "4.25 Exoneracion Hacienda",
    "MARCHAMO":            "4.28 Foto del Marchamo",
}

# Lista derivada — mantiene compatibilidad con clasificar_tipo()
TIPOS_DOCUMENTO = list(MAPA_CARPETAS.keys())

# Prefijo usado para renombrar archivos combinados al separar por tipo
# Solo aplica cuando un archivo genera múltiples destinos
PREFIJOS_TIPO = {
    "PACKING LIST": "PL",
    "INVOICE":      "INV",
}

# ---------------------------------------------------------------------------
# Maersk Bookings — monitor_maersk.py
# ---------------------------------------------------------------------------
# Carpeta dentro de la OC (bajo 4. DOCUMENTACION/) donde se guardan los bookings
MAERSK_CARPETA_BOOKING = os.getenv("MAERSK_CARPETA_BOOKING", "4.00 Booking Maersk")

# ---------------------------------------------------------------------------
# Claude API — clasificador principal
# ---------------------------------------------------------------------------
ANTHROPIC_API_KEY       = os.getenv("ANTHROPIC_API_KEY")
CLAUDE_MODELO           = os.getenv("CLAUDE_MODELO", "claude-haiku-4-5")
CLAUDE_TIMEOUT_SEGUNDOS = int(os.getenv("CLAUDE_TIMEOUT_SEGUNDOS", "30"))
CLAUDE_MAX_TOKENS       = int(os.getenv("CLAUDE_MAX_TOKENS", "2000"))
CLAUDE_MAX_CHARS_TEXTO  = int(os.getenv("CLAUDE_MAX_CHARS_TEXTO", "2000"))
# Certeza mínima (0-100): si Claude retorna menos, se activa el fallback por keywords
CLAUDE_CERTEZA_MINIMA   = int(os.getenv("CLAUDE_CERTEZA_MINIMA", "75"))

# Threshold de certeza por tipo — override del valor global para tipos con señales visuales
# inequívocas (foto del marchamo, sello fitosanitario, etc.)
CLAUDE_CERTEZA_POR_TIPO: dict[str, int] = {
    "MARCHAMO":      50,   # Foto del precinto: visualmente obvia, Claude puede dar 60 y ser correcto
    "FITOSANITARIO": 55,   # Certificado con sello oficial: formato estandarizado
    "ZOOSANITARIO":  55,   # Igual que fitosanitario
    "EXONERACION":   55,   # Formato de Hacienda CR: muy específico y reconocible
    "PRINTER":       60,   # Agencia aduanal: encabezado estándar
}

# ---------------------------------------------------------------------------
# Extracción de texto (OCR)
# ---------------------------------------------------------------------------
TESSERACT_CMD = os.getenv("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
POPPLER_PATH  = os.getenv("POPPLER_PATH",  r"C:\poppler\poppler-25.12.0\Library\bin")

# ---------------------------------------------------------------------------
# Catálogo de ejemplos reales
# ---------------------------------------------------------------------------
# Ruta a la carpeta con subcarpetas por tipo (BL/, INVOICE/, CO/, etc.)
CATALOGO_RUTA = Path(os.getenv("CATALOGO_RUTA", r"C:\Users\aux22.gg\Desktop\PROYECTOS\CATALOGO"))

# Cuántos ejemplos enviar a Claude por tipo
# Los tipos aquí listados usan 3 ejemplos; el resto usa 2
CATALOGO_EJEMPLOS_3 = {"BL", "CO", "INVOICE"}
CATALOGO_EJEMPLOS_DEFAULT = 2
CATALOGO_EJEMPLOS_MAX     = 3
# Máximo de caracteres por ejemplo (para no inflar el prompt)
CATALOGO_MAX_CHARS_EJEMPLO = int(os.getenv("CATALOGO_MAX_CHARS_EJEMPLO", "1500"))
# ---------------------------------------------------------------------------
# Notificaciones internas — reporte de inconsistencias
# ---------------------------------------------------------------------------
CORREO_ALERTAS_INCONSISTENCIAS = os.getenv("CORREO_ALERTAS_INCONSISTENCIAS", "aux23.gg@grupointeca.com")
