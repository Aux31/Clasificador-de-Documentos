# -*- coding: utf-8 -*-
"""
verificar_migracion.py
----------------------
Verifica que el entorno del proyecto CLASIFICADOR_DOCS esté correctamente
configurado en la PC nueva. No clasifica correos ni sube archivos.

Uso:
    python c:\Users\aux22.gg\Desktop\PROYECTOS\CLASIFICADOR_DOCS\clasificador_documentos\pruebas\verificar_migracion.py
"""

import sys
import importlib
from pathlib import Path

RUTA_PROYECTO = Path(__file__).resolve().parent.parent
OK   = "[OK]"
FAIL = "[FAIL]"

resultados = []

def check(nombre, ok, detalle=""):
    estado = OK if ok else FAIL
    linea = f"  {estado}  {nombre}"
    if detalle:
        linea += f"  →  {detalle}"
    print(linea)
    resultados.append(ok)

def seccion(titulo):
    print(f"\n{'='*55}")
    print(f"  {titulo}")
    print('='*55)


# ── 1. Python ───────────────────────────────────────────────
seccion("1. Python")
version = sys.version_info
check(
    f"Python {version.major}.{version.minor}.{version.micro}",
    version.major == 3 and version.minor >= 12,
    "Se requiere Python 3.12+" if not (version.major == 3 and version.minor >= 12) else ""
)

# ── 2. Librerías ────────────────────────────────────────────
seccion("2. Librerías requeridas")
libs = [
    ("anthropic",    "Claude API"),
    ("msal",         "Autenticación Graph API"),
    ("requests",     "HTTP"),
    ("dotenv",       "python-dotenv"),
    ("win32com",     "pywin32 — Outlook COM"),
    ("openpyxl",     "Excel"),
    ("bs4",          "BeautifulSoup"),
    ("PIL",          "Pillow"),
    ("psutil",       "Procesos — antivirus"),
]
for lib, uso in libs:
    try:
        importlib.import_module(lib)
        check(f"{lib} ({uso})", True)
    except ImportError:
        check(f"{lib} ({uso})", False, "No instalada")

# OCR — opcionales
seccion("3. OCR (opcionales pero recomendados)")
for lib, uso in [("pytesseract", "Tesseract OCR"), ("pdf2image", "PDF → imagen")]:
    try:
        importlib.import_module(lib)
        check(f"{lib} ({uso})", True)
    except ImportError:
        check(f"{lib} ({uso})", False, "No instalada — OCR de imágenes no funcionará")

# ── 4. Archivos del proyecto ────────────────────────────────
seccion("4. Archivos del proyecto")
archivos = [
    RUTA_PROYECTO / "configuracion" / ".env",
    RUTA_PROYECTO / "configuracion" / "ajustes.py",
    RUTA_PROYECTO / "clasificador.py",
    RUTA_PROYECTO / "clasificador_claude.py",
    RUTA_PROYECTO / "catalogo_tipos.py",
    RUTA_PROYECTO / "catalogo_ejemplos.py",
    RUTA_PROYECTO / "graph_client.py",
    RUTA_PROYECTO / "recopilador_documentos.py",
    RUTA_PROYECTO / "monitor_correos.py",
    RUTA_PROYECTO / "agente_seguridad.py",
    RUTA_PROYECTO / "notificador.py",
]
for archivo in archivos:
    check(archivo.name, archivo.exists(), str(archivo.relative_to(RUTA_PROYECTO)))

# Carpeta temporal
carpeta_temp = RUTA_PROYECTO / "temporal"
check("carpeta temporal/", carpeta_temp.exists(), "Crear manualmente si falta")

# ── 5. Variables de configuración ──────────────────────────
seccion("5. Variables de configuración (.env)")
try:
    sys.path.insert(0, str(RUTA_PROYECTO.parent))
    from clasificador_documentos.configuracion.ajustes import (
        TENANT_ID, CLIENT_ID, CLIENT_SECRET,
        SHAREPOINT_SITE_ID, SHAREPOINT_DRIVE_ID,
        ANTHROPIC_API_KEY, CLAUDE_MODELO,
        OUTLOOK_EMAIL,
        CATALOGO_RUTA,
    )
    check("GRAPH_TENANT_ID",       bool(TENANT_ID),       "configurado" if TENANT_ID else "vacío")
    check("GRAPH_CLIENT_ID",       bool(CLIENT_ID),       "configurado" if CLIENT_ID else "vacío")
    check("GRAPH_CLIENT_SECRET",   bool(CLIENT_SECRET),   "configurado" if CLIENT_SECRET else "vacío")
    check("SHAREPOINT_SITE_ID",    bool(SHAREPOINT_SITE_ID), "configurado" if SHAREPOINT_SITE_ID else "vacío")
    check("SHAREPOINT_DRIVE_ID",   bool(SHAREPOINT_DRIVE_ID), "configurado" if SHAREPOINT_DRIVE_ID else "vacío")
    check("ANTHROPIC_API_KEY",     bool(ANTHROPIC_API_KEY), "configurado" if ANTHROPIC_API_KEY else "vacío")
    check("CLAUDE_MODELO",         bool(CLAUDE_MODELO),   CLAUDE_MODELO)
    check("OUTLOOK_EMAIL",         bool(OUTLOOK_EMAIL),   OUTLOOK_EMAIL or "vacío")
    check("CATALOGO_RUTA existe",  Path(CATALOGO_RUTA).exists(), str(CATALOGO_RUTA))
except Exception as e:
    check("Carga de ajustes.py", False, str(e))

# ── 6. Claude API ───────────────────────────────────────────
seccion("6. Claude API (Anthropic)")
try:
    import anthropic
    from clasificador_documentos.configuracion.ajustes import ANTHROPIC_API_KEY, CLAUDE_MODELO
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    resp = client.messages.create(
        model=CLAUDE_MODELO,
        max_tokens=10,
        messages=[{"role": "user", "content": "hola"}]
    )
    check("Conexión Claude API", True, f"modelo={CLAUDE_MODELO}")
except Exception as e:
    check("Conexión Claude API", False, str(e)[:120])

# ── 7. Microsoft Graph API ──────────────────────────────────
seccion("7. Microsoft Graph API (SharePoint)")
try:
    import msal, requests
    from clasificador_documentos.configuracion.ajustes import TENANT_ID, CLIENT_ID, CLIENT_SECRET
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, CLIENT_SECRET, authority=authority)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    token_ok = "access_token" in result
    check("Token Graph API", token_ok, "OK" if token_ok else result.get("error_description", "error")[:80])
    if token_ok:
        headers = {"Authorization": f"Bearer {result['access_token']}"}
        r = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers, timeout=10)
        check("Endpoint Graph /me", r.status_code in (200, 401), f"HTTP {r.status_code}")
except Exception as e:
    check("Graph API", False, str(e)[:120])

# ── 8. Outlook (win32com) ───────────────────────────────────
seccion("8. Outlook local (win32com)")
try:
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    check("Outlook COM", True, f"Bandeja de entrada: {inbox.Items.Count} correos")
except Exception as e:
    check("Outlook COM", False, str(e)[:100])

# ── 9. Antivirus ────────────────────────────────────────────
seccion("9. Antivirus (EPSecurityService)")
try:
    import psutil
    servicio = "EPSecurityService"
    encontrado = any(s.name() == servicio for s in psutil.process_iter(['name']))
    check(servicio, encontrado, "en ejecución" if encontrado else "NO encontrado (puede fallar agente_seguridad)")
except Exception as e:
    check("EPSecurityService", False, str(e))

# ── 10. Tesseract (OCR) ─────────────────────────────────────
seccion("10. Tesseract OCR")
try:
    from clasificador_documentos.configuracion.ajustes import TESSERACT_CMD
    tesseract_path = Path(TESSERACT_CMD)
    check("tesseract.exe", tesseract_path.exists(), str(TESSERACT_CMD))
except Exception as e:
    check("Tesseract", False, str(e))

# ── Resumen ─────────────────────────────────────────────────
total    = len(resultados)
exitosos = sum(resultados)
fallidos = total - exitosos

print(f"\n{'='*55}")
print(f"  RESUMEN: {exitosos}/{total} verificaciones pasadas")
if fallidos:
    print(f"  {fallidos} elemento(s) requieren atención (ver [FAIL] arriba)")
else:
    print("  Todo OK — el entorno está listo")
print('='*55)
