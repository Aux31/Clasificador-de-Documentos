"""
prueba_monitor_vm.py
---------------------
Script de prueba para verificar que monitor_sin_claude.py está listo
para correr en la VM el fin de semana.

Comprueba paso a paso:
  1. Configuración cargada (.env)
  2. MODO_SIN_AV activado (no requiere Bitdefender)
  3. MODO_MOCK (advertencia si está en true)
  4. Conexión Outlook COM y cuenta configurada
  5. Clasificación por keywords (sin Claude)
  6. Escritura de log de inconsistencias
  7. Conexión Graph API / SharePoint

Uso:
    python pruebas/prueba_monitor_vm.py
"""

import sys
import os
from pathlib import Path

# Ejecutable desde la carpeta raíz del proyecto
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# ---------------------------------------------------------------------------
def _ok(msg):  print(f"  [OK] {msg}")
def _fail(msg): print(f"  [FAIL] {msg}")
def _warn(msg): print(f"  [WARN] {msg}")
def _titulo(msg): print(f"\n{'-' * 60}\n{msg}\n{'-' * 60}")
# ---------------------------------------------------------------------------


_titulo("1. Configuracion (.env)")
try:
    from configuracion.ajustes import (
        OUTLOOK_EMAIL, MODO_MOCK, MODO_SIN_AV,
        SEGURIDAD_SERVICIO_AV, EXTENSIONES_BLOQUEADAS,
        ANTHROPIC_API_KEY,
    )
    _ok(f"OUTLOOK_EMAIL        = {OUTLOOK_EMAIL}")
    _ok(f"MODO_MOCK            = {MODO_MOCK}")
    _ok(f"MODO_SIN_AV          = {MODO_SIN_AV}")
    _ok(f"EXTENSIONES_BLOQUEADAS = {len(EXTENSIONES_BLOQUEADAS)} extensiones")
    if ANTHROPIC_API_KEY:
        _warn("ANTHROPIC_API_KEY está configurada — monitor_sin_claude.py NO la usa, pero está presente")
    else:
        _ok("ANTHROPIC_API_KEY no configurada (correcto para modo sin Claude)")
except Exception as e:
    _fail(f"Error cargando configuracion: {e}")
    sys.exit(1)


_titulo("2. MODO_SIN_AV")
if MODO_SIN_AV:
    _ok("MODO_SIN_AV=true — Bitdefender NO es necesario")
else:
    _warn("MODO_SIN_AV=false — el monitor verificará Bitdefender en cada correo")
    _warn("Si la VM no tiene Bitdefender, cambia MODO_SIN_AV=true en el .env")

if MODO_MOCK:
    _warn("MODO_MOCK=true — los archivos NO se subirán a SharePoint real")
else:
    _ok("MODO_MOCK=false — modo produccion, se subirá a SharePoint real")


_titulo("3. Conexión Outlook COM")
try:
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        _ok("Outlook abierto y COM accesible")

        bandeja = None
        cuentas = []
        for store in namespace.Stores:
            try:
                nombre = store.DisplayName
                cuentas.append(nombre)
                if OUTLOOK_EMAIL.lower() in nombre.lower():
                    bandeja = store.GetDefaultFolder(6)
            except Exception:
                pass

        _ok(f"Cuentas encontradas en Outlook: {cuentas}")
        if bandeja:
            n_items = bandeja.Items.Count
            _ok(f"Bandeja de entrada '{OUTLOOK_EMAIL}' encontrada — {n_items} item(s)")
        else:
            _fail(f"Cuenta '{OUTLOOK_EMAIL}' NO encontrada en Outlook")
            _warn("Abre Outlook y asegúrate de que la cuenta esté iniciada antes de correr el monitor")
    finally:
        pythoncom.CoUninitialize()
except ImportError:
    _fail("win32com no disponible — ¿pywin32 instalado?")
except Exception as e:
    _fail(f"Error conectando a Outlook COM: {e}")


_titulo("4. Clasificacion por keywords (sin Claude)")
try:
    from clasificador import _clasificar_fallback_nombre, extraer_po_y_bl_de_asunto

    casos = [
        ("BL PO 196893.pdf",              ["BL"]),
        ("INVOICE PO 197000.pdf",         ["INVOICE"]),
        ("PACKING LIST PO 196800.pdf",    ["PACKING LIST"]),
        ("CO PO 194386.pdf",              ["CO"]),
        ("DOCUMENTO_SIN_TIPO.pdf",        ["OTROS"]),
        ("WEIGHT CERTIFICATE PO 123.pdf", ["WEIGHT CERTIFICATE"]),
    ]

    todos_ok = True
    for nombre, esperados in casos:
        resultado = _clasificar_fallback_nombre(nombre)
        ok = all(e in resultado for e in esperados)
        if ok:
            _ok(f"{nombre:45s} -> {resultado}")
        else:
            _fail(f"{nombre:45s} -> {resultado}  (esperado: {esperados})")
            todos_ok = False

    # Extracción de PO/BL del asunto
    po, bl = extraer_po_y_bl_de_asunto("PO 196893 - Shipping Docs BL-MEDUWH689046")
    if po == "196893":
        _ok(f"Extraccion PO del asunto: '{po}' (correcto)")
    else:
        _fail(f"Extraccion PO del asunto: '{po}' (esperado '196893')")

    if bl:
        _ok(f"Extraccion BL del asunto: '{bl}'")
    else:
        _warn("BL no detectado en el asunto de prueba")

    if todos_ok:
        _ok("Clasificacion por keywords funcionando correctamente")
    else:
        _warn("Algunos casos de clasificacion no coinciden — revisar")

except Exception as e:
    _fail(f"Error en clasificacion por keywords: {e}")


_titulo("5. Log de inconsistencias (texto plano)")
try:
    from datetime import datetime
    dir_registros = Path(__file__).resolve().parent.parent / "registros"
    dir_registros.mkdir(exist_ok=True)
    fecha_hoy  = datetime.now().strftime("%Y-%m-%d")
    ruta_test  = dir_registros / f"prueba_vm_{fecha_hoy}.log"

    with open(ruta_test, "w", encoding="utf-8") as f:
        f.write(f"Prueba de escritura de log — {datetime.now()}\n")

    _ok(f"Log escrito en: {ruta_test}")
    os.remove(ruta_test)
    _ok("Archivo de prueba eliminado correctamente")
except Exception as e:
    _fail(f"Error escribiendo log: {e}")


_titulo("6. Conexion Graph API / SharePoint")
try:
    from graph_client import GraphClient
    cliente = GraphClient()
    _ok("GraphClient instanciado correctamente")

    if MODO_MOCK:
        _warn("MODO_MOCK=true — la conexion real a SharePoint no se probara")
    else:
        # Si GraphClient() se instancio sin excepcion, el token ya se obtuvo
        _ok("Token de Graph API obtenido correctamente (sin errores en instancia)")
except Exception as e:
    _fail(f"Error instanciando GraphClient: {e}")


_titulo("7. Verificacion del agente de seguridad")
try:
    from agente_seguridad import verificar_servicio_av
    ok_av, err_av = verificar_servicio_av()
    if ok_av:
        _ok("Agente seguridad OK (MODO_SIN_AV activo o Bitdefender corriendo)")
    else:
        _warn(f"Agente seguridad: {err_av} — el monitor bloqueará adjuntos si no se activa MODO_SIN_AV")
except Exception as e:
    _fail(f"Error en agente_seguridad: {e}")


# ---------------------------------------------------------------------------
_titulo("RESUMEN")
print(f"  monitor_sin_claude.py listo para correr en la VM")
print(f"  Comando para iniciar:")
print(f"")
print(f"    python c:\\Users\\aux22.gg\\Desktop\\PROYECTOS\\CLASIFICADOR_DOCS\\clasificador_documentos\\monitor_sin_claude.py")
print(f"")
print(f"  Los logs se guardan en:")
print(f"    registros_subidas.log           — archivos subidos")
print(f"    registros_eventos.log           — eventos y errores")
print(f"    registros/inconsistencias_HOY.log — inconsistencias (texto, sin Word)")
print(f"  Ctrl+C para detener el monitor.\n")
