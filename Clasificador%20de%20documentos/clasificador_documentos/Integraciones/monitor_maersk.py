"""
Monitor de bookings Maersk — descarga PDFs de reservas y los sube a SharePoint.

Uso:
    python monitor_maersk.py              # Procesa todos los bookings nuevos
    python monitor_maersk.py --login      # Fuerza re-login (renovar sesión guardada)
    python monitor_maersk.py --desde 268551540   # Reanuda desde un booking específico

Protecciones anti-bloqueo:
    - Delay aleatorio entre bookings (configurable con MAERSK_DELAY_MIN/MAX en .env)
    - Pausa larga cada N bookings para simular comportamiento humano
    - Reintentos con backoff exponencial en cada descarga
    - Checkpoint incremental: si se corta, reanuda donde quedó
    - User-Agent de Chrome real para evitar detección headless

Dependencias adicionales requeridas:
    pip install playwright
    playwright install chromium
"""

import sys
import os
import io
import json
import time
import random
import base64
import argparse
import tempfile
import logging
import traceback
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Asegurar UTF-8 en stdout (Windows)
# ---------------------------------------------------------------------------
if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ---------------------------------------------------------------------------
# Path setup — igual que en monitor_correos.py
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from configuracion.ajustes import (
    ANTHROPIC_API_KEY,
    CLAUDE_MODELO,
    SHAREPOINT_CARPETA_OCS,
    MODO_MOCK,
    MAERSK_CARPETA_BOOKING,
)
from Integraciones.graph_client import GraphClient
from Nucleo.clasificador import formatear_numero_oc

# ---------------------------------------------------------------------------
# Rutas de archivos internos
# ---------------------------------------------------------------------------
_DIR              = Path(__file__).parent
_PROCESADOS_FILE  = _DIR / "maersk_procesados.txt"
_SESSION_FILE     = _DIR / "maersk_session.json"
_LOG_SUBIDAS      = Path(__file__).parent.parent / "registros_subidas_maersk.log"

MAERSK_URL_BOOKINGS = "https://www.maersk.com/bookings/ocean"

# ---------------------------------------------------------------------------
# Parámetros anti-rate-limit (configurables desde .env)
# ---------------------------------------------------------------------------
# Delay aleatorio entre cada booking (segundos)
_DELAY_MIN = float(os.getenv("MAERSK_DELAY_MIN", "2.5"))
_DELAY_MAX = float(os.getenv("MAERSK_DELAY_MAX", "5.0"))

# Cada cuántos bookings hacer una pausa larga (simula humano que va a tomar café)
_PAUSA_LARGA_CADA = int(os.getenv("MAERSK_PAUSA_CADA", "15"))
_PAUSA_LARGA_SEG  = float(os.getenv("MAERSK_PAUSA_SEG", "30.0"))

# Reintentos por descarga de PDF fallida
_MAX_REINTENTOS_DESCARGA = int(os.getenv("MAERSK_REINTENTOS", "3"))

# User-Agent de Chrome real para no ser detectado como headless
_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

# ---------------------------------------------------------------------------
# Logger de subidas (misma estructura que registros_subidas.log)
# ---------------------------------------------------------------------------
_logger_subidas = logging.getLogger("maersk_subidas")
if not _logger_subidas.handlers:
    _h = logging.FileHandler(_LOG_SUBIDAS, encoding="utf-8")
    _h.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_h)
    _logger_subidas.setLevel(logging.INFO)


def _log_subida(booking_num: str, ruta_sharepoint: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ruta_norm = ruta_sharepoint.replace("\\", "/")
    _logger_subidas.info(f"{ts} | {booking_num} | {ruta_norm}")


# ---------------------------------------------------------------------------
# Registro de bookings procesados
# ---------------------------------------------------------------------------

def _cargar_procesados() -> set:
    if _PROCESADOS_FILE.exists():
        return set(_PROCESADOS_FILE.read_text(encoding="utf-8").splitlines())
    return set()


def _marcar_procesado(booking_num: str):
    with _PROCESADOS_FILE.open("a", encoding="utf-8") as f:
        f.write(booking_num + "\n")


# ---------------------------------------------------------------------------
# Extracción de #PO con Claude API
# ---------------------------------------------------------------------------

_PROMPT_SISTEMA_PO = """\
Eres un asistente especializado en extraer datos de documentos de reservas marítimas (Maersk Booking Confirmation).

Tu única tarea es localizar el número de Orden de Compra del cliente (PO Number / Purchase Order).

Este número aparece típicamente:
- Al final del documento, en texto pequeño, en una sección de referencias o notas
- Etiquetado como: "PO Number", "PO#", "P.O.", "Purchase Order", "Customer Reference",
  "Buyer's Reference", "Customer PO", "Reference Number", o variantes similares
- El formato es numérico o alfanumérico (ej: 196893, OC-00196893, PO-2024-001)

Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional ni markdown:
{
  "numero_po": "<número exacto tal como aparece en el documento, o null si no se encuentra>",
  "certeza": <entero 0-100>,
  "campo_origen": "<nombre exacto del campo donde se encontró, ej: 'Customer Reference'>"
}
"""


def extraer_po_de_booking(ruta_pdf: str) -> str | None:
    """
    Envía el PDF del booking a Claude y retorna el número de PO extraído.
    Retorna None si no se encontró o hubo error.
    """
    try:
        import anthropic
    except ImportError:
        print("[ERROR] SDK de Anthropic no instalado. Ejecuta: pip install anthropic")
        return None

    try:
        datos = Path(ruta_pdf).read_bytes()
        b64   = base64.standard_b64encode(datos).decode("utf-8")

        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        resp = client.messages.create(
            model=CLAUDE_MODELO,
            max_tokens=300,
            system=[{
                "type": "text",
                "text": _PROMPT_SISTEMA_PO,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": "Extrae el número de PO de este Maersk Booking Confirmation.",
                    },
                ],
            }],
        )

        texto = resp.content[0].text.strip()

        # Parsear JSON con la misma estrategia robusta del clasificador
        import re
        candidatos = [texto]
        limpio = re.sub(r"^```(?:json)?\s*", "", texto).strip()
        limpio = re.sub(r"\s*```$", "", limpio)
        if limpio != texto:
            candidatos.append(limpio)
        m = re.search(r"\{.*\}", texto, re.DOTALL)
        if m and m.group() not in candidatos:
            candidatos.append(m.group())

        for c in candidatos:
            try:
                data = json.loads(c)
                po      = data.get("numero_po")
                certeza = data.get("certeza", 0)
                campo   = data.get("campo_origen", "?")
                if po:
                    print(f"  [Claude] PO: {po}  certeza: {certeza}%  campo: {campo}")
                    return str(po).strip()
                break
            except json.JSONDecodeError:
                continue

        print("  [Claude] No se encontró #PO en el booking")
        return None

    except Exception as e:
        print(f"  [ERROR] Claude API al extraer PO: {type(e).__name__}: {e}")
        return None


# ---------------------------------------------------------------------------
# Playwright — autenticación interactiva (guarda cookies para reusar)
# ---------------------------------------------------------------------------

def _login_interactivo(playwright):
    """
    Abre Chromium en modo visible para que el usuario inicie sesión manualmente.
    Guarda el estado del contexto (cookies + localStorage) en _SESSION_FILE.
    """
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page    = context.new_page()

    print()
    print("=" * 60)
    print("AUTENTICACIÓN MAERSK — INICIO DE SESIÓN MANUAL")
    print("=" * 60)
    print("Se abrió el navegador. Por favor:")
    print("  1. Inicia sesión en Maersk con tu cuenta empresarial")
    print("  2. Navega hasta ver la lista de bookings")
    print("  3. Vuelve aquí y presiona ENTER para guardar la sesión")
    print("=" * 60)

    page.goto(MAERSK_URL_BOOKINGS)

    input("\nPresiona ENTER cuando hayas iniciado sesión y veas los bookings... ")

    state = context.storage_state()
    _SESSION_FILE.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    print(f"  Sesión guardada: {_SESSION_FILE}")

    browser.close()
    return state


# ---------------------------------------------------------------------------
# Playwright — extracción de TODOS los bookings (todas las pestañas + scroll)
# ---------------------------------------------------------------------------

def _leer_total_en_tab(page, tab_texto: str) -> int:
    """
    Lee el número entre paréntesis del tab, ej: 'Ocean FCL (122)' → 122.
    Retorna 0 si no se puede leer.
    """
    try:
        texto = page.evaluate(f"""
            () => {{
                const tabs = document.querySelectorAll('[role="tab"], button, a');
                for (const t of tabs) {{
                    const txt = t.textContent || '';
                    if (txt.includes('{tab_texto}')) {{
                        const m = txt.match(/\\((\\d+)\\)/);
                        return m ? parseInt(m[1]) : 0;
                    }}
                }}
                return 0;
            }}
        """)
        return int(texto or 0)
    except Exception:
        return 0


def _quitar_filtro_fecha(page):
    """
    Elimina el filtro 'Saliendo: Siguientes 14 días' (y cualquier otro filtro activo)
    para ver TODOS los bookings sin restricción de fecha.
    """
    # Buscar chips/tags de filtro con una X para cerrar
    selectores_borrar = [
        "button[aria-label*='remove' i]",
        "button[aria-label*='clear' i]",
        "button[aria-label*='eliminar' i]",
        "[data-testid*='filter-chip'] button",
        "[data-testid*='filter-tag'] button",
        # La X dentro del chip de "Siguientes 14 días"
        "button:has-text('Siguientes 14 días') ~ button",
        "button:has-text('Next 14 days') ~ button",
    ]

    # Intentar clic en los botones de cierre de chips de filtro
    for selector in selectores_borrar:
        try:
            btns = page.query_selector_all(selector)
            for btn in btns:
                btn.click()
                page.wait_for_timeout(500)
        except Exception:
            continue

    # También intentar el botón "Clear all filters" si existe
    for selector in [
        "button:has-text('Clear all')",
        "button:has-text('Limpiar')",
        "button:has-text('Reset')",
        "button:has-text('Borrar filtros')",
    ]:
        try:
            btn = page.query_selector(selector)
            if btn:
                btn.click()
                page.wait_for_timeout(1_000)
        except Exception:
            continue


def _scroll_hasta_cargar_todo(page, total_esperado: int, tab_nombre: str = "") -> list[str]:
    """
    Hace scroll infinito / click en 'Load more' hasta que el número de
    bookings cargados en el DOM sea igual a total_esperado (o no aumenta más).

    Retorna la lista de todos los números de booking encontrados.
    """
    def _extraer_actuales() -> set:
        nums = page.evaluate("""
            () => {
                const set = new Set();
                document.querySelectorAll('a').forEach(a => {
                    const t = (a.textContent || '').trim();
                    if (/^\\d{9,}$/.test(t)) set.add(t);
                });
                return Array.from(set);
            }
        """)
        return set(nums or [])

    vistos_antes = set()
    intentos_sin_cambio = 0
    MAX_INTENTOS_SIN_CAMBIO = 4

    while True:
        actuales = _extraer_actuales()
        n = len(actuales)

        if total_esperado > 0:
            print(f"  [{tab_nombre}] Cargados: {n} / {total_esperado}", end="\r", flush=True)
        else:
            print(f"  [{tab_nombre}] Cargados: {n}", end="\r", flush=True)

        # Condición de parada: ya tenemos todos
        if total_esperado > 0 and n >= total_esperado:
            break

        # Condición de parada: nada nuevo en varios intentos
        if actuales == vistos_antes:
            intentos_sin_cambio += 1
            if intentos_sin_cambio >= MAX_INTENTOS_SIN_CAMBIO:
                break
        else:
            intentos_sin_cambio = 0
            vistos_antes = actuales

        # 1. Intentar botón "Show more" / "Load more"
        cargado_mas = False
        for selector in [
            "button:has-text('Show more')",
            "button:has-text('Load more')",
            "button:has-text('Ver más')",
            "button:has-text('Cargar más')",
            "[data-testid*='load-more']",
            "[data-testid*='show-more']",
        ]:
            try:
                btn = page.query_selector(selector)
                if btn and btn.is_visible():
                    btn.scroll_into_view_if_needed()
                    btn.click()
                    page.wait_for_timeout(1_500)
                    cargado_mas = True
                    break
            except Exception:
                continue

        # 2. Si no había botón, hacer scroll hasta el fondo para activar carga infinita
        if not cargado_mas:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1_500)

    print()  # nueva línea tras el \r
    return list(_extraer_actuales())


def _obtener_todos_los_bookings(page) -> list[str]:
    """
    Recorre TODAS las pestañas (Ocean FCL, LCL, Aéreo) y hace scroll/paginación
    en cada una para obtener el listado completo de bookings sin importar filtros.
    """
    todos = set()

    # Quitar filtros de fecha activos para ver el historial completo
    _quitar_filtro_fecha(page)
    page.wait_for_timeout(1_500)

    # Esperar que cargue la tabla inicial
    try:
        page.wait_for_selector("table", timeout=20_000)
    except Exception:
        pass
    page.wait_for_timeout(2_000)

    # Lista de pestañas a recorrer con su texto de identificación
    tabs_a_recorrer = [
        ("Ocean FCL", "Ocean FCL"),
        ("LCL",       "LCL"),
        ("Aéreo",     "Aéreo"),
        ("Air",       "Air"),  # por si el idioma es inglés
    ]

    tabs_procesados = set()

    for tab_id, tab_texto in tabs_a_recorrer:
        if tab_id in tabs_procesados:
            continue

        # Intentar hacer clic en la pestaña
        try:
            tab_btn = page.query_selector(
                f"[role='tab']:has-text('{tab_texto}'), "
                f"button:has-text('{tab_texto}'), "
                f"a:has-text('{tab_texto}')"
            )
            if tab_btn:
                total = _leer_total_en_tab(page, tab_texto)
                if total == 0:
                    # pestaña vacía, no hacer clic ni perder tiempo
                    continue
                tab_btn.click()
                page.wait_for_timeout(2_000)
                print(f"\n  Pestaña '{tab_texto}': {total} bookings esperados")
                tabs_procesados.add(tab_id)
            else:
                # Primera pestaña activa por defecto (Ocean FCL ya está activa)
                if tab_id == "Ocean FCL":
                    total = _leer_total_en_tab(page, "Ocean FCL")
                    print(f"\n  Pestaña 'Ocean FCL' (activa por defecto): {total} bookings esperados")
                else:
                    continue
        except Exception:
            continue

        # Scrollear hasta cargar todos los bookings de esta pestaña
        nums = _scroll_hasta_cargar_todo(page, total, tab_texto)
        todos.update(nums)
        print(f"  Pestaña '{tab_texto}': {len(nums)} bookings cargados")

    return list(todos)


# ---------------------------------------------------------------------------
# Playwright — descarga de PDF de un booking individual
# ---------------------------------------------------------------------------

_SELECTORES_DESCARGA = [
    # Texto exacto visible en el botón
    "button:has-text('Download')",
    "button:has-text('Descargar')",
    "button:has-text('Print')",
    # Links con texto
    "a:has-text('Download booking')",
    "a:has-text('Booking confirmation')",
    # Atributos
    "[data-testid*='download' i]",
    "[aria-label*='download' i]",
    "a[download]",
]


def _descargar_pdf_booking(page, booking_num: str, dir_temp: str) -> str | None:
    """
    Navega a la página del booking y descarga el PDF de confirmación.
    Reintenta hasta _MAX_REINTENTOS_DESCARGA veces con backoff exponencial.
    Retorna la ruta local del PDF o None si falló en todos los intentos.
    """
    url_booking = f"https://www.maersk.com/booking/{booking_num}"
    print(f"  URL: {url_booking}")

    for intento in range(1, _MAX_REINTENTOS_DESCARGA + 1):
        try:
            page.goto(url_booking, wait_until="domcontentloaded", timeout=60_000)
        except Exception as e:
            print(f"  [WARN] Timeout cargando página (intento {intento}): {e}")

        page.wait_for_timeout(3_000)

        # Intento A: botón de descarga directo en la página
        ruta_pdf = _intentar_descarga(page, booking_num, dir_temp)
        if ruta_pdf:
            return ruta_pdf

        # Intento B: pestaña/sección "Documents"
        for selector_tab in [
            "button:has-text('Documents')",
            "a:has-text('Documents')",
            "[role='tab']:has-text('Documents')",
        ]:
            try:
                tab = page.query_selector(selector_tab)
                if tab:
                    tab.click()
                    page.wait_for_timeout(2_000)
                    ruta_pdf = _intentar_descarga(page, booking_num, dir_temp)
                    if ruta_pdf:
                        return ruta_pdf
            except Exception:
                continue

        # Intento C: menú de 3 puntos (⋮)
        try:
            menu_btn = page.query_selector(
                "button[aria-label*='more' i], button:has-text('⋮'), button:has-text('...')"
            )
            if menu_btn:
                menu_btn.click()
                page.wait_for_timeout(1_000)
                ruta_pdf = _intentar_descarga(page, booking_num, dir_temp)
                if ruta_pdf:
                    return ruta_pdf
        except Exception:
            pass

        # No se encontró — esperar antes de reintentar (backoff 5s, 10s, 20s...)
        if intento < _MAX_REINTENTOS_DESCARGA:
            espera = 5 * (2 ** (intento - 1))
            print(f"  [REINTENTO {intento}/{_MAX_REINTENTOS_DESCARGA}] Esperando {espera}s antes de reintentar...")
            time.sleep(espera)

    print(f"  [WARN] No se encontró botón de descarga para booking {booking_num} tras {_MAX_REINTENTOS_DESCARGA} intentos")
    return None


def _intentar_descarga(page, booking_num: str, dir_temp: str) -> str | None:
    """
    Intenta hacer click en cada selector de descarga conocido y captura el archivo.
    """
    for selector in _SELECTORES_DESCARGA:
        try:
            btn = page.query_selector(selector)
            if not btn:
                continue

            with page.expect_download(timeout=30_000) as dl_info:
                btn.click()

            download = dl_info.value
            nombre   = download.suggested_filename or f"Booking_Maersk_{booking_num}.pdf"
            ruta     = str(Path(dir_temp) / nombre)
            download.save_as(ruta)
            print(f"  Descargado: {nombre}")
            return ruta

        except Exception:
            continue

    return None


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------

def procesar_bookings(forzar_login: bool = False, desde: str | None = None):
    """
    Escanea Maersk, descarga PDFs y sube a SharePoint.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[ERROR] Playwright no instalado. Ejecuta:")
        print("  pip install playwright")
        print("  playwright install chromium")
        return

    procesados = _cargar_procesados()
    cliente    = GraphClient()

    with sync_playwright() as pw:

        # --- Sesión ---
        if forzar_login or not _SESSION_FILE.exists():
            state = _login_interactivo(pw)
        else:
            state = json.loads(_SESSION_FILE.read_text(encoding="utf-8"))

        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(
            storage_state=state,
            accept_downloads=True,
            user_agent=_USER_AGENT,
            # Ocultar señales de automatización
            extra_http_headers={"Accept-Language": "es-CR,es;q=0.9,en;q=0.8"},
        )
        # Enmascarar WebDriver para evitar detección headless
        context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        """)
        page = context.new_page()

        try:
            # --- Navegar a lista de bookings ---
            print(f"Navegando a {MAERSK_URL_BOOKINGS}...")
            try:
                page.goto(MAERSK_URL_BOOKINGS, wait_until="networkidle", timeout=60_000)
            except Exception:
                page.goto(MAERSK_URL_BOOKINGS, timeout=60_000)
                page.wait_for_timeout(5_000)

            # Verificar sesión activa
            if any(kw in page.url.lower() for kw in ("login", "signin", "auth")):
                print("[WARN] Sesión expirada. Renovando...")
                browser.close()
                state = _login_interactivo(pw)
                browser = pw.chromium.launch(headless=True)
                context = browser.new_context(storage_state=state, accept_downloads=True)
                page    = context.new_page()
                page.goto(MAERSK_URL_BOOKINGS, wait_until="networkidle", timeout=60_000)

            # --- Extraer TODOS los bookings (todas las pestañas + scroll completo) ---
            print("Cargando todos los bookings...")
            bookings = _obtener_todos_los_bookings(page)
            print(f"Total bookings encontrados: {len(bookings)}")

            nuevos = [b for b in bookings if b not in procesados]

            # --desde: saltar bookings hasta encontrar el número indicado
            if desde and desde in nuevos:
                idx = nuevos.index(desde)
                print(f"Reanudando desde booking {desde} (posición {idx + 1}/{len(nuevos)})")
                nuevos = nuevos[idx:]
            elif desde:
                print(f"[WARN] --desde {desde} no está en la lista de nuevos; se procesa todo")

            print(f"Bookings nuevos a procesar: {len(nuevos)}")

            if not nuevos:
                print("No hay bookings nuevos para procesar.")
                return

            resumen_ok      = []
            resumen_errores = []

            with tempfile.TemporaryDirectory() as dir_temp:
                for n_procesado, booking_num in enumerate(nuevos, start=1):
                    print(f"\n{'='*55}")
                    print(f"[{n_procesado}/{len(nuevos)}] Booking: {booking_num}")

                    try:
                        # 1. Descargar PDF (con reintentos internos)
                        ruta_pdf = _descargar_pdf_booking(page, booking_num, dir_temp)
                        if not ruta_pdf:
                            resumen_errores.append(f"{booking_num}: no se pudo descargar PDF")
                        else:
                            # 2. Extraer #PO con Claude
                            numero_po = extraer_po_de_booking(ruta_pdf)
                            if not numero_po:
                                print(f"  [WARN] #PO no encontrado — guardando en SIN_PO")
                                carpeta_oc = "SIN_PO"
                            else:
                                carpeta_oc = cliente.buscar_carpeta_oc(
                                    formatear_numero_oc(numero_po)
                                )

                            # 3. Construir ruta SharePoint
                            nombre_archivo = f"Booking_Maersk_{booking_num}.pdf"
                            ruta_sp = (
                                f"{SHAREPOINT_CARPETA_OCS}/{carpeta_oc}"
                                f"/4. DOCUMENTACION/{MAERSK_CARPETA_BOOKING}"
                                f"/{nombre_archivo}"
                            )

                            if MODO_MOCK:
                                print(f"  [MOCK] Subiría: {ruta_sp}")
                            else:
                                # 4. Crear carpeta y subir
                                cliente.crear_carpeta_si_no_existe(ruta_sp)
                                ruta_final = cliente.resolver_ruta_versionada(ruta_sp)
                                cliente.subir_archivo(ruta_pdf, ruta_final)
                                _log_subida(booking_num, ruta_final)
                                print(f"  OK  PO: {numero_po or 'N/A'}  →  {ruta_final}")

                            # Checkpoint: marcar como procesado inmediatamente
                            _marcar_procesado(booking_num)
                            resumen_ok.append(booking_num)

                    except Exception as e:
                        msg = f"{booking_num}: {type(e).__name__}: {e}"
                        print(f"  [ERROR] {msg}")
                        traceback.print_exc()
                        resumen_errores.append(msg)
                        # No marcamos como procesado si hubo error — se reintentará la próxima vez

                    # --- Delay anti-rate-limit entre bookings ---
                    if n_procesado < len(nuevos):
                        delay = random.uniform(_DELAY_MIN, _DELAY_MAX)

                        # Pausa larga cada N bookings
                        if n_procesado % _PAUSA_LARGA_CADA == 0:
                            print(f"  [Pausa] {n_procesado} bookings procesados — esperando {_PAUSA_LARGA_SEG:.0f}s...")
                            time.sleep(_PAUSA_LARGA_SEG)
                        else:
                            time.sleep(delay)

        finally:
            browser.close()

    # --- Resumen final ---
    print(f"\n{'='*55}")
    print(f"RESUMEN")
    print(f"  Subidos:  {len(resumen_ok)}")
    if resumen_ok:
        for b in resumen_ok:
            print(f"    ✓ {b}")
    print(f"  Errores:  {len(resumen_errores)}")
    if resumen_errores:
        for e in resumen_errores:
            print(f"    ✗ {e}")
    print("=" * 55)


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Monitor de bookings Maersk → SharePoint"
    )
    parser.add_argument(
        "--login",
        action="store_true",
        help="Forzar re-login (renovar sesión guardada)",
    )
    parser.add_argument(
        "--desde",
        metavar="BOOKING_NUM",
        default=None,
        help="Reanudar desde un número de booking específico (útil si se cortó)",
    )
    args = parser.parse_args()

    modo = "MOCK" if MODO_MOCK else "REAL"
    print(f"Monitor Maersk Bookings [{modo}]")
    print(f"Carpeta SharePoint: {MAERSK_CARPETA_BOOKING}")
    print(f"Delay entre bookings: {_DELAY_MIN}-{_DELAY_MAX}s  |  Pausa larga cada {_PAUSA_LARGA_CADA} bookings ({_PAUSA_LARGA_SEG:.0f}s)")
    print()

    procesar_bookings(forzar_login=args.login, desde=args.desde)


if __name__ == "__main__":
    main()
