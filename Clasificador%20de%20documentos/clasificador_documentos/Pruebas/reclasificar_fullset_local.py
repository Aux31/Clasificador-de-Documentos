"""
Reprocesa PDFs locales que fueron subidos como full set sin separar.

Para cada archivo:
  1. Llama a separar_fullset() para que Claude detecte y corte los documentos.
  2. Clasifica cada fragmento con procesar_adjunto().
  3. Sube cada fragmento a su carpeta correcta en SharePoint.
  4. Registra la subida en registros_subidas.log.

Uso:
    python reclasificar_fullset_local.py
    python reclasificar_fullset_local.py --dry-run
"""

import sys
import io
import os
import logging
import tempfile
from pathlib import Path
from datetime import datetime
from urllib.parse import quote

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

_BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_BASE))
sys.path.insert(0, str(_BASE.parent.parent / "Agente_Seguridad"))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID
from Nucleo.clasificador import procesar_adjunto, formatear_numero_oc
from Nucleo.separador_fullset import separar_fullset
from Integraciones.graph_client import GraphClient

# ---------------------------------------------------------------------------
# Archivos a reprocesar: (ruta_local, numero_po, remitente)
# ---------------------------------------------------------------------------
ARCHIVOS = [
    (r"C:\Users\aux22.gg\Documents\GitHub\PL 20260513160923401.pdf", "202708", "Samuel.Delitte@ecofrost.be"),
]

# ---------------------------------------------------------------------------
# Logger
# ---------------------------------------------------------------------------
_LOG_DIR = _BASE / "Registros"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"reclasificar_fullset_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(str(_LOG_FILE), encoding="utf-8"),
    ],
)
log = logging.getLogger("reclasificar_fullset")

# Logger de subidas (mismo formato que el sistema principal)
_LOG_SUBIDAS = _BASE / "registros_subidas.log"
_logger_subidas = logging.getLogger("subidas_sharepoint")
if not _logger_subidas.handlers:
    _h = logging.FileHandler(str(_LOG_SUBIDAS), encoding="utf-8")
    _h.setFormatter(logging.Formatter("%(message)s"))
    _logger_subidas.addHandler(_h)
    _logger_subidas.setLevel(logging.INFO)


def _siguiente_consecutivo() -> int:
    if not _LOG_SUBIDAS.exists():
        return 1
    lineas = _LOG_SUBIDAS.read_text(encoding="utf-8").splitlines()
    for linea in reversed(lineas):
        parte = linea.split("|")[0].strip()
        if parte.isdigit():
            return int(parte) + 1
    return 1


def _log_subida(remitente: str, ruta_sharepoint: str):
    n  = _siguiente_consecutivo()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger_subidas.info(f"{n} | {ts} | {remitente} |  | {ruta_sharepoint}")


# ---------------------------------------------------------------------------
# Proceso principal
# ---------------------------------------------------------------------------

def procesar_archivo(ruta_pdf: str, numero_po: str, remitente: str, gc: GraphClient, dry_run: bool):
    ruta = Path(ruta_pdf)
    nombre = ruta.name
    log.info("-" * 70)
    log.info(f"Archivo : {nombre}")
    log.info(f"PO      : {numero_po}")

    with tempfile.TemporaryDirectory() as tmpdir:
        # --- 1. Separar full set ---
        log.info("Llamando a separar_fullset()...")
        fragmentos = separar_fullset(
            ruta_pdf=str(ruta),
            nombre_archivo=nombre,
            carpeta_temp=tmpdir,
        )

        if fragmentos is None:
            log.warning("separar_fullset devolvió None — Claude no detectó múltiples documentos.")
            log.warning("Procesando como documento único.")
            fragmentos_a_clasificar = [{"ruta": str(ruta), "nombre": nombre, "tipo_sugerido": None}]
        elif not fragmentos:
            log.error("separar_fullset devolvió lista vacía — error al cortar el PDF.")
            return
        else:
            log.info(f"Full set separado en {len(fragmentos)} fragmento(s):")
            for f in fragmentos:
                log.info(f"  - {f['nombre']}  ({f.get('tipo_sugerido','')}  ref={f.get('referencia','')})")
            fragmentos_a_clasificar = fragmentos

        # --- 2. Clasificar y subir cada fragmento ---
        subidos = 0
        for frag in fragmentos_a_clasificar:
            es_frag = len(fragmentos_a_clasificar) > 1
            resultados = procesar_adjunto(
                nombre_archivo=frag["nombre"],
                numero_po_asunto=numero_po,
                ruta_local=frag["ruta"],
                asunto_correo=f"Reproceso manual PO {numero_po}",
                _es_fragmento=es_frag,
                _tipo_sugerido=frag.get("tipo_sugerido"),
            )

            if not resultados:
                log.warning(f"  Sin resultados para: {frag['nombre']}")
                continue

            for r in resultados:
                ruta_sp = r["ruta_sharepoint"]

                # Reemplazar carpeta base por la real (con código de proveedor)
                oc_base = formatear_numero_oc(r["numero_po"] or numero_po)
                oc_real = gc.buscar_carpeta_oc(oc_base)
                if oc_real != oc_base:
                    ruta_sp = ruta_sp.replace(oc_base, oc_real, 1)

                log.info(f"  [{r['tipo']} — {r['certeza']}% {r['metodo_clasificacion']}]  →  {ruta_sp}")

                if r.get("inconsistencias"):
                    for inc in r["inconsistencias"]:
                        sev = inc.get("severidad", "")
                        log.warning(f"    [INCONSISTENCIA/{sev}] {inc.get('campo','')} — {inc.get('descripcion','')}")

                if dry_run:
                    log.info("    [DRY-RUN] No subido.")
                    continue

                try:
                    gc.crear_carpeta_si_no_existe(ruta_sp)
                    gc.subir_archivo(frag["ruta"], ruta_sp)
                    _log_subida(remitente, ruta_sp)
                    log.info("    [OK] Subido.")
                    subidos += 1
                except Exception as e:
                    log.error(f"    [ERROR] {e}")

        if not dry_run:
            log.info(f"Total subidos: {subidos}")


def main():
    dry_run = "--dry-run" in sys.argv
    log.info("=" * 70)
    log.info(f"RECLASIFICACIÓN DE FULL SETS — {'DRY-RUN' if dry_run else 'MODO REAL'}")
    log.info(f"Log: {_LOG_FILE}")
    log.info("=" * 70)

    gc = GraphClient()

    for ruta_pdf, numero_po, remitente in ARCHIVOS:
        procesar_archivo(ruta_pdf, numero_po, remitente, gc, dry_run)

    log.info("=" * 70)
    log.info("Proceso completado.")


if __name__ == "__main__":
    main()
