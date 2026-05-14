"""
Reprocesa PDFs locales directamente por el pipeline de clasificación y los sube a SharePoint.

Uso:
    python reprocesar_local.py <ruta_pdf> <numero_po> [<remitente>]
    python reprocesar_local.py <ruta_pdf> <numero_po> [<remitente>] --dry-run

Ejemplos:
    python "c:\\...\\reprocesar_local.py" "C:\\ruta\\12194.pdf" 201093
    python "c:\\...\\reprocesar_local.py" "C:\\ruta\\12194.pdf" 201093 "ahmet.nuri@erisflourmills.com"
    python "c:\\...\\reprocesar_local.py" "C:\\ruta\\12194.pdf" 201093 --dry-run
"""

import sys
import io
import os
import logging
from pathlib import Path
from datetime import datetime

if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

_BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_BASE))
sys.path.insert(0, str(_BASE.parent.parent / "Agente_Seguridad"))

from Nucleo.clasificador import procesar_adjunto
from Integraciones.graph_client import GraphClient

_LOG_DIR = _BASE / "Registros"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"reprocesar_local_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(str(_LOG_FILE), encoding="utf-8"),
    ],
)
log = logging.getLogger("reprocesar_local")


def reprocesar(ruta_pdf: str, numero_po: str, remitente: str = "reproceso@manual", dry_run: bool = False):
    ruta = Path(ruta_pdf)
    if not ruta.exists():
        log.error(f"No existe el archivo: {ruta_pdf}")
        return

    nombre_archivo = ruta.name
    log.info("=" * 70)
    log.info(f"Reprocesando: {nombre_archivo}")
    log.info(f"PO           : {numero_po}")
    log.info(f"Remitente    : {remitente}")
    log.info(f"Modo         : {'DRY-RUN (simulación)' if dry_run else 'REAL — se subirá a SharePoint'}")
    log.info("=" * 70)

    resultados = procesar_adjunto(
        nombre_archivo=nombre_archivo,
        numero_po_asunto=numero_po,
        ruta_local=str(ruta),
        asunto_correo=f"Reproceso manual PO {numero_po}",
    )

    if not resultados:
        log.error("procesar_adjunto no retornó resultados — revisar si se detectó la PO o si hubo error de API.")
        return

    log.info(f"\n{len(resultados)} destino(s) clasificado(s):")
    for r in resultados:
        log.info(f"  Tipo    : {r['tipo']}  ({r['certeza']}% — {r['metodo_clasificacion']})")
        log.info(f"  Destino : {r['ruta_sharepoint']}")
        if r.get("inconsistencias"):
            for inc in r["inconsistencias"]:
                log.warning(f"  [INCONSISTENCIA] {inc.get('campo','')}: {inc.get('descripcion','')}")

    if dry_run:
        log.info("\n[DRY-RUN] No se subió nada.")
        return

    gc = GraphClient()
    subidos = 0
    for r in resultados:
        ruta_sp   = r["ruta_sharepoint"]
        ruta_orig = r.get("ruta_local") or str(ruta)
        try:
            gc.crear_carpeta_si_no_existe(ruta_sp)
            gc.subir_archivo(ruta_orig, ruta_sp)
            log.info(f"  [OK] Subido: {ruta_sp}")
            subidos += 1
        except Exception as e:
            log.error(f"  [ERROR] Al subir '{ruta_sp}': {e}")

    log.info(f"\nRESUMEN: {subidos}/{len(resultados)} archivos subidos.")
    log.info(f"Log: {_LOG_FILE}")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    dry_run = "--dry-run" in sys.argv

    if len(args) < 2:
        print("Uso: python reprocesar_local.py <ruta_pdf> <numero_po> [<remitente>] [--dry-run]")
        sys.exit(1)

    ruta_pdf_arg   = args[0]
    numero_po_arg  = args[1]
    remitente_arg  = args[2] if len(args) >= 3 else "reproceso@manual"

    reprocesar(ruta_pdf_arg, numero_po_arg, remitente_arg, dry_run=dry_run)
