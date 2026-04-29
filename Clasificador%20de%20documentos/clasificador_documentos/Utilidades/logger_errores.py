"""
logger_errores.py
-----------------
Logger centralizado de problemas del clasificador de documentos.

Registra en registros_eventos.log únicamente eventos que representan
un problema: rechazos de seguridad, errores de API, fallos al guardar
adjuntos, ciclos fallidos, SKIPs, ABORTADOs, FALLBACKs, etc.

El flujo exitoso NO se toca — este módulo es solo para diagnóstico de fallos.

Uso desde cualquier módulo:
    from logger_errores import log_error, log_advertencia

    log_error("agente_seguridad", "SEG-003", "BL PO 196893.pdf", "Magic bytes no coinciden con .pdf")
    log_advertencia("clasificador_claude", "CLAUDE-004", nombre_archivo="factura.pdf", detalle="JSON inválido")
    log_evento("SKIP", "recopilador", archivo="factura.pdf", remitente="x@x.com", asunto="RE: docs", detalle="sin clasificacion")
"""

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Configuración del archivo de log
# ---------------------------------------------------------------------------
_LOG_PATH = Path(__file__).parent / "registros_eventos.log"

_logger = logging.getLogger("registros_eventos")

if not _logger.handlers:
    _handler = RotatingFileHandler(
        _LOG_PATH,
        maxBytes=5 * 1024 * 1024,   # 5 MB por archivo
        backupCount=5,               # hasta registros_eventos.log.5
        encoding="utf-8",
    )
    _handler.setFormatter(logging.Formatter("%(message)s"))
    _logger.addHandler(_handler)
    _logger.setLevel(logging.DEBUG)
    _logger.propagate = False


# ---------------------------------------------------------------------------
# Helpers internos
# ---------------------------------------------------------------------------

def _s(valor: str) -> str:
    """Sanitiza un valor para que no rompa el formato de una sola línea."""
    return str(valor).replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()


# ---------------------------------------------------------------------------
# Funciones públicas
# ---------------------------------------------------------------------------

def log_error(
    modulo: str,
    codigo: str,
    archivo: str = "-",
    detalle: str = "",
    remitente: str = "-",
    asunto: str = "-",
) -> None:
    """
    Registra un error que impidió procesar un archivo o ciclo.

    Formato: timestamp | ERROR | API_ERROR | modulo | codigo | archivo | remitente | asunto | detalle
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger.error(
        f"{ts} | ERROR | API_ERROR | {_s(modulo)} | {_s(codigo)} | {_s(archivo)} | {_s(remitente)} | {_s(asunto)} | {_s(detalle)}"
    )


def log_advertencia(
    modulo: str,
    codigo: str,
    archivo: str = "-",
    detalle: str = "",
    remitente: str = "-",
    asunto: str = "-",
) -> None:
    """
    Registra una advertencia: situación degradada pero el proceso continuó
    (ej. Claude falló pero el fallback funcionó, ZIP no se pudo descomprimir).

    Formato: timestamp | ADVERTENCIA | API_ERROR | modulo | codigo | archivo | remitente | asunto | detalle
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger.warning(
        f"{ts} | ADVERTENCIA | API_ERROR | {_s(modulo)} | {_s(codigo)} | {_s(archivo)} | {_s(remitente)} | {_s(asunto)} | {_s(detalle)}"
    )


def log_evento(
    evento: str,
    modulo: str,
    archivo: str = "-",
    detalle: str = "",
    remitente: str = "-",
    asunto: str = "-",
) -> None:
    """
    Registra un evento de proceso problemático: SKIP, ABORTADO, FALLBACK.

    Formato: timestamp | ADVERTENCIA | evento | modulo | - | archivo | remitente | asunto | detalle
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger.warning(
        f"{ts} | ADVERTENCIA | {_s(evento)} | {_s(modulo)} | - | {_s(archivo)} | {_s(remitente)} | {_s(asunto)} | {_s(detalle)}"
    )


def log_info(modulo: str, mensaje: str, archivo: str = "-") -> None:
    """
    Registra un evento informativo de flujo normal (no error, no advertencia).

    Formato: timestamp | INFO | modulo | archivo | mensaje
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger.info(f"{ts} | INFO | {_s(modulo)} | {_s(archivo)} | {_s(mensaje)}")


def log_clasificacion(
    archivo: str,
    tipo: str,
    certeza: int,
    justificacion: str,
    inconsistencias: list,
    remitente: str = "-",
    asunto: str = "-",
) -> None:
    """
    Registra el resultado de una clasificación exitosa de Claude.

    Formato principal:
        timestamp | INFO | CLASIFICACION | archivo | tipo | certeza% | remitente | asunto | justificacion

    Si hay inconsistencias, agrega una línea adicional por cada una:
        timestamp | INFO | INCONSISTENCIA | archivo | severidad | campo | descripcion
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _logger.info(
        f"{ts} | INFO | CLASIFICACION | {_s(archivo)} | {_s(tipo)} | {certeza}% "
        f"| {_s(remitente)} | {_s(asunto)} | {_s(justificacion)}"
    )
    for inc in inconsistencias:
        sev   = inc.get("severidad", "-")
        campo = inc.get("campo", "-")
        desc  = inc.get("descripcion", "-")
        _logger.info(
            f"{ts} | INFO | INCONSISTENCIA | {_s(archivo)} | {_s(sev)} | {_s(campo)} | {_s(desc)}"
        )
