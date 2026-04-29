"""
Cola de aprobacion de sugerencias de respuesta a proveedores.

Genera un archivo .txt por cada sugerencia pendiente en la carpeta
pendientes_aprobacion/. El usuario edita el texto, cambia la linea
DECISION: PENDIENTE a DECISION: APROBAR o DECISION: RECHAZAR,
y el watcher_aprobaciones.py detecta el cambio y ejecuta la accion.

Uso interno — no se llama directamente.
"""

import json
import re
from datetime import datetime
from pathlib import Path

_DIR_BASE      = Path(__file__).parent / "pendientes_aprobacion"
_DIR_HISTORIAL = _DIR_BASE / "historial"

_MARCA_DECISION = "DECISION:"
_MARCA_META     = "__META__:"  # linea oculta con JSON de metadatos


def _asegurar_carpetas():
    _DIR_BASE.mkdir(exist_ok=True)
    _DIR_HISTORIAL.mkdir(exist_ok=True)


def encolar_sugerencia(remitente: str, asunto: str, sugerencia: str, numero_po: str = "") -> Path:
    """
    Guarda la sugerencia como .txt en pendientes_aprobacion/.
    Retorna la ruta del archivo generado.
    """
    _asegurar_carpetas()

    ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
    po_safe = re.sub(r'[\\/*?:"<>|]', "", numero_po).strip() if numero_po else ""
    if po_safe:
        # Versionar: OC-XXXXXXXX_v1.txt, OC-XXXXXXXX_v2.txt, ...
        existing = list(_DIR_BASE.glob(f"OC-{po_safe}_v*.txt"))
        nums_v = []
        for f in existing:
            m = re.search(r'_v(\d+)\.txt$', f.name)
            if m:
                nums_v.append(int(m.group(1)))
        version = (max(nums_v) + 1) if nums_v else 1
        nombre = f"OC-{po_safe}_v{version}.txt"
    else:
        nombre = f"sugerencia_{ts}.txt"
    ruta = _DIR_BASE / nombre

    asunto_re = f"RE: {asunto}" if not asunto.upper().startswith("RE:") else asunto
    ts_display = datetime.now().strftime("%Y-%m-%d %H:%M")

    meta = json.dumps({
        "remitente": remitente,
        "asunto":    asunto_re,
        "ts":        ts_display,
    }, ensure_ascii=False)

    contenido = (
        "================================================================\n"
        "SUPPLIER RESPONSE SUGGESTION\n"
        "Edit the message text freely. When ready, change DECISION below.\n"
        "================================================================\n"
        "\n"
        f"To      : {remitente}\n"
        f"Subject : {asunto_re}\n"
        f"Date    : {ts_display}\n"
        "\n"
        "----------------------------------------------------------------\n"
        f"{sugerencia}\n"
        "----------------------------------------------------------------\n"
        "\n"
        "DECISION: PENDIENTE\n"
        "\n"
        "  Opciones:\n"
        "    DECISION: APROBAR   → envia el correo al proveedor\n"
        "    DECISION: RECHAZAR  → descarta sin enviar\n"
        "\n"
        f"{_MARCA_META} {meta}\n"
    )

    ruta.write_text(contenido, encoding="utf-8")
    return ruta


def leer_decision(ruta: Path) -> str:
    """
    Lee la decision del archivo .txt.
    Retorna: 'APROBAR', 'RECHAZAR' o 'PENDIENTE'.
    """
    try:
        texto = ruta.read_text(encoding="utf-8")
        for linea in texto.splitlines():
            if linea.startswith(_MARCA_DECISION):
                valor = linea[len(_MARCA_DECISION):].strip().upper()
                if valor in ("APROBAR", "RECHAZAR"):
                    return valor
        return "PENDIENTE"
    except Exception:
        return "PENDIENTE"


def leer_metadatos(ruta: Path) -> dict:
    """Extrae remitente, asunto y ts del archivo."""
    try:
        texto = ruta.read_text(encoding="utf-8")
        for linea in texto.splitlines():
            if linea.startswith(_MARCA_META):
                json_str = linea[len(_MARCA_META):].strip()
                return json.loads(json_str)
    except Exception:
        pass
    return {}


def leer_mensaje_editado(ruta: Path) -> str:
    """
    Extrae el texto entre las lineas '---' del archivo
    (el cuerpo editable de la sugerencia).
    """
    try:
        texto  = ruta.read_text(encoding="utf-8")
        lineas = texto.splitlines()
        separadores = [i for i, l in enumerate(lineas) if l.startswith("---")]
        if len(separadores) >= 2:
            inicio = separadores[0] + 1
            fin    = separadores[1]
            return "\n".join(lineas[inicio:fin]).strip()
    except Exception:
        pass
    return ""


def archivar(ruta: Path, decision: str):
    """Mueve el archivo al historial con sufijo _APROBADO o _RECHAZADO."""
    _asegurar_carpetas()
    sufijo   = "_APROBADO" if decision == "APROBAR" else "_RECHAZADO"
    destino  = _DIR_HISTORIAL / (ruta.stem + sufijo + ruta.suffix)
    # Si ya existe un archivo con ese nombre, agregar timestamp
    if destino.exists():
        ts      = datetime.now().strftime("%H%M%S")
        destino = _DIR_HISTORIAL / (ruta.stem + sufijo + f"_{ts}" + ruta.suffix)
    ruta.rename(destino)


def listar_pendientes() -> list[Path]:
    """Retorna lista de .txt en pendientes_aprobacion/ (sin historial)."""
    _asegurar_carpetas()
    return sorted(_DIR_BASE.glob("OC-*.txt")) + sorted(_DIR_BASE.glob("sugerencia_*.txt"))
