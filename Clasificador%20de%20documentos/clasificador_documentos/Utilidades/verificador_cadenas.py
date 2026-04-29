"""
verificador_cadenas.py
----------------------
Detecta si un correo es parte de una cadena (RE:/RV:/FW:) y compara
PO/BL con mensajes previos de la misma conversación en Outlook.

Uso:
    from verificador_cadenas import verificar_cadena

    incs = verificar_cadena(msg_com, asunto, po_actual, bl_actual)
    # incs: lista de dicts {campo, descripcion, severidad} — vacía si no hay problemas
"""

import re
from Utilidades.logger_errores import log_advertencia

# Prefijos que indican correo de respuesta o reenvío
_RE_CADENA = re.compile(r'^\s*(re|rv|fw|fwd|res)\s*:', re.IGNORECASE)

# Patrones para extraer PO y BL del texto
_RE_PO = re.compile(r'\b(?:PO|OC)[-#\s]*(\d{5,7})\b', re.IGNORECASE)
_RE_BL = re.compile(r'\b(?:BL|B/L|HBL|MBL)[-#\s]*([A-Z0-9]{6,20})\b', re.IGNORECASE)


def _extraer_referencias(texto: str) -> tuple:
    """Extrae todos los POs y BLs encontrados en un texto."""
    pos = set(_RE_PO.findall(texto))
    bls = set(_RE_BL.findall(texto))
    return pos, bls


def verificar_cadena(msg_com, asunto: str, po_actual, bl_actual) -> list:
    """
    Si el correo es una respuesta o reenvío, busca mensajes previos en la misma
    conversación (via ConversationID de Outlook) y compara PO/BL.

    Parámetros:
        msg_com   — objeto MailItem de Outlook (win32com). Si es None, retorna [].
        asunto    — asunto del correo actual (str).
        po_actual — PO extraído del correo actual (str o None).
        bl_actual — BL extraído del correo actual (str o None).

    Retorna lista de inconsistencias [{campo, descripcion, severidad}].
    Lista vacía = sin problemas detectados o no es cadena.
    """
    if msg_com is None:
        return []

    if not _RE_CADENA.match(asunto):
        return []  # No es respuesta ni reenvío — no hay cadena que verificar

    inconsistencias = []
    try:
        conv_id = msg_com.ConversationID
        msg_id  = msg_com.EntryID
        bandeja = msg_com.Parent  # carpeta donde vive el correo

        filtro = f"[ConversationID] = '{conv_id}'"
        items  = bandeja.Items.Restrict(filtro)

        po_cadena = set()
        bl_cadena = set()
        n_previos = 0

        for item in items:
            try:
                if item.Class != 43:          # solo MailItem
                    continue
                if item.EntryID == msg_id:    # saltar el correo actual
                    continue
                subj = item.Subject or ""
                body = (item.Body or "")[:800]  # primeras líneas son suficientes
                pos_p, bls_p = _extraer_referencias(subj + " " + body)
                po_cadena.update(pos_p)
                bl_cadena.update(bls_p)
                n_previos += 1
            except Exception:
                continue

        if n_previos == 0:
            return []  # Cadena de un solo correo — nada que comparar

        # PO del correo actual no coincide con los de la cadena
        if po_actual and po_cadena and po_actual not in po_cadena:
            inconsistencias.append({
                "campo": "PO (cadena)",
                "descripcion": (
                    f"El correo actual referencia PO {po_actual}, pero correos anteriores "
                    f"de esta cadena mencionan: {', '.join(sorted(po_cadena))}"
                ),
                "severidad": "alta",
            })

        # Múltiples POs distintos dentro de la cadena (señal de mezcla de órdenes)
        if len(po_cadena) > 1:
            inconsistencias.append({
                "campo": "POs múltiples (cadena)",
                "descripcion": (
                    f"La cadena de correos contiene referencias a múltiples POs: "
                    f"{', '.join(sorted(po_cadena))}"
                ),
                "severidad": "media",
            })

        # BL del correo actual no coincide con los de la cadena
        if bl_actual and bl_cadena and bl_actual not in bl_cadena:
            inconsistencias.append({
                "campo": "BL (cadena)",
                "descripcion": (
                    f"El correo actual referencia BL {bl_actual}, pero correos anteriores "
                    f"de esta cadena mencionan: {', '.join(sorted(bl_cadena))}"
                ),
                "severidad": "media",
            })

    except Exception as e:
        log_advertencia("verificador_cadenas", "VER-001", asunto, f"{type(e).__name__}: {e}")

    return inconsistencias
