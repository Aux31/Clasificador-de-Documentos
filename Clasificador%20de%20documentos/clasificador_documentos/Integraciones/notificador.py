"""
Notificador de Teams para el pipeline de clasificación de documentos.

Canales:
  - pruebas   → sugerencias de respuesta al proveedor, confirmaciones de subida
  - problemas → errores del pipeline, excepciones, fallos de clasificación

Uso:
    from notificador import notificar_sugerencia, notificar_error

    notificar_sugerencia(remitente, asunto, sugerencia)
    notificar_error("Falló la subida a SharePoint", detalle=str(e), paso="graph_client.subir_archivo")
"""

import requests
from datetime import datetime
from configuracion.ajustes import TEAMS_WEBHOOK_PRUEBAS, TEAMS_WEBHOOK_PROBLEMAS


def _enviar(webhook_url: str, payload: dict) -> bool:
    """Envía un payload JSON al webhook de Teams. Retorna True si fue exitoso."""
    if not webhook_url:
        return False
    try:
        r = requests.post(webhook_url, json=payload, timeout=10, verify=False)
        return r.status_code < 300
    except Exception:
        return False


def _tarjeta_sugerencia(remitente: str, asunto: str, sugerencia: str) -> dict:
    """Adaptive Card para el canal 'pruebas' con la sugerencia de respuesta al proveedor."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    # Truncar sugerencia si es muy larga (Teams tiene límite de ~28 KB por tarjeta)
    cuerpo = sugerencia if len(sugerencia) <= 2000 else sugerencia[:1997] + "..."
    return {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.4",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "📋 Sugerencia de respuesta al proveedor",
                            "weight": "Bolder",
                            "size": "Medium",
                            "color": "Accent",
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {"title": "Proveedor", "value": remitente},
                                {"title": "Asunto",    "value": asunto},
                                {"title": "Hora",      "value": ts},
                            ],
                        },
                        {"type": "TextBlock", "text": "---", "separator": True},
                        {
                            "type": "TextBlock",
                            "text": cuerpo,
                            "wrap": True,
                            "fontType": "Monospace",
                        },
                    ],
                },
            }
        ],
    }


def _tarjeta_error(mensaje: str, detalle: str = "", paso: str = "", como_resolver: str = "") -> dict:
    """Adaptive Card para el canal 'problemas' con información de diagnóstico."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    facts = [
        {"title": "Hora", "value": ts},
    ]
    if paso:
        facts.append({"title": "Paso", "value": paso})
    body = [
        {
            "type": "TextBlock",
            "text": "🚨 Error en el pipeline de documentos",
            "weight": "Bolder",
            "size": "Medium",
            "color": "Attention",
        },
        {
            "type": "TextBlock",
            "text": mensaje,
            "wrap": True,
            "weight": "Bolder",
        },
        {"type": "FactSet", "facts": facts},
    ]
    if detalle:
        body += [
            {"type": "TextBlock", "text": "Detalle técnico:", "weight": "Bolder", "separator": True},
            {
                "type": "TextBlock",
                "text": detalle[:800] + ("..." if len(detalle) > 800 else ""),
                "wrap": True,
                "fontType": "Monospace",
            },
        ]
    if como_resolver:
        body += [
            {"type": "TextBlock", "text": "Cómo resolver:", "weight": "Bolder", "color": "Warning", "separator": True},
            {"type": "TextBlock", "text": como_resolver, "wrap": True},
        ]
    return {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.4",
                    "body": body,
                },
            }
        ],
    }


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------

def notificar_sugerencia(remitente: str, asunto: str, sugerencia: str) -> bool:
    """
    Envía la sugerencia de respuesta al proveedor al canal 'pruebas' de Teams.
    Retorna True si el envío fue exitoso.
    """
    payload = _tarjeta_sugerencia(remitente, asunto, sugerencia)
    return _enviar(TEAMS_WEBHOOK_PRUEBAS, payload)


def notificar_error(
    mensaje: str,
    detalle: str = "",
    paso: str = "",
    como_resolver: str = "",
) -> bool:
    """
    Envía un error al canal 'problemas' de Teams.

    Args:
        mensaje:       Descripción corta del error (qué pasó).
        detalle:       Traza técnica o excepción (cómo pasó).
        paso:          Nombre del módulo/función donde ocurrió.
        como_resolver: Sugerencia de solución (opcional).

    Retorna True si el envío fue exitoso.
    """
    payload = _tarjeta_error(mensaje, detalle, paso, como_resolver)
    return _enviar(TEAMS_WEBHOOK_PROBLEMAS, payload)
