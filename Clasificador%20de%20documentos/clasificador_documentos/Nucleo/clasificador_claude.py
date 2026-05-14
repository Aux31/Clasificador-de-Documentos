"""
Clasificador de documentos usando Claude API (Haiku).
Envía el archivo directamente (PDF o imagen) a Claude como documento base64.
No depende de extracción de texto local ni OCR.

Este módulo es el único que conoce al SDK de Anthropic.

Optimizaciones de tokens:
  - Prompt Caching: el system prompt de clasificación se cachea en la API (~90% ahorro
    de tokens de entrada del system prompt a partir del segundo documento por sesión).
  - Separación de tareas: la detección de inconsistencias es una segunda llamada
    independiente, solo para tipos que tienen reglas definidas (INVOICE, BL, etc.).
    Documentos como CO, MARCHAMO, EXONERACION, OTROS nunca pagan esos tokens.
"""

import base64
import json
import re
import time
from pathlib import Path
from Utilidades.logger_errores import log_error, log_advertencia, log_clasificacion, log_info
from configuracion.ajustes import (
    ANTHROPIC_API_KEY,
    CLAUDE_MODELO,
    CLAUDE_TIMEOUT_SEGUNDOS,
    CLAUDE_MAX_TOKENS,
)
from Catalogo.catalogo_tipos import TIPOS_VALIDOS
from Catalogo.checklist_excel import formatear_checklist_para_prompt
from Nucleo.extractor_texto import extraer_texto
from Nucleo.validador_campos import validar_campos

# Marcador especial para indicar fallo de API al llamador
_ERROR_API     = "__API_ERROR__"
_TIPO_INVALIDO = "__INVALID_TYPE__"

# ---------------------------------------------------------------------------
# Helper de reintentos con backoff exponencial
# ---------------------------------------------------------------------------
_MAX_REINTENTOS  = 3
_PAUSA_BASE_SEG  = 2.0   # 2s → 4s → 8s


def _llamar_api(client, **kwargs):
    """
    Llama a client.messages.create con reintentos y backoff exponencial.

    Solo reintenta errores transitorios (servidor ocupado, timeout, red).
    Errores permanentes (401, 400, ImportError) se propagan de inmediato.
    """
    try:
        from anthropic import APIStatusError, APIConnectionError, APITimeoutError
    except ImportError:
        raise  # si el SDK no está instalado, propagar

    ultimo_error = None
    for intento in range(1, _MAX_REINTENTOS + 1):
        try:
            return client.messages.create(**kwargs)

        except APIStatusError as e:
            # 529 = overloaded, 429 = rate limit → transitorios
            # 400 = prompt inválido, 401 = key inválida → permanentes
            if e.status_code in (429, 529) and intento < _MAX_REINTENTOS:
                espera = _PAUSA_BASE_SEG * (2 ** (intento - 1))
                print(f"  [CLAUDE] HTTP {e.status_code} — reintento {intento}/{_MAX_REINTENTOS} en {espera:.0f}s...")
                time.sleep(espera)
                ultimo_error = e
            else:
                raise

        except (APIConnectionError, APITimeoutError) as e:
            if intento < _MAX_REINTENTOS:
                espera = _PAUSA_BASE_SEG * (2 ** (intento - 1))
                print(f"  [CLAUDE] {type(e).__name__} — reintento {intento}/{_MAX_REINTENTOS} en {espera:.0f}s...")
                time.sleep(espera)
                ultimo_error = e
            else:
                raise

    raise ultimo_error


# ---------------------------------------------------------------------------
# Helper de parseo robusto de JSON
# ---------------------------------------------------------------------------

def _parsear_json_respuesta(texto: str, fallback: dict) -> dict:
    """
    Intenta parsear JSON de la respuesta de Claude con múltiples estrategias.

    Estrategias en orden:
      1. json.loads directo
      2. Limpiar markdown (```json ... ```) y reintentar
      3. Extraer el bloque {...} más externo con regex y reintentar
      4. Reemplazar saltos de línea literales dentro de strings y reintentar

    Si todas fallan, retorna fallback sin lanzar excepción.
    """
    candidatos = [texto]

    # Estrategia 2: quitar bloques markdown
    limpio = re.sub(r'^```(?:json)?\s*', '', texto.strip())
    limpio = re.sub(r'\s*```$', '', limpio)
    if limpio != texto:
        candidatos.append(limpio)

    # Estrategia 3: extraer el bloque {...} más externo
    m = re.search(r'\{.*\}', texto, re.DOTALL)
    if m and m.group() not in candidatos:
        candidatos.append(m.group())

    for c in candidatos:
        try:
            return json.loads(c)
        except json.JSONDecodeError:
            pass

    # Estrategia 4: reemplazar saltos de línea DENTRO de strings JSON
    # Trabaja sobre el candidato más limpio disponible
    base = candidatos[-1]
    try:
        # Sustituye \n reales dentro de valores de string por el escape \n
        def _escapar_saltos(m_str):
            return m_str.group(0).replace('\n', '\\n').replace('\r', '\\r')
        reparado = re.sub(r'"[^"\\]*(?:\\.[^"\\]*)*"', _escapar_saltos, base, flags=re.DOTALL)
        return json.loads(reparado)
    except (json.JSONDecodeError, Exception):
        pass

    return fallback


# Tipos para los que vale la pena detectar inconsistencias numéricas/lógicas.
# El resto (CO, MARCHAMO, FITOSANITARIO, PRINTER, EXONERACION, OTROS, etc.)
# no tiene reglas definidas → se omite la segunda llamada.
_TIPOS_CON_INCONSISTENCIAS = {
    "INVOICE", "PL + INV", "PACKING LIST",
    "MBL", "HBL", "BL", "AWB",
    "CO", "WEIGHT CERTIFICATE", "QUALITY CERTIFICATE", "FOB LETTER",
}

# ---------------------------------------------------------------------------
# Prompt de sistema — SOLO clasificación (sin reglas de inconsistencias).
# Se cachea en la API de Anthropic para reutilización entre documentos.
# ---------------------------------------------------------------------------
_PROMPT_SISTEMA_CLASIFICACION = """\
Eres un clasificador experto en documentos de comercio exterior para una empresa importadora costarricense.
Recibirás el documento adjunto directamente (imagen o PDF) junto con:
  1. Nombre del archivo
  2. Asunto del correo en que llegó el documento

Determina a qué tipo corresponde el documento. Los tipos válidos son exactamente estos:
MBL, HBL, BL, AWB, INVOICE, PACKING LIST, PL + INV, CO, WEIGHT CERTIFICATE,
QUALITY CERTIFICATE, FUMIGATION CERTIFICATE, FITOSANITARIO, ZOOSANITARIO, FOB LETTER,
PRINTER, EXONERACION, MARCHAMO, OTROS.

Para identificar el tipo, usa el siguiente checklist de campos que debe contener cada documento.
Si el documento contiene la mayoría de los campos listados para un tipo, clasifícalo como ese tipo.

{checklist}

Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional, sin markdown.
Formato exacto:
{{
  "tipo": "<TIPO_EXACTO>",
  "certeza": <0-100>,
  "justificacion": "<1-2 oraciones breves>",
  "nombre_proveedor": "<nombre de la empresa emisora del documento>",
  "es_borrador": true | false,
  "numero_po": "<número de OC/PO que aparece en el documento, solo dígitos, o null si no aparece>"
}}

Reglas:
- "tipo" debe ser exactamente uno de los nombres del catálogo (mayúsculas, respeta espacios y signos).
- "certeza" es un entero 0-100:
    - 80-100: el documento claramente pertenece a ese tipo
    - 50-79: hay indicios moderados pero cierta ambigüedad
    - 0-49: poca evidencia, clasificación a riesgo
- "justificacion" explica en 1-2 oraciones los elementos clave que determinaron el tipo.
- Para "PL + INV": úsalo SOLO si el documento tiene TANTO precios/montos COMO pesos/empaque físico.
- Para distinguir MBL/HBL/BL: si el emisor es una naviera directa → MBL o BL; si es un agente de carga → HBL.
- EUR.1, Form A y GSP Certificate son certificados de origen preferenciales → clasifícalos como "CO".
- Si el documento no encaja en ningún tipo conocido, usa "OTROS".
- Clasifica el documento por lo que EL DOCUMENTO ES, no por los datos que referencia o menciona internamente.
  Algunos documentos contienen números de factura, números de BL, imágenes de facturas u otros datos
  de embarque, pero no son ese tipo de documento. Los siguientes deben ir siempre a "OTROS":
    - Account Statement / Estado de Cuenta / Statement of Account: resumen de saldos o pagos entre partes,
      aunque incluya datos de facturas o montos.
    - DAE / Declaración de Exportación / Export Declaration / Documento de Acompañamiento de Exportación:
      trámite aduanero del país exportador, aunque contenga número de BL, contenedor o shipper.
    - Export Statement: resumen de exportación emitido por el proveedor, no es una factura ni un BL.
    - SOA / Statement of Account: igual que Account Statement → OTROS.
    - Request for Payment / Solicitud de Pago: comunicación de cobro del proveedor, no es una factura comercial.
    - Health Certificate / Certificado de Salud / Sanitary Certificate / Certificado Sanitario /
      Certificate of Health / Health and Sanitary Certificate: certificado que acredita que la mercancía
      cumple requisitos sanitarios o de salud pública. NO es un Certificate of Origin (CO) aunque tenga
      estructura similar o mencione el país de origen. Siempre → OTROS.
- IMPORTANTE: "CO" (Certificate of Origin) acredita el PAÍS DE FABRICACIÓN para fines aduanales o
  arancelarios. Un certificado sanitario o de salud acredita que el producto es APTO PARA CONSUMO o
  cumple normas de salud — son documentos distintos aunque ambos sean "certificados". Si el documento
  dice "Health Certificate", "Sanitary Certificate", "Certificate of Health" o similar, usa "OTROS".
- "nombre_proveedor": nombre de la empresa que emitió el documento (shipper, exporter, seller o equivalente).
  Si el documento es emitido por una naviera o agente de carga, usa el nombre del shipper.
  Si no se puede determinar con certeza, usa "".
- "es_borrador": true si el documento en sí está marcado como borrador, draft, preliminary, non-negotiable,
  specimen, copy, muestra o equivalente en cualquier idioma. También true para un CO (Certificate of
  Origin) que no tenga sello oficial, firma o estampilla de la autoridad emisora (un CO sin sello es
  siempre un borrador, aunque no diga "draft" explícitamente). false en caso contrario.
  REGLA CRÍTICA PARA CO: Para un CO, el valor por defecto es true (borrador). Solo pon false si puedes
  confirmar visualmente la presencia de un sello, estampilla o sello húmedo de la autoridad emisora
  (cámara de comercio, aduana, organismo gubernamental, etc.) en el documento. La duda siempre va a true.
  IMPORTANTE: el asunto del correo NO debe influir en este campo — evalúa únicamente el contenido del documento.
  ADVERTENCIA ESPECIAL: algunos proveedores usan plantillas de correo donde el asunto siempre incluye
  frases como "DRAFT DOCUMENT FOR [EMPRESA] SALES CONTRACT" o "DRAFT INVOICE" aunque el documento adjunto
  sea la factura o el packing list definitivos. Si el documento adjunto NO tiene la palabra DRAFT, BORRADOR,
  PRELIMINARY, SPECIMEN o NON-NEGOTIABLE impresa o estampada sobre el propio documento, responde false.
  En caso de duda para facturas e INVOICE: prefiere false (definitivo).
- "numero_po": si en el cuerpo del documento aparece un número de orden de compra (Purchase Order, PO, OC), extrae solo los dígitos. Si no aparece, usa null.\
"""

# ---------------------------------------------------------------------------
# Prompt de sistema — SOLO detección de inconsistencias (segunda llamada).
# Se invoca solo para tipos en _TIPOS_CON_INCONSISTENCIAS.
# También se cachea.
# ---------------------------------------------------------------------------
_PROMPT_SISTEMA_INCONSISTENCIAS = """\
Eres un revisor experto en documentos de comercio exterior para una empresa importadora costarricense.
Se te proporcionará un documento ya clasificado. Tu única tarea es detectar inconsistencias internas.
Revisa SOLO lo que el documento declara explícitamente — no inferas datos ausentes como error.

Contexto de negocio: la empresa importadora está en Costa Rica. Los puertos de destino esperados
para carga marítima son Puerto Limón (CRLIM) y Puerto Caldera (CRCAL). Aeropuerto: SJO.

Tipos de inconsistencias a buscar según el tipo de documento:

INVOICE / PL + INV:
  - Suma de líneas (qty × unit price) no coincide con el subtotal declarado
  - Subtotal + impuestos/cargos no coincide con el total declarado
  - Fechas contradictorias (ej: fecha de embarque anterior a la fecha de factura)
  - Números de PO o referencia distintos entre el encabezado y el cuerpo
  - Monedas mezcladas sin conversión explicada (ej: líneas en USD y EUR sin tipo de cambio)
  - Cantidad total de cajas/piezas en líneas ≠ cantidad total declarada en resumen
  - INCOTERMS presentes pero no son un término válido Incoterms 2020
    (válidos: EXW, FCA, FAS, FOB, CFR, CIF, CPT, CIP, DAP, DPU, DDP)
  - Moneda de pago diferente a la moneda de precio unitario sin explicación

PACKING LIST / PL + INV:
  - Peso neto mayor que el peso bruto en alguna línea o en el total
  - Suma de pesos individuales no coincide con el total declarado
  - Número de cajas/unidades por línea no coincide con el total declarado
  - Descripción de mercancía diferente entre el encabezado y el detalle
  NOTA IMPORTANTE sobre pesos: Muchos proveedores usan notación europea donde la coma (,) es separador
  de miles y el punto (.) es decimal. Por ejemplo, "288,000 kg" puede significar 288 kg (no 288,000 kg).
  Antes de marcar una diferencia de peso como inconsistencia, evalúa si los valores son físicamente
  razonables para el tipo de mercancía y el número de cajas/contenedores declarados. Una diferencia
  aparentemente abismal (ej: 64,890 kg de tara para 100 cajas) probablemente es un problema de notación,
  no un error real. Solo marcar si la diferencia es irrazonable incluso considerando notación europea.

MBL / HBL / BL:
  - Puerto de carga y puerto de descarga son el mismo
  - Fecha de embarque posterior a la fecha de llegada declarada
  - Número de contenedor con formato inválido (debe ser 4 letras + 7 dígitos)
  - Shipper y consignee son la misma entidad
  - Número de BL con prefijo de naviera distinto al emisor del documento
  - Puerto de descarga que no sea Costa Rica cuando el contexto lo indica
    (esperar: Puerto Limón, CRLIM, Puerto Caldera, CRCAL, Limón, Caldera)
  - Cantidad de contenedores en texto ≠ cantidad indicada en el resumen del BL
  - Datos de contenedor, lugar o valor completados con "X" en lugar de valores reales
  NOTA: Es completamente normal que la Date of Issue del BL sea posterior al Shipped on Board Date,
  ya que la naviera emite el documento después de que la mercancía embarca. NO marcar esto como error.

AWB:
  - Aeropuerto de origen igual al aeropuerto de destino
  - Peso chargeable menor que el peso bruto declarado
  - Fecha de emisión posterior a la fecha de vuelo

CO (Certificate of Origin):
  - País de origen declarado diferente al país del emisor del certificado
  - Descripción de mercancía inconsistente con la factura referenciada (si se menciona)
  - Certificado marcado como "no válido", "draft", "non valid", "specimen" o equivalente en cualquier idioma
  - Fecha de emisión del certificado posterior a la fecha de embarque declarada
  - Número de factura o referencia mencionado en el CO diferente al del BL o PO del correo

WEIGHT CERTIFICATE / QUALITY CERTIFICATE:
  - Peso neto mayor que el peso bruto
  - Fecha de inspección/emisión posterior a la fecha de embarque (si se menciona)
  - Contenedor u orden referenciada en el certificado que no coincide con el encabezado del documento

FOB LETTER:
  - Valor FOB declarado con moneda distinta a la de la factura referenciada (si se menciona)
  - Número de contenedor o factura referenciado no corresponde al formato estándar
  - Valor FOB mayor que el valor CIF si ambos aparecen en el mismo documento

Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional, sin markdown.
IMPORTANTE: Los valores de "campo" y "descripcion" deben ser strings en una sola línea — sin saltos de línea (\n) dentro del texto.
Formato exacto:
{{
  "inconsistencias": [
    {{
      "campo": "<nombre del campo o sección afectada>",
      "descripcion": "<explicación concisa de la inconsistencia con los valores reales que se contradicen>",
      "severidad": "alta" | "media" | "baja"
    }}
  ]
}}

Severidades: "alta" = dato erróneo que invalida el documento o impide el despacho aduanal,
"media" = discrepancia importante que requiere corrección,
"baja" = anomalía menor o posible error tipográfico.
Si no hay inconsistencias, devuelve: {{"inconsistencias": []}}
No incluyas ningún campo adicional en el JSON.\
"""

# ---------------------------------------------------------------------------
# Prompt de sistema — validación CRUZADA entre documentos del mismo correo.
# Tercera llamada opcional, se ejecuta solo si hay ≥ 2 documentos con texto.
# ---------------------------------------------------------------------------
_PROMPT_SISTEMA_CRUZADO = """\
Eres un revisor experto en comercio exterior para una empresa importadora costarricense.
Se te proporcionan extractos de texto de múltiples documentos de un mismo embarque, ya clasificados.
Tu única tarea es detectar INCONSISTENCIAS ENTRE documentos — campos que deberían coincidir pero no lo hacen.

Campos que deben ser consistentes entre documentos del mismo embarque:

1. NÚMERO DE CONTENEDOR
   Debe ser idéntico en: BL/MBL/HBL, INVOICE, PACKING LIST, CO, WEIGHT CERTIFICATE.
   Pequeñas diferencias de mayúsculas/minúsculas no son inconsistencias.

2. NÚMERO DE BL / BOOKING
   Si la INVOICE o PACKING LIST referencian un número de BL o booking, debe coincidir con el BL emitido.

3. NÚMERO DE FACTURA (Invoice Number)
   Si el PACKING LIST o BL referencian un número de factura, debe coincidir con el de la INVOICE.

4. SHIPPER / EXPORTADOR
   El nombre o razón social del exportador en INVOICE debe coincidir con el del BL.
   Diferencias menores de formato (abreviaciones, Ltd vs Limited) no son inconsistencias.

5. CONSIGNEE / IMPORTADOR
   El nombre del importador en INVOICE debe coincidir con el del BL.

6. PAÍS DE ORIGEN
   El país de origen en CO debe coincidir con el mencionado en INVOICE y BL.

7. DESCRIPCIÓN DE MERCANCÍA
   La descripción en INVOICE, PACKING LIST, BL y CO debe ser compatible.
   Solo reportar si los productos son distintos o hay contradicción directa — las diferencias
   de nivel de detalle (más o menos descriptivo) no son inconsistencias.

8. CANTIDAD TOTAL (cajas, unidades, piezas)
   El total de unidades/cajas en INVOICE debe coincidir con el total en PACKING LIST.
   Diferencias de redondeo menores al 1% no son inconsistencias.

9. PESO TOTAL BRUTO
   El peso bruto total en PACKING LIST debe coincidir con el del WEIGHT CERTIFICATE (si existe).
   Diferencias de redondeo menores al 1% no son inconsistencias.

10. NÚMERO DE PO / OC
    Si aparece en múltiples documentos, debe ser el mismo número.

11. INCOTERMS
    Los términos de comercio declarados en INVOICE deben coincidir con los del BL o FOB LETTER
    si ambos los mencionan explícitamente.

12. FECHA DE FACTURA VS FECHA DE EMBARQUE
    La fecha de la INVOICE no debe ser posterior a la fecha de embarque del BL en más de 30 días.
    Una INVOICE fechada después del embarque más de un mes es sospechosa.

Solo reporta inconsistencias que puedas observar directamente en los textos proporcionados.
Si un campo no aparece en alguno de los documentos, no lo reportes como inconsistencia.
No inferas ni supongas valores ausentes.

Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional, sin markdown.
IMPORTANTE: Los valores de todos los campos string deben estar en una sola línea — sin saltos de línea (\n) dentro del texto.
Formato exacto:
{{
  "inconsistencias_cruzadas": [
    {{
      "campo": "<nombre del campo inconsistente>",
      "documentos": "<tipos de documento involucrados, ej: BL y INVOICE>",
      "descripcion": "<qué declara cada documento y por qué es inconsistente, con valores concretos>",
      "severidad": "alta" | "media" | "baja"
    }}
  ]
}}

Severidades: "alta" = discrepancia que invalida el despacho aduanal,
"media" = discrepancia importante que requiere corrección al proveedor,
"baja" = diferencia menor o posible error tipográfico.
Si no hay inconsistencias entre documentos, devuelve: {{"inconsistencias_cruzadas": []}}
No incluyas ningún campo adicional en el JSON.\
"""

_PROMPT_USUARIO = """\
Nombre del archivo: {nombre_archivo}
Asunto del correo: {asunto_correo}

Analiza el documento adjunto y clasifícalo.\
"""

_PROMPT_USUARIO_INCONSISTENCIAS = """\
Nombre del archivo: {nombre_archivo}
Tipo de documento: {tipo_documento}

Analiza el documento adjunto y detecta inconsistencias internas.\
"""

# Media types soportados para envío directo a Claude
_MEDIA_TYPES_PDF   = {"application/pdf"}
_MEDIA_TYPES_IMAGE = {"image/jpeg", "image/png", "image/gif", "image/webp"}
_EXT_MEDIA = {
    ".pdf":  "application/pdf",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".tiff": "image/png",   # se convierte vía pillow si es necesario
    ".tif":  "image/png",
    ".bmp":  "image/png",
}


def _construir_prompt_sistema_clasificacion() -> str:
    """Construye el prompt de clasificación con el checklist. Resultado estable entre llamadas → cacheable."""
    return _PROMPT_SISTEMA_CLASIFICACION.format(
        checklist=formatear_checklist_para_prompt(),
    )


def _construir_prompt_sistema_inconsistencias() -> str:
    """Prompt de detección de inconsistencias. Estático → cacheable."""
    return _PROMPT_SISTEMA_INCONSISTENCIAS


def _construir_bloque_documento(ruta_local: str) -> dict | None:
    """
    Lee el archivo y devuelve el bloque de contenido para la API de Anthropic.
    Soporta PDF e imágenes. Retorna None si la extensión no es soportada.
    """
    ext = Path(ruta_local).suffix.lower()
    media_type = _EXT_MEDIA.get(ext)
    if not media_type:
        return None

    # Imágenes TIFF/BMP no soportadas nativamente — convertir a PNG con pillow
    if ext in (".tiff", ".tif", ".bmp"):
        try:
            from PIL import Image
            import io
            img = Image.open(ruta_local).convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            data_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")
        except Exception as e:
            log_advertencia("clasificador_claude", "CLAUDE-012", Path(ruta_local).name,
                            f"No se pudo convertir imagen {ext}: {e}")
            return None
    else:
        with open(ruta_local, "rb") as f:
            data_b64 = base64.standard_b64encode(f.read()).decode("utf-8")

    if media_type == "application/pdf":
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": media_type, "data": data_b64},
        }
    else:
        return {
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": data_b64},
        }


def _detectar_inconsistencias(
    client,
    nombre_archivo: str,
    tipo_documento: str,
    bloque_doc: dict | None,
    texto_extraido: str,
) -> list:
    """
    Segunda llamada a Claude — solo detección de inconsistencias.
    Se invoca únicamente si tipo_documento está en _TIPOS_CON_INCONSISTENCIAS.

    Returns:
        Lista de dicts [{"campo": ..., "descripcion": ..., "severidad": ...}].
        Lista vacía si no hay inconsistencias o si falla la llamada.
    """
    texto_usuario = _PROMPT_USUARIO_INCONSISTENCIAS.format(
        nombre_archivo=nombre_archivo,
        tipo_documento=tipo_documento,
    )
    if texto_extraido:
        texto_usuario += f"\n\nContenido del documento:\n{texto_extraido[:4000]}"

    if bloque_doc is not None:
        contenido_mensaje = [bloque_doc, {"type": "text", "text": texto_usuario}]
    else:
        contenido_mensaje = [{"type": "text", "text": texto_usuario}]

    try:
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=600,
            system=[{
                "type": "text",
                "text": _construir_prompt_sistema_inconsistencias(),
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{"role": "user", "content": contenido_mensaje}],
        )
        texto_respuesta = respuesta.content[0].text.strip()
        fallback_inc = {"inconsistencias": []}
        datos = _parsear_json_respuesta(texto_respuesta, fallback_inc)
        if datos is fallback_inc:
            log_advertencia("clasificador_claude", "CLAUDE-008", nombre_archivo,
                            f"JSON inválido en detección de inconsistencias. "
                            f"Motivo: parseo fallido. Respuesta cruda: {texto_respuesta[:300]!r}")
        inconsistencias = datos.get("inconsistencias", [])
        if not isinstance(inconsistencias, list):
            return []
        return [
            inc for inc in inconsistencias
            if isinstance(inc, dict) and "descripcion" in inc
        ]
    except Exception as e:
        log_advertencia("clasificador_claude", "CLAUDE-008", nombre_archivo,
                        f"Error en detección de inconsistencias. "
                        f"Motivo: {type(e).__name__}: {e}")
        return []


def clasificar_con_claude(
    nombre_archivo: str,
    ruta_local: str,
    asunto_correo: str = "",
    texto_documento: str = "",   # mantenido por compatibilidad, ya no se usa
) -> tuple[str, int, str, list, str]:
    """
    Clasifica un documento usando Claude — envía el archivo directamente (PDF o imagen).
    Si el tipo clasificado admite revisión de inconsistencias, realiza una segunda
    llamada separada solo para esa tarea.

    Args:
        nombre_archivo   : Nombre del archivo adjunto (con extensión).
        ruta_local       : Ruta al archivo en disco para enviar a la API.
        asunto_correo    : Asunto del correo en que llegó (opcional).
        texto_documento  : Ignorado. Mantenido por compatibilidad con código existente.

    Returns:
        (tipo, certeza, justificacion, inconsistencias, texto_extraido, nombre_proveedor) donde:
          - tipo             : string exacto del catálogo, o "_ERROR_API" / "_TIPO_INVALIDO" en caso de fallo.
          - certeza          : entero 0-100. 0 si hubo fallo.
          - justificacion    : string con la justificación de Claude. "" si hubo fallo.
          - inconsistencias  : lista de dicts [{"campo": ..., "descripcion": ..., "severidad": ...}].
                               Lista vacía si no hay inconsistencias o si hubo fallo.
          - texto_extraido   : texto del documento (para validación cruzada). "" si no disponible.
          - nombre_proveedor : nombre de la empresa emisora extraído del documento. "" si no disponible.
    """
    if not ANTHROPIC_API_KEY:
        log_error("clasificador_claude", "CLAUDE-001", nombre_archivo, "ANTHROPIC_API_KEY no configurada")
        return (_ERROR_API, 0, "", [], "", "", False, None)

    bloque_doc = _construir_bloque_documento(ruta_local)
    texto_extraido = ""
    if bloque_doc is None:
        ext = Path(ruta_local).suffix.lower()
        if ext in (".docx", ".doc", ".xlsx", ".xls"):
            texto_extraido = extraer_texto(ruta_local)
            if not texto_extraido.strip():
                log_advertencia("clasificador_claude", "CLAUDE-009", nombre_archivo,
                                f"No se pudo extraer texto de archivo {ext}")
                return (_ERROR_API, 0, "", [], "", "", False, None)
            log_info("clasificador_claude", f"Clasificando por texto extraído: {ext}", nombre_archivo)
        else:
            log_advertencia("clasificador_claude", "CLAUDE-010", nombre_archivo,
                            f"Formato no soportado para envío directo: {ext}")
            return (_ERROR_API, 0, "", [], "", "", False, None)

    texto_usuario = _PROMPT_USUARIO.format(
        nombre_archivo=nombre_archivo,
        asunto_correo=asunto_correo or "(no disponible)",
    )
    if texto_extraido:
        texto_usuario += f"\n\nContenido del documento:\n{texto_extraido[:4000]}"

    try:
        import anthropic

        client = anthropic.Anthropic(
            api_key=ANTHROPIC_API_KEY,
            timeout=float(CLAUDE_TIMEOUT_SEGUNDOS),
        )

        if bloque_doc is not None:
            contenido_mensaje = [bloque_doc, {"type": "text", "text": texto_usuario}]
        else:
            contenido_mensaje = [{"type": "text", "text": texto_usuario}]

        # --- Llamada 1: clasificación con prompt cacheado ---
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=500,
            system=[{
                "type": "text",
                "text": _construir_prompt_sistema_clasificacion(),
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{"role": "user", "content": contenido_mensaje}],
        )

        texto_respuesta = respuesta.content[0].text.strip()
        datos = _parsear_json_respuesta(texto_respuesta, {})
        if not datos:
            raise json.JSONDecodeError("No se pudo parsear respuesta de clasificación", texto_respuesta, 0)
        tipo             = str(datos.get("tipo", "")).strip()
        certeza          = int(datos.get("certeza", 0))
        justificacion    = str(datos.get("justificacion", "")).strip()
        nombre_proveedor = str(datos.get("nombre_proveedor", "")).strip()
        es_borrador      = bool(datos.get("es_borrador", False))
        _po_raw = re.sub(r'\D', '', str(datos.get("numero_po") or ""))
        numero_po_claude = _po_raw or None

        if tipo not in TIPOS_VALIDOS:
            log_advertencia("clasificador_claude", "CLAUDE-002", nombre_archivo, f"Tipo desconocido en respuesta: '{tipo}'")
            return (_TIPO_INVALIDO, 0, "", [], "", "", False, None)

        certeza = max(0, min(100, certeza))

        # Extraer texto de todos los PDFs (necesario para validación cruzada y determinista)
        if not texto_extraido and bloque_doc is not None and Path(ruta_local).suffix.lower() == ".pdf":
            texto_extraido = extraer_texto(ruta_local)

        # --- Llamada 2: inconsistencias (solo si aplica al tipo) ---
        if tipo in _TIPOS_CON_INCONSISTENCIAS:
            inconsistencias_claude = _detectar_inconsistencias(
                client=client,
                nombre_archivo=nombre_archivo,
                tipo_documento=tipo,
                bloque_doc=bloque_doc,
                texto_extraido=texto_extraido,
            )
            inconsistencias_codigo = validar_campos(tipo, texto_extraido)
            inconsistencias = inconsistencias_claude + inconsistencias_codigo

            if inconsistencias_codigo:
                log_info("clasificador_claude",
                         f"Validador de código: {len(inconsistencias_codigo)} inconsistencia(s) detectadas",
                         nombre_archivo)
        else:
            inconsistencias = []

        if inconsistencias:
            altas  = sum(1 for i in inconsistencias if i.get("severidad") == "alta")
            medias = sum(1 for i in inconsistencias if i.get("severidad") == "media")
            bajas  = sum(1 for i in inconsistencias if i.get("severidad") == "baja")
            log_info("clasificador_claude",
                     f"Inconsistencias detectadas: {len(inconsistencias)} (altas={altas}, medias={medias}, bajas={bajas})",
                     nombre_archivo)

        log_clasificacion(
            archivo=nombre_archivo,
            tipo=tipo,
            certeza=certeza,
            justificacion=justificacion,
            inconsistencias=inconsistencias,
            asunto=asunto_correo,
        )

        return (tipo, certeza, justificacion, inconsistencias, texto_extraido, nombre_proveedor, es_borrador, numero_po_claude)

    except ImportError:
        log_error("clasificador_claude", "CLAUDE-003", nombre_archivo, "Librería 'anthropic' no instalada")
        return (_ERROR_API, 0, "", [], "", "", False, None)
    except json.JSONDecodeError as e:
        log_advertencia("clasificador_claude", "CLAUDE-004", nombre_archivo, f"JSON inválido en respuesta: {e}")
        return (_ERROR_API, 0, "", [], "", "", False, None)
    except Exception as e:
        log_error("clasificador_claude", "CLAUDE-005", nombre_archivo, f"{type(e).__name__}: {e}")
        return (_ERROR_API, 0, "", [], "", "", False, None)


def es_fallo(tipo: str) -> bool:
    """Retorna True si el tipo indica un fallo de la API (no un tipo real del catálogo)."""
    return tipo.startswith("__")


# Tipos que no aportan campos cruzables (imágenes, docs sin campos de embarque)
_TIPOS_SIN_CRUCE = {"MARCHAMO", "EXONERACION", "PRINTER", "OTROS"}

# Máximo de caracteres enviados por documento en la validación cruzada
_MAX_CHARS_CRUCE = 2500


def detectar_inconsistencias_cruzadas(documentos: list[dict]) -> list[dict]:
    """
    Detecta inconsistencias ENTRE múltiples documentos de un mismo embarque.

    Args:
        documentos: Lista de dicts [{"tipo": ..., "nombre_archivo": ..., "texto": ...}].
                    Solo se procesan los que tienen texto extraído y tipo cruzable.

    Returns:
        Lista de inconsistencias en formato estándar {"campo", "descripcion", "severidad"}.
        Lista vacía si hay menos de 2 documentos con texto o si falla la API.
    """
    if not ANTHROPIC_API_KEY:
        return []

    docs_utiles = [
        d for d in documentos
        if d.get("texto", "").strip() and d.get("tipo") not in _TIPOS_SIN_CRUCE
    ]

    if len(docs_utiles) < 2:
        return []

    bloques = []
    for doc in docs_utiles:
        tipo    = doc.get("tipo", "DESCONOCIDO")
        nombre  = doc.get("nombre_archivo", "")
        texto   = doc.get("texto", "").strip()[:_MAX_CHARS_CRUCE]
        bloques.append(f"=== {tipo}: {nombre} ===\n{texto}")

    mensaje_usuario = (
        f"Se adjuntan {len(docs_utiles)} documento(s) del mismo embarque.\n"
        f"Detecta inconsistencias entre ellos.\n\n"
        + "\n\n".join(bloques)
    )

    try:
        import anthropic

        client = anthropic.Anthropic(
            api_key=ANTHROPIC_API_KEY,
            timeout=float(CLAUDE_TIMEOUT_SEGUNDOS),
        )
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=800,
            system=[{
                "type": "text",
                "text": _PROMPT_SISTEMA_CRUZADO,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{"role": "user", "content": mensaje_usuario}],
        )
        texto_respuesta = respuesta.content[0].text.strip()
        fallback_cruzado = {"inconsistencias_cruzadas": []}
        datos = _parsear_json_respuesta(texto_respuesta, fallback_cruzado)
        if datos is fallback_cruzado:
            log_advertencia("clasificador_claude", "CLAUDE-006", "-",
                            f"JSON inválido en validación cruzada. "
                            f"Motivo: parseo fallido. Respuesta cruda: {texto_respuesta[:300]!r}")
        cruzadas = datos.get("inconsistencias_cruzadas", [])
        if not isinstance(cruzadas, list):
            return []

        resultado = []
        for inc in cruzadas:
            if not isinstance(inc, dict) or "descripcion" not in inc:
                continue
            docs_str   = inc.get("documentos", "")
            campo_base = inc.get("campo", "")
            campo_final = f"{campo_base} ({docs_str})" if docs_str else campo_base
            resultado.append({
                "campo":       campo_final,
                "descripcion": inc.get("descripcion", ""),
                "severidad":   inc.get("severidad", "media"),
            })

        if resultado:
            log_info("clasificador_claude",
                     f"Validación cruzada: {len(resultado)} inconsistencia(s) detectadas entre documentos")
        else:
            log_info("clasificador_claude", "Validación cruzada: sin inconsistencias entre documentos")

        return resultado

    except Exception as e:
        log_advertencia("clasificador_claude", "CLAUDE-007", "-",
                        f"Error en validación cruzada. Motivo: {type(e).__name__}: {e}")
        return []


_PROMPT_RESPUESTA_PROVEEDOR = """\
You are a foreign trade document reviewer for a Costa Rican importing company.
The document "{nombre_archivo}" (type: {tipo_documento}) has been reviewed and contains the following issues:

{lista_inconsistencias}

Write a short, direct email to the supplier. Rules:
- Start with this exact line: "*** This is an automated notification generated by our document review system. ***"
- Then "Dear team,"
- Immediately after the greeting, write this exact sentence: "Please review the following issues found in the shipment documents and provide corrections at your earliest convenience."
- Then list the issues as numbered items. Use ONLY the document type as the section header (e.g., "INVOICE:") — never the filename.
- Each item: 1-2 sentences explaining the discrepancy with relevant figures or details, then end with a specific action request (e.g., "Please reissue...", "Please confirm...", "Please provide..."). Be clear but concise — not a single line, not a paragraph.
- End with 2-3 polite closing sentences: acknowledge that these corrections are necessary to move forward with customs clearance, thank the supplier for their attention and prompt cooperation, and express willingness to assist if they have any questions.
- No signature block.
- Tone: professional but direct.
- Language: English.

Reply only with the email body text, no markdown, no quotes.\
"""


def generar_respuesta_proveedor_consolidada(
    remitente: str,
    asunto: str,
    documentos: list[dict],
    nombre_proveedor: str = "",
) -> str:
    """
    Genera una sugerencia de correo al proveedor consolidando las inconsistencias
    de múltiples documentos de un mismo correo.

    Args:
        remitente        : Correo del remitente.
        asunto           : Asunto del correo recibido.
        documentos       : Lista de dicts con claves nombre_archivo, tipo, inconsistencias.
        nombre_proveedor : Nombre de la empresa proveedora extraído de los documentos.

    Returns:
        Texto del correo sugerido, o cadena vacía si falla la API.
    """
    if not ANTHROPIC_API_KEY or not documentos:
        return ""

    # Los documentos con inconsistencia de full set (PDF concatenado) van siempre primero
    _TIPO_FULLSET = "Documentos enviados concatenados en un solo PDF"
    docs_ordenados = sorted(
        documentos,
        key=lambda d: 0 if any(
            inc.get("campo") == _TIPO_FULLSET
            for inc in d.get("inconsistencias") or []
        ) else 1,
    )

    bloques = []
    for doc in docs_ordenados:
        lineas = [f"Documento: {doc['nombre_archivo']} (tipo: {doc['tipo']})"]
        for i, inc in enumerate(doc["inconsistencias"], 1):
            sev   = inc.get("severidad", "").upper()
            campo = inc.get("campo", "")
            desc  = inc.get("descripcion", "")
            lineas.append(f"  {i}. [{sev}] {campo}: {desc}")
        bloques.append("\n".join(lineas))

    destinatario = nombre_proveedor if nombre_proveedor else "supplier team"
    prompt = (
        "You are a foreign trade document reviewer for a Costa Rican importing company.\n"
        f"Documents received from {nombre_proveedor + ' ' if nombre_proveedor else ''}"
        f"<{remitente}> (subject: \"{asunto}\") were reviewed "
        "and the following issues were found:\n\n"
        + "\n\n".join(bloques)
        + "\n\nWrite ONE short, direct email to the supplier. Rules:\n"
        "- Start with this exact line: '*** This is an automated notification generated by our document review system. ***'\n"
        f"- Then 'Dear {destinatario},'\n"
        "- Immediately after the greeting, write this exact sentence: "
        "'Please review the following issues found in the shipment documents and provide corrections at your earliest convenience.'\n"
        "- Then list issues grouped by document type. Use ONLY the document type as the section header (e.g., 'INVOICE:', 'PACKING LIST:', 'CO:') — never the filename.\n"
        "- IMPORTANT: If any issue is about documents being sent merged/concatenated in a single PDF, list that issue FIRST, before any other issues, under the header 'DOCUMENT SUBMISSION FORMAT:'.\n"
        "- Each item: 1-2 sentences explaining the discrepancy with relevant figures or details, then end with a specific action request "
        "(e.g., 'Please reissue...', 'Please confirm...', 'Please provide...'). Be clear but concise — not a single line, not a paragraph.\n"
        "- End with 2-3 polite closing sentences: acknowledge that these corrections are necessary to move forward with customs clearance, thank the supplier for their attention and prompt cooperation, and express willingness to assist if they have any questions.\n"
        "- No signature block.\n"
        "- Tone: professional but direct. Language: English.\n\n"
        "Reply only with the email body text, no markdown, no quotes."
    )

    try:
        import anthropic

        client = anthropic.Anthropic(
            api_key=ANTHROPIC_API_KEY,
            timeout=float(CLAUDE_TIMEOUT_SEGUNDOS),
        )
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )
        return respuesta.content[0].text.strip()

    except Exception as e:
        log_advertencia("clasificador_claude", "CLAUDE-011", "-",
                        f"No se pudo generar respuesta consolidada al proveedor. Motivo: {type(e).__name__}: {e}")
        return ""


def contrastar_sugerencia_vs_inconsistencias(
    sugerencia: str,
    documentos: list[dict],
) -> str:
    """
    Verifica que el borrador de respuesta al proveedor cubra todas las inconsistencias
    detectadas. Si falta alguna, la agrega al final marcada como [AGREGADO].

    No hace ninguna llamada a la API — es comparación local de strings.

    Args:
        sugerencia  : Texto del borrador generado por Claude.
        documentos  : Lista de dicts con claves nombre_archivo, tipo, inconsistencias.

    Returns:
        El mismo texto si todo está cubierto, o el texto con las inconsistencias
        faltantes agregadas al final.
    """
    if not sugerencia or not documentos:
        return sugerencia

    try:
        sugerencia_lower = sugerencia.lower()
        faltantes = []

        for doc in documentos:
            for inc in doc.get("inconsistencias", []):
                campo = inc.get("campo", "")
                desc  = inc.get("descripcion", "")
                sev   = inc.get("severidad", "").upper()

                # Extraer palabras clave: tokens de 4+ chars del campo y descripción
                tokens = re.findall(r"[a-záéíóúüñ]{4,}", (campo + " " + desc).lower())
                # Considerar cubierta si al menos 2 tokens distintos aparecen en el borrador
                tokens_unicos = list(dict.fromkeys(tokens))  # deduplica, preserva orden
                encontrados = sum(1 for t in tokens_unicos if t in sugerencia_lower)

                if encontrados < min(2, len(tokens_unicos)):
                    faltantes.append({
                        "tipo":   doc.get("tipo", ""),
                        "campo":  campo,
                        "desc":   desc,
                        "sev":    sev,
                    })

        if not faltantes:
            return sugerencia

        # Agregar bloque con las inconsistencias no cubiertas
        lineas = ["\n\n[AGREGADO — inconsistencias no cubiertas en el borrador anterior:]"]
        tipo_actual = None
        for f in faltantes:
            if f["tipo"] != tipo_actual:
                tipo_actual = f["tipo"]
                lineas.append(f"\n{tipo_actual}:")
            lineas.append(f"  - [{f['sev']}] {f['campo']}: {f['desc']}")

        log_advertencia(
            "clasificador_claude", "CLAUDE-012", "-",
            f"Contraste: {len(faltantes)} inconsistencia(s) no cubiertas en el borrador — agregadas al .txt"
        )
        return sugerencia + "\n".join(lineas)

    except Exception as e:
        log_advertencia(
            "clasificador_claude", "CLAUDE-013", "-",
            f"contrastar_sugerencia_vs_inconsistencias falló (borrador sin cambios): {e}"
        )
        return sugerencia


def generar_respuesta_proveedor(
    nombre_archivo: str,
    tipo_documento: str,
    inconsistencias: list[dict],
) -> str:
    """
    Genera una sugerencia de correo al proveedor basada en las inconsistencias detectadas.

    Args:
        nombre_archivo  : Nombre del archivo con inconsistencias.
        tipo_documento  : Tipo clasificado del documento.
        inconsistencias : Lista de dicts con claves campo, descripcion, severidad.

    Returns:
        Texto del correo sugerido, o cadena vacía si falla la API.
    """
    if not ANTHROPIC_API_KEY or not inconsistencias:
        return ""

    lineas = []
    for i, inc in enumerate(inconsistencias, 1):
        sev   = inc.get("severidad", "").upper()
        campo = inc.get("campo", "")
        desc  = inc.get("descripcion", "")
        lineas.append(f"{i}. [{sev}] {campo}: {desc}")

    lista_txt = "\n".join(lineas)
    prompt = _PROMPT_RESPUESTA_PROVEEDOR.format(
        nombre_archivo=nombre_archivo,
        tipo_documento=tipo_documento,
        lista_inconsistencias=lista_txt,
    )

    try:
        import anthropic

        client = anthropic.Anthropic(
            api_key=ANTHROPIC_API_KEY,
            timeout=float(CLAUDE_TIMEOUT_SEGUNDOS),
        )
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )
        return respuesta.content[0].text.strip()

    except Exception as e:
        log_advertencia("clasificador_claude", "CLAUDE-011", nombre_archivo,
                        f"No se pudo generar respuesta al proveedor. Motivo: {type(e).__name__}: {e}")
        return ""
