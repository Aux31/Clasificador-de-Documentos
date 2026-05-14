"""
Validaciones deterministas de campos en documentos de comercio exterior.
Complementa la detección de Claude con reglas verificables en código.

Estas validaciones se ejecutan después de Claude y sus resultados se fusionan
con los de la llamada a la API. Solo se implementan reglas con bajo riesgo de
falso positivo — patrones con formato estándar definido y unívoco.

Las validaciones marcadas con [SA] derivan directamente del Shipping Agreement
firmado con MERCASA (referencias/Shipping Agreement.docx).
"""

import re

# ---------------------------------------------------------------------------
# Constantes del Shipping Agreement (SA) — fuente: Shipping Agreement.docx
# ---------------------------------------------------------------------------
# Datos del consignatario obligatorio en el BL [SA Annex 1 - Payment Process]
_SA_CONSIGNATARIO_NOMBRE = "MERCADEO DE ARTICULOS DE CONSUMO"
_SA_CONSIGNATARIO_ID     = "3-101-137584"

# Idioma requerido para la factura [SA cláusula 5.2]
_SA_INVOICE_PALABRAS_ESPANOL = [
    "FACTURA", "CANTIDAD", "PRECIO", "VALOR", "DESCRIPCION", "DESCRIPCIÓN",
    "PESO", "UNIDAD", "TOTAL", "PROVEEDOR", "CLIENTE", "BULTOS", "NETO", "BRUTO",
]

# Palabras que indican flete explícito en el BL [SA cláusula 4, punto 3]
_SA_BL_PALABRAS_FLETE = [
    "FREIGHT PREPAID", "FREIGHT COLLECT", "FREIGHT PRE-PAID",
    "FLETE PREPAGADO", "FLETE POR COBRAR", "PREPAID", "COLLECT",
]

# Número de sello — patterns para detectarlo en texto [SA cláusula 3.3]
_RE_SELLO = re.compile(
    r'\b(?:SEAL|MARCHAMO|PRECINTO|SELLO)\s*(?:NO\.?|NUMBER|#|NUM\.?)?\s*[:\-]?\s*([A-Z0-9]{4,15})\b',
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Contenedor: estándar BIC ISO 6346
# Formato: 3 letras (owner code) + 1 letra (categoría: U/J/Z) + 6 dígitos + 1 check digit
# Ejemplos válidos: TEMU6803617, MSCU1234568, HLXU9876543
# ---------------------------------------------------------------------------
_RE_CANDIDATO_CONT = re.compile(r'\b([A-Z]{4}\d{7})\b')
_CATEGORIAS_VALIDAS = {"U", "J", "Z"}

# Valores numéricos de letras para el algoritmo de check digit ISO 6346
# (omite los valores que son múltiplos de 11: se salta el 11, 22, 33)
_ISO6346_LETRA_VALORES: dict[str, int] = {
    'A': 10, 'B': 12, 'C': 13, 'D': 14, 'E': 15, 'F': 16, 'G': 17,
    'H': 18, 'I': 19, 'J': 20, 'K': 21, 'L': 23, 'M': 24, 'N': 25,
    'O': 26, 'P': 27, 'Q': 28, 'R': 29, 'S': 30, 'T': 31, 'U': 32,
    'V': 34, 'W': 35, 'X': 36, 'Y': 37, 'Z': 38,
}


def _calcular_check_digit_iso6346(contenedor: str) -> int | None:
    """
    Calcula el dígito verificador ISO 6346 de un número de contenedor.

    Args:
        contenedor: 11 caracteres en mayúsculas (4 letras + 6 dígitos + 1 check).

    Returns:
        Dígito calculado (0-9), o None si algún carácter tiene valor indefinido.
    """
    if len(contenedor) != 11:
        return None
    suma = 0
    for i, c in enumerate(contenedor[:10]):
        if c.isdigit():
            v = int(c)
        else:
            v = _ISO6346_LETRA_VALORES.get(c)
            if v is None:
                return None
        suma += v * (2 ** i)
    return (suma % 11) % 10  # si resto es 10, el estándar usa 0


def _validar_contenedores(texto_upper: str) -> list[dict]:
    """
    Detecta números que parecen contenedores (4 letras + 7 dígitos) pero cuya
    4ª letra no es U, J ni Z según el estándar BIC ISO 6346.

    No valida el dígito verificador aquí — eso lo hace _validar_check_digit.
    """
    inconsistencias = []
    vistos = set()
    for m in _RE_CANDIDATO_CONT.finditer(texto_upper):
        num = m.group(1)
        if num in vistos:
            continue
        vistos.add(num)
        categoria = num[3]
        if categoria not in _CATEGORIAS_VALIDAS:
            inconsistencias.append({
                "campo": "Número de contenedor",
                "descripcion": (
                    f"El número '{num}' no cumple el formato BIC estándar: "
                    f"la 4ª letra debe ser U (carga), J (equipo detachable) o Z (trailer/chasis), "
                    f"no '{categoria}'. Verificar si es un error tipográfico."
                ),
                "severidad": "media",
            })
    return inconsistencias


def _validar_check_digit_contenedores(texto_upper: str) -> list[dict]:
    """
    Verifica el dígito verificador ISO 6346 para cada número de contenedor
    con formato válido (4 letras + 7 dígitos, 4ª letra U/J/Z).

    Solo reporta cuando el dígito calculado difiere del declarado — nunca
    reporta si el cálculo es imposible (letras fuera de tabla, etc.).
    """
    inconsistencias = []
    vistos = set()
    for m in _RE_CANDIDATO_CONT.finditer(texto_upper):
        num = m.group(1)
        if num in vistos:
            continue
        vistos.add(num)
        if num[3] not in _CATEGORIAS_VALIDAS:
            continue  # ya reportado por _validar_contenedores

        check_declarado = int(num[-1])
        check_calculado = _calcular_check_digit_iso6346(num)

        if check_calculado is not None and check_declarado != check_calculado:
            inconsistencias.append({
                "campo": "Número de contenedor (dígito verificador)",
                "descripcion": (
                    f"El número '{num}' tiene dígito verificador incorrecto: "
                    f"el último dígito es '{check_declarado}' pero el cálculo ISO 6346 "
                    f"da '{check_calculado}'. Posible error tipográfico en el número de contenedor."
                ),
                "severidad": "alta",
            })
    return inconsistencias


# ---------------------------------------------------------------------------
# Pesos: peso neto total > peso bruto total
# Aplica a PACKING LIST y WEIGHT CERTIFICATE
# Busca líneas de resumen que contengan "TOTAL" + "GROSS" y "TOTAL" + "NET"
# ---------------------------------------------------------------------------
_RE_TOTAL_GROSS = re.compile(
    r'TOTAL\s*(?:GROSS\s*)?WEIGHT[:\s]*([0-9]+(?:[.,][0-9]+)?)\s*(?:KGS?|MT|T\b)',
    re.IGNORECASE,
)
_RE_TOTAL_NET = re.compile(
    r'TOTAL\s*NET\s*WEIGHT[:\s]*([0-9]+(?:[.,][0-9]+)?)\s*(?:KGS?|MT|T\b)',
    re.IGNORECASE,
)


def _normalizar_numero(s: str) -> float:
    """Convierte '1,234.56' o '1.234,56' a float. Asume punto como decimal si hay ambos."""
    s = s.strip()
    if ',' in s and '.' in s:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(',', '')
        else:
            s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def _validar_pesos_totales(texto_upper: str) -> list[dict]:
    """
    Detecta si el peso neto total declarado supera el peso bruto total.
    Solo actúa si encuentra exactamente un valor de cada tipo en el texto.
    """
    brutos = _RE_TOTAL_GROSS.findall(texto_upper)
    netos  = _RE_TOTAL_NET.findall(texto_upper)

    if len(brutos) == 1 and len(netos) == 1:
        gw = _normalizar_numero(brutos[0])
        nw = _normalizar_numero(netos[0])
        if gw > 0 and nw > 0 and nw > gw:
            return [{
                "campo": "Pesos totales (neto > bruto)",
                "descripcion": (
                    f"El peso neto total ({brutos[0]} GW / {netos[0]} NW) es mayor que el "
                    f"peso bruto total, lo cual es físicamente imposible. "
                    f"Verificar si los valores están invertidos."
                ),
                "severidad": "alta",
            }]
    return []


# ---------------------------------------------------------------------------
# [SA cláusula 5.2] Factura en español
# Aplica a INVOICE y PL + INV
# ---------------------------------------------------------------------------

def _validar_invoice_en_espanol(texto_upper: str) -> list[dict]:
    """
    Detecta si la factura parece estar exclusivamente en inglés.
    Solo reporta cuando no encuentra ninguna palabra típica del español y sí
    encuentra al menos dos señales claras de inglés comercial. Umbral
    conservador para evitar falsos positivos en facturas bilingües.
    """
    palabras_espanol_encontradas = sum(
        1 for p in _SA_INVOICE_PALABRAS_ESPANOL if p in texto_upper
    )
    if palabras_espanol_encontradas >= 2:
        return []

    _PALABRAS_INGLES = [
        "COMMERCIAL INVOICE", "QUANTITY", "UNIT PRICE", "AMOUNT",
        "DESCRIPTION OF GOODS", "SHIPPER", "CONSIGNEE",
    ]
    palabras_ingles = sum(1 for p in _PALABRAS_INGLES if p in texto_upper)
    if palabras_ingles >= 2 and palabras_espanol_encontradas == 0:
        return [{
            "campo": "Idioma de la factura [SA cláusula 5.2]",
            "descripcion": (
                "La factura parece estar en inglés. El Shipping Agreement (cláusula 5.2) "
                "requiere que las facturas estén en español e incluyan los códigos de "
                "producto MERCASA. Confirmar con el proveedor."
            ),
            "severidad": "media",
        }]
    return []


# ---------------------------------------------------------------------------
# [SA Annex 1 - Payment Process] Consignatario correcto en el BL
# Aplica a BL, MBL, HBL, AWB
# ---------------------------------------------------------------------------

def _validar_consignatario_bl(texto_upper: str) -> list[dict]:
    """
    Verifica que el BL mencione el nombre legal y el ID jurídico de MERCASA
    como consignatario. Solo reporta si el texto tiene suficiente contenido
    de BL (al menos una de las palabras clave de transporte) para evitar
    falsos positivos en documentos cortos.
    """
    _PALABRAS_BL = ["BILL OF LADING", "B/L", "CONSIGNEE", "SHIPPER", "PORT OF"]
    tiene_estructura_bl = any(p in texto_upper for p in _PALABRAS_BL)
    if not tiene_estructura_bl:
        return []

    falta_nombre = _SA_CONSIGNATARIO_NOMBRE not in texto_upper
    falta_id     = _SA_CONSIGNATARIO_ID not in texto_upper

    if falta_nombre or falta_id:
        faltantes = []
        if falta_nombre:
            faltantes.append(f"nombre legal '{_SA_CONSIGNATARIO_NOMBRE}'")
        if falta_id:
            faltantes.append(f"cédula jurídica '{_SA_CONSIGNATARIO_ID}'")
        return [{
            "campo": "Consignatario en el BL [SA Annex 1]",
            "descripcion": (
                f"El BL no contiene el {' ni el '.join(faltantes)} de MERCASA en el campo "
                f"consignatario. El Shipping Agreement exige que figure exactamente: "
                f"'MERCADEO DE ARTICULOS DE CONSUMO S.A.' con cédula jurídica 3-101-137584."
            ),
            "severidad": "alta",
        }]
    return []


# ---------------------------------------------------------------------------
# [SA cláusula 4, punto 3] Flete impreso en el BL
# Aplica a BL, MBL, HBL, AWB
# ---------------------------------------------------------------------------

def _validar_flete_en_bl(texto_upper: str) -> list[dict]:
    """
    Verifica que el BL indique explícitamente si el flete es prepagado o por cobrar.
    Solo actúa si el texto tiene estructura reconocible de BL.
    """
    _PALABRAS_BL = ["BILL OF LADING", "B/L", "CONSIGNEE", "SHIPPER", "PORT OF LOADING"]
    tiene_estructura_bl = any(p in texto_upper for p in _PALABRAS_BL)
    if not tiene_estructura_bl:
        return []

    tiene_flete = any(p in texto_upper for p in _SA_BL_PALABRAS_FLETE)
    if not tiene_flete:
        return [{
            "campo": "Condición de flete en el BL [SA cláusula 4 punto 3]",
            "descripcion": (
                "El BL no indica explícitamente si el flete es 'Freight Prepaid' o "
                "'Freight Collect'. El Shipping Agreement (cláusula 4, punto 3) exige "
                "que el BL sea emitido 'with the Ocean Freight Printed'. Verificar con "
                "la naviera o el proveedor."
            ),
            "severidad": "alta",
        }]
    return []


# ---------------------------------------------------------------------------
# [SA cláusula 3.3] Número de sello/marchamo en Packing List
# Aplica a PACKING LIST, PL + INV
# ---------------------------------------------------------------------------

def _validar_sello_en_packing(texto_upper: str) -> list[dict]:
    """
    Verifica que el Packing List declare al menos un número de sello/marchamo
    del contenedor. El SA exige que se informe el número correcto de los
    precintos en el Packing List o en documento certificado aparte.
    Solo reporta si el texto tiene estructura de Packing List reconocible.
    """
    _PALABRAS_PL = ["PACKING LIST", "PACKING DETAILS", "GROSS WEIGHT", "NET WEIGHT",
                    "CARTON", "CONTAINER", "TOTAL PACKAGES"]
    tiene_estructura_pl = any(p in texto_upper for p in _PALABRAS_PL)
    if not tiene_estructura_pl:
        return []

    tiene_sello = bool(_RE_SELLO.search(texto_upper))
    if not tiene_sello:
        return [{
            "campo": "Número de sello/marchamo en Packing List [SA cláusula 3.3]",
            "descripcion": (
                "El Packing List no declara el número de sello (marchamo/precinto) del "
                "contenedor. El Shipping Agreement (cláusula 3.3) exige informar el número "
                "correcto de los precintos del contenedor. Solicitar al proveedor que lo "
                "incluya en el Packing List o en documento certificado adjunto."
            ),
            "severidad": "media",
        }]
    return []


# ---------------------------------------------------------------------------
# Registro: tipo → lista de funciones de validación aplicables
# ---------------------------------------------------------------------------
_VALIDACIONES_CONTENEDOR = [_validar_contenedores, _validar_check_digit_contenedores]
_VALIDACIONES_BL_SA     = [_validar_consignatario_bl, _validar_flete_en_bl]

_VALIDACIONES: dict[str, list] = {
    # Documentos de transporte — contenedor + validaciones SA del BL
    "BL":               _VALIDACIONES_CONTENEDOR + _VALIDACIONES_BL_SA,
    "MBL":              _VALIDACIONES_CONTENEDOR + _VALIDACIONES_BL_SA,
    "HBL":              _VALIDACIONES_CONTENEDOR + _VALIDACIONES_BL_SA,
    "AWB":              _VALIDACIONES_BL_SA,  # AWB no tiene número de contenedor ISO
    # Documentos comerciales
    "INVOICE":          _VALIDACIONES_CONTENEDOR + [_validar_invoice_en_espanol],
    "PACKING LIST":     _VALIDACIONES_CONTENEDOR + [_validar_pesos_totales, _validar_sello_en_packing],
    "PL + INV":         _VALIDACIONES_CONTENEDOR + [_validar_pesos_totales, _validar_sello_en_packing,
                                                    _validar_invoice_en_espanol],
    # Certificados — pueden referenciar el contenedor
    "CO":               _VALIDACIONES_CONTENEDOR,
    "WEIGHT CERTIFICATE": _VALIDACIONES_CONTENEDOR + [_validar_pesos_totales],
    "FOB LETTER":       _VALIDACIONES_CONTENEDOR,
}


def validar_campos(tipo_documento: str, texto: str) -> list[dict]:
    """
    Ejecuta validaciones deterministas para el tipo de documento dado.

    Args:
        tipo_documento: Tipo clasificado (ej. "BL", "MBL", "INVOICE").
        texto: Texto extraído del documento. Puede estar vacío.

    Returns:
        Lista de inconsistencias [{campo, descripcion, severidad}].
        Lista vacía si no hay reglas definidas para el tipo o no hay texto.
    """
    if not texto:
        return []

    funciones = _VALIDACIONES.get(tipo_documento, [])
    if not funciones:
        return []

    texto_upper = texto.upper()
    resultado = []
    for fn in funciones:
        resultado.extend(fn(texto_upper))
    return resultado
