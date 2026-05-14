"""
Separador de PDFs con múltiples documentos ("Full Set").

Flujo:
  1. Claude API recibe el PDF completo como documento base64 y detecta si contiene
     varios documentos concatenados (entiende texto e imágenes/scaneados).
  2. Si los detecta, devuelve los rangos de páginas y el tipo de cada uno.
  3. PyMuPDF recorta el PDF original en archivos individuales.
  4. Retorna la lista de rutas generadas para que el pipeline las clasifique normalmente.

Dependencia requerida:
    pip install pymupdf
"""

import base64
import re
from pathlib import Path

from Utilidades.logger_errores import log_error, log_advertencia
from configuracion.ajustes import (
    ANTHROPIC_API_KEY,
    CLAUDE_MODELO,
    CLAUDE_TIMEOUT_SEGUNDOS,
)
from Nucleo.clasificador_claude import _llamar_api, _parsear_json_respuesta

# ---------------------------------------------------------------------------
# Prompt para detección de segmentos
# ---------------------------------------------------------------------------

_PROMPT_SISTEMA_FULLSET = """\
Eres un experto en documentos de comercio exterior.
Recibirás un PDF completo — puede contener texto digital o imágenes escaneadas.
Tu tarea es determinar si el PDF contiene UN SOLO documento o VARIOS documentos concatenados.

Tipos de documentos que puedes encontrar (úsalos EXACTAMENTE como aparecen aquí):
- BL          → Bill of Lading / Conocimiento de Embarque / Connaissement
- MBL         → Master Bill of Lading
- HBL         → House Bill of Lading
- AWB         → Air Waybill / Guía Aérea
- INVOICE     → Factura Comercial / Commercial Invoice / Proforma Invoice
- PACKING LIST → Lista de Empaque / Packing List / Colisage
- CO          → Certificado de Origen / Certificate of Origin / Form A / EUR.1
- FITOSANITARIO → Certificado Fitosanitario / Phytosanitary Certificate / Plant Health Certificate
- ZOOSANITARIO → Certificado Zoosanitario / Veterinary Certificate / Animal Health Certificate
- HEALTH CERTIFICATE → Certificado de Salud / Health Certificate / Sanitary Certificate / Certificate of Sanitary Inspection / Official Health Attestation — para productos alimenticios, cárnicos, lácteos, pesqueros o cualquier producto de origen animal que requiera inspección sanitaria
- WEIGHT CERTIFICATE → Certificado de Peso / Weight Certificate / Weight and Quality Certificate / Inspection Certificate emitido por empresa de surveying (ej: General Survey, SGS, Bureau Veritas)
- QUALITY CERTIFICATE → Certificado de Calidad / Quality Certificate / Certificate of Analysis / Analysis Report / Reporte de Análisis de Laboratorio
- FUMIGATION CERTIFICATE → Certificado de Fumigación / Fumigation Certificate — documento que certifica tratamiento de fumigación (fosfuro de aluminio, metil bromuro, etc.) aplicado a la carga
- OTROS       → SOLO si el documento NO encaja en ninguno de los anteriores (ej: carta explicativa del proveedor, carta de crédito, seguro, instrucciones de embarque)

IMPORTANTE — Guía para distinguir tipos similares:
- Si menciona "phytosanitary", "plant health", "plantas", "vegetales", "semillas", "frutas", "granos" → FITOSANITARIO
- Si menciona "veterinary", "animal health", "zoosanitario", "ganado", "carne", "aves", "pesca" → ZOOSANITARIO
- Si menciona "health certificate", "sanitary certificate", "certificate of health", "inspección sanitaria",
  "food safety", "hygiene", "apto para consumo humano", "human consumption" → HEALTH CERTIFICATE
- Si menciona "certificate of origin", "origen", "preferential", "Form A", "GSP", "EUR.1" → CO
- Si menciona "fumigation", "fumigación", "fumigant", "aluminium phosphide", "methyl bromide",
  "fosfuro de aluminio", "bromuro de metilo", "fumigado", "disinfestation" → FUMIGATION CERTIFICATE
- Si menciona "weight certificate", "weight and quality", "weight & quality", "surveyor", "inspection co.",
  "tally", "superintendent", "GAFTA", "FOSFA", "General Survey", "SGS", "Bureau Veritas", "Cotecna"
  en el contexto de verificar peso o calidad de la carga → WEIGHT CERTIFICATE
- Si menciona "analysis report", "certificate of analysis", "alveograph", "farinograph", "extensograph",
  "rheology", "protein", "moisture", "ash", "gluten", "wet gluten", "quality report" → QUALITY CERTIFICATE
- USA "OTROS" SOLO si el documento es genuinamente inclasificable

IMPORTANTE — Reglas para QUALITY CERTIFICATE con reportes de laboratorio:
Los proveedores de materias primas (harinas, cereales, etc.) suelen incluir múltiples reportes de
laboratorio en un mismo full set: Analysis Report, Certificate of Analysis, Alveograph, Farinograph,
Extensograph, etc. Cada uno de estos es un documento independiente con su propio formulario y emisor.
REGLA CRÍTICA: agrupa en un SOLO QUALITY CERTIFICATE únicamente las páginas que sean continuación
directa del mismo formulario (ej: página 1/2 y 2/2 del mismo reporte). Si cambia el tipo de ensayo,
el formulario, el emisor o el número de muestra/lote, ES UN NUEVO QUALITY CERTIFICATE independiente.
Ejemplos de cortes dentro de "calidad":
- Analysis Report (Eris) → Certificate of Analysis (Eris) → Alveograph (Chopin) → Farinograph (Brabender)
  → Extensograph (Brabender) → cada uno es un QUALITY CERTIFICATE distinto aunque sean de la misma carga.
- Un Fumigation Certificate es SIEMPRE independiente del FITOSANITARIO, aunque ambos mencionen fumigación.

Reglas para identificar cortes entre documentos:
1. Cambio de tipo de documento (ej: de BL a INVOICE, de CO a FITOSANITARIO) → SIEMPRE es un corte.
2. Aparición de un nuevo encabezado/título de documento con nuevo emisor o número de referencia distinto
3. Cambio de emisor o de estructura de formulario
4. CADA TIPO DE DOCUMENTO ES SIEMPRE INDEPENDIENTE — Una INVOICE es solo la factura. Un CO es
   solo el certificado de origen. Un PACKING LIST es solo la lista de empaque. Un BL es solo el
   conocimiento de embarque. Un FUMIGATION CERTIFICATE es solo el certificado de fumigación.
   Ningún documento puede "contener" páginas de otro tipo distinto.
   Compartir número de PO, número de contenedor, número de factura, nombre de shipper u cualquier
   otro dato NO convierte un documento en anexo de otro — cada tipo se separa siempre.
5. PÁGINAS CONTINUAS DEL MISMO FORMULARIO — Las únicas páginas que se agrupan con un documento
   son aquellas que son continuación directa del mismo formulario: "Page 2", "Page 3", etc. del
   mismo documento (misma estructura, mismo emisor, mismo número de referencia principal y mismo
   tipo). Las páginas de términos y condiciones al dorso de un BL son parte del mismo BL.
   EXCEPCIÓN CRÍTICA: si una página etiquetada "Annex", "Anexo" o similar es de un TIPO DIFERENTE
   al documento anterior (ej: un certificado adjunto a una factura), se separa obligatoriamente.
6. NUNCA solapar páginas entre documentos — cada página pertenece a EXACTAMENTE UN documento.
   Si una página marca el inicio de un nuevo documento (cambio de tipo, nuevo formulario), asígnala
   al NUEVO documento como primera página. El documento anterior termina en la página ANTERIOR.
   Ejemplo: si la página 2 es el inicio del Quality Certificate, entonces INVOICE=[1] y QUALITY CERT=[2].
   NUNCA pongas la misma página en dos documentos distintos.
7. Verifica que la suma de todas las páginas asignadas = número total de páginas del PDF (sin huecos)
8. Un BL/MBL/HBL/AWB es SIEMPRE un documento independiente. Un CO es SIEMPRE un documento
   independiente. Un FUMIGATION CERTIFICATE es SIEMPRE independiente. Ninguno puede ser anexo
   de otro tipo bajo ninguna circunstancia.
9. LÍMITE DE PÁGINAS POR TIPO — Si el PDF tiene muchas páginas, es casi seguro un full set:
   - Un CO/Certificate of Origin raramente excede 3 páginas. Si ves más de 3 páginas con aspecto
     de CO, es muy probable que sean COs distintos o documentos de tipos diferentes concatenados.
   - Una INVOICE raramente excede 5 páginas. Un PACKING LIST raramente excede 8 páginas.
   - Un BL/AWB raramente excede 4 páginas (incluyendo términos y condiciones al dorso).
   - Un FUMIGATION CERTIFICATE raramente excede 2 páginas.
   - Un QUALITY CERTIFICATE individual (Analysis Report, Alveograph, Farinograph, Extensograph)
     raramente excede 2 páginas. Si hay más de 2 páginas de "calidad", es casi seguro que son
     varios reportes distintos concatenados — busca cortes entre ellos.
   - Si el PDF supera las 10 páginas en total, sé ESPECIALMENTE escrupuloso en buscar cortes.
     Con 10+ páginas es estadísticamente improbable que sea un solo documento.
10. MISMO TIPO, DISTINTOS DOCUMENTOS — Que dos páginas sean del mismo tipo NO significa que sean
    el mismo documento. Si aparece un nuevo formulario del mismo tipo (nuevo número de certificado,
    nueva fecha, nuevo emisor, nueva referencia principal), es un documento nuevo e independiente.
    Ejemplo: dos COs distintos con números de certificado diferentes son DOS documentos separados,
    aunque ambos sean para el mismo contenedor o la misma PO.

FORMATO DE RESPUESTA:
Responde ÚNICAMENTE con un objeto JSON válido. Sin texto adicional, sin markdown, sin explicaciones.

Si el PDF contiene UN SOLO documento:
{
  "es_fullset": false,
  "tipo": "<tipo del documento, usando los tipos listados arriba>",
  "referencia": "<número de referencia principal del documento, o \"\" si no aparece>",
  "es_continuacion": <true si este bloque es continuación directa del documento anterior indicado en el contexto, false si es un documento nuevo>,
  "razon": "<explicación breve>"
}

Si el PDF contiene MÚLTIPLES documentos:
{
  "es_fullset": true,
  "documentos": [
    {
      "paginas": [1, 2],
      "tipo": "BL",
      "referencia": "266241142",
      "descripcion": "Bill of Lading Maersk"
    },
    {
      "paginas": [3],
      "tipo": "INVOICE",
      "referencia": "FDC2026000000205",
      "descripcion": "Commercial Invoice OBA Food"
    }
  ]
}

Reglas del JSON:
- "paginas" son números enteros empezando en 1 (no en 0)
- "tipo" debe ser exactamente uno de los tipos listados arriba (BL, MBL, HBL, AWB, INVOICE, PACKING LIST, CO, FITOSANITARIO, ZOOSANITARIO, HEALTH CERTIFICATE, WEIGHT CERTIFICATE, QUALITY CERTIFICATE, FUMIGATION CERTIFICATE, OTROS)
- "referencia" es el número principal del documento (BL number, invoice number, cert number, etc.) — usa "" si no aparece
- "descripcion" es una frase corta y descriptiva del contenido real
- Los rangos de páginas deben cubrir TODAS las páginas del PDF sin solapamientos ni huecos\
"""


# ---------------------------------------------------------------------------
# Llamada a Claude para detectar segmentos (PDF enviado directamente)
# ---------------------------------------------------------------------------

# Páginas máximas por bloque (sin contar la página de overlap del bloque anterior).
# 8 páginas permite bloques granulares para PDFs de 40-50 páginas con múltiples juegos de BL.
_MAX_PAGINAS_POR_BLOQUE = 8


def _pdf_a_b64(ruta_pdf: str) -> str | None:
    """Lee un PDF y lo devuelve como base64, o None si falla."""
    try:
        with open(ruta_pdf, "rb") as f:
            return base64.standard_b64encode(f.read()).decode("utf-8")
    except Exception as e:
        print(f"  [FULLSET] No se pudo leer el PDF: {e}")
        return None


def _recortar_pdf_temporal(doc_original, paginas_inicio: int, paginas_fin: int) -> str | None:
    """
    Recorta páginas [paginas_inicio, paginas_fin] (base 1) del doc en un archivo temporal.
    Devuelve la ruta del archivo temporal, o None si falla.
    El llamador es responsable de eliminar el archivo cuando termine.
    """
    import tempfile
    try:
        import fitz
        doc_bloque = fitz.open()
        for idx in range(paginas_inicio - 1, paginas_fin):  # base 0
            doc_bloque.insert_pdf(doc_original, from_page=idx, to_page=idx)
        # delete=False + cerrar antes de guardar para evitar bloqueo en Windows
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        ruta_tmp = tmp.name
        tmp.close()  # cerrar el handle antes de que fitz escriba
        doc_bloque.save(ruta_tmp)
        doc_bloque.close()
        return ruta_tmp
    except Exception as e:
        print(f"  [FULLSET] Error recortando bloque temporal: {e}")
        return None


def _llamar_claude_sobre_pdf_b64(
    client,
    pdf_b64: str,
    nombre_archivo: str,
    num_paginas_bloque: int,
    pagina_inicio: int,
    total_paginas_pdf: int,
    contexto_anterior: dict | None = None,
    paginas_overlap: int = 0,
) -> dict | None:
    """
    Llama a Claude con un PDF (en base64) y devuelve el dict parseado, o None si falla.
    pagina_inicio indica en qué página del PDF original empieza este bloque (base 1).
    contexto_anterior: dict con 'tipo' y 'referencia' del último segmento del bloque previo.
    paginas_overlap: número de páginas de solapamiento incluidas AL INICIO del PDF enviado
                     (pertenecen al bloque anterior — no deben aparecer en la respuesta JSON).
    """
    pagina_real_inicio = pagina_inicio  # primera página real (sin overlap)
    pagina_real_fin    = pagina_inicio + num_paginas_bloque - 1

    # Las páginas reales dentro del PDF enviado a Claude van desde (overlap+1) hasta (overlap+tam_bloque)
    pagina_bloque_inicio = paginas_overlap + 1   # primera página real dentro del bloque enviado
    pagina_bloque_fin    = paginas_overlap + num_paginas_bloque

    if contexto_anterior:
        ctx_texto = (
            f"\nCONTEXTO: El bloque ANTERIOR terminó con un documento de tipo '{contexto_anterior['tipo']}'"
            + (f" con referencia '{contexto_anterior['referencia']}'" if contexto_anterior.get('referencia') else "")
            + ".\n"
        )
    else:
        ctx_texto = ""

    if paginas_overlap > 0:
        overlap_texto = (
            f"IMPORTANTE: Las primeras {paginas_overlap} página(s) de este PDF son de CONTEXTO (del bloque anterior) "
            f"— NO las incluyas en tu respuesta JSON. Solo clasifica las páginas {pagina_bloque_inicio} a {pagina_bloque_fin} "
            f"del PDF enviado (que corresponden a las páginas {pagina_real_inicio}–{pagina_real_fin} del PDF completo).\n"
            f"Los números de página en tu JSON deben ser {pagina_bloque_inicio} a {pagina_bloque_fin} (dentro de este PDF).\n"
        )
    else:
        overlap_texto = (
            f"Los números de página en tu respuesta JSON deben corresponder a las páginas de ESTE BLOQUE "
            f"(1 a {num_paginas_bloque}), no del PDF completo.\n"
        )

    try:
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=2000,
            system=_PROMPT_SISTEMA_FULLSET,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": (
                            f"Archivo: {nombre_archivo}\n"
                            f"Total de páginas del PDF completo: {total_paginas_pdf}\n"
                            f"Este bloque cubre las páginas {pagina_real_inicio} a {pagina_real_fin} del PDF completo ({num_paginas_bloque} páginas reales).\n"
                            f"{overlap_texto}"
                            f"{ctx_texto}"
                            f"La primera página real (pág. {pagina_bloque_inicio} de este PDF) puede ser:\n"
                            f"  a) Continuación directa del documento anterior (misma estructura, mismo número de referencia), O\n"
                            f"  b) El inicio de un documento NUEVO (aunque sea del mismo tipo).\n"
                            f"Indica 'es_continuacion': true SOLO si la pág. {pagina_bloque_inicio} es genuinamente el mismo documento continuando.\n\n"
                            f"¿Las páginas reales ({pagina_bloque_inicio}–{pagina_bloque_fin}) contienen uno o varios documentos?"
                        ),
                    },
                ],
            }],
        )
        texto_respuesta = respuesta.content[0].text.strip()
        datos = _parsear_json_respuesta(texto_respuesta, fallback={})
        if not datos:
            return datos

        # Ajustar números de página: Claude devuelve páginas relativas al PDF enviado (con overlap).
        # Necesitamos páginas relativas al bloque real (sin overlap), base 1 dentro del bloque.
        if paginas_overlap > 0 and datos.get("es_fullset") and datos.get("documentos"):
            for seg in datos["documentos"]:
                # Quitar páginas de overlap y rebasar al rango real del bloque
                seg["paginas"] = [p - paginas_overlap for p in seg.get("paginas", []) if p > paginas_overlap]
        return datos
    except Exception as e:
        print(f"  [FULLSET] Error en API Claude ({type(e).__name__}): {e}")
        log_error("separador_fullset", "FULL-002", nombre_archivo, f"{type(e).__name__}: {e}")
        return None


def _detectar_segmentos_con_claude(
    ruta_pdf: str,
    nombre_archivo: str,
    num_paginas: int | None = None,
    doc_fitz=None,
) -> dict | None:
    """
    Envía el PDF completo a Claude API como documento base64.
    Si el PDF es demasiado grande (excede el límite de tokens), lo analiza en bloques
    de _MAX_PAGINAS_POR_BLOQUE páginas y consolida los resultados.

    Returns:
        Dict con la respuesta de Claude, o None si falla.
        Ejemplos:
          {"es_fullset": False, "razon": "..."}
          {"es_fullset": True, "documentos": [...]}
    """
    if not ANTHROPIC_API_KEY:
        print("  [FULLSET] ANTHROPIC_API_KEY no configurada")
        return None

    try:
        import anthropic
        client = anthropic.Anthropic(
            api_key=ANTHROPIC_API_KEY,
            timeout=float(CLAUDE_TIMEOUT_SEGUNDOS),
        )
    except Exception as e:
        print(f"  [FULLSET] Error inicializando cliente Anthropic: {e}")
        return None

    # --- Intento 1: enviar el PDF completo ---
    pdf_b64 = _pdf_a_b64(ruta_pdf)
    if pdf_b64 is None:
        return None

    # Ajustar el texto del mensaje según el número de páginas: si hay 3+ páginas sin
    # indicador en el nombre, ser más explícito con Claude para que revise con cuidado.
    _pags_str = str(num_paginas) if num_paginas else "desconocido"
    if num_paginas and num_paginas >= 3:
        _pregunta_usuario = (
            f"Archivo: {nombre_archivo}\n"
            f"Total de páginas del PDF: {_pags_str}\n\n"
            f"ATENCIÓN: Este PDF tiene {_pags_str} páginas. En documentos de comercio exterior "
            f"es muy frecuente que los proveedores concatenen varios documentos en un solo PDF "
            f"(BL, Invoice, Packing List, Certificados, etc.). El nombre del archivo puede ser "
            f"genérico y no indicar que sea un full set.\n"
            f"Revisa CADA PÁGINA con cuidado. ¿Este PDF contiene un solo documento o varios documentos concatenados?"
        )
    else:
        _pregunta_usuario = (
            f"Archivo: {nombre_archivo}\n"
            f"Total de páginas del PDF: {_pags_str}\n\n"
            f"¿Este PDF contiene un solo documento o varios documentos concatenados?"
        )

    try:
        respuesta = _llamar_api(
            client,
            model=CLAUDE_MODELO,
            max_tokens=2000,
            system=_PROMPT_SISTEMA_FULLSET,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": _pregunta_usuario,
                    },
                ],
            }],
        )
        texto_respuesta = respuesta.content[0].text.strip()
        datos = _parsear_json_respuesta(texto_respuesta, fallback={})
        if datos:
            return datos
        print(f"  [FULLSET] JSON invalido en respuesta de Claude")
        log_advertencia("separador_fullset", "FULL-001", nombre_archivo, "JSON invalido o vacío")
        return None

    except Exception as e:
        # Si el error es por exceso de tokens, intentar análisis por bloques
        es_too_long = "too long" in str(e).lower() or "prompt is too long" in str(e).lower()
        if not es_too_long:
            print(f"  [FULLSET] Error en API Claude ({type(e).__name__}): {e}")
            log_error("separador_fullset", "FULL-002", nombre_archivo, f"{type(e).__name__}: {e}")
            return None

    # --- Intento 2: análisis por bloques ---
    if num_paginas is None or doc_fitz is None:
        print(f"  [FULLSET] PDF demasiado grande y no hay doc_fitz disponible para bloques — abortando")
        return None

    print(f"  [FULLSET] PDF demasiado grande para una sola llamada — analizando en bloques de {_MAX_PAGINAS_POR_BLOQUE} páginas")

    import os
    import tempfile

    todos_los_segmentos = []   # lista de dicts con páginas ya ajustadas al PDF original
    es_fullset_global   = False

    bloque_num = 0
    pagina_inicio = 1
    while pagina_inicio <= num_paginas:
        pagina_fin   = min(pagina_inicio + _MAX_PAGINAS_POR_BLOQUE - 1, num_paginas)
        bloque_num  += 1
        tam_bloque   = pagina_fin - pagina_inicio + 1

        # Incluir 1 página de solapamiento del bloque anterior (contexto visual para detectar
        # si la primera página del bloque es continuación o documento nuevo)
        pagina_overlap = pagina_inicio - 1 if pagina_inicio > 1 else None
        pagina_recorte_inicio = pagina_overlap if pagina_overlap else pagina_inicio
        paginas_overlap_en_bloque = 1 if pagina_overlap else 0

        print(f"  [FULLSET] Bloque {bloque_num}: páginas {pagina_inicio}–{pagina_fin}" +
              (f" (overlap pág. {pagina_overlap})" if pagina_overlap else ""))

        ruta_tmp = _recortar_pdf_temporal(doc_fitz, pagina_recorte_inicio, pagina_fin)
        if ruta_tmp is None:
            pagina_inicio = pagina_fin + 1
            continue

        # Pasar contexto del último segmento para que Claude detecte continuaciones vs. documentos nuevos
        ctx_anterior = None
        if todos_los_segmentos:
            ultimo = todos_los_segmentos[-1]
            ctx_anterior = {"tipo": ultimo["tipo"], "referencia": ultimo.get("referencia", "")}

        try:
            b64_bloque = _pdf_a_b64(ruta_tmp)
            if b64_bloque is None:
                pagina_inicio = pagina_fin + 1
                continue

            datos_bloque = _llamar_claude_sobre_pdf_b64(
                client, b64_bloque, nombre_archivo,
                num_paginas_bloque=tam_bloque,
                pagina_inicio=pagina_inicio,
                total_paginas_pdf=num_paginas,
                contexto_anterior=ctx_anterior,
                paginas_overlap=paginas_overlap_en_bloque,
            )
        finally:
            try:
                os.unlink(ruta_tmp)
            except Exception:
                pass

        if not datos_bloque:
            # Bloque sin respuesta válida → tratar sus páginas como un único segmento OTROS
            todos_los_segmentos.append({
                "paginas": list(range(pagina_inicio, pagina_fin + 1)),
                "tipo": "OTROS",
                "referencia": "",
                "descripcion": f"Bloque {bloque_num} sin clasificar",
            })
            pagina_inicio = pagina_fin + 1
            continue

        if not datos_bloque.get("es_fullset", False):
            # Bloque es un único documento — usar tipo/referencia que Claude devolvió,
            # y solo fusionar con el segmento anterior si Claude indicó es_continuacion=true
            razon           = datos_bloque.get("razon", "")
            tipo_bloque     = datos_bloque.get("tipo") or (ctx_anterior["tipo"] if ctx_anterior else "OTROS")
            ref_bloque      = datos_bloque.get("referencia", "")
            es_continuacion = datos_bloque.get("es_continuacion", False)

            print(f"  [FULLSET] Bloque {bloque_num}: documento único — {tipo_bloque} ref={ref_bloque!r} continuacion={es_continuacion} ({razon})")

            if es_continuacion and ctx_anterior and todos_los_segmentos:
                # Extender el segmento anterior con las páginas de este bloque
                todos_los_segmentos[-1]["paginas"].extend(range(pagina_inicio, pagina_fin + 1))
                todos_los_segmentos[-1]["paginas"].sort()
                print(f"  [FULLSET]   → Fusionado con segmento anterior ({todos_los_segmentos[-1]['tipo']})")
            else:
                es_fullset_global = True
                todos_los_segmentos.append({
                    "paginas":     list(range(pagina_inicio, pagina_fin + 1)),
                    "tipo":        tipo_bloque,
                    "referencia":  ref_bloque,
                    "descripcion": f"{tipo_bloque} páginas {pagina_inicio}–{pagina_fin}",
                })
        else:
            es_fullset_global = True
            for seg in datos_bloque.get("documentos", []):
                # Ajustar páginas del bloque (base 1 dentro del bloque) al PDF original
                paginas_ajustadas = [p + pagina_inicio - 1 for p in seg.get("paginas", [])]
                todos_los_segmentos.append({
                    "paginas":     paginas_ajustadas,
                    "tipo":        seg.get("tipo", "OTROS"),
                    "referencia":  seg.get("referencia", ""),
                    "descripcion": seg.get("descripcion", ""),
                })
                print(f"    Segmento: págs {paginas_ajustadas} → {seg.get('tipo')} ({seg.get('referencia', '')})")

        pagina_inicio = pagina_fin + 1

    # Si ningún bloque fue fullset, consolidar todos como un único documento
    if not es_fullset_global:
        return {"es_fullset": False, "razon": "Ningún bloque contiene múltiples documentos"}

    # Fusionar segmentos contiguos del mismo tipo si el último del bloque anterior
    # y el primero del siguiente son el mismo documento partido por el corte de bloque
    segmentos_fusionados = _fusionar_segmentos_contiguos(todos_los_segmentos, num_paginas)

    # Corregir solapamientos y huecos antes de devolver
    segmentos_fusionados = _resolver_solapamientos_y_huecos(segmentos_fusionados, num_paginas)

    return {"es_fullset": True, "documentos": segmentos_fusionados}


def _resolver_solapamientos_y_huecos(documentos: list[dict], num_paginas_total: int) -> list[dict]:
    """
    Corrige solapamientos y huecos en los rangos de páginas devueltos por Claude.

    Solapamientos: cuando la última página del doc A coincide con la primera del doc B,
    esa página pertenece al doc B (el nuevo documento empieza ahí). Se elimina del doc A.

    Huecos: páginas del PDF que quedaron sin asignar se adjudican al documento
    cuya última página es la inmediatamente anterior.
    """
    if not documentos:
        return documentos

    # Paso 1 — resolver solapamientos entre documentos consecutivos
    correcciones = 0
    for i in range(len(documentos) - 1):
        pags_a = documentos[i]["paginas"]
        pags_b = documentos[i + 1]["paginas"]
        if not pags_a or not pags_b:
            continue
        ultima_a  = max(pags_a)
        primera_b = min(pags_b)
        if ultima_a >= primera_b:
            antes = list(documentos[i]["paginas"])
            documentos[i]["paginas"] = [p for p in pags_a if p < primera_b]
            correcciones += 1
            print(f"  [FULLSET] Solapamiento corregido: {documentos[i]['tipo']} {antes} -> {documentos[i]['paginas']} (pag {primera_b} cedida a {documentos[i+1]['tipo']})")

    # Eliminar documentos que quedaron sin páginas tras la corrección
    documentos = [d for d in documentos if d["paginas"]]

    # Paso 2 — rellenar huecos (páginas no asignadas)
    asignadas = {p for d in documentos for p in d["paginas"]}
    huecos = [p for p in range(1, num_paginas_total + 1) if p not in asignadas]
    if huecos:
        print(f"  [FULLSET] Páginas sin asignar detectadas: {huecos} — asignando al doc anterior")
        for p in sorted(huecos):
            # Asignar al documento cuya última página es p-1
            asignado = False
            for doc in documentos:
                if doc["paginas"] and max(doc["paginas"]) == p - 1:
                    doc["paginas"].append(p)
                    doc["paginas"].sort()
                    asignado = True
                    break
            if not asignado:
                # Fallback: asignar al último documento
                documentos[-1]["paginas"].append(p)
                documentos[-1]["paginas"].sort()

    if correcciones:
        print(f"  [FULLSET] {correcciones} solapamiento(s) corregido(s) automáticamente")

    return documentos


def _fusionar_segmentos_contiguos(segmentos: list[dict], num_paginas_total: int) -> list[dict]:
    """
    Fusiona segmentos consecutivos que sean del mismo tipo y cuyas páginas sean contiguas,
    para corregir el caso en que un documento quede partido en el borde entre dos bloques.
    Solo fusiona si ambos segmentos tienen referencia vacía o la misma referencia.
    """
    if not segmentos:
        return segmentos

    resultado = [segmentos[0].copy()]
    for seg in segmentos[1:]:
        prev = resultado[-1]
        misma_referencia = (
            not prev["referencia"] and not seg["referencia"]
        ) or (prev["referencia"] and prev["referencia"] == seg["referencia"])
        paginas_contiguas = (
            prev["paginas"] and seg["paginas"] and
            max(prev["paginas"]) + 1 == min(seg["paginas"])
        )
        if prev["tipo"] == seg["tipo"] and misma_referencia and paginas_contiguas:
            prev["paginas"] = prev["paginas"] + seg["paginas"]
            print(f"  [FULLSET] Fusionados segmentos contiguos: {prev['tipo']} → págs {prev['paginas']}")
        else:
            resultado.append(seg.copy())

    return resultado


# ---------------------------------------------------------------------------
# Recorte del PDF por rangos de páginas
# ---------------------------------------------------------------------------

def _recortar_pdf(
    doc_original,
    paginas: list[int],
    ruta_salida: str,
    ruta_pdf: str,
) -> bool:
    """
    Extrae las páginas indicadas del PDF original ya abierto y las guarda en ruta_salida.

    Args:
        doc_original: Objeto fitz.Document ya abierto.
        paginas: Lista de números de página (base 1).

    Returns:
        True si tuvo éxito, False si falló.
    """
    try:
        import fitz

        doc_nuevo = fitz.open()

        for num_pagina in paginas:
            idx = num_pagina - 1  # PyMuPDF usa base 0
            if 0 <= idx < len(doc_original):
                doc_nuevo.insert_pdf(doc_original, from_page=idx, to_page=idx)
            else:
                print(f"  [FULLSET] Página {num_pagina} fuera de rango — ignorada")

        doc_nuevo.save(ruta_salida)
        doc_nuevo.close()
        return True

    except Exception as e:
        print(f"  [FULLSET] Error recortando PDF: {e}")
        log_error("separador_fullset", "FULL-003", ruta_pdf, f"Error al recortar: {e}")
        return False


# ---------------------------------------------------------------------------
# Función principal pública
# ---------------------------------------------------------------------------

def separar_fullset(
    ruta_pdf: str,
    nombre_archivo: str,
    carpeta_temp: str | None = None,
    _es_fragmento: bool = False,
) -> list[dict] | None:
    """
    Detecta si un PDF es un "Full Set" y lo separa en documentos individuales.

    Args:
        ruta_pdf        : Ruta al PDF original en disco.
        nombre_archivo  : Nombre del archivo (para logs y nombres de salida).
        carpeta_temp    : Carpeta donde guardar los PDFs separados.
                          Si es None, usa el mismo directorio que el original.

    Returns:
        None  — si el PDF NO es un full set (procesarlo normalmente).
        []    — si es un full set pero falló la separación (error).
        Lista de dicts — un dict por fragmento, con:
            {
              "ruta":        "ruta/al/fragmento.pdf",
              "nombre":      "nombre_del_fragmento.pdf",
              "tipo_sugerido": "BL",          # sugerencia de Claude (orientativa)
              "referencia":  "266241142",      # número de referencia del documento
              "descripcion": "Bill of Lading Maersk",
            }
    """
    ruta = Path(ruta_pdf)

    if ruta.suffix.lower() != ".pdf":
        return None  # Solo aplica a PDFs

    # Los fragmentos ya cortados no se re-analizan como full set
    if _es_fragmento:
        return None

    # Abrir el PDF una sola vez — se reutiliza para contar páginas y recortar fragmentos
    doc_fitz = None
    num_paginas = None
    try:
        import fitz
        doc_fitz = fitz.open(ruta_pdf)
        num_paginas = len(doc_fitz)
    except Exception:
        pass  # Si falla, dejamos que Claude decida

    # Detectar si el nombre indica explícitamente que es un full set.
    # Patrones: "FULL SET", "FULLSET", o el nombre termina en " SET" (ej: "12194 Set.pdf")
    _nombre_upper = nombre_archivo.upper()
    _stem_upper   = Path(nombre_archivo).stem.upper()
    _es_fullset_por_nombre = (
        "FULL SET" in _nombre_upper
        or "FULLSET" in _nombre_upper
        or _stem_upper.endswith(" SET")
    )

    print(f"  [FULLSET] Analizando: {nombre_archivo}" + (f" ({num_paginas} pags.)" if num_paginas else "")
          + (" [nombre indica FULL SET]" if _es_fullset_por_nombre else ""))

    if num_paginas == 1:
        if doc_fitz:
            doc_fitz.close()
        print(f"  [FULLSET] 1 sola pagina — no es full set")
        return None

    # Enviar el PDF completo a Claude para detección.
    # - Nombre indica FULL SET → 3 intentos (Claude puede dudar en expedientes complejos).
    # - PDF de 3+ páginas sin indicador de nombre → 2 intentos: un solo intento no es suficiente
    #   para descartar un full set cuando hay varias páginas.
    # - PDF de 2 páginas sin indicador → 1 intento (un BL de 2 págs. es el caso más común).
    if _es_fullset_por_nombre:
        _max_intentos = 3
    elif num_paginas and num_paginas >= 3:
        _max_intentos = 2
    else:
        _max_intentos = 1

    resultado = None
    for _intento in range(1, _max_intentos + 1):
        if _intento > 1:
            _razon_reintento = "nombre indica FULL SET" if _es_fullset_por_nombre else f"PDF de {num_paginas} pags."
            print(f"  [FULLSET] Reintento {_intento}/{_max_intentos} ({_razon_reintento})...")
        resultado = _detectar_segmentos_con_claude(ruta_pdf, nombre_archivo, num_paginas, doc_fitz)
        if resultado is not None and resultado.get("es_fullset", False):
            break  # Detectado correctamente — no reintentar
        if resultado is not None and not resultado.get("es_fullset", False) and _max_intentos == 1:
            break  # PDF de 2 páginas sin indicador: confiar en el primer intento

    if resultado is None:
        if doc_fitz:
            doc_fitz.close()
        if _es_fullset_por_nombre:
            log_advertencia("separador_fullset", "FULL-004", nombre_archivo,
                            "Claude no respondió tras múltiples intentos — el nombre indica FULL SET, revisar manualmente")
            print(f"  [FULLSET] ADVERTENCIA: nombre indica FULL SET pero Claude no respondió — procesando como documento único, revisar manualmente")
        else:
            print(f"  [FULLSET] Claude no respondió — procesando como documento único")
        return None

    if not resultado.get("es_fullset", False):
        if doc_fitz:
            doc_fitz.close()
        razon = resultado.get("razon", "")
        if _es_fullset_por_nombre:
            log_advertencia("separador_fullset", "FULL-005", nombre_archivo,
                            f"Nombre indica FULL SET pero Claude no detectó múltiples documentos tras {_max_intentos} intento(s) — revisar manualmente. Razón: {razon}")
            print(f"  [FULLSET] ADVERTENCIA: nombre indica FULL SET pero Claude no detectó separaciones tras {_max_intentos} intento(s) — revisar manualmente")
        else:
            print(f"  [FULLSET] Documento único detectado. Razón: {razon}")
        return None

    # 3. Es un full set — obtener los segmentos
    documentos = resultado.get("documentos", [])
    if not documentos:
        if doc_fitz:
            doc_fitz.close()
        print(f"  [FULLSET] Claude indicó fullset pero no retornó segmentos")
        return None

    # Corregir solapamientos y huecos antes de recortar
    if num_paginas:
        documentos = _resolver_solapamientos_y_huecos(documentos, num_paginas)

    print(f"  [FULLSET] Full Set detectado: {len(documentos)} documento(s)")
    for i, doc in enumerate(documentos, 1):
        print(f"    {i}. Paginas {doc.get('paginas')} -> {doc.get('tipo')} ({doc.get('referencia', '')})")

    # 4. Recortar el PDF en fragmentos (reutilizando el doc ya abierto)
    if carpeta_temp is None:
        carpeta_temp = str(ruta.parent)
    carpeta_salida = Path(carpeta_temp)

    stem_original = ruta.stem
    fragmentos = []

    try:
        for i, doc in enumerate(documentos, 1):
            paginas_doc = [p for p in doc.get("paginas", []) if 1 <= p <= (num_paginas or 9999)]
            tipo         = doc.get("tipo", "OTROS")
            referencia   = doc.get("referencia", "")
            descripcion  = doc.get("descripcion", "")

            if not paginas_doc:
                print(f"  [FULLSET] Segmento {i} sin páginas definidas — ignorado")
                continue

            # Nombre del fragmento: ORIGINAL_STEM_01_BL_266241142.pdf
            ref_clean = re.sub(r'[^\w\-]', '', referencia)[:20]
            nombre_fragmento = f"{stem_original}_{i:02d}_{tipo}"
            if ref_clean:
                nombre_fragmento += f"_{ref_clean}"
            nombre_fragmento += ".pdf"

            ruta_salida = carpeta_salida / nombre_fragmento

            exito = _recortar_pdf(doc_fitz, paginas_doc, str(ruta_salida), ruta_pdf)
            if exito:
                print(f"  [FULLSET] Generado: {nombre_fragmento} ({len(paginas_doc)} pág.)")
                fragmentos.append({
                    "ruta":           str(ruta_salida),
                    "nombre":         nombre_fragmento,
                    "tipo_sugerido":  tipo,
                    "referencia":     referencia,
                    "descripcion":    descripcion,
                })
            else:
                print(f"  [FULLSET] Error al generar fragmento {i} — saltado")
    finally:
        if doc_fitz:
            doc_fitz.close()

    if not fragmentos:
        print(f"  [FULLSET] No se generó ningún fragmento válido")
        return []

    return fragmentos
