# CLASIFICADOR_DOCS — Documentación del Proyecto

Sistema automatizado de clasificación y organización de documentos de comercio exterior recibidos por correo electrónico. El sistema monitorea un buzón de Outlook, clasifica cada adjunto usando Claude API (IA), detecta inconsistencias en los documentos, y los sube organizados a SharePoint.

---

## Tabla de Contenidos

1. [Descripción general](#1-descripción-general)
2. [Arquitectura del sistema](#2-arquitectura-del-sistema)
3. [Estructura de carpetas](#3-estructura-de-carpetas)
4. [Flujo completo del pipeline](#4-flujo-completo-del-pipeline)
5. [Módulos principales](#5-módulos-principales)
6. [Catálogo de tipos de documentos](#6-catálogo-de-tipos-de-documentos)
7. [Clasificación con Claude API](#7-clasificación-con-claude-api)
8. [Separador de Full Set](#8-separador-de-full-set)
9. [Validaciones de calidad](#9-validaciones-de-calidad)
10. [Generación de respuestas al proveedor](#10-generación-de-respuestas-al-proveedor)
11. [Cola de aprobación](#11-cola-de-aprobación)
12. [Integración con SharePoint (Graph API)](#12-integración-con-sharepoint-graph-api)
13. [Seguridad de adjuntos](#13-seguridad-de-adjuntos)
14. [Generación de reportes Word](#14-generación-de-reportes-word)
15. [Configuración](#15-configuración)
16. [Dependencias](#16-dependencias)
17. [Comandos de ejecución](#17-comandos-de-ejecución)
18. [Códigos de error y logs](#18-códigos-de-error-y-logs)

---

## 1. Descripción general

El sistema resuelve el problema de procesar manualmente decenas de correos diarios de proveedores internacionales que adjuntan documentación de importación (BL, Facturas, Packing Lists, Certificados de Origen, etc.).

**Qué hace el sistema:**

1. Escucha en tiempo real la bandeja de entrada de Outlook mediante eventos COM (no hace polling).
2. Por cada correo nuevo, extrae los adjuntos (filtrando imágenes inline y firmas).
3. Pasa cada adjunto por el agente de seguridad (Bitdefender).
4. Clasifica el tipo de documento usando Claude API como motor de IA.
5. Detecta inconsistencias en los documentos (números de contenedor inválidos, pesos erróneos, etc.).
6. Sube el archivo a la carpeta correcta en SharePoint bajo el número de OC correspondiente.
7. Si hay inconsistencias, genera un borrador de respuesta al proveedor en formato Word y lo encola para aprobación humana.

---

## 2. Arquitectura del sistema

```
Outlook (buzón)
    │
    ▼  (evento COM ItemAdd — tiempo real)
monitor_correos.py
    │
    ├── agente_seguridad.py  (Bitdefender + validación extensión/tamaño)
    │
    ├── clasificador.py  (orquestador de clasificación)
    │       ├── separador_fullset.py  (detecta PDFs con múltiples docs)
    │       ├── clasificador_claude.py  (Claude API — clasificación + inconsistencias)
    │       │       └── validador_campos.py  (validaciones deterministas adicionales)
    │       └── Fallback: keywords en nombre y contenido del archivo
    │
    ├── graph_client.py  (sube a SharePoint via Microsoft Graph API)
    │
    ├── clasificador_claude.py  (genera respuesta al proveedor)
    │       └── contrastar_sugerencia_vs_inconsistencias()
    │
    ├── cola_aprobacion.py  (encola sugerencia para aprobación humana)
    │
    └── generador_reporte.py  (genera reporte Word de inconsistencias)
```

---

## 3. Estructura de carpetas

```
CLASIFICADOR_DOCS/
├── clasificador_documentos/
│   ├── Aprobaciones/           # Cola de aprobación de respuestas a proveedores
│   │   ├── cola_aprobacion.py
│   │   ├── watcher_aprobaciones.py
│   │   └── pendientes_aprobacion/  # Archivos .txt pendientes de revisar
│   ├── Catalogo/               # Tipos de documentos y ejemplos
│   │   ├── catalogo_tipos.py   # Definición de los 17 tipos de documentos
│   │   ├── catalogo_ejemplos.py
│   │   └── checklist_excel.py  # Checklists por tipo de documento
│   ├── configuracion/
│   │   ├── .env                # Variables secretas (NO subir a Git)
│   │   └── ajustes.py          # Carga .env y expone constantes globales
│   ├── Herramientas/
│   │   └── generar_catalogo_txt.py
│   ├── Inconsistencias/        # Archivos .txt de inconsistencias detectadas
│   ├── Integraciones/
│   │   ├── graph_client.py     # Microsoft Graph API (SharePoint + correos)
│   │   ├── monitor_correos.py  # Punto de entrada — escucha Outlook COM
│   │   ├── monitor_sin_claude.py  # Versión sin IA (solo keywords)
│   │   ├── notificador.py      # Alertas Teams
│   │   └── procesados.txt      # Registro de EntryIDs ya procesados
│   ├── Nucleo/
│   │   ├── clasificador.py         # Orquestador del pipeline de clasificación
│   │   ├── clasificador_claude.py  # Motor Claude API
│   │   ├── extractor_texto.py      # Extracción de texto (PDF, Word, Excel, OCR)
│   │   ├── limpiar_duplicados.py
│   │   ├── recopilador_documentos.py
│   │   ├── separador_fullset.py    # Separación de PDFs multi-documento
│   │   └── validador_campos.py     # Validaciones deterministas
│   ├── Pruebas/                # Scripts de prueba manuales
│   ├── Registros/              # Reportes Word generados (.docx)
│   ├── Reportes/
│   │   └── generador_reporte.py   # Genera reportes Word de inconsistencias
│   └── Utilidades/
│       ├── logger_errores.py       # Logger centralizado
│       ├── programar_monitor.py
│       ├── reprocesar_adjunto.py
│       └── verificador_cadenas.py  # Detecta conflictos en hilos de correo
└── Diagramas/
    └── Docs/                   # Diagramas de flujo (.drawio)
```

---

## 4. Flujo completo del pipeline

A continuación se describe paso a paso lo que ocurre cuando llega un correo nuevo:

### Paso 1 — Detección del correo (monitor_correos.py)

- El monitor está suscrito al evento COM `ItemAdd` de la bandeja de entrada de Outlook.
- Al recibir un `MailItem`, verifica que no haya sido procesado antes (compara `EntryID` y una clave secundaria `remitente|asunto|fecha` en `procesados.txt`).
- Extrae el asunto y remitente, e intenta detectar el número de PO y BL desde el asunto con expresiones regulares.
- Verifica el hilo de correos previos en busca de conflictos de PO/BL.

### Paso 2 — Verificación del Antivirus

- Antes de tocar cualquier adjunto, verifica que Bitdefender esté activo.
- Si el AV no está corriendo, el correo se aborta sin procesar.

### Paso 3 — Iteración de adjuntos

Por cada adjunto del correo:

- Se filtra si es imagen inline (firma, logo). Criterios: patrón `image001.png`, imagen < 100 KB, `ContentId` presente, tipo OLE embebido.
- Se filtra si la extensión está bloqueada (`.exe`, `.js`, `.bat`, `.vbs`, etc.).
- Se guarda en un archivo temporal en disco.
- Se pasa por el **Agente de Seguridad** (tamaño, extensión permitida, tiempo de espera para que el AV escanee).
- Si es ZIP o RAR, se descomprime y se procesa cada archivo interno individualmente.

### Paso 4 — Clasificación del documento (clasificador.py → clasificador_claude.py)

1. Se extrae el número de PO (obligatorio) desde: asunto del correo > nombre del archivo.
2. Si el archivo es PDF, se llama a `separador_fullset.py` para detectar si contiene múltiples documentos concatenados.
3. Se envía el archivo directamente a **Claude API** (modelo configurable, por defecto Haiku) como documento base64.
4. Claude devuelve: tipo de documento, porcentaje de certeza, justificación, e inconsistencias detectadas.
5. Si la certeza de Claude supera el umbral mínimo configurado (`CLAUDE_CERTEZA_MINIMA`), se acepta la clasificación.
6. Si no supera el umbral o Claude falla, se aplica **fallback por keywords** en el nombre del archivo.
7. Si el fallback por nombre tampoco clasifica, se aplica **fallback por keywords en el contenido** del documento.

### Paso 5 — Validaciones deterministas (validador_campos.py)

Después de la clasificación con Claude, se ejecutan validaciones de código adicionales sobre el texto del documento:

- Verificación del formato ISO 6346 de números de contenedor (4 letras + 7 dígitos, categoría U/J/Z).
- Verificación del dígito verificador ISO 6346.
- Detección de peso neto mayor al peso bruto (físicamente imposible).

Estas inconsistencias se fusionan con las detectadas por Claude.

### Paso 6 — Subida a SharePoint (graph_client.py)

- Se construye la ruta destino: `OC´s/cmer-OC-{numero_oc_8dig}/4. DOCUMENTACION/{carpeta_por_tipo}/{nombre_archivo}`.
- Se verifica si la ruta ya fue subida antes (evita duplicados).
- Se crea la estructura de carpetas si no existe.
- Se sube el archivo.
- Se registra la subida en `registros_subidas.log`.

### Paso 7 — Respuesta al proveedor (si hay inconsistencias)

- Si algún documento tuvo inconsistencias, se llama a Claude API para generar un borrador de respuesta al proveedor.
- Se ejecuta una segunda llamada de contraste para verificar que la respuesta mencione todas las inconsistencias.
- Se genera un reporte Word en `Registros/` con los detalles coloreados por severidad.
- La sugerencia de respuesta se encola en `Aprobaciones/pendientes_aprobacion/` para revisión humana.

---

## 5. Módulos principales

### monitor_correos.py
- Punto de entrada del sistema.
- Usa COM Events (`win32com.client.DispatchWithEvents`) para recibir correos en tiempo real sin polling.
- Al iniciar, hace un escaneo inicial de la bandeja procesando correos no vistos (de más reciente a más antiguo) hasta encontrar el primero ya marcado como procesado.
- Mantiene `procesados.txt` con EntryIDs y claves secundarias para resistir cambios de EntryID.

**Comando:**
```
python c:\Users\aux22.gg\Desktop\PROYECTOS\CLASIFICADOR_DOCS\clasificador_documentos\Integraciones\monitor_correos.py
```

### clasificador.py
- Orquestador central del pipeline de clasificación.
- Exporta `procesar_adjunto()` como función pública.
- Contiene las expresiones regulares para extraer PO y BL del asunto/nombre.
- Gestiona el flujo Claude → fallback nombre → fallback contenido.
- Construye las rutas destino en SharePoint con `construir_ruta()`.

### clasificador_claude.py
- Único módulo que conoce el SDK de Anthropic.
- Envía el archivo como base64 directamente a Claude (sin extracción de texto previa).
- Usa **Prompt Caching** para el system prompt: ~90% de ahorro en tokens a partir del segundo documento en la misma sesión.
- Separación de tareas: la detección de inconsistencias es una segunda llamada independiente, solo para los tipos de documentos que la requieren.
- Implementa reintentos con backoff exponencial (hasta 3 intentos) para errores HTTP 429, 529 y errores de red.
- También contiene:
  - `generar_respuesta_proveedor_consolidada()` — redacta el correo de respuesta.
  - `contrastar_sugerencia_vs_inconsistencias()` — verifica que la respuesta sea completa.

### extractor_texto.py
- Extrae texto plano de: PDF (digital), PDF escaneado (OCR con Tesseract), Word (.doc/.docx), Excel (.xlsx/.xls), Imágenes (.jpg/.png/.tiff).
- Para PDFs, primero intenta extracción digital con `pdfplumber`; si no hay texto seleccionable, aplica OCR automáticamente.
- OCR en 3 idiomas: español + inglés + chino simplificado (`spa+eng+chi_sim`).
- Nota: la extracción de texto es usada principalmente por los fallbacks. Claude recibe el archivo directamente como documento base64.

### separador_fullset.py
- Detecta PDFs que contienen múltiples documentos concatenados ("Full Set de embarque").
- Envía el PDF completo a Claude API con un prompt especializado.
- Claude devuelve los rangos de páginas y el tipo de cada segmento.
- Usa `PyMuPDF (fitz)` para cortar el PDF original en archivos individuales.
- Cada fragmento se procesa luego de forma independiente por el clasificador.

### validador_campos.py
- Validaciones deterministas que complementan a Claude.
- Implementa el algoritmo de dígito verificador ISO 6346 para números de contenedor.
- Detecta pesos físicamente imposibles (neto > bruto).
- Se aplica solo a los tipos de documento que tienen reglas: BL, MBL, HBL, INVOICE, PACKING LIST, CO, WEIGHT CERTIFICATE, FOB LETTER.

### graph_client.py
- Integración con Microsoft Graph API.
- Autenticación via MSAL (OAuth2 Client Credentials).
- Métodos principales: `buscar_carpeta_oc()`, `crear_carpeta_si_no_existe()`, `subir_archivo()`.
- Tiene modo MOCK para pruebas sin conexión real.

---

## 6. Catálogo de tipos de documentos

El sistema reconoce los siguientes 17 tipos de documentos:

| # | Tipo | Carpeta SharePoint |
|---|---|---|
| 1 | MBL | 4.05 BL-AWB-Porte definitivo |
| 2 | HBL | 4.05 BL-AWB-Porte definitivo |
| 3 | BL | 4.05 BL-AWB-Porte definitivo |
| 4 | AWB | 4.05 BL-AWB-Porte definitivo |
| 5 | INVOICE | 4.02 Factura Definitiva |
| 6 | PACKING LIST | 4.27 Packing list definitivo |
| 7 | PL + INV | 4.27 Packing list definitivo |
| 8 | CO | 4.10 Certificado Origen definitivo (COO) |
| 9 | WEIGHT CERTIFICATE | OTROS |
| 10 | QUALITY CERTIFICATE | OTROS |
| 11 | FITOSANITARIO | 4.12 Aprob Borr Cert fito origen |
| 12 | ZOOSANITARIO | 4.18 Certifi Zoos Origen |
| 13 | FOB LETTER | OTROS |
| 14 | PRINTER | 4.24 Aceptacion documentos de agencia |
| 15 | EXONERACION | 4.25 Exoneracion Hacienda |
| 16 | MARCHAMO | 4.28 Foto del Marchamo |
| 17 | OTROS | OTROS |

El catálogo completo con las descripciones semánticas que usa Claude está en [Catalogo/catalogo_tipos.py](clasificador_documentos/Catalogo/catalogo_tipos.py).

---

## 7. Clasificación con Claude API

### Motor principal

El clasificador usa **claude-haiku** (configurable en `.env`) para clasificar documentos. El modelo recibe el archivo directamente como documento base64, lo que le permite leer tanto texto digital como documentos escaneados sin necesidad de OCR previo.

### Estrategia de tres niveles

```
1. Claude API (certeza >= umbral configurado)
        ↓ si falla o certeza baja
2. Fallback keywords en nombre del archivo
        ↓ si resultado es OTROS
3. Fallback keywords en contenido del documento
```

### Umbrales de certeza

- `CLAUDE_CERTEZA_MINIMA`: umbral global (configurable en `.env`, default 70%).
- `CLAUDE_CERTEZA_POR_TIPO`: umbrales específicos por tipo (para tipos que requieren mayor precisión).

### Prompt Caching

El system prompt de clasificación (que incluye el catálogo completo de tipos) se marca con `cache_control: {"type": "ephemeral"}`. Esto permite que la API de Anthropic reutilice el procesamiento del prompt en llamadas sucesivas dentro de la misma sesión, reduciendo el costo en tokens de entrada en aproximadamente un 90%.

### Detección de inconsistencias por Claude

Claude detecta en el mismo documento:
- Campos faltantes o incorrectos (números de contenedor, fechas, puertos).
- Inconsistencias entre campos (ej: puerto de carga igual al de descarga).
- Formatos no estándar.

Esta detección es una segunda llamada independiente, solo para tipos de documento que la requieren (no se hace para MARCHAMO, EXONERACION, OTROS, etc.).

---

## 8. Separador de Full Set

Cuando un proveedor envía todos los documentos de un embarque en un único PDF ("Full Set"), el sistema los separa automáticamente.

**Flujo:**

1. Claude recibe el PDF completo y analiza si contiene múltiples documentos.
2. Si detecta segmentos, retorna un JSON con rangos de páginas y tipos:
   ```json
   {
     "es_fullset": true,
     "segmentos": [
       {"tipo": "INVOICE", "paginas": [1, 2]},
       {"tipo": "PACKING LIST", "paginas": [3, 3]},
       {"tipo": "BL", "paginas": [4, 6]}
     ]
   }
   ```
3. PyMuPDF corta el PDF original en archivos individuales temporales.
4. Cada fragmento se procesa de forma independiente por el clasificador normal.
5. Caso especial `PL + INV`: cuando ya existe un INVOICE separado en el full set, el fragmento `PL + INV` se sube únicamente como `PACKING LIST`.

---

## 9. Validaciones de calidad

Las validaciones deterministas en `validador_campos.py` complementan a Claude con reglas de negocio verificables en código:

### Validación de número de contenedor (ISO 6346)

**Formato estándar:** 4 letras + 7 dígitos (ej: `TEMU6803617`)

- Las primeras 3 letras son el código del propietario (owner code).
- La 4ª letra es la categoría del equipo: `U` (carga), `J` (equipo detachable), `Z` (trailer/chasis).
- Los siguientes 6 dígitos son el número de serie.
- El último dígito es el dígito verificador calculado por el algoritmo ISO 6346.

Si el dígito verificador declarado no coincide con el calculado, se reporta como inconsistencia de severidad **alta**.

### Validación de pesos

Detecta cuando el peso neto total declarado supera al peso bruto total. Esto es físicamente imposible (el peso bruto incluye el embalaje, por lo tanto siempre es mayor al neto).

### Niveles de severidad

| Nivel | Color en reporte | Etiqueta |
|---|---|---|
| alta | Rojo | CRÍTICO |
| media | Naranja | ADVERTENCIA |
| baja | Azul | MENOR |

---

## 10. Generación de respuestas al proveedor

Cuando se detectan inconsistencias en los documentos, el sistema genera automáticamente un borrador de correo de respuesta al proveedor:

1. **Primera llamada a Claude:** genera el texto completo de la respuesta en el idioma del proveedor (detectado automáticamente), solicitando las correcciones necesarias.
2. **Segunda llamada a Claude (contraste):** verifica que el borrador mencione todas las inconsistencias detectadas. Si falta alguna, la agrega.
3. El borrador se guarda como archivo `.txt` en `Aprobaciones/pendientes_aprobacion/`.
4. Un humano debe revisar, editar si necesita, y cambiar `DECISION: PENDIENTE` a `DECISION: APROBAR` o `DECISION: RECHAZAR`.
5. `watcher_aprobaciones.py` detecta el cambio y ejecuta la acción.

---

## 11. Cola de aprobación

**Archivo:** `Aprobaciones/cola_aprobacion.py`

Cada sugerencia se guarda como un archivo `.txt` con el formato:

```
================================================================
DECISION: PENDIENTE
Para: proveedor@empresa.com
Asunto: RE: PO 196893 - Shipping Documents
Timestamp: 2026-03-27 14:30

[Texto del borrador de respuesta]

================================================================
```

Para aprobar: cambiar `DECISION: PENDIENTE` → `DECISION: APROBAR` y guardar el archivo.
El `watcher_aprobaciones.py` detecta el cambio y envía el correo via Graph API.

---

## 12. Integración con SharePoint (Graph API)

### Ruta de destino

Los documentos se suben siguiendo la estructura:

```
Centro de Documentación Grupo Inteca Digitalizado/
└── Proveedores/
    └── Transacciones con proveedores internacionales/
        └── OC´s/
            └── cmer-OC-{numero_oc_8_digitos}/
                └── 4. DOCUMENTACION/
                    └── {carpeta_por_tipo}/
                        └── {nombre_archivo}
```

Ejemplo: `OC´s/cmer-OC-00196893/4. DOCUMENTACION/4.05 BL-AWB-Porte definitivo/BL PO 196893.pdf`

### Búsqueda fuzzy de carpeta OC

El número de OC en SharePoint puede tener variaciones de nombre (mayúsculas, espacios, guiones). `buscar_carpeta_oc()` hace una búsqueda case-insensitive y retorna el nombre real de la carpeta para evitar crear duplicados.

### Anti-duplicado

Antes de subir, se consulta `registros_subidas.log` para verificar si la ruta exacta ya fue subida en una sesión anterior. Esto evita subir el mismo archivo dos veces si el monitor se reinicia.

---

## 13. Seguridad de adjuntos

El agente de seguridad (`agente_seguridad.py`, en `Agente_Seguridad/`) verifica cada adjunto antes de clasificarlo:

1. **Extensión:** solo se procesan extensiones en la lista blanca (`.pdf`, `.xlsx`, `.docx`, `.zip`, `.rar`, etc.).
2. **Tamaño:** máximo configurable via `.env` (default 50 MB).
3. **Antivirus:** espera `N` segundos para que Bitdefender escanee el archivo antes de procesarlo.
4. **Extensiones bloqueadas directamente** (antes del AV): `.exe`, `.scr`, `.js`, `.vbs`, `.bat`, `.cmd`, `.ps1`.

Si el AV (`EPSecurityService`) no está activo, el procesamiento del correo completo se aborta.

Para entornos de prueba sin AV disponible, usar `MODO_SIN_AV=true` en el `.env`.

---

## 14. Generación de reportes Word

**Archivo:** `Reportes/generador_reporte.py`

Cada vez que se detectan inconsistencias, se genera un reporte `.docx` en `Registros/` con:

- Logo de la empresa en el encabezado.
- Fecha, remitente y asunto del correo.
- Un bloque por cada documento con inconsistencias, coloreado por severidad.
- Borrador de respuesta al proveedor al final del documento.

Los reportes se nombran con timestamp: `reporte_YYYY-MM-DD_HHmmss.docx`.

---

## 15. Configuración

Todas las variables sensibles van en `configuracion/.env`. **Este archivo NO debe subirse a Git.**

### Variables principales

| Variable | Descripción |
|---|---|
| `GRAPH_TENANT_ID` | ID del tenant de Azure AD |
| `GRAPH_CLIENT_ID` | Client ID de la aplicación registrada en Azure |
| `GRAPH_CLIENT_SECRET` | Secret de la aplicación Azure |
| `OUTLOOK_EMAIL` | Correo a monitorear |
| `SHAREPOINT_SITE_ID` | ID del sitio de SharePoint |
| `SHAREPOINT_DRIVE_ID` | ID de la biblioteca de documentos |
| `ANTHROPIC_API_KEY` | API key de Anthropic (Claude) |
| `CLAUDE_MODELO` | Modelo a usar (ej: `claude-haiku-4-5-20251001`) |
| `CLAUDE_CERTEZA_MINIMA` | Umbral mínimo de certeza (0-100, default 70) |
| `CLAUDE_TIMEOUT_SEGUNDOS` | Timeout para llamadas a Claude |
| `CLAUDE_MAX_TOKENS` | Tokens máximos en la respuesta |
| `TEAMS_WEBHOOK_PRUEBAS` | Webhook de Teams para notificaciones informativas |
| `TEAMS_WEBHOOK_PROBLEMAS` | Webhook de Teams para errores críticos |
| `MODO_MOCK` | `true` para pruebas sin conectar a Graph API real |
| `MODO_SIN_AV` | `true` para pruebas en entornos sin Bitdefender |
| `SEGURIDAD_TAMANO_MAX_MB` | Tamaño máximo de adjunto en MB |
| `TESSERACT_CMD` | Ruta al ejecutable de Tesseract para OCR |
| `POPPLER_PATH` | Ruta a Poppler para conversión PDF→imagen en OCR |

---

## 16. Dependencias

Requiere **Python 3.12**. Instalar con:

```
pip install anthropic msal requests python-dotenv pdfplumber pymupdf python-docx openpyxl xlrd pytesseract pdf2image pillow pywin32 pythoncom
```

| Librería | Uso |
|---|---|
| `anthropic` | Claude API — clasificación e inconsistencias |
| `msal` | Autenticación OAuth2 con Microsoft Graph |
| `requests` | HTTP para Graph API |
| `python-dotenv` | Carga del `.env` |
| `pdfplumber` | Extracción de texto de PDFs digitales |
| `pymupdf (fitz)` | Corte de PDFs en el separador de Full Set |
| `python-docx` | Generación de reportes Word |
| `openpyxl` | Lectura de archivos Excel .xlsx |
| `xlrd` | Lectura de archivos Excel .xls |
| `pytesseract` | OCR para PDFs escaneados e imágenes |
| `pdf2image` | Conversión PDF→imagen para OCR |
| `Pillow` | Procesamiento de imágenes para OCR |
| `pywin32` | COM automation de Outlook |
| `pythoncom` | Inicialización del entorno COM |

---

## 17. Comandos de ejecución

### Iniciar el monitor en tiempo real

```
python c:\Users\aux22.gg\Desktop\PROYECTOS\CLASIFICADOR_DOCS\clasificador_documentos\Integraciones\monitor_correos.py
```

Requisitos previos:
- Outlook abierto y con la cuenta configurada iniciada sesión.
- Bitdefender activo (o `MODO_SIN_AV=true` en `.env`).
- `MODO_MOCK=false` en `.env` para subir a SharePoint real.

### Scripts de prueba

| Script | Descripción |
|---|---|
| `Pruebas/prueba_pipeline_completo.py` | Prueba el pipeline completo con un correo de muestra |
| `Pruebas/prueba_claude_local.py` | Prueba la clasificación Claude con un archivo local |
| `Pruebas/prueba_fullset.py` | Prueba el separador de Full Set |
| `Pruebas/prueba_separador.py` | Prueba básica del separador |
| `Pruebas/prueba_contraste.py` | Prueba la generación y contraste de respuesta al proveedor |
| `Pruebas/prueba_hoy.py` | Reprocesa los correos del día actual |
| `Pruebas/reprocesar_pendientes.py` | Reprocesa correos que fallaron |
| `Pruebas/subir_fullset.py` | Sube manualmente un full set |

---

## 18. Códigos de error y logs

### Archivos de log

| Archivo | Contenido |
|---|---|
| `Utilidades/registros_eventos.log` | Eventos del sistema (BASURA, CADENA, etc.) |
| `Nucleo/registros_subidas.log` | Registro de archivos subidos a SharePoint |
| `Integraciones/registros/borradores_YYYY-MM-DD.log` | Borradores de respuesta generados |

### Códigos de error

| Código | Módulo | Descripción |
|---|---|---|
| `MON-000` | monitor_correos | No se pudo acceder al adjunto N del correo |
| `MON-001` | monitor_correos | No se pudo guardar el adjunto en disco |
| `MON-002` | monitor_correos | Error genérico procesando el mensaje |
| `MON-003` | monitor_correos | Error en el listener COM (OnItemAdd) |
| `MON-004` | monitor_correos | Cuenta de Outlook no encontrada |
| `MON-005` | monitor_correos | Error generando reporte Word de inconsistencias |
| `CLAS-001` | clasificador | No se encontró número de PO |
| `CLAS-002` | clasificador | Claude falló — usando fallback por nombre |
| `CLAS-003` | clasificador | Certeza de Claude baja — usando fallback por nombre |

---

*Documentación generada el 2026-04-01.*
