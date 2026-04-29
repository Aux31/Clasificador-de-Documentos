"""
agente_seguridad.py  —  Agente de seguridad compartido
=======================================================
Módulo único usado por todos los proyectos del workspace.

Verificaciones disponibles (se activan según los parámetros recibidos):
    SEG-001  servicio_av       — EPSecurityService (Bitdefender) en RUNNING
    SEG-002  extension         — extensión en lista blanca (solo si extensiones_ok)
    SEG-003  magic_bytes       — contenido real coincide con extensión (solo si extensiones_ok)
    SEG-004  tamano            — archivo no supera máximo en MB (solo si tamano_max_mb)
    SEG-005  esperar_av        — tras N segundos, archivo sigue en disco
    SEG-006  estabilidad       — tamaño del archivo no cambia por N segundos
    SEG-007  dominio_fuente    — URL usa HTTPS y pertenece al dominio esperado
    SEG-008  hash_integridad   — SHA-256 coincide con el valor registrado

Parámetros de ejecutar():
    ruta_local        str   — ruta absoluta al archivo en disco                   (siempre)
    nombre_orig       str   — nombre original del archivo (para extension/magic)   (opcional)
    extensiones_ok    set   — lista blanca de extensiones; si None, omite SEG-002/003
    tamano_max_mb     float — límite en MB; si None, omite SEG-004
    url_fuente        str   — URL usada para descargar el archivo; si None, omite SEG-007
    dominio_esperado  str   — dominio válido contra url_fuente; si None, omite SEG-007
    hash_esperado     str   — SHA-256 esperado; si None, omite SEG-008
    nombre_proyecto   str   — nombre para el logger (default "seguridad")
    espera_av         int   — segundos de espera para análisis AV (default 3)

Retorna:
    {
        "resultado"    : "aprobado" | "rechazado",
        "codigo_error" : str | None,
        "descripcion"  : str | None,
    }

Uso — SHAREPOINT:
    from agente_seguridad_compartido.agente_seguridad import ejecutar, verificar_servicio_av
    ok, err = verificar_servicio_av()
    res = ejecutar(
        ruta_local       = adjunto["ruta_local"],
        nombre_orig      = nombre,
        extensiones_ok   = SEGURIDAD_EXTENSIONES_PERMITIDAS,
        tamano_max_mb    = SEGURIDAD_TAMANO_MAX_MB,
        nombre_proyecto  = "sharepoint",
    )

Uso — NAVIERAS:
    from agente_seguridad_compartido.agente_seguridad import ejecutar, verificar_servicio_av
    ok, err = verificar_servicio_av()
    res = ejecutar(
        ruta_local        = ruta_tabla_temporal,
        url_fuente        = url_fuente,
        dominio_esperado  = dominio_esperado,
        hash_esperado     = hash_tabla_temporal,
        nombre_proyecto   = nombre_proyecto,
    )
"""

import os
import sys
import time
import json
import hashlib
import logging
import subprocess
from pathlib import Path
from urllib.parse import urlparse

_logger_seg = logging.getLogger("agente_seguridad")

# ---------------------------------------------------------------------------
#  Códigos de error y sus descripciones
# ---------------------------------------------------------------------------

ERROR_SERVICIO_AV_INACTIVO   = "SEG-001"
ERROR_EXTENSION_NO_PERMITIDA = "SEG-002"
ERROR_MAGIC_BYTES_INVALIDO   = "SEG-003"
ERROR_TAMANO_EXCEDIDO        = "SEG-004"
ERROR_ARCHIVO_ELIMINADO_AV   = "SEG-005"
ERROR_ARCHIVO_INESTABLE      = "SEG-006"
ERROR_DOMINIO_INVALIDO       = "SEG-007"
ERROR_PROTOCOLO_INVALIDO     = "SEG-007b"
ERROR_HASH_NO_COINCIDE       = "SEG-008"
ERROR_ARCHIVO_NO_EXISTE      = "SEG-009"
ERROR_ARCHIVO_MODIFICADO_AV  = "SEG-010"
RESULTADO_DUPLICADO          = "duplicado"

_DESCRIPCION_ERROR = {
    ERROR_SERVICIO_AV_INACTIVO:   "Bitdefender no está activo",
    ERROR_EXTENSION_NO_PERMITIDA: "Extensión de archivo no permitida",
    ERROR_MAGIC_BYTES_INVALIDO:   "Contenido del archivo no coincide con su extensión",
    ERROR_TAMANO_EXCEDIDO:        "Archivo supera el límite de tamaño configurado",
    ERROR_ARCHIVO_ELIMINADO_AV:   "Bitdefender eliminó el archivo (posible amenaza)",
    ERROR_ARCHIVO_INESTABLE:      "El archivo no se estabilizó en disco",
    ERROR_DOMINIO_INVALIDO:       "Dominio de la URL no es el esperado",
    ERROR_PROTOCOLO_INVALIDO:     "La URL no usa HTTPS",
    ERROR_HASH_NO_COINCIDE:       "Hash SHA-256 no coincide con el registrado",
    ERROR_ARCHIVO_NO_EXISTE:      "Archivo temporal no encontrado",
    ERROR_ARCHIVO_MODIFICADO_AV:  "El AV modificó el archivo",
}

# ---------------------------------------------------------------------------
#  Magic bytes: extensión → lista de (offset, firma_esperada)
# ---------------------------------------------------------------------------

_MAGIC: dict[str, list[tuple[int, bytes]]] = {
    ".pdf":  [(0, b"%PDF")],
    ".xlsx": [(0, b"PK\x03\x04")],
    ".xls":  [(0, b"\xd0\xcf\x11\xe0")],
    ".docx": [(0, b"PK\x03\x04")],
    ".doc":  [(0, b"\xd0\xcf\x11\xe0")],
    ".pptx": [(0, b"PK\x03\x04")],
    ".ppt":  [(0, b"\xd0\xcf\x11\xe0")],
    ".zip":  [(0, b"PK\x03\x04")],
    ".msg":  [(0, b"\xd0\xcf\x11\xe0")],
    ".jpg":  [(0, b"\xff\xd8\xff")],
    ".jpeg": [(0, b"\xff\xd8\xff")],
    ".png":  [(0, b"\x89PNG")],
    ".bmp":  [(0, b"BM")],
    ".webp": [(8, b"WEBP")],
    # GIF/TIFF — múltiples firmas válidas, se aceptan sin verificar bytes
    ".gif":  [],
    ".tiff": [],
    ".tif":  [],
    # Texto plano — sin firma binaria, se aceptan sin verificar bytes
    ".txt":  [],
    ".csv":  [],
    ".xml":  [],
    ".json": [],
}

# ---------------------------------------------------------------------------
#  Logger interno (print simple, sin dependencia del nucleo de NAVIERAS)
# ---------------------------------------------------------------------------

def _log(mensaje: str, nombre_proyecto: str = "seguridad") -> None:
    _logger_seg.debug(f"  [SEG:{nombre_proyecto}] {mensaje}")


# ---------------------------------------------------------------------------
#  Utilidades internas
# ---------------------------------------------------------------------------

def _servicio_activo(nombre_servicio: str) -> bool:
    try:
        r = subprocess.run(
            ["sc", "query", nombre_servicio],
            capture_output=True, timeout=10,
        )
        return "RUNNING" in r.stdout.decode(errors="replace")
    except Exception:
        return False


def _calcular_hash(ruta: str) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


def _esperar_estabilidad(
    ruta: str,
    segundos_estables: int = 2,
    timeout: int = 30,
    nombre_proyecto: str = "seguridad",
) -> bool:
    """
    Espera hasta que el tamaño del archivo no cambie durante
    `segundos_estables` consecutivos.
    """
    _log(f"Esperando estabilidad del archivo (máx. {timeout}s)...", nombre_proyecto)
    inicio         = time.time()
    tamaño_prev    = -1
    tiempo_estable = 0.0

    while True:
        if not os.path.isfile(ruta):
            return False
        tamaño_actual = os.path.getsize(ruta)
        if tamaño_actual == tamaño_prev:
            tiempo_estable += 1
            if tiempo_estable >= segundos_estables:
                _log(f"Archivo estable: {tamaño_actual:,} bytes", nombre_proyecto)
                return True
        else:
            tiempo_estable = 0
            tamaño_prev    = tamaño_actual

        if time.time() - inicio > timeout:
            return False
        time.sleep(1)


# ---------------------------------------------------------------------------
#  Verificaciones individuales (exportadas para uso externo)
# ---------------------------------------------------------------------------

_SERVICIO_AV_DEFAULT = "EPSecurityService"


def verificar_no_duplicado(
    ruta_local: str,
    ruta_registro: str,
    nombre_proyecto: str = "seguridad",
) -> tuple[bool, str | None, str | None]:
    """
    Calcula el SHA-256 del archivo y lo busca en el registro de hashes ya subidos.

    Returns:
        (es_nuevo, codigo_error, hash_calculado)
        - es_nuevo=True  → el archivo no existía antes, continuar con el pipeline
        - es_nuevo=False → duplicado detectado, retorna RESULTADO_DUPLICADO
        - hash_calculado → el hash SHA-256, para guardarlo tras subir (si es nuevo)
    """
    hash_archivo = _calcular_hash(ruta_local)

    registro = {}
    ruta_reg = Path(ruta_registro)
    if ruta_reg.exists():
        try:
            registro = json.loads(ruta_reg.read_text(encoding="utf-8"))
        except Exception:
            pass

    if hash_archivo in registro:
        ruta_previa = registro[hash_archivo]
        _log(f"Duplicado detectado — ya subido como: {ruta_previa}", nombre_proyecto)
        return False, RESULTADO_DUPLICADO, hash_archivo

    return True, None, hash_archivo


def verificar_servicio_av(nombre_servicio: str = _SERVICIO_AV_DEFAULT) -> tuple[bool, str | None]:
    """Comprueba que Bitdefender esté en estado RUNNING."""
    if _servicio_activo(nombre_servicio):
        return True, None
    return False, ERROR_SERVICIO_AV_INACTIVO


def verificar_extension(
    nombre_orig: str,
    extensiones_ok: set,
) -> tuple[bool, str | None]:
    ext = Path(nombre_orig).suffix.lower()
    if ext not in extensiones_ok:
        return False, ERROR_EXTENSION_NO_PERMITIDA
    return True, None


def verificar_magic_bytes(
    ruta_local: str,
    nombre_orig: str,
) -> tuple[bool, str | None]:
    ext = Path(nombre_orig).suffix.lower()
    firmas = _MAGIC.get(ext, [])

    if not firmas:
        return True, None  # texto plano, sin firma binaria

    try:
        with open(ruta_local, "rb") as f:
            cabecera = f.read(8)
    except Exception:
        return False, ERROR_MAGIC_BYTES_INVALIDO

    for offset, firma in firmas:
        if cabecera[offset: offset + len(firma)] == firma:
            return True, None

    return False, ERROR_MAGIC_BYTES_INVALIDO


def verificar_tamano(
    ruta_local: str,
    tamano_max_mb: float,
) -> tuple[bool, str | None]:
    try:
        tamano_mb = os.path.getsize(ruta_local) / (1024 * 1024)
    except Exception:
        return False, ERROR_TAMANO_EXCEDIDO

    if tamano_mb > tamano_max_mb:
        return False, ERROR_TAMANO_EXCEDIDO
    return True, None


def esperar_av(
    ruta_local: str,
    espera_av: int = 3,
    nombre_proyecto: str = "seguridad",
) -> tuple[bool, str | None]:
    """Espera N segundos y verifica que el archivo siga existiendo."""
    _log(f"Esperando {espera_av}s para análisis Bitdefender...", nombre_proyecto)
    time.sleep(espera_av)
    if not os.path.isfile(ruta_local):
        return False, ERROR_ARCHIVO_ELIMINADO_AV
    return True, None


def verificar_dominio_fuente(
    url_fuente: str,
    dominio_esperado: str,
    nombre_proyecto: str = "seguridad",
) -> tuple[bool, str | None]:
    parsed = urlparse(url_fuente)

    if parsed.scheme != "https":
        _log(f"Protocolo inválido: '{parsed.scheme}'. Se requiere HTTPS.", nombre_proyecto)
        return False, ERROR_PROTOCOLO_INVALIDO

    dominio_real = parsed.netloc.lower().lstrip("www.")
    dominio_esp  = dominio_esperado.lower().lstrip("www.")

    if dominio_real != dominio_esp and not dominio_real.endswith("." + dominio_esp):
        _log(f"Dominio inválido: '{parsed.netloc}'. Se esperaba: '{dominio_esperado}'.", nombre_proyecto)
        return False, ERROR_DOMINIO_INVALIDO

    _log(f"Dominio verificado: {parsed.netloc}", nombre_proyecto)
    return True, None


def verificar_hash_integridad(
    ruta_local: str,
    hash_esperado: str,
    espera_av_seg: int = 4,
    estabilidad_seg: int = 2,
    estabilidad_timeout: int = 30,
    nombre_proyecto: str = "seguridad",
) -> tuple[bool, str | None]:
    """
    Verificación completa de integridad:
      1. Existencia del archivo
      2. Estabilidad en disco
      3. SHA-256 inicial
      4. Esperar análisis AV
      5. Verificar que el AV no actuó
      6. Comparar hash con el esperado
    """
    # 1 — existencia
    if not os.path.isfile(ruta_local):
        _log(f"Archivo no encontrado: {ruta_local}", nombre_proyecto)
        return False, ERROR_ARCHIVO_NO_EXISTE

    # 2 — estabilidad
    if not _esperar_estabilidad(ruta_local, estabilidad_seg, estabilidad_timeout, nombre_proyecto):
        _log("El archivo no se estabilizó.", nombre_proyecto)
        return False, ERROR_ARCHIVO_INESTABLE

    # 3 — hash inicial
    hash_antes = _calcular_hash(ruta_local)
    _log(f"SHA-256 inicial: {hash_antes}", nombre_proyecto)

    # 4 — esperar AV
    _log(f"Esperando {espera_av_seg}s para análisis AV...", nombre_proyecto)
    time.sleep(espera_av_seg)

    # 5 — verificar que el AV no actuó
    if not os.path.isfile(ruta_local):
        _log("El archivo fue eliminado por el AV.", nombre_proyecto)
        return False, ERROR_ARCHIVO_ELIMINADO_AV

    hash_despues = _calcular_hash(ruta_local)
    if hash_despues != hash_antes:
        _log("El AV modificó el archivo.", nombre_proyecto)
        return False, ERROR_ARCHIVO_MODIFICADO_AV

    # 6 — comparar con el hash registrado
    if hash_despues != hash_esperado:
        _log(f"Hash no coincide.\n  Esperado : {hash_esperado}\n  Calculado: {hash_despues}", nombre_proyecto)
        return False, ERROR_HASH_NO_COINCIDE

    _log(f"Hash SHA-256 verificado: {hash_despues}", nombre_proyecto)
    return True, None


# ---------------------------------------------------------------------------
#  Función principal — punto de entrada para todos los proyectos
# ---------------------------------------------------------------------------

def ejecutar(
    ruta_local:       str,
    nombre_orig:      str | None  = None,
    extensiones_ok:   set | None  = None,
    tamano_max_mb:    float | None = None,
    url_fuente:       str | None  = None,
    dominio_esperado: str | None  = None,
    hash_esperado:    str | None  = None,
    nombre_proyecto:  str         = "seguridad",
    espera_av:        int         = 3,
    ruta_registro:    str | None  = None,
) -> dict:
    """
    Ejecuta las verificaciones de seguridad habilitadas según los parámetros.

    Siempre ejecuta:
        SEG-001  servicio_av
        SEG-005  esperar_av    (salvo que hash_esperado esté presente, que incluye la espera)

    Condicionales:
        SEG-002  extension        si extensiones_ok
        SEG-003  magic_bytes      si extensiones_ok y nombre_orig
        SEG-004  tamano           si tamano_max_mb
        SEG-007  dominio_fuente   si url_fuente y dominio_esperado
        SEG-008  hash_integridad  si hash_esperado (incluye SEG-006 estabilidad y espera AV interna)
    """
    verificaciones: list[tuple[str, object]] = []
    _hash_calculado: list[str] = []  # contenedor mutable para capturar el hash del duplicado check

    # SEG-DUP — verificar duplicado antes de todo (evita el sleep del AV si ya existe)
    if ruta_registro:
        def _check_dup():
            es_nuevo, cod, h = verificar_no_duplicado(ruta_local, ruta_registro, nombre_proyecto)
            if h:
                _hash_calculado.append(h)
            return es_nuevo, cod
        verificaciones.append(("no_duplicado", _check_dup))

    # SEG-001 — siempre
    verificaciones.append(
        ("servicio_av", lambda: verificar_servicio_av())
    )

    # SEG-002 y SEG-003 — solo si hay lista blanca
    if extensiones_ok is not None and nombre_orig:
        verificaciones.append(
            ("extension",   lambda: verificar_extension(nombre_orig, extensiones_ok))
        )
        verificaciones.append(
            ("magic_bytes", lambda: verificar_magic_bytes(ruta_local, nombre_orig))
        )

    # SEG-004 — solo si hay límite de tamaño
    if tamano_max_mb is not None:
        verificaciones.append(
            ("tamano", lambda: verificar_tamano(ruta_local, tamano_max_mb))
        )

    # SEG-007 — solo si hay URL y dominio esperado
    if url_fuente and dominio_esperado:
        verificaciones.append(
            ("dominio_fuente", lambda: verificar_dominio_fuente(url_fuente, dominio_esperado, nombre_proyecto))
        )

    # SEG-008 — hash (incluye estabilidad + espera AV interna)
    if hash_esperado:
        verificaciones.append(
            ("hash_integridad", lambda: verificar_hash_integridad(
                ruta_local, hash_esperado,
                espera_av_seg=espera_av,
                nombre_proyecto=nombre_proyecto,
            ))
        )
    else:
        # SEG-005 — esperar AV simple (solo cuando no hay verificación de hash)
        verificaciones.append(
            ("esperar_av", lambda: esperar_av(ruta_local, espera_av, nombre_proyecto))
        )

    # Ejecutar todas las verificaciones en orden
    for nombre_check, fn in verificaciones:
        aprobado, codigo_error = fn()
        if not aprobado:
            if codigo_error == RESULTADO_DUPLICADO:
                return {
                    "resultado":    "duplicado",
                    "codigo_error": None,
                    "descripcion":  "Archivo idéntico ya subido anteriormente",
                    "hash":         _hash_calculado[0] if _hash_calculado else None,
                }
            descripcion = _DESCRIPCION_ERROR.get(codigo_error, "")
            _log(f"RECHAZADO [{codigo_error}] {nombre_check}: {descripcion}", nombre_proyecto)
            return {
                "resultado":    "rechazado",
                "codigo_error": codigo_error,
                "descripcion":  descripcion,
                "hash":         None,
            }
        _log(f"OK {nombre_check}", nombre_proyecto)

    _log("APROBADO", nombre_proyecto)
    return {
        "resultado":    "aprobado",
        "codigo_error": None,
        "descripcion":  None,
        "hash":         _hash_calculado[0] if _hash_calculado else None,
    }


# ---------------------------------------------------------------------------
#  Script independiente para pruebas
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso:  python agente_seguridad.py <ruta_local> [nombre_original]")
        print("Ej:   python agente_seguridad.py \"C:/Temp/adj_abc.pdf\" \"BL PO 196893.pdf\"")
        sys.exit(1)

    ruta  = sys.argv[1]
    nomre = sys.argv[2] if len(sys.argv) > 2 else Path(ruta).name

    print(f"\nVerificando servicio AV...")
    ok_av, err_av = verificar_servicio_av()
    if not ok_av:
        _logger_seg.debug(f"  [SEG] ADVERTENCIA — {_DESCRIPCION_ERROR[err_av]}")
    else:
        _logger_seg.debug("  [SEG] OK servicio_av")

    res = ejecutar(ruta_local=ruta, nombre_orig=nomre)
    print("\n--- Resultado ---")
    for k, v in res.items():
        print(f"  {k}: {v}")
