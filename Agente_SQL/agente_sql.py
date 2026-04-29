"""
agente_sql.py  —  Agente SQL compartido
========================================
Módulo único usado por todos los proyectos del workspace.
No tiene dependencias de NAVIERAS ni de ningún otro proyecto.

Operaciones que realiza:
    1. Conectar a SQL Server vía pyodbc (Windows Auth o SQL Auth)
    2. MERGE/UPSERT de tabla_temporal.xlsx hacia la tabla destino
    3. Limpiar tabla_temporal.xlsx si todo fue exitoso
    4. Rollback automático si algo falla a mitad

Parámetros de ejecutar():
    ruta_tabla_temporal  str   — ruta absoluta al Excel temporal            (siempre)
    tabla_destino        str   — nombre de la tabla en SQL Server           (siempre)
    clave_primaria       str   — columna PK para el MERGE (default "container_number")
    nombre_proyecto      str   — nombre para el logger (default "sql")
    servidor             str   — host SQL Server
    puerto               int   — puerto (default 1433)
    base_datos           str   — nombre de la base de datos
    odbc_driver          str   — nombre del driver ODBC
    auth                 str   — "windows" o "sql"
    usuario              str   — solo para auth="sql"
    password             str   — solo para auth="sql"
    mapa_track           dict  — mapeo columnas recopilador → TRACK_* para PURCHCONTAINERS
                                 (si tabla_destino.upper() == "PURCHCONTAINERS")

Retorna:
    {
        "resultado_carga"    : "exitoso" | "fallido",
        "filas_insertadas"   : int,
        "filas_actualizadas" : int,
        "filas_sin_cambios"  : int,
        "detalles_cambios"   : dict,
        "contenedores_nuevos": list,
    }

Uso — NAVIERAS:
    from Agente_SQL.agente_sql import ejecutar
    res = ejecutar(
        ruta_tabla_temporal = ruta,
        tabla_destino       = config.tabla_destino_sql,
        clave_primaria      = config.clave_primaria_sql,
        nombre_proyecto     = nombre_proyecto,
        servidor            = DB_SERVIDOR,
        puerto              = DB_PUERTO,
        base_datos          = DB_NOMBRE,
        odbc_driver         = DB_ODBC_DRIVER,
        auth                = DB_AUTH,
        mapa_track          = MAPA_TRACK.get(nombre_proyecto, {}),
    )
"""

import sys
import time
from pathlib import Path
from datetime import datetime as _dt, datetime, date


# ---------------------------------------------------------------------------
#  Logger interno (print simple, sin dependencia del nucleo de ningún proyecto)
# ---------------------------------------------------------------------------

def _log(mensaje: str, nombre_proyecto: str = "sql") -> None:
    print(f"  [SQL:{nombre_proyecto}] {mensaje}")


# ---------------------------------------------------------------------------
#  Conexión
# ---------------------------------------------------------------------------

def _construir_connection_string(
    servidor: str,
    puerto: int,
    base_datos: str,
    odbc_driver: str,
    auth: str,
    usuario: str = "",
    password: str = "",
) -> str:
    if auth == "windows":
        return (
            f"DRIVER={{{odbc_driver}}};"
            f"SERVER={servidor},{puerto};"
            f"DATABASE={base_datos};"
            "Trusted_Connection=yes;"
            "TrustServerCertificate=yes;"
        )
    return (
        f"DRIVER={{{odbc_driver}}};"
        f"SERVER={servidor},{puerto};"
        f"DATABASE={base_datos};"
        f"UID={usuario};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )


def _conectar(conn_str: str, nombre_proyecto: str):
    try:
        import pyodbc
    except ImportError:
        raise RuntimeError("pyodbc no está instalado. Ejecuta: pip install pyodbc")

    esperas = [1, 2, 4]
    ultimo_error = None
    for intento, espera in enumerate(esperas, start=1):
        try:
            conn = pyodbc.connect(conn_str, timeout=15)
            _log("Conexión a SQL Server establecida.", nombre_proyecto)
            return conn
        except Exception as e:
            ultimo_error = e
            if intento < len(esperas):
                _log(f"Intento {intento} fallido: {e}. Reintentando en {espera}s...", nombre_proyecto)
                time.sleep(espera)

    raise RuntimeError(f"No se pudo conectar a SQL Server tras {len(esperas)} intentos: {ultimo_error}")


# ---------------------------------------------------------------------------
#  Columnas datetime de la tabla destino
# ---------------------------------------------------------------------------

def _columnas_datetime(conn, tabla_destino: str) -> set:
    cursor = conn.cursor()
    cursor.execute(
        "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = ? AND DATA_TYPE IN ('datetime', 'datetime2', 'date', 'smalldatetime')",
        tabla_destino,
    )
    return {row[0].strip() for row in cursor.fetchall()}


def _parsear_fecha(valor: str):
    if not valor or not str(valor).strip():
        return None
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return _dt.strptime(str(valor).strip(), fmt)
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
#  Lectura del Excel temporal
# ---------------------------------------------------------------------------

def _leer_excel(ruta: str, cols_fecha: set = None) -> tuple:
    import openpyxl
    wb   = openpyxl.load_workbook(ruta, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))

    if not rows:
        return [], []

    columnas  = [str(c).strip() if c is not None else "" for c in rows[0]]
    cols_fecha = cols_fecha or set()
    filas     = []
    for row in rows[1:]:
        fila = []
        for col_nombre, v in zip(columnas, row):
            if v in ("", "—"):
                v = None
            elif col_nombre in cols_fecha and isinstance(v, str):
                v = _parsear_fecha(v)
            fila.append(v)
        filas.append(tuple(fila))

    wb.close()
    return columnas, filas


# ---------------------------------------------------------------------------
#  Normalización para comparaciones
# ---------------------------------------------------------------------------

_CAMPOS_IGNORAR_CAMBIOS = {"fuente", "fecha_consulta"}
_TRACK_IGNORAR_CAMBIOS  = {"TRACK_FUENTE", "TRACK_FECHA_CONSULTA"}

# Mapeo campo recopilador → columna SIT_* en PURCHTABLE, por proyecto
_MAPA_PURCHTABLE = {
    "maersk":  {"arrival_date": "SIT_ARRIVALDATEAJUSTED", "eta_final": "SIT_AJUSTDEPARTURE"},
    "cma_cgm": {"discharge_date": "SIT_ARRIVALDATEAJUSTED", "vessel_eta": "SIT_AJUSTDEPARTURE"},
    "msc":     {"arrival_date": "SIT_ARRIVALDATEAJUSTED"},
}


def _normalizar(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, (datetime, date)):
        return str(valor).strip()
    return str(valor).strip()


# ---------------------------------------------------------------------------
#  MERGE / UPSERT genérico
# ---------------------------------------------------------------------------

def _ejecutar_merge(
    conn,
    columnas: list,
    filas: list,
    tabla_destino: str,
    clave_primaria: str,
    nombre_proyecto: str,
) -> tuple:
    cursor = conn.cursor()

    cols_sin_pk     = [c for c in columnas if c != clave_primaria]
    set_clause      = ", ".join(f"[{c}] = ?" for c in cols_sin_pk)
    cols_str        = ", ".join(f"[{c}]" for c in columnas)
    cols_sin_pk_str = ", ".join(f"[{c}]" for c in cols_sin_pk)
    placeholders    = ", ".join(["?"] * len(columnas))
    pk_idx          = columnas.index(clave_primaria)

    filas_insertadas   = 0
    filas_actualizadas = 0
    filas_sin_cambios  = 0
    detalles_cambios   = {}
    contenedores_nuevos = []

    for fila in filas:
        pk_valor   = fila[pk_idx]
        otros_vals = tuple(v for i, v in enumerate(fila) if i != pk_idx)

        cursor.execute(
            f"SELECT COUNT(1) FROM [{tabla_destino}] WHERE [{clave_primaria}] = ?",
            pk_valor,
        )
        existe = cursor.fetchone()[0] > 0

        if existe:
            cursor.execute(
                f"SELECT {cols_sin_pk_str} FROM [{tabla_destino}] WHERE [{clave_primaria}] = ?",
                pk_valor,
            )
            registro_actual = dict(zip(cols_sin_pk, cursor.fetchone()))

            cambios = {}
            for col, val_nuevo in zip(cols_sin_pk, otros_vals):
                if col in _CAMPOS_IGNORAR_CAMBIOS:
                    continue
                val_actual = registro_actual.get(col)
                if _normalizar(val_actual) != _normalizar(val_nuevo):
                    cambios[col] = {"antes": val_actual, "despues": val_nuevo}

            if cambios:
                cursor.execute(
                    f"UPDATE [{tabla_destino}] SET {set_clause} WHERE [{clave_primaria}] = ?",
                    (*otros_vals, pk_valor),
                )
                filas_actualizadas += 1
                detalles_cambios[str(pk_valor)] = cambios
            else:
                filas_sin_cambios += 1
        else:
            cursor.execute(
                f"INSERT INTO [{tabla_destino}] ({cols_str}) VALUES ({placeholders})",
                fila,
            )
            filas_insertadas += 1
            contenedores_nuevos.append(str(pk_valor))

    _log(
        f"MERGE completado — Insertadas: {filas_insertadas} | "
        f"Actualizadas: {filas_actualizadas} | Sin cambios: {filas_sin_cambios}",
        nombre_proyecto,
    )
    return filas_insertadas, filas_actualizadas, filas_sin_cambios, detalles_cambios, contenedores_nuevos


# ---------------------------------------------------------------------------
#  MERGE especial para PURCHCONTAINERS
# ---------------------------------------------------------------------------

def _ejecutar_merge_purchcontainers(
    conn,
    columnas: list,
    filas: list,
    nombre_proyecto: str,
    mapa_track: dict,
) -> tuple:
    if not mapa_track:
        _log(f"No hay mapa TRACK_ para el proyecto '{nombre_proyecto}'. Nada que actualizar.", nombre_proyecto)
        return 0, 0, 0, {}, []

    cols_fecha = _columnas_datetime(conn, "PURCHCONTAINERS")
    cursor = conn.cursor()

    try:
        idx_cont = columnas.index("container_number")
    except ValueError:
        _log("El Excel temporal no tiene columna 'container_number'.", nombre_proyecto)
        return 0, 0, 0, {}, []

    track_cols = list(mapa_track.values()) + ["TRACK_FUENTE", "TRACK_FECHA_CONSULTA"]
    set_clause = ", ".join(f"[{c}] = ?" for c in track_cols)

    filas_actualizadas = 0
    filas_sin_cambios  = 0
    detalles_cambios   = {}

    for fila in filas:
        container_number = str(fila[idx_cont]).strip().upper() if fila[idx_cont] else None
        if not container_number:
            continue

        cursor.execute(
            "SELECT COUNT(1) FROM [PURCHCONTAINERS] "
            "WHERE REPLACE([CONTAINERID], ' ', '') = ?",
            container_number,
        )
        if cursor.fetchone()[0] == 0:
            _log(f"CONTAINERID no encontrado en PURCHCONTAINERS: {container_number} — se omite.", nombre_proyecto)
            continue

        select_cols = ", ".join(f"[{c}]" for c in track_cols)
        cursor.execute(
            f"SELECT {select_cols} FROM [PURCHCONTAINERS] "
            f"WHERE REPLACE([CONTAINERID], ' ', '') = ?",
            container_number,
        )
        fila_actual = cursor.fetchone()
        registro_actual = dict(zip(track_cols, fila_actual)) if fila_actual else {}

        nuevos_track = {}
        for col_recop, col_track in mapa_track.items():
            if col_recop in columnas:
                val = fila[columnas.index(col_recop)]
                if col_track in cols_fecha and isinstance(val, str):
                    val = _parsear_fecha(val)
                nuevos_track[col_track] = val
            else:
                nuevos_track[col_track] = None

        nuevos_track["TRACK_FUENTE"]         = nombre_proyecto
        nuevos_track["TRACK_FECHA_CONSULTA"] = datetime.now()

        cambios = {}
        for col_track, val_nuevo in nuevos_track.items():
            if col_track in _TRACK_IGNORAR_CAMBIOS:
                continue
            val_actual = registro_actual.get(col_track)
            if _normalizar(val_actual) != _normalizar(val_nuevo):
                cambios[col_track] = {"antes": val_actual, "despues": val_nuevo}

        vals_update = [nuevos_track[c] for c in track_cols]
        cursor.execute(
            f"UPDATE [PURCHCONTAINERS] SET {set_clause} "
            f"WHERE REPLACE([CONTAINERID], ' ', '') = ?",
            (*vals_update, container_number),
        )

        if cambios:
            filas_actualizadas += 1
            detalles_cambios[container_number] = cambios
        else:
            filas_sin_cambios += 1

    _log(
        f"MERGE PURCHCONTAINERS — Actualizadas: {filas_actualizadas} | Sin cambios: {filas_sin_cambios}",
        nombre_proyecto,
    )
    return 0, filas_actualizadas, filas_sin_cambios, detalles_cambios, []


def _ejecutar_merge_purchtable(
    conn,
    columnas: list,
    filas: list,
    nombre_proyecto: str,
) -> tuple:
    """
    Actualiza PURCHTABLE via stored procedure usp_ActualizarShipStatusPurchTable.

    Firma del SP:
        EXEC dbo.usp_ActualizarShipStatusPurchTable
            @purchid        VARCHAR  — número de OC
            @arrival_date   DATE     — fecha de arribo
            @departure_date DATE     — fecha de salida/ETA
            @fuente         VARCHAR  — código de naviera (nombre_proyecto)
    """
    mapa = _MAPA_PURCHTABLE.get(nombre_proyecto, {})
    if not mapa:
        _log(f"No hay mapa PURCHTABLE para '{nombre_proyecto}'.", nombre_proyecto)
        return 0, 0, 0, {}, []

    cursor = conn.cursor()

    try:
        idx_purchid = columnas.index("purchid")
    except ValueError:
        _log("El Excel temporal no tiene columna 'purchid'.", nombre_proyecto)
        return 0, 0, 0, {}, []

    # Determinar qué columnas del recopilador corresponden a arrival y departure
    col_arrival   = None
    col_departure = None
    for col_recop, col_sit in mapa.items():
        sit_upper = col_sit.upper()
        if "ARRIVALDATEAJUSTED" in sit_upper:
            col_arrival = col_recop
        elif "AJUSTDEPARTURE" in sit_upper:
            col_departure = col_recop

    filas_actualizadas = 0
    filas_sin_cambios  = 0
    detalles_cambios   = {}

    for fila in filas:
        purchid = str(fila[idx_purchid]).strip() if fila[idx_purchid] else None
        if not purchid:
            continue

        arrival_val   = None
        departure_val = None

        if col_arrival and col_arrival in columnas:
            val = fila[columnas.index(col_arrival)]
            if val not in (None, "", "—"):
                arrival_val = _parsear_fecha(str(val)) if isinstance(val, str) else val

        if col_departure and col_departure in columnas:
            val = fila[columnas.index(col_departure)]
            if val not in (None, "", "—"):
                departure_val = _parsear_fecha(str(val)) if isinstance(val, str) else val

        _log(f"SP — purchid={repr(purchid)} arrival={repr(arrival_val)} departure={repr(departure_val)}", nombre_proyecto)

        if arrival_val is None and departure_val is None:
            _log(f"PURCHID '{purchid}' — sin fechas para actualizar, se omite.", nombre_proyecto)
            filas_sin_cambios += 1
            continue

        cursor.execute(
            "EXEC dbo.usp_ActualizarShipStatusPurchTable ?, ?, ?, ?",
            purchid,
            arrival_val,
            departure_val,
            "CMER",
        )

        filas_actualizadas += 1
        detalles_cambios[purchid] = {
            "SIT_ARRIVALDATEAJUSTED": {"antes": None, "despues": arrival_val},
            "SIT_AJUSTDEPARTURE":     {"antes": None, "despues": departure_val},
        }

    _log(
        f"SP PURCHTABLE — Ejecutadas: {filas_actualizadas} | Sin fechas: {filas_sin_cambios}",
        nombre_proyecto,
    )
    return 0, filas_actualizadas, filas_sin_cambios, detalles_cambios, []


# ---------------------------------------------------------------------------
#  Función principal — punto de entrada para todos los proyectos
# ---------------------------------------------------------------------------

def ejecutar(
    ruta_tabla_temporal: str,
    tabla_destino: str,
    clave_primaria: str = "container_number",
    nombre_proyecto: str = "sql",
    servidor: str = "localhost",
    puerto: int = 1433,
    base_datos: str = "",
    odbc_driver: str = "ODBC Driver 18 for SQL Server",
    auth: str = "windows",
    usuario: str = "",
    password: str = "",
    mapa_track: dict = None,
) -> dict:
    """
    Carga los datos del Excel temporal en SQL Server.

    Retorna:
        {
            "resultado_carga"    : "exitoso" | "fallido",
            "filas_insertadas"   : int,
            "filas_actualizadas" : int,
            "filas_sin_cambios"  : int,
            "detalles_cambios"   : dict,
            "contenedores_nuevos": list,
        }
    """
    _log("=== Agente SQL iniciado ===", nombre_proyecto)

    if not Path(ruta_tabla_temporal).exists():
        _log("El Excel temporal no existe.", nombre_proyecto)
        return {"resultado_carga": "fallido", "filas_insertadas": 0, "filas_actualizadas": 0,
                "filas_sin_cambios": 0, "detalles_cambios": {}, "contenedores_nuevos": []}

    conn_str = _construir_connection_string(servidor, puerto, base_datos, odbc_driver, auth, usuario, password)
    conn = None
    try:
        conn = _conectar(conn_str, nombre_proyecto)
        cols_fecha = _columnas_datetime(conn, tabla_destino)
        _log(f"Columnas datetime detectadas: {cols_fecha}", nombre_proyecto)

        columnas, filas = _leer_excel(ruta_tabla_temporal, cols_fecha)
        _log(f"Registros leídos del Excel temporal: {len(filas)}", nombre_proyecto)

        if not filas:
            _log("El Excel temporal está vacío. No hay datos para cargar.", nombre_proyecto)
            return {"resultado_carga": "fallido", "filas_insertadas": 0, "filas_actualizadas": 0,
                    "filas_sin_cambios": 0, "detalles_cambios": {}, "contenedores_nuevos": []}

        if tabla_destino.upper() == "PURCHCONTAINERS":
            filas_insertadas, filas_actualizadas, filas_sin_cambios, detalles_cambios, contenedores_nuevos = \
                _ejecutar_merge_purchcontainers(conn, columnas, filas, nombre_proyecto, mapa_track or {})
        elif tabla_destino.upper() == "PURCHTABLE":
            filas_insertadas, filas_actualizadas, filas_sin_cambios, detalles_cambios, contenedores_nuevos = \
                _ejecutar_merge_purchtable(conn, columnas, filas, nombre_proyecto)
        else:
            filas_insertadas, filas_actualizadas, filas_sin_cambios, detalles_cambios, contenedores_nuevos = \
                _ejecutar_merge(conn, columnas, filas, tabla_destino, clave_primaria, nombre_proyecto)

        conn.commit()
        _log("Commit realizado correctamente.", nombre_proyecto)

        try:
            Path(ruta_tabla_temporal).unlink()
            _log(f"Archivo temporal eliminado: {ruta_tabla_temporal}", nombre_proyecto)
        except Exception as e:
            _log(f"No se pudo eliminar el archivo temporal: {e}", nombre_proyecto)

        _log("=== Agente SQL: EXITOSO ===", nombre_proyecto)
        return {
            "resultado_carga":      "exitoso",
            "filas_insertadas":     filas_insertadas,
            "filas_actualizadas":   filas_actualizadas,
            "filas_sin_cambios":    filas_sin_cambios,
            "detalles_cambios":     detalles_cambios,
            "contenedores_nuevos":  contenedores_nuevos,
        }

    except Exception as e:
        _log(f"Error durante la carga SQL: {e}", nombre_proyecto)
        if conn:
            try:
                conn.rollback()
                _log("Rollback ejecutado.", nombre_proyecto)
            except Exception:
                pass
        return {"resultado_carga": "fallido", "filas_insertadas": 0, "filas_actualizadas": 0,
                "filas_sin_cambios": 0, "detalles_cambios": {}, "contenedores_nuevos": []}

    finally:
        if conn:
            conn.close()


# ---------------------------------------------------------------------------
#  Script independiente para pruebas
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Uso: python agente_sql.py <ruta_excel> <tabla_destino> <servidor> <base_datos> [clave_primaria]")
        print("Ej:  python agente_sql.py C:/tmp/tabla.xlsx contenedores_pil 172.26.1.56 GIAX2012_Prod")
        sys.exit(1)

    resultado = ejecutar(
        ruta_tabla_temporal = sys.argv[1],
        tabla_destino       = sys.argv[2],
        servidor            = sys.argv[3],
        base_datos          = sys.argv[4],
        clave_primaria      = sys.argv[5] if len(sys.argv) > 5 else "container_number",
        nombre_proyecto     = "prueba_sql",
        auth                = "windows",
        odbc_driver         = "ODBC Driver 18 for SQL Server",
        puerto              = 1433,
    )
    print("\n--- Resultado del agente SQL ---")
    for k, v in resultado.items():
        print(f"  {k}: {v}")
