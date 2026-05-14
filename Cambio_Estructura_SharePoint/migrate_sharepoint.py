# -*- coding: utf-8 -*-
"""
migrate_sharepoint.py — Migración de estructura de carpetas por OC en SharePoint.

Modos:
  --mode dry-run   Detecta escenario y reporta acciones sin modificar nada.
  --mode backup    Genera snapshot del estado actual (requiere dry-run previo).
  --mode execute   Ejecuta la migración (requiere backup reciente < 30 min).

Argumentos:
  --oc    Nombre de una OC específica (opcional). Si se omite, procesa todas las del listado.

OCs configuradas:
  Prueba_1.1, Prueba_2.1, Prueba_3.1, Prueba_4.1
"""

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ---------------------------------------------------------------------------
# sys.path para importar desde el proyecto
# ---------------------------------------------------------------------------
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Clasificador de documentos", "clasificador_documentos"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Clasificador%20de%20documentos", "clasificador_documentos"))

from configuracion.ajustes import SHAREPOINT_DRIVE_ID, SHAREPOINT_CARPETA_OCS
from Integraciones.graph_client import GraphClient

_GRAPH_BASE   = "https://graph.microsoft.com/v1.0"
SCRIPT_DIR    = Path(__file__).parent
CONFIG_FILE   = SCRIPT_DIR / "migration-config.json"
DRY_RUN_FILE  = SCRIPT_DIR / "migration_sp_dry_run.json"
BACKUP_DIR    = SCRIPT_DIR / "backups"
BACKUP_DIR.mkdir(exist_ok=True)
LOGS_DIR      = SCRIPT_DIR / "logs"
LOGS_DIR.mkdir(exist_ok=True)
REGISTRO_FILE = SCRIPT_DIR / "ocs_migradas.json"
BACKUP_MAX_AGE_MINUTES = 30

# OCs de prueba a procesar
OCS_PRUEBA = ["Prueba_1", "Prueba_2", "cmer-OC-00194453-IITC000", "cmer-OC-00191063-TCHC001", "cmer-OC-00195231-PAINL000", "cmer-OC-00196719-PAINL000"]


# ---------------------------------------------------------------------------
# Carga de configuracion
# ---------------------------------------------------------------------------

def load_config() -> dict:
    if not CONFIG_FILE.exists():
        sys.exit(f"[ERROR] No se encontro {CONFIG_FILE}")
    with CONFIG_FILE.open(encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Helpers Graph API
# ---------------------------------------------------------------------------

def _listar_carpetas(gc: GraphClient, item_id: str) -> list[dict]:
    """Lista solo carpetas (no archivos) hijas directas de un item."""
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/children?$select=id,name,folder&$top=999"
    items = []
    while url:
        data = gc._get(url)
        items.extend(h for h in data.get("value", []) if "folder" in h)
        url = data.get("@odata.nextLink")
    return items


def _listar_todos(gc: GraphClient, item_id: str) -> list[dict]:
    """Lista todos los hijos (carpetas y archivos) de un item."""
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}/children?$select=id,name,folder,file,size&$top=999"
    items = []
    while url:
        data = gc._get(url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _obtener_item(gc: GraphClient, ruta: str) -> dict | None:
    encoded = quote(ruta, safe="/")
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/root:/{encoded}:"
    gc._renovar_token_si_necesario()
    r = gc.session.get(url, headers=gc._headers, timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def _crear_carpeta(gc: GraphClient, parent_id: str, nombre: str) -> dict:
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{parent_id}/children"
    return gc._post(url, {"name": nombre, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"})


def _tiene_archivos_recursivo(gc: GraphClient, item_id: str) -> bool:
    """Retorna True si la carpeta (o alguna subcarpeta) contiene al menos un archivo."""
    hijos = _listar_todos(gc, item_id)
    for h in hijos:
        if "file" in h:
            return True
        if "folder" in h:
            if _tiene_archivos_recursivo(gc, h["id"]):
                return True
    return False


def _eliminar_carpeta(gc: GraphClient, item_id: str) -> None:
    url = f"{_GRAPH_BASE}/drives/{SHAREPOINT_DRIVE_ID}/items/{item_id}"
    gc._renovar_token_si_necesario()
    r = gc.session.delete(url, headers=gc._headers, timeout=30)
    r.raise_for_status()


# ---------------------------------------------------------------------------
# Deteccion de escenario
# ---------------------------------------------------------------------------

def detect_scenario(carpetas_existentes: set[str], config: dict) -> int:
    """
    1 — Solo estructura vieja/desconocida
    2 — Ambas estructuras (nueva + vieja/desconocida)
    3 — Solo estructura nueva (o vacia)
    """
    new_set      = set(config["new_structure"]["folders"])
    preserve_set = set(config["old_structure"]["folders_to_preserve"])
    safe_set     = new_set | preserve_set

    has_unknown = bool(carpetas_existentes - safe_set)
    has_new     = bool(carpetas_existentes & new_set)

    if has_unknown and has_new:
        return 2
    if has_unknown and not has_new:
        return 1
    return 3


def folders_to_delete(carpetas_existentes: set[str], config: dict) -> tuple[list[str], list[str]]:
    new_set      = set(config["new_structure"]["folders"])
    preserve_set = set(config["old_structure"]["folders_to_preserve"])
    safe_set     = new_set | preserve_set

    to_delete = sorted(carpetas_existentes - safe_set)
    to_keep   = sorted(carpetas_existentes & safe_set)
    return to_delete, to_keep


# ---------------------------------------------------------------------------
# Modo dry-run
# ---------------------------------------------------------------------------

def _preview_new_structure(gc: GraphClient, oc_id: str, config: dict) -> tuple[list[str], list[str]]:
    """
    Compara la estructura nueva esperada contra lo que existe en SharePoint.
    Retorna (to_create, subs_to_delete): subcarpetas a crear y subcarpetas desconocidas a borrar.
    """
    folders    = config["new_structure"]["folders"]
    subfolders = config["new_structure"]["subfolders"]

    carpetas_raiz = {c["name"]: c["id"] for c in _listar_carpetas(gc, oc_id)}
    to_create     = []
    subs_to_delete = []

    for folder in folders:
        if folder not in carpetas_raiz:
            to_create.append(folder)
            for sub in subfolders.get(folder, []):
                to_create.append(f"{folder}/{sub}")
        else:
            expected      = set(subfolders.get(folder, []))
            hijos         = _listar_carpetas(gc, carpetas_raiz[folder])
            sub_existentes = {c["name"] for c in hijos}
            for sub in subfolders.get(folder, []):
                if sub not in sub_existentes:
                    to_create.append(f"{folder}/{sub}")
            for sub in sub_existentes:
                if sub not in expected:
                    subs_to_delete.append(f"{folder}/{sub}")

    return to_create, subs_to_delete


def run_dry_run(gc: GraphClient, oc_filter: str | None, config: dict) -> None:
    ocs    = _resolver_ocs(oc_filter)
    report = {"generated_at": _now_iso(), "ocs": {}}

    for nombre_oc in ocs:
        ruta_oc  = f"{SHAREPOINT_CARPETA_OCS}/{nombre_oc}"
        oc_item  = _obtener_item(gc, ruta_oc)
        if not oc_item:
            print(f"\n[ERROR] No se encontro la carpeta: {ruta_oc}")
            continue

        oc_id         = oc_item["id"]
        carpetas_sp   = _listar_carpetas(gc, oc_id)
        nombres       = {c["name"] for c in carpetas_sp}
        scenario      = detect_scenario(nombres, config)
        to_del, to_keep = folders_to_delete(nombres, config)

        if scenario == 3:
            to_del  = []
            to_keep = sorted(nombres)

        to_create, subs_to_delete = _preview_new_structure(gc, oc_id, config)

        entry = {
            "scenario":         scenario,
            "action":           config["migration"]["scenarios"][str(scenario)],
            "to_delete":        to_del,
            "subs_to_delete":   subs_to_delete,
            "to_keep":          to_keep,
            "to_create":        to_create,
        }
        report["ocs"][nombre_oc] = entry

        print(f"\n{'-'*60}")
        print(f"  OC: {nombre_oc}  ->  Escenario {scenario}")
        print(f"  {entry['action']}")
        all_to_delete = to_del + subs_to_delete
        if all_to_delete:
            print(f"  BORRARIA ({len(all_to_delete)}):")
            for f in all_to_delete:
                print(f"    - {f}")
        else:
            print("  BORRARIA: (ninguna)")
        print(f"  CONSERVARIA ({len(to_keep)}): {', '.join(to_keep) or '(ninguna)'}")
        if to_create:
            print(f"  CREARIA ({len(to_create)}):")
            for f in to_create:
                print(f"    + {f}")
        else:
            print("  CREARIA: (ninguna — estructura nueva completa)")

    DRY_RUN_FILE.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n[OK] Reporte guardado en {DRY_RUN_FILE}")


# ---------------------------------------------------------------------------
# Modo backup
# ---------------------------------------------------------------------------

def run_backup(gc: GraphClient, oc_filter: str | None, config: dict) -> None:
    if not DRY_RUN_FILE.exists():
        sys.exit(f"[ERROR] Ejecuta --mode dry-run primero.")

    ocs      = _resolver_ocs(oc_filter)
    ts       = _timestamp()
    snapshot = {"generated_at": _now_iso(), "ocs": {}}

    for nombre_oc in ocs:
        ruta_oc  = f"{SHAREPOINT_CARPETA_OCS}/{nombre_oc}"
        oc_item  = _obtener_item(gc, ruta_oc)
        if not oc_item:
            print(f"[WARN] No se encontro: {ruta_oc}")
            continue

        oc_entry = {}
        carpetas = _listar_carpetas(gc, oc_item["id"])
        for carpeta in sorted(carpetas, key=lambda c: c["name"]):
            archivos = [h["name"] for h in _listar_todos(gc, carpeta["id"]) if "file" in h]
            oc_entry[carpeta["name"]] = sorted(archivos)
        snapshot["ocs"][nombre_oc] = oc_entry

    backup_file = BACKUP_DIR / f"migration_sp_backup_{ts}.json"
    backup_file.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"[OK] Backup guardado en {backup_file}")
    return backup_file


# ---------------------------------------------------------------------------
# Modo execute
# ---------------------------------------------------------------------------

def run_execute(gc: GraphClient, oc_filter: str | None, config: dict, backup_file: Path | None = None) -> None:

    ocs = _resolver_ocs(oc_filter)
    ts  = _timestamp()
    log = {"generated_at": _now_iso(), "ocs": {}}

    for nombre_oc in ocs:
        ruta_oc  = f"{SHAREPOINT_CARPETA_OCS}/{nombre_oc}"
        oc_item  = _obtener_item(gc, ruta_oc)
        if not oc_item:
            print(f"\n[ERROR] No se encontro: {ruta_oc}")
            continue

        oc_id         = oc_item["id"]
        carpetas_sp   = _listar_carpetas(gc, oc_id)
        nombres       = {c["name"] for c in carpetas_sp}
        id_map        = {c["name"]: c["id"] for c in carpetas_sp}

        scenario           = detect_scenario(nombres, config)
        to_del, to_keep    = folders_to_delete(nombres, config)
        oc_log             = {"scenario": scenario, "operations": []}

        print(f"\n{'='*60}")
        print(f"  OC: {nombre_oc}  ->  Escenario {scenario}")

        # Escenario 3: solo crear subcarpetas faltantes
        if scenario == 3:
            created = _ensure_new_structure(gc, oc_id, id_map, config, oc_log)
            print(f"  Sin borrados, {created} carpetas creadas")
            log["ocs"][nombre_oc] = oc_log
            _actualizar_registro(nombre_oc, oc_log, LOGS_DIR / f"migration_sp_execution_log_{ts}.json", backup_file)
            continue

        # Escenarios 1 y 2: borrar carpetas viejas vacias
        # Pre-verificar cuales tienen archivos ANTES de borrar cualquiera,
        # para evitar que borrar subcarpetas primero vacíe al padre.
        con_archivos = set()
        for folder_name in to_del:
            folder_id = id_map.get(folder_name)
            if folder_id and _tiene_archivos_recursivo(gc, folder_id):
                con_archivos.add(folder_name)

        for folder_name in to_del:
            folder_id = id_map.get(folder_name)
            op = {"folder": folder_name, "action": "delete"}

            if not folder_id:
                op.update(result="skip", detail="Ya no existe")
            elif folder_name in con_archivos:
                op.update(result="warning", detail="Carpeta con archivos — NO borrada")
                print(f"  [WARN] {folder_name}: tiene archivos, omitida")
            else:
                try:
                    _eliminar_carpeta(gc, folder_id)
                    op.update(result="ok", detail="Borrada")
                    print(f"  [DEL]  {folder_name}")
                    time.sleep(0.2)
                except Exception as e:
                    op.update(result="error", detail=str(e))
                    print(f"  [ERR]  {folder_name}: {e}")

            oc_log["operations"].append(op)

        for folder_name in to_keep:
            oc_log["operations"].append(
                {"folder": folder_name, "action": "keep", "result": "ok"}
            )

        # Refrescar mapa de carpetas despues de borrar
        carpetas_sp = _listar_carpetas(gc, oc_id)
        id_map      = {c["name"]: c["id"] for c in carpetas_sp}

        created = _ensure_new_structure(gc, oc_id, id_map, config, oc_log)
        print(f"  {created} carpetas nuevas creadas")

        log["ocs"][nombre_oc] = oc_log
        _actualizar_registro(nombre_oc, oc_log, LOGS_DIR / f"migration_sp_execution_log_{ts}.json", backup_file)

    log_file = LOGS_DIR / f"migration_sp_execution_log_{ts}.json"
    log_file.write_text(
        json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n[OK] Log guardado en {log_file}")


# ---------------------------------------------------------------------------
# Utilidades internas
# ---------------------------------------------------------------------------

def _ensure_new_structure(gc: GraphClient, oc_id: str, id_map: dict,
                          config: dict, oc_log: dict) -> int:
    """
    Crea subcarpetas faltantes y borra las desconocidas (sin archivos) dentro
    de cada carpeta de new_structure. Retorna cantidad de carpetas creadas.
    """
    created    = 0
    folders    = config["new_structure"]["folders"]
    subfolders = config["new_structure"]["subfolders"]

    for folder in folders:
        if folder not in id_map:
            nuevo = _crear_carpeta(gc, oc_id, folder)
            id_map[folder] = nuevo["id"]
            oc_log["operations"].append({"folder": folder, "action": "create", "result": "ok"})
            print(f"  [NEW]  {folder}")
            created += 1
            time.sleep(0.15)

        folder_id       = id_map[folder]
        expected_subs   = set(subfolders.get(folder, []))
        hijos           = _listar_carpetas(gc, folder_id)
        sub_existentes  = {c["name"]: c["id"] for c in hijos}

        # Borrar subcarpetas desconocidas que no tengan archivos
        for sub_name, sub_id in sub_existentes.items():
            if sub_name not in expected_subs:
                if _tiene_archivos_recursivo(gc, sub_id):
                    oc_log["operations"].append({"folder": f"{folder}/{sub_name}", "action": "delete", "result": "warning", "detail": "tiene archivos — NO borrada"})
                    print(f"  [WARN] {folder}/{sub_name}: tiene archivos, omitida")
                else:
                    try:
                        _eliminar_carpeta(gc, sub_id)
                        oc_log["operations"].append({"folder": f"{folder}/{sub_name}", "action": "delete", "result": "ok"})
                        print(f"  [DEL]  {folder}/{sub_name}")
                        time.sleep(0.2)
                    except Exception as e:
                        oc_log["operations"].append({"folder": f"{folder}/{sub_name}", "action": "delete", "result": "error", "detail": str(e)})
                        print(f"  [ERR]  {folder}/{sub_name}: {e}")

        # Crear subcarpetas faltantes
        for sub in subfolders.get(folder, []):
            if sub in sub_existentes:
                oc_log["operations"].append({"folder": f"{folder}/{sub}", "action": "skip", "result": "ya existe"})
            else:
                _crear_carpeta(gc, folder_id, sub)
                oc_log["operations"].append({"folder": f"{folder}/{sub}", "action": "create", "result": "ok"})
                print(f"  [NEW]  {folder}/{sub}")
                created += 1
                time.sleep(0.15)

    return created


def _resolver_ocs(oc_filter: str | None) -> list[str]:
    if oc_filter:
        if oc_filter not in OCS_PRUEBA:
            sys.exit(f"[ERROR] OC '{oc_filter}' no esta en la lista. Opciones: {OCS_PRUEBA}")
        return [oc_filter]
    return OCS_PRUEBA


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def _find_recent_backup() -> Path | None:
    cutoff     = datetime.now().timestamp() - BACKUP_MAX_AGE_MINUTES * 60
    candidates = sorted(BACKUP_DIR.glob("migration_sp_backup_*.json"), reverse=True)
    for f in candidates:
        if f.stat().st_mtime >= cutoff:
            return f
    return None


# ---------------------------------------------------------------------------
# Registro de OCs migradas
# ---------------------------------------------------------------------------

def _actualizar_registro(nombre_oc: str, oc_log: dict, log_file: Path, backup_file: Path | None) -> None:
    if REGISTRO_FILE.exists():
        registro = json.loads(REGISTRO_FILE.read_text(encoding="utf-8"))
    else:
        registro = {"descripcion": "Registro de OCs con cambio de estructura ejecutado", "ocs": []}

    # Carpetas omitidas por tener archivos
    omitidas = [
        op["folder"] for op in oc_log.get("operations", [])
        if op.get("result") == "warning" and "tiene archivos" in op.get("detail", "")
    ]
    # Errores
    errores = [
        f"{op['folder']}: {op.get('detail', '')}" for op in oc_log.get("operations", [])
        if op.get("result") == "error"
    ]

    entrada = {
        "oc":           nombre_oc,
        "fecha":        datetime.now().strftime("%Y-%m-%d"),
        "escenario":    oc_log.get("scenario"),
        "log":          log_file.name,
        "backup":       f"backups/{backup_file.name}" if backup_file else None,
        "carpetas_con_archivos_omitidas": omitidas,
    }
    if errores:
        entrada["errores"] = errores

    # Reemplazar entrada existente si ya fue migrada antes
    registro["ocs"] = [e for e in registro["ocs"] if e["oc"] != nombre_oc]
    registro["ocs"].append(entrada)

    REGISTRO_FILE.write_text(
        json.dumps(registro, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  [REG]  Registro actualizado en {REGISTRO_FILE.name}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Migracion de estructura de carpetas en SharePoint")
    parser.add_argument("--mode", required=True, choices=["dry-run", "backup", "execute", "backup-execute"])
    parser.add_argument("--oc",   default=None,
                        help="Nombre de una OC especifica (opcional): 'prueba 1', 'prueba 2' o 'prueba 3'")
    args = parser.parse_args()

    config = load_config()
    gc     = GraphClient()

    if args.mode == "dry-run":
        run_dry_run(gc, args.oc, config)
    elif args.mode == "backup":
        run_backup(gc, args.oc, config)
    elif args.mode == "execute":
        run_execute(gc, args.oc, config)
    elif args.mode == "backup-execute":
        bf = run_backup(gc, args.oc, config)
        run_execute(gc, args.oc, config, backup_file=bf)


if __name__ == "__main__":
    main()
