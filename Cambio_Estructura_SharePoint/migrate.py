# -*- coding: utf-8 -*-
"""
migrate.py — Migración de estructura de carpetas por OC en SharePoint local.

Modos:
  --mode dry-run   Detecta escenario y reporta acciones sin modificar nada.
  --mode backup    Genera snapshot del estado actual (requiere dry-run previo).
  --mode execute   Ejecuta la migración (requiere backup reciente < 30 min).

Argumentos:
  --path  Ruta base donde están las OCs  (default: test_migration/)
  --oc    Nombre de una OC específica    (opcional)
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

# Forzar UTF-8 en stdout para consolas Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ── constantes ───────────────────────────────────────────────────────────────

SCRIPT_DIR   = Path(__file__).parent
CONFIG_FILE  = SCRIPT_DIR / "migration-config.json"
DRY_RUN_FILE = SCRIPT_DIR / "migration_dry_run.json"
BACKUP_MAX_AGE_MINUTES = 30


# ── carga de configuración ───────────────────────────────────────────────────

def load_config() -> dict:
    if not CONFIG_FILE.exists():
        sys.exit(f"[ERROR] No se encontró {CONFIG_FILE}")
    with CONFIG_FILE.open(encoding="utf-8") as f:
        return json.load(f)


# ── detección de escenario ───────────────────────────────────────────────────

def detect_scenario(oc_path: Path, config: dict) -> int:
    """
    Detecta el escenario de migración de una OC comparando sus carpetas
    contra new_structure.folders y old_structure.folders_to_delete.

    Retorna:
        1 — Solo estructura vieja (hay carpetas old, NO hay carpetas new)
        2 — Ambas estructuras    (hay carpetas old Y carpetas new)
        3 — Solo estructura nueva (NO hay carpetas old, hay carpetas new o está vacía)
    """
    new_set = set(config["new_structure"]["folders"])
    old_set = set(config["old_structure"]["folders_to_delete"])

    existing = {p.name for p in oc_path.iterdir() if p.is_dir()}

    has_old = bool(existing & old_set)
    has_new = bool(existing & new_set)

    if has_old and has_new:
        return 2
    if has_old and not has_new:
        return 1
    # has_new (with or without files) or completely empty → no touch
    return 3


# ── helpers compartidos ──────────────────────────────────────────────────────

def find_ocs(base: Path, oc_filter: str | None, config: dict) -> list[Path]:
    import re
    pattern = re.compile(config["new_structure"]["root_pattern"])
    ocs = sorted(
        p for p in base.iterdir()
        if p.is_dir() and pattern.match(p.name)
    )
    if oc_filter:
        ocs = [p for p in ocs if p.name == oc_filter]
        if not ocs:
            sys.exit(f"[ERROR] OC '{oc_filter}' no encontrada en {base}")
    return ocs


def folders_to_delete_for_oc(oc_path: Path, config: dict) -> tuple[list[str], list[str]]:
    """
    Retorna (to_delete, to_keep) para una OC según sus carpetas existentes.

    to_delete: carpetas presentes en old_structure.folders_to_delete
               que NO están en new_structure.folders ni en folders_to_preserve.
    to_keep:   carpetas que se conservan (new o preserve).
    """
    old_set      = set(config["old_structure"]["folders_to_delete"])
    new_set      = set(config["new_structure"]["folders"])
    preserve_set = set(config["old_structure"]["folders_to_preserve"])
    safe_set     = new_set | preserve_set

    existing = {p.name for p in oc_path.iterdir() if p.is_dir()}

    to_delete = sorted((existing & old_set) - safe_set)
    to_keep   = sorted(existing & safe_set)
    return to_delete, to_keep


# ── modo dry-run ─────────────────────────────────────────────────────────────

def run_dry_run(base: Path, oc_filter: str | None, config: dict) -> None:
    ocs    = find_ocs(base, oc_filter, config)
    report = {"generated_at": _now_iso(), "base_path": str(base), "ocs": {}}

    for oc_path in ocs:
        scenario              = detect_scenario(oc_path, config)
        to_delete, to_keep    = folders_to_delete_for_oc(oc_path, config)
        all_folders           = sorted(p.name for p in oc_path.iterdir() if p.is_dir())

        if scenario == 3:
            to_delete = []
            to_keep   = all_folders

        entry = {
            "scenario":  scenario,
            "action":    config["migration"]["scenarios"][str(scenario)],
            "to_delete": to_delete,
            "to_keep":   to_keep,
        }
        report["ocs"][oc_path.name] = entry

        print(f"\n{'-'*60}")
        print(f"  OC: {oc_path.name}  ->  Escenario {scenario}")
        print(f"  {entry['action']}")
        if to_delete:
            print(f"  BORRARÍA ({len(to_delete)}):")
            for f in to_delete:
                print(f"    - {f}")
        else:
            print("  BORRARÍA: (ninguna)")
        print(f"  CONSERVARÍA ({len(to_keep)}): {', '.join(to_keep) or '(ninguna)'}")

    DRY_RUN_FILE.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n[OK] Reporte guardado en {DRY_RUN_FILE}")


# ── modo backup ──────────────────────────────────────────────────────────────

def run_backup(base: Path, oc_filter: str | None, config: dict) -> None:
    if not DRY_RUN_FILE.exists():
        sys.exit(f"[ERROR] Ejecuta --mode dry-run primero. No existe {DRY_RUN_FILE}")

    ocs     = find_ocs(base, oc_filter, config)
    ts      = _timestamp()
    snapshot = {"generated_at": _now_iso(), "base_path": str(base), "ocs": {}}

    for oc_path in ocs:
        oc_entry = {}
        for folder in sorted(oc_path.iterdir()):
            if folder.is_dir():
                oc_entry[folder.name] = sorted(
                    f.name for f in folder.iterdir() if f.is_file()
                )
        snapshot["ocs"][oc_path.name] = oc_entry

    backup_file = SCRIPT_DIR / f"migration_backup_{ts}.json"
    backup_file.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"[OK] Backup guardado en {backup_file}")


# ── modo execute ─────────────────────────────────────────────────────────────

def run_execute(base: Path, oc_filter: str | None, config: dict) -> None:
    backup_file = _find_recent_backup()
    if not backup_file:
        sys.exit(
            f"[ERROR] No existe un backup generado en los últimos "
            f"{BACKUP_MAX_AGE_MINUTES} minutos. Ejecuta --mode backup primero."
        )

    ocs     = find_ocs(base, oc_filter, config)
    ts      = _timestamp()
    log     = {"generated_at": _now_iso(), "backup_used": backup_file.name, "ocs": {}}

    for oc_path in ocs:
        scenario           = detect_scenario(oc_path, config)
        to_delete, to_keep = folders_to_delete_for_oc(oc_path, config)
        oc_log             = {"scenario": scenario, "operations": []}

        # Escenario 3: solo crear subfolders faltantes, sin borrar nada
        if scenario == 3:
            created = _ensure_new_structure(oc_path, config, oc_log)
            print(f"  {oc_path.name}: Escenario 3 — sin borrados, {created} carpetas creadas")
            log["ocs"][oc_path.name] = oc_log
            continue

        # Escenarios 1 y 2: borrar carpetas viejas
        for folder_name in to_delete:
            folder_path = oc_path / folder_name
            op = {"folder": folder_name, "action": "delete"}

            if not folder_path.exists():
                op.update(result="skip", detail="Ya no existe")
            elif _has_files(folder_path):
                op.update(result="warning",
                          detail="Carpeta con archivos — NO borrada")
                print(f"  [WARN] {oc_path.name}/{folder_name}: tiene archivos, omitida")
            else:
                try:
                    os.rmdir(folder_path)
                    op.update(result="ok", detail="Borrada")
                    print(f"  [DEL]  {oc_path.name}/{folder_name}")
                except OSError as e:
                    op.update(result="error", detail=str(e))
                    print(f"  [ERR]  {oc_path.name}/{folder_name}: {e}")

            oc_log["operations"].append(op)

        for folder_name in to_keep:
            oc_log["operations"].append(
                {"folder": folder_name, "action": "keep", "result": "ok"}
            )

        # Crear toda la estructura nueva (folders + subfolders) que falte
        created = _ensure_new_structure(oc_path, config, oc_log)
        print(f"  {oc_path.name}: {created} carpetas nuevas creadas")

        log["ocs"][oc_path.name] = oc_log

    log_file = SCRIPT_DIR / f"migration_execution_log_{ts}.json"
    log_file.write_text(
        json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"\n[OK] Log de ejecución guardado en {log_file}")


# ── utilidades internas ──────────────────────────────────────────────────────

def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def _has_files(folder: Path) -> bool:
    return any(True for _ in folder.rglob("*") if _.is_file())

def _ensure_new_structure(oc_path: Path, config: dict, oc_log: dict) -> int:
    """
    Crea todos los folders y subfolders de new_structure que no existan.
    Registra cada creación en oc_log. Retorna el número de carpetas creadas.
    """
    created = 0
    folders    = config["new_structure"]["folders"]
    subfolders = config["new_structure"]["subfolders"]

    for folder in folders:
        folder_path = oc_path / folder
        if not folder_path.exists():
            folder_path.mkdir(parents=True, exist_ok=True)
            oc_log["operations"].append(
                {"folder": folder, "action": "create", "result": "ok"}
            )
            created += 1

        for sub in subfolders.get(folder, []):
            sub_path = folder_path / sub
            if not sub_path.exists():
                sub_path.mkdir(parents=True, exist_ok=True)
                oc_log["operations"].append(
                    {"folder": f"{folder}/{sub}", "action": "create", "result": "ok"}
                )
                created += 1

    return created


def _find_recent_backup() -> Path | None:
    cutoff = datetime.now().timestamp() - BACKUP_MAX_AGE_MINUTES * 60
    candidates = sorted(SCRIPT_DIR.glob("migration_backup_*.json"), reverse=True)
    for f in candidates:
        if f.stat().st_mtime >= cutoff:
            return f
    return None


# ── entry point ──────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description="Migración de estructura de carpetas por OC")
    parser.add_argument("--mode",   required=True, choices=["dry-run", "backup", "execute"])
    parser.add_argument("--path",   default="test_migration/",
                        help="Ruta base donde están las OCs")
    parser.add_argument("--oc",     default=None,
                        help="Nombre de una OC específica (opcional)")
    args = parser.parse_args()

    config   = load_config()
    base     = Path(args.path)
    if not base.is_absolute():
        base = SCRIPT_DIR / base
    if not base.exists():
        sys.exit(f"[ERROR] La ruta base no existe: {base}")

    if args.mode == "dry-run":
        run_dry_run(base, args.oc, config)
    elif args.mode == "backup":
        run_backup(base, args.oc, config)
    elif args.mode == "execute":
        run_execute(base, args.oc, config)


if __name__ == "__main__":
    main()
