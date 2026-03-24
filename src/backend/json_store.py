import json
import os
import shutil


def read_json_with_backup(path, default_value):
    """Read JSON file with fallback to <file>.bak if primary is unreadable."""
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        bak_path = f"{path}.bak"
        try:
            with open(bak_path, encoding="utf-8") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return default_value


def atomic_write_json(path, payload):
    """Write JSON atomically and keep a one-file backup of previous content."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_path = f"{path}.tmp"
    bak_path = f"{path}.bak"

    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())

    if os.path.exists(path):
        shutil.copy2(path, bak_path)
    os.replace(tmp_path, path)
