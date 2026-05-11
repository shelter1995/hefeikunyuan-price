from __future__ import annotations

import os
from pathlib import Path


def _strip_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        return value[1:-1]
    return value


def load_env_file(path: str | Path = ".env", overwrite: bool = False) -> Path | None:
    """
    Load KEY=VALUE pairs from a .env file into process environment.
    Returns resolved Path when file exists, otherwise None.
    """
    env_path = Path(path)
    if not env_path.exists() or not env_path.is_file():
        return None

    for raw_line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.lower().startswith("export "):
            line = line[7:].strip()
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = _strip_quotes(value)
        if not key:
            continue
        if overwrite or key not in os.environ:
            os.environ[key] = value

    return env_path.resolve()
