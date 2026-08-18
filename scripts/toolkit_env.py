"""
Load toolkit-local defaults from config files (gitignored).

Precedence: existing OS environment variables always win; files only fill gaps.

Files (under toolkit root / config/):
  local.settings.json  — paths, Jira site, email, default issue key (optional)
  secrets.local.env    — KEY=VALUE lines, typically JIRA_API_TOKEN

See config/local.settings.example.json and config/secrets.local.example.env
"""

from __future__ import annotations

import json
import os
from pathlib import Path

_SNAKE_TO_ENV = {
    "jira_base_url": "JIRA_BASE_URL",
    "jira_email": "JIRA_EMAIL",
    "jira_api_version": "JIRA_API_VERSION",
    "local_forge_venv_python": "LOCAL_FORGE_VENV_PYTHON",
    "default_issue_key": "DEFAULT_ISSUE_KEY",
}


def toolkit_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _parse_dotenv(path: Path) -> dict[str, str]:
    out: dict[str, str] = {}
    if not path.is_file():
        return out
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        k, _, v = line.partition("=")
        k = k.strip()
        v = v.strip().strip('"').strip("'")
        if k:
            out[k] = v
    return out


def bootstrap_toolkit_env() -> None:
    """
    Merge config/local.settings.json and config/secrets.local.env into os.environ
    only for keys that are not already set.
    """
    root = toolkit_root()
    cfg = root / "config"

    settings_path = cfg / "local.settings.json"
    if settings_path.is_file():
        try:
            data = json.loads(settings_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Invalid JSON in {settings_path}: {e}") from e
        if not isinstance(data, dict):
            raise RuntimeError(f"{settings_path} must be a JSON object")
        for snake, val in data.items():
            if val is None:
                continue
            env_key = _SNAKE_TO_ENV.get(str(snake))
            if not env_key:
                continue
            if env_key in os.environ and os.environ[env_key].strip():
                continue
            os.environ[env_key] = str(val).strip()

    secrets_path = cfg / "secrets.local.env"
    for k, v in _parse_dotenv(secrets_path).items():
        if k in os.environ and os.environ[k].strip():
            continue
        os.environ[k] = v


def python_for_toolkit() -> str:
    """Interpreter path: explicit env, then activated venv, then toolkit-local local_forge_venv."""
    bootstrap_toolkit_env()
    py = os.environ.get("LOCAL_FORGE_VENV_PYTHON", "").strip()
    if py:
        return py
    venv = os.environ.get("VIRTUAL_ENV", "").strip()
    if venv:
        for name in ("python", "python3"):
            cand = Path(venv) / "bin" / name
            if cand.is_file():
                return str(cand)
    cand = toolkit_root() / "local_forge_venv" / "bin" / "python"
    if cand.is_file():
        return str(cand)
    return os.environ.get("PYTHON", "python3")


def default_issue_key() -> str | None:
    bootstrap_toolkit_env()
    k = os.environ.get("DEFAULT_ISSUE_KEY", "").strip()
    return k.upper() if k else None
