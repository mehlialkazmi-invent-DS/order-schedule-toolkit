#!/usr/bin/env python3
"""
One-shot pipeline: jira_input.py then order_sch_main.py.

Uses only the stdlib in this file; delegates to local_forge_venv (see config/local.settings.json).

From toolkit root (after `source local_forge_venv/bin/activate`, or with local_forge_venv_python in local.settings.json):
  python3 scripts/run_pipeline.py TBRCS-659
  python3 scripts/run_pipeline.py              # uses default_issue_key from config/local.settings.json

Both steps use python_for_toolkit() (activated VIRTUAL_ENV wins after LOCAL_FORGE_VENV_PYTHON).
order_sch_main.py reads the datastore via Forge Anvil's DataWorkbench — no Spark needed.
Requires config for REST: local.settings.json + secrets.local.env (see *.example.*).
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from toolkit_env import (  # noqa: E402
    bootstrap_toolkit_env,
    default_issue_key,
    python_for_toolkit,
    toolkit_root,
)


def main() -> int:
    bootstrap_toolkit_env()
    issue = sys.argv[1] if len(sys.argv) > 1 else default_issue_key()
    if not issue:
        print(
            "Usage: python3 scripts/run_pipeline.py [ISSUE-KEY]\n"
            "Or set default_issue_key in config/local.settings.json",
            file=sys.stderr,
        )
        return 1
    issue = issue.strip().upper()
    py = python_for_toolkit()
    root = toolkit_root()
    r1 = subprocess.run([py, str(root / "scripts" / "jira_input.py"), issue], cwd=str(root))
    if r1.returncode:
        return r1.returncode
    r2 = subprocess.run(
        [py, str(root / "scripts" / "order_sch_main.py"), "--issue-key", issue],
        cwd=str(root),
    )
    return r2.returncode


if __name__ == "__main__":
    raise SystemExit(main())
