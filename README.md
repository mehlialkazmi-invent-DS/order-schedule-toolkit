# Order schedule toolkit

Small, shareable workflow for tbretail **order schedule** updates from **Jira** Excel attachments.

## Prerequisite: Forge Anvil

`scripts/order_sch_main.py` reads the datastore (order schedule prod file + fiscal calendar)
through **Forge Anvil**'s `DataWorkbench` — no Spark session or `customer-pipeline-tbretail`
checkout needed. You need:

- **Azure CLI**, logged in (`az login`) with access to the `tbretail` datastore.
- Network access to Invent's private PyPI index (VPN if required).
- Repo: <https://github.com/inventanalytics/forge-anvil> (source, docs, and the `workbench`
  Jupyter CLI — not required to just run this toolkit, only if you want the interactive
  notebook/Jupyter workflow described there).

`local_forge_venv` (see Quick start below) installs `invent-forge-anvil` for you; you don't need
to clone `forge-anvil` itself unless you want to read its docs or use its `workbench` CLI directly.

## Quick start

1. Clone this folder as its own repo (or copy into your monorepo).
2. **Optional — one-time local inputs (recommended for the Claude Code skill's "just run"):**
   - Copy `config/local.settings.example.json` → `config/local.settings.json` (edit paths, Jira site, email, optional `default_issue_key`).
   - Copy `config/secrets.local.example.env` → `config/secrets.local.env` (set `JIRA_API_TOKEN`). Both files are gitignored.
3. **Create `local_forge_venv`** (one-time; this is the only venv this toolkit needs):
   ```bash
   uv venv local_forge_venv
   uv pip install --python local_forge_venv/bin/python \
     --index https://pypi.euwest1.prod.inventanalytics.com/ \
     invent-forge-anvil pandas numpy openpyxl
   ```
4. **One-shot pipeline** (stdlib launcher; real work uses `local_forge_venv_python` from config):
   ```bash
   python3 scripts/run_pipeline.py TBRCS-659
   # or: python3 scripts/run_pipeline.py   # if default_issue_key is in local.settings.json
   ```
5. If you skipped step 2, export Jira credentials instead (**`jira_input.py` uses REST to download the attachment**):
   - `JIRA_BASE_URL` (e.g. `https://your-site.atlassian.net`), `JIRA_EMAIL`, `JIRA_API_TOKEN` — create a token at https://id.atlassian.com/manage-profile/security/api-tokens
6. From the **toolkit root** (use `local_forge_venv`'s `python` if not using `run_pipeline.py`):
   ```bash
   python scripts/jira_input.py YOUR-123
   ```
   If you cannot use REST in the shell but have a saved file + issue text:
   ```bash
   python scripts/jira_input.py YOUR-123 --excel-path /path/to/file.xlsx --summary "..." --description-file /path/to/desc.txt
   ```
7. Either edit `scripts/order_sch_main.ipynb` (set `JIRA_ISSUE_KEY`, select the `local_forge_venv` kernel) **or** run the same pipeline as a script:
   ```bash
   python scripts/order_sch_main.py --issue-key YOUR-123
   ```

Artifacts:

- **Backups:** `local_backup/`
- **Jira:** `jira_downloads/raw/`, `jira_downloads/cleaned/` (includes `*_pipeline_meta.json` for **IS_FISCAL**)
- **Upload files (local only):** `output/` and `output/final/`

Nothing in this toolkit writes to the **datastore**.

## Claude Code skill

The skill lives at `.claude/commands/order-schedule-pipeline.md`. Claude Code loads **project**
commands when this folder is (part of) the open workspace.

### How to test the skill

1. **Open the toolkit root** (or a parent workspace that contains it) in Claude Code.
2. Ensure the skill file exists: `.claude/commands/order-schedule-pipeline.md`.
3. Start a chat and prompt with a phrase that matches the skill description, for example:
   - "Run the **order schedule pipeline** for **YOUR-123** using the toolkit."
   - Or mention "order schedule Jira attachment" and the issue key.
4. Confirm the agent: runs or instructs `python scripts/jira_input.py YOUR-123`, reads `jira_downloads/cleaned/YOUR-123_pipeline_meta.json`, and walks through the merge step, using `local_forge_venv` throughout (never `pip install`s into another interpreter).

A synced copy also lives at the workspace root's `.claude/commands/order-schedule-pipeline.md` so it's discoverable from outside this folder — after editing one, copy the change to the other so they stay identical.

## Legacy notebooks

Older experiments may still live under `customers/tbretail/order schedule tasks/scripts/` (e.g. `order_schedule_v2.ipynb`, duplicate checker). This toolkit is the maintained entry point.
