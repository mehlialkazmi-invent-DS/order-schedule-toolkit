---
name: order-schedule-pipeline
description: "For tbretail order-schedule-toolkit: run_pipeline.py from toolkit root (cd) under local_forge_venv; jira_input then order_sch_main. Reads the datastore via Forge Anvil's DataWorkbench (Azure, customer_name=tbretail) — no Spark/customer-pipeline checkout needed. Fiscal cadence by planner intent: 3 rows (first week each fiscal month in-quarter, 4-5-4), 1 row quarterly (91), 1 row semi-annual (182). Anchor weekday + fiscal_cal; 28/7 in Excel are not the driver. TBRCS-693. Read Planner notes."
---

# Order schedule pipeline

The pipeline downloads and cleans Excel from Jira, merges the submission with production schedules from the datastore (read-only), and writes local CSVs plus text and HTML merge reports. Work from the `order-schedule-toolkit/` clone root.

Do not write to the datastore; only read prod and fiscal calendar inputs.

**Canonical copy:** This skill lives at `.claude/commands/order-schedule-pipeline.md` inside `customers/tbretail/order-schedule-toolkit/`. A synced mirror lives at the workspace root's `.claude/commands/order-schedule-pipeline.md` for discovery outside this folder — after editing one, copy the change to the other so they stay identical.

## Prerequisite: Forge Anvil

Reading the datastore requires **Forge Anvil**'s `DataWorkbench` (<https://github.com/inventanalytics/forge-anvil>), installed into this toolkit's own `local_forge_venv` (see [Config and first run](#config-and-first-run)). You also need Azure CLI logged in (`az login`) with access to the `tbretail` datastore, and network access to Invent's private PyPI index (VPN if required) the first time you install. No Spark session and no `customer-pipeline-tbretail` checkout are needed anymore.

---

## Contents

| Section | What |
|--------|------|
| [Repository layout](#repository-layout) | Paths and artifacts |
| [Config and first run](#config-and-first-run) | Settings, secrets, venv |
| [Preflight](#preflight-if-config-is-missing) | What to ask the user |
| [Jira access](#jira-access) | REST, MCP, local Excel |
| [Excel layout (title row)](#excel-layout-title-row--header-detection) | Workbooks with a row above column names; `read_submission_excel` |
| [All attachments](#all-attachments-excel-screenshots-inline-media) | Excel vs images; 0-row merges; screenshots with destination codes |
| [Python interpreter](#python-interpreter-local_forge_venv) | local_forge_venv resolution |
| [Agent checklist](#agent-execution-checklist) | Step-by-step when invoked |
| [Pairs and destinations](#pairs-and-destinations) | How pairs and blank destinations work |
| [Pipeline behavior](#pipeline-behavior) | Fiscal, frequency, merge rules |
| [Fiscal calendar cadence](#fiscal-calendar-cadence-supply-chain) | 91 vs 182, weekdays, agent guidance |
| [Fiscal months vs civil dates (454)](#fiscal-months-vs-civil-dates-454) | Last-Wed or first-week anchors in **Fiscal_Year_Month**; validate in `fiscal_cal` |
| [Planner notes (Excel)](#planner-notes-excel) | Optional 6th column; not in CSV—read raw workbook for intent |
| [Fiscal cadence patterns](#fiscal-cadence-patterns-planner-intent) | 3-row 4-5-4 in-quarter vs 1-row quarterly (91) vs semi-annual (182) |
| [Uniform date gaps (fiscal)](#uniform-date-gaps-fiscal) | How merge detects quarterly vs in-quarter from civil spacing |
| [Every fiscal quarter, first week](#every-fiscal-quarter-first-week-of-the-month) | One seed row; 4-5-4 weeks sum to 91 per quarter |
| [Pitfalls](#pitfalls-and-verification) | Common mistakes from real runs |
| [Troubleshooting](#troubleshooting) | When things fail |

---

## Repository layout

### Config (gitignored)

| Path | Role |
|------|------|
| `config/local.settings.json` | Paths, `jira_base_url`, `jira_email`, optional `default_issue_key`, optional `local_forge_venv_python` (from `local.settings.example.json`) |
| `config/secrets.local.env` | `JIRA_API_TOKEN=` (from `secrets.local.example.env`) |

`scripts/toolkit_env.py` loads these into `os.environ` without overriding keys already set.

### Scripts

| Script | Role |
|--------|------|
| `scripts/run_pipeline.py` | Stdlib launcher: `jira_input` then `order_sch_main` using configured venv |
| `scripts/jira_input.py` | Jira REST (or `--excel-path`) → clean CSV + `{KEY}_pipeline_meta.json` |
| `scripts/order_sch_main.py` | Forge Anvil DataWorkbench merge → `output/` (skips pairs that already match prod) |
| `scripts/order_sch_main.ipynb` | Same as script; `bootstrap_toolkit_env()` in config cell |

### Data directories

| Path | Role |
|------|------|
| `local_backup/` | Timestamped backups from datastore reads |
| `jira_downloads/raw/` | Raw Excel from REST or copy |
| `jira_downloads/cleaned/` | See [Cleaned outputs](#cleaned-outputs) below |
| `output/final/` | `order_schedule_input_prod.csv` (typical upload candidate; overwritten each run) |

### Cleaned outputs

| Artifact | Purpose |
|----------|---------|
| `{KEY}_cleaned.csv` | Cleaned rows; overwritten after merge step with merge-ready input |
| `{KEY}_cleaned_applied.csv` | Only submission rows for pairs that changed prod |
| `{KEY}_pipeline_meta.json` | `is_fiscal`, `fiscal_anchor_weekday`, `rationale`, paths, counters |
| `{KEY}_jira_context.txt` | Summary + description snapshot |
| `{KEY}_merge_report.txt` | Human-readable log + embedded CSV sections |
| `{KEY}_merge_report.html` | Interactive tables (Tabulator inlined at generation). Slack/Jira in-app preview is not interactive; download the file and open it in a desktop browser (Chrome, Edge, Firefox). |

---

## Config and first run

### One-time setup (user, not in chat)

1. Copy `config/local.settings.example.json` → `config/local.settings.json` and edit paths, Jira site, email, optional `default_issue_key`.
2. Copy `config/secrets.local.example.env` → `config/secrets.local.env` and set `JIRA_API_TOKEN` ([Atlassian: API tokens](https://id.atlassian.com/manage-profile/security/api-tokens)).

### When the skill runs

1. Read (never print token values) whether `local.settings.json` and `secrets.local.env` exist and look valid.
2. If something is missing, ask only for what is absent. Never ask the user to paste `JIRA_API_TOKEN` into chat.
3. Issue key: use the key from the user message (e.g. `TBRCS-659`), else `default_issue_key` from JSON, else ask.

### Run commands (toolkit root)

**Always `cd` into the `order-schedule-toolkit` directory first.** Running `python3 scripts/run_pipeline.py` from a parent folder (e.g. workspace root) fails because `./scripts/run_pipeline.py` is not there.

**One-time setup:** create `local_forge_venv` (this toolkit's only venv) and install its deps from the private index:

```bash
cd "<order-schedule-toolkit>"
uv venv local_forge_venv
uv pip install --python local_forge_venv/bin/python \
  --index https://pypi.euwest1.prod.inventanalytics.com/ \
  invent-forge-anvil pandas numpy openpyxl
```

Option A — venv already activated:

```bash
cd "<order-schedule-toolkit>"
source local_forge_venv/bin/activate
python3 scripts/run_pipeline.py TBRCS-659
# or omit key if default_issue_key is set:
python3 scripts/run_pipeline.py
```

After `source`, `VIRTUAL_ENV` is set; the launcher uses that interpreter even if `local_forge_venv_python` is unset in JSON.

Option B — no activation: `python_for_toolkit()` defaults to `<toolkit_root>/local_forge_venv/bin/python` automatically once it exists; no config needed. Override with `local_forge_venv_python` in `local.settings.json` only if you keep the venv elsewhere. Then:

```bash
python3 scripts/run_pipeline.py TBRCS-659
```

---

## Preflight (if config is missing)

Ask only for what you still need:

| Need | Why |
|------|-----|
| Jira issue key | e.g. `TBRCS-659` |
| Paths / site / email | Or point to creating `local.settings.json` from the example |
| `JIRA_API_TOKEN` | Only in `secrets.local.env` or shell env — not in chat |
| Excel file name | Optional `--attachment "exact name"` when several workbooks on the issue (default: all Excel attachments merged) |
| Fiscal ambiguity | `--fiscal` / `--no-fiscal` |

Do not `pip install` into arbitrary interpreters. Subprocesses must use `local_forge_venv` (`VIRTUAL_ENV`, or `local_forge_venv_python`, or the toolkit-local `local_forge_venv/bin/python` default).

---

## Jira access

### A — Jira REST (recommended)

`jira_input.py` loads the issue and downloads each `.xlsx` / `.xls` via attachment `content` URLs (Basic auth: email + token). Rows from all files merge into one `{KEY}_cleaned.csv`. **PNG/JPEG/PDF and other non-Excel attachments are ignored by this script**—see [All attachments](#all-attachments-excel-screenshots-inline-media).

- Multiple workbooks: all are downloaded (sorted by filename), cleaned, concatenated; exact duplicate rows across files removed.
- Single workbook only: `--attachment "exact filename"`.
- Local file only: `--excel-path` (one file).

Manual env example (prefer config files in practice):

```bash
export JIRA_BASE_URL="https://YOUR-SITE.atlassian.net"
export JIRA_EMAIL="you@company.com"
export JIRA_API_TOKEN="your_api_token"
# optional Server/DC: export JIRA_API_VERSION=2

TBPY="<order-schedule-toolkit>/local_forge_venv/bin/python"
cd "<order-schedule-toolkit>"
"$TBPY" scripts/jira_input.py TBRCS-659
# "$TBPY" scripts/jira_input.py TBRCS-659 --attachment "exact.xlsx"
```

A Jira UI login in your editor does not populate the shell; `local.settings.json` and `secrets.local.env` are what make terminal runs work.

### Excel layout (title row / header detection)

Some submissions put a **banner or title in row 1** (e.g. a banner code like `JAB`) and the **real column headers on row 2** (`order_group_description`, `destination_code`, …). If `pandas` reads with the default first row as the header, the loader sees bogus columns such as `JAB`, `Unnamed: 1`, … and **`jira_input.py` fails** with:

`Submission missing required columns: [...]`

**Implementation:** `scripts/jira_input.py` uses `read_submission_excel()`: it reads the workbook with **header row index 0, then 1, then 2**, and picks the first result where all `ORDER_COLUMNS` are present. Agents do not need a flag for normal title-row workbooks.

**If it still fails:** open the `.xlsx` in Excel—confirm the standard five columns exist on **one** contiguous header row; remove merged title blocks or move headers to row 1, or build `--manual-csv`.

**Real ticket pattern (TBRCS-685):** single sheet `order groups`, title row + headers, fiscal **91** rows with spacer blank lines—cleaning drops blank rows; merge **kept ≥3 fiscal lines per pair as-is**; **prod row count unchanged** (replacements only)—see [Interpreting merge deltas](#interpreting-merge-deltas).

**Real ticket pattern (TBRCS-687):** nine new MSP D30 order groups in one workbook. Requesters sometimes enter **three anchor dates** spaced by `order_frequency` (e.g. 28 days) **across** coat / pant / vest rows as if staggered launches. If the business intent is **one shared first order date for every order group on the ticket**, the default non-fiscal rule is still **earliest date per `(order_group, destination)` pair**—it does **not** automatically collapse different pairs to one global min date. **Normalize in Excel** or use **`jira_input.py ISSUE --manual-csv corrected.csv`** so every row carries that **same** `order_schedule_date`. **Typo in `order_group_description`** (e.g. `..._BUE_PANT_FLT` vs `..._BLUE_PANT_FLT`) is a **different prod key** than the corrected spelling; merge **updates only rows present in the submission**, so a typo row already published stays in prod unless you **drop it** from the upload file or datastore before/after merge. After editing, run `jira_input` with `--manual-csv` then `order_sch_main` (same venv); `run_pipeline.py` does not pass through `--manual-csv`.

**Real ticket pattern (TBRCS-693):** **dual-banner** — two workbooks. **MSP:** [Every fiscal quarter, first week](#every-fiscal-quarter-first-week-of-the-month) — Excel may list four **91-day**-spaced rows, but prod needs **one seed row** only (`6/3/2026`, `order_frequency` 91); each fiscal quarter’s first Wednesday is derived downstream ([uniform gaps](#uniform-date-gaps-fiscal)). **TMW:** three rows for **4-5-4 week rhythm inside one quarter** (first Wed per fiscal month); **uneven** civil gaps (35/28) → keep all **3** submitted rows. Does not change pre-693 rules (≥3 fiscal rows kept when gaps are not uniform, last/first Wed validation, etc.).

**Real ticket pattern (TBRCS-753):** copy-pasted template workbook — filename and ticket said new group `KNG_W0_DANNECRAFT_FLT`, but the sheet's `order_group_description` still read a **different** order group (`KNG_W0_MAGID_FLT`) left over from whatever ticket the template was copied from. Same row also stacked **six dates in one `order_schedule_date` cell** (`8/24/2026, 9/28/2026, ...`) and put free text in `order_frequency` (`LAST MONDAY OF EVERY FISCAL MONTH`) instead of a number — `jira_input.py` correctly parsed **0** rows rather than guessing. **Planner notes** said `SAME ORDER CALENDAR AS KNG_W0_CBC_FLT`, matching the ticket description ("same monthly schedule as existing order group KNG_W0_CBC_FLT"). Resolution: do **not** try to salvage the malformed row — instead **look up the referenced existing order group's actual prod rows** (`grep` the order group in the prod CSV / backup) and treat that as the source of truth for shape (row count, blank vs literal destination, `order_frequency`, review calendar), confirm the intended new group name with the requester/user, then build a `--manual-csv` with that shape under the corrected name and the earliest N submitted dates. Confirm the group name and the "same as X" target with the user before writing the manual CSV — don't silently pick one.

### B — Atlassian MCP

Use MCP to read issue summary, description, and **full attachment lists** (including image filenames) when REST is not in the terminal or when you need to **inspect screenshots** for destination codes before or after a pipeline run. MCP does not replace `jira_input.py` Excel downloads unless `JIRA_*` is also set.

### C — Local Excel only

```bash
"$TBPY" scripts/jira_input.py TBRCS-659 \
  --excel-path "/path/to/file.xlsx" \
  --summary "..." \
  --description-text "..."   # or --description-file
```

### D — No Excel on the issue (manual CSV)

When there is no workbook, build a CSV with the same columns as the toolkit (`order_group_description`, `destination_code`, `order_schedule_date`, `order_frequency`, `order_review_calendar`) from the ticket description/summary (agent drafts rows). Then:

```bash
"$TBPY" scripts/jira_input.py TBRCS-659 --manual-csv /path/to/drafted.csv
```

With `JIRA_*` set, summary/description are still loaded from the issue for fiscal inference. Without Jira REST, pass `--summary` and `--description-text` (or `--description-file`).

---

## Planner notes (Excel)

Many submission templates add a sixth column **`Planner notes`** (or similar). `jira_input.py` ingests only the five `ORDER_COLUMNS`; notes are **not** in `{KEY}_cleaned.csv`. Agents should still open the **raw** workbook (or `jira_downloads/raw/`) and read notes before QA or stakeholder replies.

| Note (examples) | Typical meaning | Validation |
|-----------------|-----------------|------------|
| *Quarterly first week of the month* / *every fiscal quarter, first week* | **One prod row** per pair: earliest date; quarterly first-Wed is derived from fiscal calendar + `order_frequency` **91** | [Every fiscal quarter, first week](#every-fiscal-quarter-first-week-of-the-month); [Uniform gaps](#uniform-date-gaps-fiscal) |
| *Every 28 days on a 4-5-4 Calendar* | **Three rows** in one quarter (first anchor weekday per **consecutive** fiscal month); **`OrdReview_28D_*`** is review naming | Keep **3** rows when gaps are uneven; uniform **28-day** gaps → seed + expand to 3 ([Uniform gaps](#uniform-date-gaps-fiscal)) |

Notes can clarify intent when Jira summary/description is vague (e.g. “change order review calendar” only).

---

## Fiscal cadence patterns (planner intent)

**Yes — this is in the skill and in `order_sch_main.py`.** For fiscal tickets, infer **pattern from planner intent and dates**, not from literal **7** / **28** in Excel or `OrdReview_28D_*` review-calendar names. A fiscal quarter is **13 weeks (91 days)** because 4‑5‑4 weeks add up to 91; that does **not** mean every schedule with `order_frequency` **91** needs four prod rows.

**Always:** pick the **anchor weekday** (e.g. Wednesday) and use **`fiscal_cal`** (`Fiscal_Year_Month`, first vs last occurrence of that weekday in the month — see [Planner notes](#planner-notes-excel)).

| Planner wants | Prod rows | Typical `order_frequency` | What the dates represent |
|-----------------|-----------|---------------------------|---------------------------|
| **First week of every fiscal month** inside **one quarter** (capture **4‑5‑4** week rhythm) | **3** | **91** on each row (fiscal cadence flag, not “91 civil days apart”) | First anchor weekday in each of **3 consecutive** fiscal months (e.g. `6/3`, `7/8`, `8/5`). Civil gaps may be **uneven** (35 / 28) — still **3 rows**. |
| **Every fiscal quarter** (first week of month, rolling through the year) | **1** | **91** | One seed; each later quarter’s first-week anchor is derived downstream. Four Excel rows with **91 civil days** between them are redundant. |
| **Semi-annual** (same idea across half-year / every two fiscal months) | **1** | **182** | One seed; **182** = two fiscal months in the 91-day cadence system. Uniform **182-day** spacing in Excel → one row only (same rule as quarterly). |
| **Semi-annual in-half-year** with anchors in **every other fiscal month** inside the window | **3** | **182** | Like 4‑5‑4 but stride **2** across fiscal months — keep submitted rows or expand from seed per merge rules. |

**TBRCS-693 mapping**

- **MSP** — every fiscal quarter, first week → **1 row**, `91`.
- **TMW** — first week each fiscal month **in one quarter** (4‑5‑4) → **3 rows** per destination; uneven civil gaps; `OrdReview_28D_*` is naming only.

**What does *not* drive fiscal row count**

- `order_frequency` **7** or **28** in the workbook (unless the whole ticket is non-fiscal).
- Civil “every 28 days” between order dates when the real intent is fiscal-month anchors.
- Counting four Excel lines when uniform gaps are **91 days** apart (quarterly) — use **one** seed.

**Merge toolkit (summary)**

| Signal | Merge behavior |
|--------|----------------|
| **≥3** fiscal rows, **non-uniform** gaps | Keep all rows (in-quarter 4‑5‑4, TBRCS-693 TMW) |
| **Uniform 91-day** (or multiple of 91) gaps | **1 seed row** only (quarterly / semi-annual seed) |
| **Uniform 28-day** gaps | Earliest row + expand to **3** consecutive fiscal months |
| Otherwise | Default fiscal expand from earliest row |

Earlier ticket learnings (TBRCS-685/687, attachments, last/first Wed, `prod_row_delta`, title-row Excel, etc.) are unchanged.

---

## Uniform date gaps (fiscal)

Implementation detail for [Fiscal cadence patterns](#fiscal-cadence-patterns-planner-intent). When consecutive **civil** gaps between submitted dates are **equal**, merge uses this table:

| Uniform civil gap | Planner intent | Rows in prod |
|-------------------|----------------|--------------|
| **91** (or multiple of **91**, e.g. **182**) | Quarterly or semi-annual **across** the year — one anchor per period | **1 row** (seed only) |
| **28** | Often in-quarter 4‑5‑4 when planner typed even civil spacing | **3 rows** (seed + expand to 3 consecutive fiscal months) |
| **7 / 14** | Civil weekly / biweekly (usually non-fiscal) | Earliest row only |

**Non-uniform gaps** with **≥3** fiscal-cadence rows → **keep all** (first week of each fiscal month in-quarter; 4‑5‑4 weeks need not space evenly on the civil calendar).

**Do not confuse:** three in-quarter rows often have **`order_frequency` 91** on each line — that marks **fiscal cadence**, not “each row is 91 civil days after the previous.”

---

## Every fiscal quarter, first week of the month

**Annual quarterly** ordering (TBRCS-693 MSP) — see [Fiscal cadence patterns](#fiscal-cadence-patterns-planner-intent):

1. **One prod row** — earliest date, `order_frequency` **91**.
2. **Anchor weekday** = first occurrence in that **fiscal month** (e.g. first Wednesday).
3. Later quarters are **not** separate prod lines; 4‑5‑4 weeks per quarter still sum to **91 days**, which is why **91** is the right frequency even with a single row.

Excel may show four dates **91 civil days** apart (`6/3` → `9/2` → `12/2` → `3/3/2027`) for documentation; merge keeps **only `6/3/2026`**.

**Contrast — in-quarter (TMW):** three rows, same **91** on each row, **three consecutive fiscal months** in one quarter — that is the 4‑5‑4 **monthly** pattern, not quarterly across the year.

---

## All attachments (Excel, screenshots, inline media)

`jira_input.py` **only** downloads and parses **`.xlsx` / `.xls`** attachments. It does **not** OCR images, read PDFs, or pull inline pasted images from the issue description. Agents must still **inventory and interpret every attachment** on the ticket before trusting the pipeline output.

### Why this matters

- Requesters often attach an **empty template** workbook **and** a **screenshot** (or photo) of a filled grid. The script merges **all** Excel files on the issue (sorted by filename); if every workbook is empty, **cleaned row count is 0** and prod does not change—even though the screenshot shows valid **destination codes**, dates, and review calendars.
- **Destination codes** frequently appear **only** on the screenshot or in an image embedded in the description (Jira ADF `media` / `mediaSingle`), not in the first Excel upload.
- After a filled workbook is attached (e.g. `order_schedule_submission-<OrderGroup>.xlsx`), a **template** may remain; non-empty files still produce rows; exact duplicate rows across files are dropped.

### Agent workflow

1. **List all attachments** (Jira REST fields or **Atlassian MCP**): note every `.xlsx`/`.xls` **and** every image (`png`, `jpg`, …), PDF, etc.
2. **Open non-Excel assets** (MCP, issue in browser, or downloaded image): read **order group**, **destination_code** (per row if applicable), **dates**, **frequency**, **order_review_calendar** when present.
3. **Run `run_pipeline.py`** as usual when at least one workbook has data; use `--attachment "exact filename.xlsx"` if multiple workbooks exist and you must **exclude** an empty template or pick one submission file.
4. If **cleaned rows = 0** (or data clearly disagrees with the screenshot): do **not** stop at “merge unchanged.” Build rows from the **screenshot + description** and run `jira_input.py … --manual-csv` (see [D — No Excel](#d--no-excel-on-the-issue-manual-csv)), or ask the requester to attach a single filled Excel. Align **review calendar** strings with similar rows in prod when the screenshot uses informal text (e.g. “Weekly on Monday”) vs canonical codes.

### Pitfall (concrete)

| Situation | Risk | Action |
|-----------|------|--------|
| Template `.xlsx` + screenshot with real destinations | **0 rows** merged; missed destinations | Read the image(s); wait for filled workbook or use `--manual-csv`; optionally `--attachment` to skip blank template-only runs when other files exist |

---

## Python interpreter (local_forge_venv)

Resolve in order (first match wins):

1. `LOCAL_FORGE_VENV_PYTHON` / `local_forge_venv_python` in JSON
2. `$VIRTUAL_ENV/bin/python` after `source local_forge_venv/bin/activate`
3. `<order-schedule-toolkit>/local_forge_venv/bin/python` (default; no config needed once the venv exists)
4. `PYTHON` or `python3` on `PATH` (only trust if that is already the venv)

`local_forge_venv` needs `invent-forge-anvil`, `pandas`, `numpy`, `openpyxl` — see [Prerequisite: Forge Anvil](#prerequisite-forge-anvil) and [Config and first run](#config-and-first-run) for the one-time setup command. Reading the datastore also needs `az login` (Azure CLI) with access to the `tbretail` datastore — no Spark session, no `customer-pipeline-tbretail` checkout.

Merge only (after `jira_input`):

```bash
"$TBPY" scripts/order_sch_main.py --issue-key TBRCS-659
```

Notebook: Python: Select Interpreter → same `local_forge_venv/bin/python`.

---

## Agent execution checklist

1. Confirm config files exist; resolve issue key from message, `default_issue_key`, or ask.
2. **Attachments pass:** list **all** issue attachments (Excel + images/PDFs). If the description references a screenshot or pasted table, view it—**destination codes** and dates are often there before a filled workbook exists (see [All attachments](#all-attachments-excel-screenshots-inline-media)).
3. **`cd` to toolkit root**, then run `python3 scripts/run_pipeline.py [ISSUE-KEY]` (or `jira_input` then `order_sch_main` with the same venv). Use `--attachment "filename.xlsx"` when only one workbook should drive the merge.
4. If **cleaned rows = 0** or output conflicts with what you saw in images, reconcile with `--manual-csv` or a corrected Excel before declaring done.
5. Open `{KEY}_pipeline_meta.json`: check `is_fiscal`, `fiscal_anchor_weekday` (0 = Mon … 6 = Sun), `rationale`. If the ticket is clearly 91-day / fiscal but `is_fiscal` is false, stop and fix (see [Pitfalls](#pitfalls-and-verification), or `--fiscal`).
6. Read console after merge: **non-uniform** ≥3-row fiscal pairs kept as-is; **uniform 91-day** → one seed row only; **uniform 28-day** → seed + 3-row expand. Non-fiscal always earliest per pair.
7. Report **`pairs_applied`** and **`prod_row_delta`** together ([Interpreting merge deltas](#interpreting-merge-deltas)); delta 0 can still mean successful in-place updates; **positive delta** means prod gained rows (e.g. 3→4 fiscal lines per pair).
8. Point to `output/final/order_schedule_input_prod.csv` and note that each full run overwrites that file for the issue just processed.
9. Classify fiscal intent using [Fiscal cadence patterns](#fiscal-cadence-patterns-planner-intent) (3-row 4‑5‑4 vs 1-row quarterly **91** vs 1-row semi-annual **182**). Read **Planner notes** ([Fiscal months vs civil dates](#fiscal-months-vs-civil-dates-454)).

---

## Pairs and destinations

- A **pair** is `(order_group_description, destination_code)` for grouping, merge, and skip logic.
- **Blank, NA, or missing destination** is not imputed to a fake code. Empty cells stay empty; literals like `NA` or `Missing` in Excel are kept as typed. Internally, empty destinations share one sentinel key so they match each other, but the pipeline does not invent a destination.
- **Infer `is_fiscal`** from Jira text and/or Excel: any row with `order_frequency` that is a **multiple of 91 and ≥ 91** (91 quarterly rhythm, 182 half-year / every-two-fiscal-months, etc.) triggers fiscal mode unless the issue explicitly opts out of fiscal.

---

## Pipeline behavior

### Fiscal detection (`jira_input.py`)

| Source | Effect |
|--------|--------|
| Issue summary + description | Keywords / phrases → fiscal (including 91/182 and “every two fiscal months” style wording) |
| Excel: `order_frequency` in `{91, 182, 273, …}` (multiples of 91) | Fiscal on unless issue text explicitly opts out (non-fiscal phrases) |
| `fiscal_anchor_weekday` in meta | Issue text anchor wins if parsed; else inferred from dates on those fiscal-cadence rows, not assumed Monday |
| CLI | `--fiscal` / `--no-fiscal` override |

### Order frequency (`jira_input.py`)

`order_frequency` is always coerced to an **integer** day count:

- Plain integers (`14`, `30`, …)
- Phrases (e.g. biweekly → 14, weekly → 7; fiscal phrases like “every two fiscal months” → 182)
- Suffix codes: `14D`, `7D`
- `N days` / `N day`

Unparseable values use `--frequency-default` (default 30). For **fiscal** cadences, prefer multiples of **7** so the weekday does not drift (warnings in merge if not).

### Merge with prod (`order_sch_main.py`)

| Mode | Before merge |
|------|----------------|
| Non-fiscal | **One row per pair** (earliest date) |
| Fiscal | **Keep as-is** when **≥3** rows and gaps **not** uniform. **Uniform 91-day** (quarterly) → **one seed row** only. **Uniform 28-day** → seed + expand **3** rows (stride 1 / freq 91). Else default expand from earliest |

Per-pair rule: if prod already has the **same multiset** of schedule lines (dates, frequency, review calendar; row order ignored), **skip** (no change).

### Interpreting merge deltas

- **`prod_row_delta`** in `{KEY}_pipeline_meta.json` is **after minus before** row count in the combined prod CSV.
- **Delta +0** does **not** mean “nothing happened”: fiscal (or other) updates can **replace** existing lines for the same pairs without changing total row count (e.g. schedule date changes only).
- **Delta +N** (e.g. +5 on TBRCS-693) often means the submission has **more schedule lines per pair** than prod had (four fiscal quarters vs three historical rows; non-fiscal pair expanded to three fiscal rows per destination).
- Always read **`pairs_applied`**, **`applied_merge_rows`**, and `{KEY}_merge_report.txt` — not only the delta.

**Order-group re-entry (blank destination only):** if the submission has **only** blank-destination rows for an `order_group_description` (no row with a non-blank destination for that group), compare that multiset to **all** prod rows for that order group (any destination). If different, **replace all** prod rows for that order group with the submission rows. If the submission mixes blank and non-blank destinations for the same group, only normal per-pair rules apply.

Prod-only pairs are unchanged when not targeted.

Artifacts:

- `{KEY}_merge_report.txt`: full ticket input, merge input, applied rows, attachments list.
- `{KEY}_merge_report.html`: self-contained after run; generation needs outbound HTTPS once (Tabulator fetch). Viewing offline is fine after download.
- `pipeline_meta.json`: extended with merge paths, row counts, deltas.

Re-run after real schedule changes; identical data should mostly skip and stable counts.

---

## Fiscal calendar cadence (supply chain)

Context: recurring orders aligned to **fiscal** months/weeks (4-4-5 / 4-5-4 / 5-4-4 style environments) and fixed weekdays.

1. **7-day multiples:** Cadences that must land on the same weekday use multiples of 7 (7, 14, 28, 56, **91**, **182**, …). This limits calendar drift (e.g. Wednesday staying Wednesday).
2. **91 / 182 symmetry:** A fiscal quarter = **13 weeks = 91 days** (4‑5‑4 weeks). **In-quarter:** **3 rows**, first week of each fiscal month, usually `order_frequency` **91**. **Quarterly across year:** **1 row**, **91**. **Semi-annual across year:** **1 row**, **182** (two fiscal months per 91-day unit). See [Fiscal cadence patterns](#fiscal-cadence-patterns-planner-intent).
3. **Semi-annual symmetry (182 days):** A half-year is **26 weeks (182 days)**. “Every two fiscal months” is often **three** lines with frequency **182**, spaced across fiscal months (stride 2 vs consecutive months for 91).
4. **8-9-9 style gaps:** In two-month fiscal cycles, gaps between events are not always identical because fiscal months can be 4 or 5 weeks; an **8–9–9 week** style cadence over the half-year is normal.

**Agent / user tasks**

- “Every month (fiscal)” → suggest **3** start dates with frequency **91** (verify weekday consistency).
- “Every two months (fiscal)” → suggest **3** start dates with frequency **182** (months 1, 3, 5 style spacing in fiscal order).
- Always check **day-of-week** on provided dates against the stated anchor (e.g. Wednesday in week 3).
- If the ticket or Excel already has **three** correct fiscal lines for a pair, **do not re-expand**; the toolkit keeps them and only warns on mixed weekdays.

### Fiscal months vs civil dates (454)

Stakeholders often assume **“last Wednesday of every month”** means **civil** May / June / July. In **4‑5‑4 (and variants)**, “month” means **`Fiscal_Year_Month`** in `fiscal_cal`, not the wall calendar. A single fiscal month can **span two civil months** (e.g. a 5‑week fiscal month may run from late May into early July).

**How to read three anchors with `order_frequency` 91**

- **91** = one **fiscal quarter** in weeks (13 weeks = 3 fiscal months × 4/5 weeks).
- A common quarterly pattern is **three order dates**, one per **consecutive** `Fiscal_Year_Month`, each on the **same weekday** (e.g. Wednesday), often **the last occurrence of that weekday inside that fiscal month’s date range** — **not** the last Wednesday of civil March/April/May.
- **Every fiscal quarter, first week of the month** (TBRCS-693 MSP): **first** anchor weekday in the fiscal month, then **every 3 fiscal months** — civil gaps **91 days** ([Every fiscal quarter, first week](#every-fiscal-quarter-first-week-of-the-month)). This is **not** three consecutive months from one `91` seed.
- **Inside one quarter:** three consecutive fiscal months, first week each (TBRCS-693 TMW) — may have **uneven** civil gaps (35 then 28); preserve when submitted, do not force 91-day spacing.
- Within a quarter, fiscal months alternate **4-week / 5-week / 4-week** (depending on where the 5-week month sits in the year). Example (FY2026 windows from `fiscal_cal`): **M202604** (4 weeks) → **M202605** (5 weeks) → **M202606** (4 weeks), with **last Wednesday** in each window on **5/27**, **7/1**, **7/29** respectively — **7/1** is still “last Wed of **that** fiscal month” even though civil June’s last Wednesday would be different.

**Do not** validate these rows by **equal Gregorian spacing** (e.g. “every 28 days”). Gaps between anchors **will** look uneven on a civil calendar even when fiscal logic is correct.

**`order_review_calendar`** strings like `OrdReview_28D_*` refer to **review** cadence naming in prod; they are **not** a claim that **order** dates are 28 days apart when **`order_frequency` is 91**.

**Verification (agent / analyst)**

- Source: datastore **`…/one_time_uploads/fiscal_cal`** (same schema the merge step reads); columns include at least **`date`**, **`Fiscal_Year_Month`**, **`Week`** (see `order_sch_main.generate_fiscal_calendar_schedule`).
- For each candidate `order_schedule_date`, confirm its **`Fiscal_Year_Month`**, count **distinct `Week`** values in that month → **4 or 5** week month.
- For **last-Wed (or last anchor weekday) in fiscal month** patterns: filter `fiscal_cal` to that `Fiscal_Year_Month` and anchor weekday; take **max** `date`.
- For **first week of fiscal month** patterns (notes like “first week of the month”): same filter; take **min** `date` (first anchor weekday in that month).
- Ask which pattern applies if notes and dates are ambiguous; do not assume last-Wed when notes say first week.

---

## Pitfalls and verification

The tooling encodes most of this; still verify `pipeline_meta` and logs.

| Pitfall | Impact | Mitigation |
|---------|--------|------------|
| Fiscal cadence in Excel (91, 182, …) treated as non-fiscal because Jira text is silent | Earliest-per-pair collapses fiscal rows | Multiples of 91 in `order_frequency` force fiscal unless explicit non-fiscal text; read `is_fiscal` and `rationale` |
| Assuming Monday anchor | Wrong fiscal dates vs workbook (e.g. Wednesday) | Check `fiscal_anchor_weekday`; inference from fiscal-cadence row dates |
| Emitting **four** prod rows for uniform **91-day** quarterly pattern | Over-specifies; one anchor + `order_frequency` 91 is enough | **One seed row** only ([Uniform date gaps](#uniform-date-gaps-fiscal)) |
| Collapsing TMW **three uneven-gap** rows to one seed | Loses 4-5-4 in-quarter months | Keep **≥3** rows when gaps are **not** uniform |
| Expanding uniform **91-day** gaps to multiple prod rows | Same as above | Seed only; no calendar multi-expand |
| Expecting HTML report to work inside Slack/Jira | Previews often disable JS | Download file → open in Chrome, Edge, or Firefox |
| Generating HTML with no network | Tabulator embed fails | Run merge with HTTPS, or rely on `.txt` report only |
| Two issues in a row | `output/final/...` is only the last run | Archive or upload after each key if both matter |
| Trusting only Excel; screenshot has destinations template lacks | **0 rows** or wrong pairs | [All attachments](#all-attachments-excel-screenshots-inline-media): read images; `--manual-csv` or correct workbook; `--attachment` to select file |
| Title row above Excel headers (`JAB`, `Unnamed: …`) | `ValueError` missing required columns | Toolkit tries header rows 0–2 automatically; if still failing, fix workbook layout or `--manual-csv` ([Excel layout](#excel-layout-title-row--header-detection)) |
| **`prod_row_delta` is 0** but ticket changed schedules | Misread as “no merge” | Check `pairs_applied` and merge report—in-place replacements ([Interpreting merge deltas](#interpreting-merge-deltas)) |
| Judging 91-day fiscal dates by **civil** “28 days apart” or civil “last Wed of month” | False QA failures; wrong stakeholder explanation | Use **`Fiscal_Year_Month`** and **last or first** anchor weekday in period per Planner notes ([Fiscal months vs civil dates](#fiscal-months-vs-civil-dates-454)) |
| Validating fiscal rows as “must be exactly three” | Rejecting valid four-quarter submissions | **≥3** same-frequency rows are kept as-is; four rows is normal for annual fiscal coverage (TBRCS-693) |
| Ignoring **`Planner notes`**; trusting Jira “review calendar only” | Misread anchor rule (first week vs last Wed) | Read notes in raw Excel ([Planner notes](#planner-notes-excel)); verify in `fiscal_cal` |
| **`prod_row_delta` > 0** assumed duplicate/error | Good merges that add fiscal lines per pair | Check `pairs_applied` and per-pair before/after counts ([TBRCS-693](#excel-layout-title-row--header-detection)) |
| Staggered dates across **different** order groups but stakeholders want **one** shared start date | Per-pair earliest leaves different pairs on different dates | Normalize all rows to that date in Excel or `--manual-csv` ([TBRCS-687](#excel-layout-title-row--header-detection)) |
| **`order_group_description` typo** merged to prod (`BUE` vs `BLUE`) | Two keys; corrected submission does not delete the typo row | Remove orphan row from upload CSV or prod before publish; fix spelling before merge ([TBRCS-687](#excel-layout-title-row--header-detection)) |
| Copy-pasted template workbook names the **wrong order group** (doesn't match ticket title), and/or has free-text `order_frequency`/stacked dates in one cell | `jira_input.py` parses **0** rows, or (if not caught) uploads under the **wrong** group name | Don't trust the workbook literally: confirm the real group name with the user; when the ask is "same calendar as existing group X", `grep` X's actual prod rows as the source of truth and build `--manual-csv` from that shape ([TBRCS-753](#excel-layout-title-row--header-detection)) |
| Judging a "last `<weekday>` of the month" date against the **civil** calendar (e.g. "August's last Monday is the 31st, not the 24th") | False QA failures — a 4-week fiscal month can end mid-civil-month, handing the last few civil days to the next (5-week) fiscal month | Look up the actual `Fiscal_Year_Month` window for the candidate date in `fiscal_cal` (min/max `date` for that `Fiscal_Year_Month`) and take the **max** Monday (or anchor weekday) **inside that window** — do not eyeball the wall calendar ([Fiscal months vs civil dates](#fiscal-months-vs-civil-dates-454)) |
| Azure CLI token expired mid-run (`order_sch_main.py` fails with `DefaultAzureCredential`/`AADSTS50173`) | Datastore read blocked; merge step can't run | Not a toolkit bug — ask the user to `az login --tenant "<tenant-id>" --scope "https://storage.azure.com/.default"` (tenant id is in the error output), then re-run just `order_sch_main.py` (no need to redo `jira_input`) |

Explicit non-fiscal wording in Jira still overrides Excel fiscal cadence for the flag; fix the ticket text or use `--fiscal` if that override is wrong.

---

## Troubleshooting

| Symptom | Check |
|---------|--------|
| 401 / 403 on Jira | `JIRA_BASE_URL`, site for the issue, email matches token owner, token scope |
| No REST in terminal | `secrets.local.env` + `local.settings.json`, or `--excel-path` |
| DataWorkbench / datastore errors | `az login` done (Azure CLI), `invent-forge-anvil` installed in `local_forge_venv`, private-index/VPN reachable if (re)installing |
| `can't open file '.../scripts/run_pipeline.py'` | Wrong working directory | `cd` to `order-schedule-toolkit` root ([Run commands](#run-commands-toolkit-root)) |
| `Submission missing required columns` / columns like `Unnamed:` | Title row or wrong sheet | [Excel layout](#excel-layout-title-row--header-detection); verify sheet and header row |
