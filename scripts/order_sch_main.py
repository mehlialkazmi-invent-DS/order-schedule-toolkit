#!/usr/bin/env python3
"""
Order schedule — main pipeline (script form of order_sch_main.ipynb).

Prereq: run `python jira_input.py <ISSUE-KEY>` first (cleaned CSV + pipeline_meta.json).

Reads the datastore through Forge Anvil's DataWorkbench (Azure, customer_name="tbretail") —
no Spark session or customer-pipeline checkout needed. Run under local_forge_venv, from toolkit
root:
  python scripts/order_sch_main.py
Or with overrides:
  python scripts/order_sch_main.py --issue-key PROJ-456

Reads datastore only; writes local_backup/, output/ — never the datastore.

After building `new_for_merge` (non-fiscal: earliest date per pair; fiscal: if a pair already has
≥3 rows with the same fiscal cadence frequency (91, 182, …), those Excel rows are kept; otherwise
calendar expansion from the earliest submission date, stride by frequency/91, weekday inferred
from submission then `pipeline_meta`),
**overwrites** `jira_downloads/cleaned/{KEY}_cleaned.csv` with those merge-input rows.
Also writes **`{KEY}_cleaned_applied.csv`**: only submission rows whose pairs **changed** prod
(skipped duplicates are excluded). Skipped pairs log **already exists in prod — not doing anything**.
Writes **`{KEY}_merge_report.txt`** and **`{KEY}_merge_report.html`** (interactive tables: filter, sort, paginate; HTML inlines Tabulator at generation so the file is portable/offline — open in a desktop browser; Slack/Jira previews usually skip JS).

Merge rule: pair = (order_group_description, destination_code); blank destination is a real value
(not imputed). Per pair, replace only if the schedule multiset differs; same lines → skip. If the
submission uses only blank destination for an order_group, compare to all prod rows for that
order_group and replace all of them when the combined schedule differs.

Non-fiscal: one row per pair (earliest date).
Fiscal: **≥3** rows with **non-uniform** gaps → keep submission (4-5-4 within a quarter, e.g. TMW).
**Uniform 91-day** (or multiple-of-91 quarterly) gaps → **one seed row only** (first Wed each fiscal
quarter is derived downstream; do not emit four rows). **Uniform 28-day** gaps → earliest row then
expand to **3** fiscal months (stride 1 / frequency 91). Else default fiscal expansion from seed.
"""

from __future__ import annotations

import argparse
import calendar
import html as html_module
import io
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
TOOLKIT_ROOT = HERE.parent

# Defaults (override with CLI flags)
DEFAULT_JIRA_ISSUE_KEY = "PROJ-123"

CUSTOMER_NAME = "tbretail"
DATASTORE_CONTAINER = "invent-tbretail-datastore"
ORDER_DATA_SOURCE = f"{DATASTORE_CONTAINER}/one_time_uploads/order_schedule_input_prod/order_schedule_input_prod.csv"
FISCAL_CAL_SOURCE = f"{DATASTORE_CONTAINER}/one_time_uploads/fiscal_cal"

BACKUP_DIR = TOOLKIT_ROOT / "local_backup"

ORDER_COLUMNS = [
    "order_group_description",
    "destination_code",
    "order_schedule_date",
    "order_frequency",
    "order_review_calendar",
]


def _backup_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _path_rel_toolkit(path: Path, root: Path) -> str:
    try:
        return str(path.resolve().relative_to(root.resolve()))
    except ValueError:
        return str(path)


def _df_records_for_json(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    out = df.replace({np.nan: None})
    return out.to_dict(orient="records")


def _json_for_script_embed(obj) -> str:
    s = json.dumps(obj, default=str, ensure_ascii=True)
    return s.replace("</script>", "<\\/script>")


TABULATOR_VERSION = "6.2.5"
TABULATOR_CSS_URL = f"https://unpkg.com/tabulator-tables@{TABULATOR_VERSION}/dist/css/tabulator.min.css"
TABULATOR_JS_URL = f"https://unpkg.com/tabulator-tables@{TABULATOR_VERSION}/dist/js/tabulator.min.js"


def _fetch_tabulator_for_embed(*, timeout_s: int = 120) -> tuple[str, str]:
    """Download Tabulator minified CSS+JS once at report generation; embedded in HTML for offline sharing."""
    req = urllib.request.Request(
        TABULATOR_CSS_URL,
        headers={"User-Agent": "order-schedule-toolkit-merge-report/1.0"},
    )
    with urllib.request.urlopen(req, timeout=timeout_s) as r:
        tab_css = r.read().decode("utf-8")
    req = urllib.request.Request(
        TABULATOR_JS_URL,
        headers={"User-Agent": "order-schedule-toolkit-merge-report/1.0"},
    )
    with urllib.request.urlopen(req, timeout=timeout_s) as r:
        tab_js = r.read().decode("utf-8")
    return tab_css, tab_js


_MERGE_REPORT_PAGE_CSS = """
    :root { font-family: system-ui, Segoe UI, Roboto, sans-serif; background: #f4f5f7; color: #172b4d; }
    body { max-width: 1400px; margin: 0 auto; padding: 1.25rem; }
    h1 { font-size: 1.35rem; margin: 0 0 0.25rem; }
    .sub { color: #5e6c84; font-size: 0.9rem; margin-bottom: 0.35rem; }
    .share-hint { background: #fffae6; border: 1px solid #ffe380; border-radius: 6px; padding: 0.65rem 0.85rem; font-size: 0.85rem; margin-bottom: 1rem; color: #4a4113; }
    .cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 0.75rem; margin: 1rem 0; }
    .card { background: #fff; border-radius: 8px; padding: 0.85rem; box-shadow: 0 1px 2px rgba(0,0,0,.06); }
    .card .label { font-size: 0.72rem; text-transform: uppercase; color: #5e6c84; letter-spacing: .04em; }
    .card .val { font-size: 1.35rem; font-weight: 600; margin-top: 0.2rem; }
    .card .val.delta-pos { color: #0b6e4f; }
    .card .val.delta-neg { color: #bf2600; }
    .paths { background: #fff; border-radius: 8px; padding: 1rem; margin: 1rem 0; font-size: 0.85rem; line-height: 1.5; }
    .paths code { background: #f0f1f3; padding: 0.1rem 0.35rem; border-radius: 4px; }
    .tabs { display: flex; gap: 0.35rem; flex-wrap: wrap; margin: 1rem 0 0.5rem; }
    .tabs button {
      border: none; background: #dfe1e6; padding: 0.5rem 1rem; border-radius: 6px 6px 0 0;
      cursor: pointer; font-weight: 500; color: #42526e;
    }
    .tabs button.active { background: #fff; color: #0747a6; box-shadow: 0 -1px 0 #fff; }
    .tabs button:hover:not(.active) { background: #c1c7d0; }
    .panel { display: none; background: #fff; border-radius: 0 8px 8px 8px; padding: 1rem; box-shadow: 0 1px 3px rgba(0,0,0,.08); margin-bottom: 1.5rem; }
    .panel.active { display: block; }
    .panel h3 { margin: 0 0 0.75rem; font-size: 1rem; }
    .hint { font-size: 0.8rem; color: #5e6c84; margin-bottom: 0.5rem; }
    .tabulator { font-size: 0.8rem; }
"""


def write_merge_report(
    *,
    toolkit_root: Path,
    issue_key: str,
    prod_row_count_before: int,
    prod_row_count_after: int,
    ticket_full_input: pd.DataFrame,
    merge_ready: pd.DataFrame,
    applied: pd.DataFrame,
    prod_backup_path: Path,
    final_output_path: Path,
    cleaned_input_path: Path,
    cleaned_applied_path: Path,
    datastore_prod_path: str,
    pairs_applied: int,
    pairs_skipped: int,
    source_attachments: list[str] | None = None,
) -> Path:
    """
    Human-readable log: prod row counts before/after, paths, and embedded CSVs for
    full ticket input, merge input, and applied-only submission rows.
    """
    report = toolkit_root / "jira_downloads" / "cleaned" / f"{issue_key}_merge_report.txt"
    delta = prod_row_count_after - prod_row_count_before
    buf_ticket = io.StringIO()
    ticket_full_input.to_csv(buf_ticket, index=False)
    buf_in = io.StringIO()
    merge_ready.to_csv(buf_in, index=False)
    buf_ap = io.StringIO()
    applied.to_csv(buf_ap, index=False)
    rel = lambda p: _path_rel_toolkit(p, toolkit_root)
    lines = [
        f"issue_key={issue_key}",
        f"generated_utc={datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}",
        "",
    ]
    if source_attachments:
        lines.append("=== Jira Excel source file(s) (combined into cleaned CSV before this run) ===")
        for fn in source_attachments:
            lines.append(f"  - {fn}")
        lines.append("")
    lines.extend(
        [
            "=== Final combined file (local order_schedule_input_prod) ===",
            f"Row count BEFORE this merge (prod read from datastore): {prod_row_count_before:,}",
            f"Row count AFTER this merge (written to final output below): {prod_row_count_after:,}",
            f"Delta (after minus before): {delta:+,}",
            f"Datastore source (read-only): {datastore_prod_path}",
            f"Prod snapshot backup (before merge): {rel(prod_backup_path)}",
            f"Final output CSV: {rel(final_output_path)}",
            "",
            "=== Submission: FULL ticket input (all rows from Jira — all .xlsx combined by jira_input) ===",
            "Before earliest-per-pair / fiscal shrink.",
            f"Row count: {len(ticket_full_input)}",
            "",
            "--- CSV (full ticket) ---",
            buf_ticket.getvalue().rstrip(),
            "",
            "=== Submission: merge INPUT (after earliest-per-pair / fiscal) ===",
            f"File: {rel(cleaned_input_path)}",
            f"Row count: {len(merge_ready)}",
            "",
            "--- CSV (merge input) ---",
            buf_in.getvalue().rstrip(),
            "",
            "=== Submission: ADDED to prod (applied only; duplicates skipped separately) ===",
            f"File: {rel(cleaned_applied_path)}",
            f"Pairs applied (changed prod): {pairs_applied}",
            f"Pairs skipped (already in prod — not doing anything): {pairs_skipped}",
            f"Row count in applied file: {len(applied)}",
            "",
            "--- CSV (applied rows only) ---",
            buf_ap.getvalue().rstrip(),
            "",
        ]
    )
    report.write_text("\n".join(lines), encoding="utf-8")
    return report


_MERGE_REPORT_APP_JS = """
    const payload = JSON.parse(document.getElementById("merge-payload").textContent);
    const cols = payload.columns.map(f => ({ title: f, field: f, headerFilter: "input", headerFilterPlaceholder: "filter…", minWidth: 90 }));

    function tabOpts(data) {
      return {
        data,
        columns: cols,
        layout: "fitColumns",
        pagination: "local",
        paginationSize: 25,
        paginationSizeSelector: [10, 25, 50, 100, true],
        movableColumns: true,
        height: "56vh",
      };
    }

    new Tabulator("#tbl-ticket", tabOpts(payload.ticket));
    new Tabulator("#tbl-merge", tabOpts(payload.merge));
    new Tabulator("#tbl-applied", tabOpts(payload.applied));

    document.querySelectorAll(".tabs button").forEach(btn => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".tabs button").forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
        btn.classList.add("active");
        document.getElementById(btn.dataset.tab).classList.add("active");
      });
    });
"""


def write_merge_report_html(
    *,
    toolkit_root: Path,
    issue_key: str,
    prod_row_count_before: int,
    prod_row_count_after: int,
    ticket_full_input: pd.DataFrame,
    merge_ready: pd.DataFrame,
    applied: pd.DataFrame,
    prod_backup_path: Path,
    final_output_path: Path,
    cleaned_input_path: Path,
    cleaned_applied_path: Path,
    datastore_prod_path: str,
    pairs_applied: int,
    pairs_skipped: int,
    source_attachments: list[str] | None = None,
) -> Path:
    """
    Single-file interactive HTML: Tabulator CSS+JS are fetched at generation time and inlined so the
    attachment works offline after download. Generating this file needs outbound HTTPS once; viewing does not.

    Slack/Jira inline previews typically do not run JavaScript — recipients should download the HTML and open
    it in Chrome, Edge, or Firefox for filters, sort, and tabs.
    """
    try:
        tab_css, tab_js = _fetch_tabulator_for_embed()
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise RuntimeError(
            "Could not download Tabulator for merge report HTML (need outbound HTTPS to unpkg.com). "
            "Retry when online, or run from a machine that can reach the internet."
        ) from e
    tab_js = tab_js.replace("</script>", "<\\/script>")

    out_path = toolkit_root / "jira_downloads" / "cleaned" / f"{issue_key}_merge_report.html"
    delta = prod_row_count_after - prod_row_count_before
    rel = lambda p: html_module.escape(_path_rel_toolkit(p, toolkit_root))
    esc = html_module.escape

    payload = {
        "ticket": _df_records_for_json(ticket_full_input),
        "merge": _df_records_for_json(merge_ready),
        "applied": _df_records_for_json(applied),
        "columns": list(ORDER_COLUMNS),
    }
    payload_json = _json_for_script_embed(payload)

    att_ul = ""
    if source_attachments:
        items = "".join(f"<li>{esc(fn)}</li>" for fn in source_attachments)
        att_ul = f"<h2>Excel source file(s) on ticket</h2><ul>{items}</ul>"

    title = f"Order schedule merge — {issue_key}"
    gen_time = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    delta_cls = "delta-pos" if delta >= 0 else "delta-neg"

    body = f"""  <h1>{esc(title)}</h1>
  <p class="sub">Generated {esc(gen_time)} · Interactive tables: type in column headers to filter, click headers to sort.</p>
  <p class="share-hint"><strong>Sharing:</strong> Download this file and open it in a desktop browser. Slack and Jira previews usually do not run JavaScript, so tables will not be interactive there.</p>

  <div class="cards">
    <div class="card"><div class="label">Prod rows before</div><div class="val">{prod_row_count_before:,}</div></div>
    <div class="card"><div class="label">Prod rows after</div><div class="val">{prod_row_count_after:,}</div></div>
    <div class="card"><div class="label">Delta</div><div class="val {delta_cls}">{delta:+,}</div></div>
    <div class="card"><div class="label">Pairs applied</div><div class="val">{pairs_applied}</div></div>
    <div class="card"><div class="label">Pairs skipped (duplicate)</div><div class="val">{pairs_skipped}</div></div>
    <div class="card"><div class="label">Full ticket rows</div><div class="val">{len(ticket_full_input)}</div></div>
    <div class="card"><div class="label">Merge-input rows</div><div class="val">{len(merge_ready)}</div></div>
    <div class="card"><div class="label">Applied rows</div><div class="val">{len(applied)}</div></div>
  </div>

  {att_ul}

  <div class="paths">
    <strong>Datastore (read-only):</strong> <code>{esc(datastore_prod_path)}</code><br/>
    <strong>Prod backup:</strong> <code>{rel(prod_backup_path)}</code><br/>
    <strong>Final output:</strong> <code>{rel(final_output_path)}</code><br/>
    <strong>Merge-input CSV:</strong> <code>{rel(cleaned_input_path)}</code><br/>
    <strong>Applied CSV:</strong> <code>{rel(cleaned_applied_path)}</code>
  </div>

  <div class="tabs" role="tablist">
    <button type="button" class="active" data-tab="p-ticket">Full ticket input</button>
    <button type="button" data-tab="p-merge">After earliest / fiscal</button>
    <button type="button" data-tab="p-applied">Applied to prod</button>
  </div>

  <div id="p-ticket" class="panel active">
    <h3>All rows from Jira (every Excel combined)</h3>
    <p class="hint">Same data as <code>jira_input</code> cleaned output before shrink. Use header filters to search.</p>
    <div id="tbl-ticket"></div>
  </div>
  <div id="p-merge" class="panel">
    <h3>Merge input (one row per pair non-fiscal, or fiscal expansion)</h3>
    <p class="hint">What was compared against prod for skip vs replace.</p>
    <div id="tbl-merge"></div>
  </div>
  <div id="p-applied" class="panel">
    <h3>Rows that actually updated prod</h3>
    <p class="hint">Pairs that matched prod are omitted here.</p>
    <div id="tbl-applied"></div>
  </div>

  <script type="application/json" id="merge-payload">{payload_json}</script>
"""

    html_doc = (
        "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n"
        '  <meta charset="utf-8"/>\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1"/>\n'
        f"  <title>{esc(title)}</title>\n"
        "  <!-- Tabulator embedded at generation time; file works offline once saved. -->\n"
        "  <style>\n"
        + tab_css
        + "\n  </style>\n  <style>\n"
        + _MERGE_REPORT_PAGE_CSS
        + "\n  </style>\n</head>\n<body>\n"
        + body
        + "  <script>\n"
        + tab_js
        + "\n  </script>\n  <script>\n"
        + _MERGE_REPORT_APP_JS
        + "\n  </script>\n</body>\n</html>\n"
    )
    out_path.write_text(html_doc, encoding="utf-8")
    return out_path


def format_schedule_date(ts):
    t = pd.Timestamp(ts)
    return f"{t.month}/{t.day}/{t.year}"


def _weekday_to_int(x):
    if isinstance(x, int):
        if 0 <= x <= 6:
            return x
        raise ValueError("anchor_weekday int must be 0–6 (Mon=0, Sun=6).")
    full = "monday tuesday wednesday thursday friday saturday sunday".split()
    short = "mon tue wed thu fri sat sun".split()
    extra = {"tues": 1, "thur": 3, "thurs": 3}
    k = str(x).strip().lower()
    if k in extra:
        return extra[k]
    if k in short:
        return short.index(k)
    if k in full:
        return full.index(k)
    raise ValueError(f"Unknown anchor_weekday {x!r}; use 0–6 or a weekday name.")


def _is_fiscal_cadence_days(n: int) -> bool:
    return n >= 91 and n % 91 == 0


def generate_fiscal_calendar_schedule(
    calendar_df,
    order_df,
    *,
    frequency_col="order_frequency",
    output_frequency: int = 91,
    n_periods=3,
    anchor_weekday=0,
    date_col="order_schedule_date",
    group_cols=("order_group_description", "destination_code"),
):
    """
    Build ``n_periods`` schedule lines from fiscal calendar. ``output_frequency`` 91 uses consecutive
    fiscal months; 182 uses every second fiscal month (half-year symmetry); stride = freq // 91.
    Anchor is the earliest ``order_schedule_date`` in the group; weekday filter picks the recurring
    weekday (e.g. Wednesday) on/after anchor in each chosen fiscal month.
    """
    date_key = "date" if "date" in calendar_df.columns else "Date"
    fy = "Fiscal_Year_Month"
    if date_key not in calendar_df.columns or fy not in calendar_df.columns:
        raise ValueError(f"Calendar needs columns {date_key!r} and {fy!r}.")

    wd = _weekday_to_int(anchor_weekday)
    wname = calendar.day_name[wd]

    if output_frequency % 91 != 0:
        print(
            f"Warning: fiscal output_frequency={output_frequency} is not a multiple of 91; "
            "using stride 1 (quarterly month picks)."
        )
        stride = 1
    else:
        stride = max(1, output_frequency // 91)

    if output_frequency % 7 != 0:
        print(
            f"Warning: order_frequency={output_frequency} is not a multiple of 7; "
            "fiscal cadence usually uses 7-day multiples to avoid weekday drift."
        )

    if order_df.empty:
        return order_df.copy()
    for c in group_cols:
        if c not in order_df.columns:
            raise ValueError(f"order_df missing column {c!r}")

    cal = calendar_df[[date_key, fy]].copy()
    cal["_d"] = pd.to_datetime(cal[date_key], errors="coerce")
    cal = cal.dropna(subset=["_d"])

    work = order_df.copy()
    work["_a"] = pd.to_datetime(work[date_col], errors="coerce")
    if work["_a"].isna().any():
        raise ValueError(f"Invalid or missing {date_col}")

    built = []
    for _, grp in work.groupby(list(group_cols), dropna=False, sort=False):
        anchor = grp["_a"].min()
        template = grp.loc[grp["_a"].idxmin()].drop("_a")

        m = cal.loc[(cal["_d"].dt.weekday == wd) & (cal["_d"] >= anchor)]
        month_starts = m.groupby(fy, sort=False)["_d"].min().sort_values()
        idxs = list(range(0, len(month_starts), stride))[:n_periods]
        dates = month_starts.iloc[idxs] if len(month_starts) else month_starts

        if len(dates) < n_periods:
            g = {c: grp[c].iloc[0] for c in group_cols}
            print(f"Warning: only {len(dates)} fiscal {wname} anchor(s) for {g} (wanted {n_periods}).")

        for d in dates.to_numpy():
            row = template.copy()
            row[date_col] = format_schedule_date(d)
            row[frequency_col] = output_frequency
            built.append(row)

    return pd.DataFrame(built)


def generate_91_day_fiscal_schedule(
    calendar_df,
    order_df,
    *,
    frequency_col="order_frequency",
    output_frequency=91,
    n_periods=3,
    anchor_weekday=0,
    date_col="order_schedule_date",
    group_cols=("order_group_description", "destination_code"),
):
    """Backward-compatible name; calls ``generate_fiscal_calendar_schedule``."""
    return generate_fiscal_calendar_schedule(
        calendar_df,
        order_df,
        frequency_col=frequency_col,
        output_frequency=output_frequency,
        n_periods=n_periods,
        anchor_weekday=anchor_weekday,
        date_col=date_col,
        group_cols=group_cols,
    )


def _infer_anchor_weekday_from_dates(dates: pd.Series) -> int | None:
    """Mon=0 … Sun=6 from calendar dates. None if nothing parseable."""
    dts = pd.to_datetime(dates, errors="coerce").dropna()
    if dts.empty:
        return None
    wds = dts.dt.weekday
    if wds.nunique() == 1:
        return int(wds.iloc[0])
    return int(wds.mode().iloc[0])


def _warn_if_mixed_fiscal_weekdays(gf: pd.DataFrame) -> None:
    dts = pd.to_datetime(gf["order_schedule_date"], errors="coerce").dropna()
    if dts.empty:
        return
    if dts.dt.weekday.nunique() > 1:
        print(
            "Warning: preserved fiscal rows use different weekdays — verify anchor consistency "
            "(e.g. all Wednesdays)."
        )


def _uniform_civil_gap_days(dates: pd.Series) -> int | None:
    """Return the common consecutive day gap when sorted dates are evenly spaced; else None."""
    dts = pd.to_datetime(dates, errors="coerce").dropna().sort_values()
    if len(dts) < 2:
        return None
    gaps = [(dts.iloc[i + 1] - dts.iloc[i]).days for i in range(len(dts) - 1)]
    if not gaps or len(set(gaps)) != 1:
        return None
    return int(gaps[0])


def _fiscal_uniform_gap_is_quarterly_seed_only(gap_days: int) -> bool:
    """True when uniform spacing means one anchor row (every fiscal quarter), not multi-row expand."""
    return gap_days % 91 == 0


def build_fiscal_merge_from_submission(
    calendar_df: pd.DataFrame,
    order_df: pd.DataFrame,
    *,
    meta_anchor_weekday: int,
    n_periods: int = 3,
) -> pd.DataFrame:
    """
    Per (order_group, destination): if at least ``n_periods`` fiscal-cadence rows share one
    frequency and civil gaps are **not** uniform, keep submission as-is (454 week pattern in-quarter).
    Uniform **91-day** (quarterly) gaps → **one seed row** only. Uniform **28-day** gaps → seed then
    expand to ``n_periods`` consecutive fiscal months. Otherwise expand from earliest row.
    """
    work = order_df[ORDER_COLUMNS].copy()
    work["_d"] = pd.to_datetime(work["order_schedule_date"], errors="coerce")
    work["_dest_key"] = _norm_dest(work["destination_code"])
    parts: list[pd.DataFrame] = []
    n_preserve = 0
    n_expand = 0
    n_seed_only = 0
    for (_, _), grp in work.groupby(["order_group_description", "_dest_key"], dropna=False, sort=False):
        grp = grp.sort_values("_d")
        fq = pd.to_numeric(grp["order_frequency"], errors="coerce")
        m_f = fq.notna() & (fq >= 91) & (fq % 91 == 0)
        g_f = grp.loc[m_f]
        uniform_gap = _uniform_civil_gap_days(g_f["_d"]) if len(g_f) >= 2 else None
        if len(g_f) >= n_periods and uniform_gap is None:
            freqs = g_f["order_frequency"].astype(int)
            if freqs.nunique() == 1:
                _warn_if_mixed_fiscal_weekdays(g_f)
                parts.append(g_f[ORDER_COLUMNS].reset_index(drop=True))
                n_preserve += 1
                continue
        if len(g_f) > 0:
            wk = _infer_anchor_weekday_from_dates(g_f["_d"])
            if wk is None:
                wk = int(meta_anchor_weekday)
            seed = g_f.nsmallest(1, "_d")[ORDER_COLUMNS].copy()
            out_freq = int(pd.to_numeric(seed["order_frequency"].iloc[0], errors="coerce"))
            if not _is_fiscal_cadence_days(out_freq):
                out_freq = 91
            if len(g_f) >= 2 and uniform_gap is not None and _fiscal_uniform_gap_is_quarterly_seed_only(
                uniform_gap
            ):
                parts.append(seed.reset_index(drop=True))
                n_seed_only += 1
                continue
            expand_periods = n_periods
            if len(g_f) >= 2 and uniform_gap == 28:
                out_freq = 91
                expand_periods = n_periods
        else:
            wk = _infer_anchor_weekday_from_dates(grp["_d"])
            if wk is None:
                wk = int(meta_anchor_weekday)
            seed = grp.nsmallest(1, "_d")[ORDER_COLUMNS].copy()
            out_freq = int(pd.to_numeric(seed["order_frequency"].iloc[0], errors="coerce"))
            if not _is_fiscal_cadence_days(out_freq):
                out_freq = 91
            expand_periods = n_periods
        exp = generate_fiscal_calendar_schedule(
            calendar_df,
            seed,
            anchor_weekday=wk,
            n_periods=expand_periods,
            output_frequency=out_freq,
        )
        parts.append(exp)
        n_expand += 1

    if not parts:
        return pd.DataFrame(columns=ORDER_COLUMNS)
    if n_seed_only:
        print(
            f"Fiscal: one seed row only for {n_seed_only} pair(s) "
            "(uniform quarterly / 91-day gaps — no multi-row expand)."
        )
    if n_preserve and n_expand:
        print(
            f"Fiscal: kept submission fiscal-cadence rows for {n_preserve} pair(s) (≥{n_periods} per pair); "
            f"calendar-expanded {n_expand} pair(s)."
        )
    elif n_preserve and not n_expand and not n_seed_only:
        print(
            f"Fiscal: kept submission fiscal rows as-is for {n_preserve} pair(s) "
            f"(≥{n_periods} rows, non-uniform gaps, same frequency)."
        )
    elif n_preserve:
        print(
            f"Fiscal: kept submission as-is for {n_preserve} pair(s); "
            f"expanded or seed-only for other pair(s)."
        )
    else:
        print(
            f"Fiscal: calendar expansion for {n_expand} pair(s); "
            "weekday from submission when possible; frequency 91/182/… from seed row or 91 default."
        )
    return pd.concat(parts, ignore_index=True)


def get_earliest_schedule_per_group(order_df):
    df = order_df.copy()
    df["order_schedule_date"] = pd.to_datetime(df["order_schedule_date"], errors="coerce")
    idx = (
        df.groupby("order_group_description")["order_schedule_date"]
        .idxmin()
        .dropna()
        .astype(int)
    )
    return df.loc[idx].reset_index(drop=True)


def expand_destination_codes(order_df, destination_codes):
    if not destination_codes or len(destination_codes) == 0:
        return order_df
    columns = list(order_df.columns)
    if "destination_code" not in columns:
        columns.append("destination_code")
    df = order_df.copy()
    expanded_rows = []
    for _, row in df.iterrows():
        for dest in destination_codes:
            new_row = row.copy()
            new_row["destination_code"] = dest
            expanded_rows.append(new_row)
    return pd.DataFrame(expanded_rows, columns=columns)


def _norm_dest(s: pd.Series) -> pd.Series:
    out = s.astype(object)
    mask = out.isna() | (out.astype(str).str.strip() == "")
    out = out.astype(str)
    out.loc[mask] = "__NO_DEST__"
    return out


def get_earliest_schedule_per_pair(order_df: pd.DataFrame) -> pd.DataFrame:
    """
    One row per (order_group_description, destination_code): keep the row with the earliest
    order_schedule_date. Uses the same destination sentinel as merge (`_norm_dest`).

    For **non-fiscal** runs: Excel often lists multiple future dates spaced by order_frequency;
    prod expects one anchor date per pair.
    """
    df = order_df.copy()
    df["_dest_key"] = _norm_dest(df["destination_code"])
    df["_d"] = pd.to_datetime(df["order_schedule_date"], errors="coerce")
    if df["_d"].isna().any():
        raise ValueError("order_schedule_date has values that cannot be parsed as dates")
    idx = (
        df.groupby(["order_group_description", "_dest_key"], dropna=False)["_d"]
        .idxmin()
        .dropna()
        .astype(int)
    )
    out = df.loc[idx].drop(columns=["_dest_key", "_d"])
    return out.reset_index(drop=True)


def _normalized_schedule_comparison_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize ORDER_COLUMNS for multiset equality (row order ignored)."""
    out = df[list(ORDER_COLUMNS)].copy()
    out["order_schedule_date"] = pd.to_datetime(out["order_schedule_date"], errors="coerce")
    out["order_frequency"] = pd.to_numeric(out["order_frequency"], errors="coerce").astype(int)
    for col in ("order_review_calendar", "order_group_description", "destination_code"):
        out[col] = out[col].map(
            lambda x: "" if (pd.isna(x) or str(x).strip() == "") else str(x).strip()
        )
    return out.sort_values(by=list(ORDER_COLUMNS), kind="mergesort").reset_index(drop=True)


def _schedules_equal(old_sub: pd.DataFrame, new_sub: pd.DataFrame) -> bool:
    """True if prod and submission rows match for this pair (same multiset of schedule lines)."""
    if len(old_sub) != len(new_sub):
        return False
    if old_sub.empty:
        return True
    o = _normalized_schedule_comparison_frame(old_sub)
    n = _normalized_schedule_comparison_frame(new_sub)
    if not o["order_schedule_date"].dt.normalize().equals(n["order_schedule_date"].dt.normalize()):
        return False
    if not o["order_frequency"].equals(n["order_frequency"]):
        return False
    if not (o["order_review_calendar"] == n["order_review_calendar"]).all():
        return False
    if not (o["order_group_description"] == n["order_group_description"]).all():
        return False
    if not (o["destination_code"] == n["destination_code"]).all():
        return False
    return True


def merge_with_pair_updates(
    old: pd.DataFrame, new_rows: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, list[tuple[str, str]], int]:
    """
    Pair key is (order_group_description, destination_code) with blank/NA destinations normalized
    to one internal sentinel (not imputed to a fake code).

    Per (order_group, destination): if prod already has the same multiset of schedule lines, skip.

    If the submission has only blank-destination rows for an order_group (no row with a non-blank
    destination for that group), compare against all prod rows for that order_group (any destination).
    When schedules differ, replace every prod row for that order_group with the submission rows.
    Otherwise normal per-pair replace.

    Returns:
        (merged, applied submission rows, skipped labels, number of pair- or group-level applies)
    """
    empty_applied = pd.DataFrame(columns=ORDER_COLUMNS)

    old = old.copy()
    new_rows = new_rows.copy()

    if new_rows.empty:
        print("merge: no rows in Jira submission; output unchanged from prod.")
        return old.drop(columns=["_dest_key"], errors="ignore"), empty_applied, [], 0

    old["_dest_key"] = _norm_dest(old["destination_code"])
    new_rows["_dest_key"] = _norm_dest(new_rows["destination_code"])
    new_rows["_og_s"] = new_rows["order_group_description"].astype(str).str.strip()

    nonempty_ogd = set(
        new_rows.loc[new_rows["_dest_key"] != "__NO_DEST__", "_og_s"].unique()
    )
    ogd_empty_only: set[str] = set()
    for og, g in new_rows.groupby("_og_s", sort=False):
        if og in nonempty_ogd:
            continue
        if len(g) > 0 and (g["_dest_key"] == "__NO_DEST__").all():
            ogd_empty_only.add(og)

    skipped_ogd: list[tuple[str, str]] = []
    apply_ogd: set[str] = set()
    for og in sorted(ogd_empty_only):
        m_new = new_rows["_og_s"] == og
        new_cmp = new_rows.loc[m_new, ORDER_COLUMNS]
        m_old = old["order_group_description"].astype(str).str.strip() == og
        old_cmp = old.loc[m_old, ORDER_COLUMNS]
        if _schedules_equal(old_cmp, new_cmp):
            skipped_ogd.append((og, "(order_group, all destinations)"))
        else:
            apply_ogd.add(og)

    pairs = new_rows[["order_group_description", "_dest_key"]].drop_duplicates()
    replace_keys: list[tuple[str, str]] = []
    skipped_pairs: list[tuple[str, str]] = []

    for _, pr in pairs.iterrows():
        ogd_raw, dk = pr["order_group_description"], pr["_dest_key"]
        ogd_s = str(ogd_raw).strip()
        if ogd_s in ogd_empty_only:
            continue
        m_old = (old["_dest_key"] == dk) & (old["order_group_description"].astype(str).str.strip() == ogd_s)
        m_new = (new_rows["_dest_key"] == dk) & (
            new_rows["order_group_description"].astype(str).str.strip() == ogd_s
        )
        old_sub = old.loc[m_old, ORDER_COLUMNS]
        new_sub = new_rows.loc[m_new, ORDER_COLUMNS]
        if _schedules_equal(old_sub, new_sub):
            if len(new_sub) > 0:
                dc = new_sub["destination_code"].iloc[0]
                if pd.isna(dc) or str(dc).strip() == "":
                    dest_disp = "(empty)"
                else:
                    dest_disp = str(dc).strip()
            else:
                dest_disp = "(empty)"
            skipped_pairs.append((ogd_s, dest_disp))
        else:
            replace_keys.append((ogd_s, dk))

    replace_set = set(replace_keys)
    n_pairs = len(pairs)

    for ogd_s, dest_disp in skipped_pairs + skipped_ogd:
        print(
            f"  skip: order_group={ogd_s!r} destination={dest_disp!r}: "
            "already exists in prod — not doing anything."
        )

    n_apply = len(apply_ogd) + len(replace_set)
    if (skipped_pairs or skipped_ogd) and not replace_set and not apply_ogd:
        print(f"merge: no changes applied — all {len(skipped_pairs) + len(skipped_ogd)} group/pair(s) already exist in prod.")
    elif (skipped_pairs or skipped_ogd) and n_apply:
        print(
            f"merge: applying {n_apply} group/pair update(s); "
            f"{len(skipped_pairs) + len(skipped_ogd)} skipped (duplicate of prod)."
        )
    elif n_apply:
        print(f"merge: applying {n_apply} group/pair update(s) (new or changed vs prod).")
    elif n_pairs == 0:
        print("merge: no rows in Jira submission; output unchanged from prod.")

    old["_og_s"] = old["order_group_description"].astype(str).str.strip()
    old["_pair_key"] = list(zip(old["order_group_description"].astype(str).str.strip(), old["_dest_key"]))
    drop_old = old["_og_s"].isin(apply_ogd) | old["_pair_key"].isin(replace_set)
    kept = old.loc[~drop_old].drop(columns=["_dest_key", "_pair_key", "_og_s"], errors="ignore")

    new_rows["_pair_key"] = list(
        zip(new_rows["order_group_description"].astype(str).str.strip(), new_rows["_dest_key"])
    )
    applied_parts: list[pd.DataFrame] = []
    for og in sorted(apply_ogd):
        m = new_rows["_og_s"] == og
        applied_parts.append(
            new_rows.loc[m].drop(columns=["_dest_key", "_pair_key", "_og_s"], errors="ignore")
        )
    pair_applied = new_rows.loc[new_rows["_pair_key"].isin(replace_set)].drop(
        columns=["_dest_key", "_pair_key", "_og_s"], errors="ignore"
    )
    if len(pair_applied):
        applied_parts.append(pair_applied)

    if applied_parts:
        new_applied = pd.concat(applied_parts, ignore_index=True).reindex(columns=ORDER_COLUMNS)
    else:
        new_applied = empty_applied

    merged = pd.concat([kept, new_applied], ignore_index=True)
    applied_only = new_applied.copy() if len(new_applied) else empty_applied
    skipped_all = skipped_pairs + skipped_ogd
    return merged, applied_only, skipped_all, n_apply


def parse_args() -> argparse.Namespace:
    from toolkit_env import bootstrap_toolkit_env

    bootstrap_toolkit_env()

    key_env = os.environ.get("DEFAULT_ISSUE_KEY", "").strip()
    default_key = key_env.upper() if key_env else DEFAULT_JIRA_ISSUE_KEY

    p = argparse.ArgumentParser(description="Merge Jira-cleaned order schedule with datastore prod file.")
    p.add_argument(
        "--issue-key",
        default=default_key,
        help="Jira key; must match jira_input.py output files (default: DEFAULT_ISSUE_KEY env / config)",
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()
    jira_issue_key: str = str(args.issue_key).upper()

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    cleaned_jira_csv = TOOLKIT_ROOT / "jira_downloads" / "cleaned" / f"{jira_issue_key}_cleaned.csv"
    pipeline_meta_json = TOOLKIT_ROOT / "jira_downloads" / "cleaned" / f"{jira_issue_key}_pipeline_meta.json"
    export_file_path = TOOLKIT_ROOT / "output" / "order_schedule_export_main.csv"
    final_export_file_path = TOOLKIT_ROOT / "output" / "final" / "order_schedule_input_prod.csv"

    is_fiscal = False
    fiscal_anchor_weekday = 0
    pipeline_source_attachments: list[str] | None = None
    if pipeline_meta_json.exists():
        meta = json.loads(pipeline_meta_json.read_text(encoding="utf-8"))
        is_fiscal = bool(meta.get("is_fiscal", False))
        fiscal_anchor_weekday = int(meta.get("fiscal_anchor_weekday", 0))
        sa = meta.get("source_attachments")
        if isinstance(sa, list):
            pipeline_source_attachments = [str(x) for x in sa]
        print("Loaded pipeline meta:", meta)
    else:
        print(
            "No pipeline_meta.json — using IS_FISCAL=False. Run: python scripts/jira_input.py",
            jira_issue_key,
        )

    from forge_anvil.data_workbench import DataWorkbench

    dw = DataWorkbench(storage_config={"cloud_provider": "azure", "customer_name": CUSTOMER_NAME})

    old_df = dw.storage.read(source=ORDER_DATA_SOURCE).df()
    # DuckDB auto-detects order_schedule_date as a date column (unlike the old Spark
    # inferSchema read, which left it as a plain string) — restore the M/D/YYYY format
    # the rest of this file already uses (see format_schedule_date()) so unchanged rows
    # round-trip identically instead of silently becoming ISO dates.
    old_df["order_schedule_date"] = pd.to_datetime(
        old_df["order_schedule_date"], format="mixed"
    ).map(format_schedule_date)
    print(f"Loaded {len(old_df):,} rows from {ORDER_DATA_SOURCE}")

    order_bak = BACKUP_DIR / f"order_schedule_input_prod_{_backup_stamp()}.csv"
    old_df.to_csv(order_bak, index=False)
    print("Backup:", order_bak)

    cal_df = dw.storage.read(source=FISCAL_CAL_SOURCE).df()
    print(f"Fiscal calendar rows: {len(cal_df):,} from {FISCAL_CAL_SOURCE}")

    cal_bak = BACKUP_DIR / f"fiscal_cal_{_backup_stamp()}.parquet"
    cal_df.to_parquet(cal_bak, index=False)
    print("Backup:", cal_bak)

    if not cleaned_jira_csv.exists():
        raise FileNotFoundError(
            f"Missing {cleaned_jira_csv}. From toolkit root run: python scripts/jira_input.py {jira_issue_key}"
        )

    new_orders = pd.read_csv(cleaned_jira_csv)
    print("Cleaned Jira CSV:", cleaned_jira_csv, "rows:", len(new_orders))

    missing = [c for c in ORDER_COLUMNS if c not in new_orders.columns]
    if missing:
        raise SystemExit(f"Missing columns: {missing}")
    _fq = pd.to_numeric(new_orders["order_frequency"], errors="coerce")
    if not _fq.notna().all():
        raise SystemExit("order_frequency has nulls; re-run jira_input.py")
    if not np.allclose(_fq, _fq.astype(int)):
        raise SystemExit("order_frequency must be whole numbers; re-run jira_input.py")
    new_orders = new_orders.copy()
    new_orders["order_frequency"] = _fq.astype(int)
    if not new_orders["order_schedule_date"].notna().all():
        raise SystemExit("Bad dates; re-run jira_input.py")

    new_orders = new_orders[ORDER_COLUMNS].copy()
    ticket_full_input = new_orders.copy()

    # --- optional (uncomment if needed) ---
    # new_orders = get_earliest_schedule_per_group(new_orders)
    # new_orders["order_schedule_date"] = pd.to_datetime(
    #     new_orders["order_schedule_date"], errors="coerce"
    # ).map(lambda x: format_schedule_date(x) if pd.notna(x) else x)
    # new_orders = expand_destination_codes(new_orders, ["TMW_0020", "TMW_1091"])

    print("Rows for merge:", len(new_orders))

    for c in ORDER_COLUMNS:
        if c not in old_df.columns:
            raise ValueError(f"Production file missing column {c!r}")

    new_for_merge = new_orders.copy()
    if is_fiscal:
        new_for_merge = build_fiscal_merge_from_submission(
            cal_df, new_for_merge, meta_anchor_weekday=fiscal_anchor_weekday, n_periods=3
        )
    else:
        n_before = len(new_for_merge)
        new_for_merge = get_earliest_schedule_per_pair(new_for_merge)
        n_after = len(new_for_merge)
        if n_after < n_before:
            print(
                f"Non-fiscal: earliest date per (order_group, destination): "
                f"{n_before} -> {n_after} row(s) for merge."
            )
        else:
            print("Non-fiscal: one row per pair (no extra dated rows to drop).")

    merge_ready = new_for_merge[ORDER_COLUMNS].copy()
    merge_ready.to_csv(cleaned_jira_csv, index=False)
    print(f"Wrote merge-ready submission ({len(merge_ready)} rows) -> {cleaned_jira_csv}")

    updated, applied_submission, skipped_pairs, n_pairs_applied = merge_with_pair_updates(
        old_df, new_for_merge
    )

    applied_csv = TOOLKIT_ROOT / "jira_downloads" / "cleaned" / f"{jira_issue_key}_cleaned_applied.csv"
    applied_submission.to_csv(applied_csv, index=False)
    print(
        f"Wrote applied submission rows only ({len(applied_submission)} rows, "
        f"{n_pairs_applied} pair(s)) -> {applied_csv}"
    )

    print("Output shape:", updated.shape)

    export_file_path.parent.mkdir(parents=True, exist_ok=True)
    final_export_file_path.parent.mkdir(parents=True, exist_ok=True)
    updated.to_csv(export_file_path, index=False)
    updated.to_csv(final_export_file_path, index=False)
    print("Wrote (local only, not datastore):", export_file_path)
    print("Wrote (local only, not datastore):", final_export_file_path)

    merge_report_path = write_merge_report(
        toolkit_root=TOOLKIT_ROOT,
        issue_key=jira_issue_key,
        prod_row_count_before=len(old_df),
        prod_row_count_after=len(updated),
        ticket_full_input=ticket_full_input,
        merge_ready=merge_ready,
        applied=applied_submission,
        prod_backup_path=order_bak,
        final_output_path=final_export_file_path,
        cleaned_input_path=cleaned_jira_csv,
        cleaned_applied_path=applied_csv,
        datastore_prod_path=ORDER_DATA_SOURCE,
        pairs_applied=n_pairs_applied,
        pairs_skipped=len(skipped_pairs),
        source_attachments=pipeline_source_attachments,
    )
    merge_report_html_path = write_merge_report_html(
        toolkit_root=TOOLKIT_ROOT,
        issue_key=jira_issue_key,
        prod_row_count_before=len(old_df),
        prod_row_count_after=len(updated),
        ticket_full_input=ticket_full_input,
        merge_ready=merge_ready,
        applied=applied_submission,
        prod_backup_path=order_bak,
        final_output_path=final_export_file_path,
        cleaned_input_path=cleaned_jira_csv,
        cleaned_applied_path=applied_csv,
        datastore_prod_path=ORDER_DATA_SOURCE,
        pairs_applied=n_pairs_applied,
        pairs_skipped=len(skipped_pairs),
        source_attachments=pipeline_source_attachments,
    )
    print(f"Wrote merge report (text) -> {merge_report_path}")
    print(f"Wrote merge report (HTML) -> {merge_report_html_path}")
    print(
        f"Final file row counts: before merge {len(old_df):,} -> after merge {len(updated):,} "
        f"(delta {len(updated) - len(old_df):+,})"
    )

    if pipeline_meta_json.exists():
        meta_out = json.loads(pipeline_meta_json.read_text(encoding="utf-8"))
        meta_out["merge_input_rows"] = len(merge_ready)
        meta_out["applied_merge_rows"] = len(applied_submission)
        meta_out["pairs_applied"] = n_pairs_applied
        meta_out["pairs_skipped_duplicate_prod"] = len(skipped_pairs)
        meta_out["prod_rows_before_merge"] = len(old_df)
        meta_out["prod_rows_after_merge"] = len(updated)
        meta_out["prod_row_delta"] = len(updated) - len(old_df)
        meta_out["merge_report_txt"] = str(merge_report_path.relative_to(TOOLKIT_ROOT))
        meta_out["merge_report_html"] = str(merge_report_html_path.relative_to(TOOLKIT_ROOT))
        pipeline_meta_json.write_text(json.dumps(meta_out, indent=2), encoding="utf-8")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
