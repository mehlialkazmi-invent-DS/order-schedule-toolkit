#!/usr/bin/env python3
"""
Fetch order-schedule submission from a Jira issue (Excel attachment), clean / auto-fix,
infer fiscal mode from summary/description and from Excel (order_frequency 91/182/… → fiscal),
and write local artifacts only.
Does not write to the datastore.

Default path: Jira REST API loads the issue and downloads Excel attachment(s) via each
attachment `content` URL (Basic auth: email + API token). **One** `.xlsx`/`.xls` → that file;
**several** → all are downloaded, cleaned, and **concatenated** into one `{KEY}_cleaned.csv`
(exact duplicate rows across files removed). Use **`--attachment`** to process **only** one
named file.

Environment (Jira Cloud typical), or use config files (see toolkit README):
  JIRA_BASE_URL   e.g. https://your-domain.atlassian.net
  JIRA_EMAIL      login email
  JIRA_API_TOKEN  API token from Atlassian account settings
  Optional: config/local.settings.json + config/secrets.local.env (loaded automatically; env wins)

Optional:
  JIRA_API_VERSION  default 3 (use 2 for some Server/DC instances)

Usage:
  python jira_input.py PROJ-123
  python jira_input.py PROJ-123 --attachment "schedule.xlsx"   # only this file if many on issue
  python jira_input.py PROJ-123 --fiscal         # force 91-day fiscal expansion in notebook
  python jira_input.py PROJ-123 --no-fiscal      # force off

  # When Jira is available only via Cursor (saved attachment + text from the issue):
  python jira_input.py PROJ-123 --excel-path ~/Downloads/submission.xlsx \\
      --summary "..." --description-file ~/tmp/issue_desc.txt

  # No Excel on issue: CSV you built from ticket text (agent-drafted rows):
  python jira_input.py PROJ-123 --manual-csv ~/tmp/order_schedule_from_ticket.csv

Order frequency (order_frequency column): plain integers, phrases (e.g. biweekly → 14),
and Excel-style day codes such as 14D / 14d (meaning 14 days). Same for N days / N day.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shutil
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd

try:
    HERE = Path(__file__).resolve().parent
except NameError:
    HERE = Path.cwd()

if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

TOOLKIT_ROOT = HERE.parent

ORDER_COLUMNS = [
    "order_group_description",
    "destination_code",
    "order_schedule_date",
    "order_frequency",
    "order_review_calendar",
]


def read_submission_excel(rp: Path) -> pd.DataFrame:
    """Load a submission workbook. Some uploads have a title row above the real header; try header offsets."""
    for header_row in (0, 1, 2):
        df = pd.read_excel(rp, engine="openpyxl", header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        if all(c in df.columns for c in ORDER_COLUMNS):
            return df
    return pd.read_excel(rp, engine="openpyxl")


_FREQUENCY_PHRASES = (
    ("every other week", 14),
    ("every two weeks", 14),
    ("fortnightly", 14),
    ("fortnight", 14),
    ("bi weekly", 14),
    ("biweekly", 14),
    ("twice a month", 15),
    ("twice monthly", 15),
    ("semi monthly", 15),
    ("twice a week", 4),
    ("twice weekly", 4),
    ("twice per week", 4),
    ("2x weekly", 4),
    ("biannual", 182),
    ("bi annual", 182),
    ("semi annually", 182),
    ("semiannual", 182),
    ("semi annual", 182),
    ("every six months", 182),
    ("every two fiscal months", 182),
    ("every 2 fiscal months", 182),
    ("every other fiscal month", 182),
    ("fiscal every two months", 182),
    ("fiscal bi-monthly", 182),
    ("fiscal bimonthly", 182),
    ("two fiscal months", 182),
    ("every two months", 60),
    ("bi monthly", 60),
    ("bimonthly", 60),
    ("quarterly", 91),
    ("every quarter", 91),
    ("91 day", 91),
    ("91 days", 91),
    ("yearly", 365),
    ("annually", 365),
    ("annual", 365),
    ("every year", 365),
    ("biennial", 730),
    ("every two years", 730),
    ("monthly", 30),
    ("every month", 30),
    ("once a month", 30),
    ("per month", 30),
    ("weekly", 7),
    ("every week", 7),
    ("once a week", 7),
    ("per week", 7),
    ("daily", 1),
    ("every day", 1),
    ("each day", 1),
)


def format_schedule_date(ts: pd.Timestamp) -> str:
    t = pd.Timestamp(ts)
    return f"{t.month}/{t.day}/{t.year}"


def coerce_order_frequency(val):
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, np.integer)):
        return int(val)
    if isinstance(val, float):
        if np.isnan(val):
            return np.nan
        if val == int(val):
            return int(val)
        return val
    s = str(val).strip()
    if not s:
        return np.nan
    # Excel / planners often use "14D" for a 14-day cadence (not parsed by to_numeric).
    m_nd = re.match(r"^(\d+)\s*[Dd]\s*$", s)
    if m_nd:
        return int(m_nd.group(1))
    m_n_days = re.match(r"^(\d+)\s*days?$", s, re.IGNORECASE)
    if m_n_days:
        return int(m_n_days.group(1))
    num = pd.to_numeric(s, errors="coerce")
    if pd.notna(num) and num == int(num):
        return int(num)
    t = re.sub(r"\s+", " ", s.lower().replace("-", " ").replace("_", " ")).strip()
    for phrase, days in _FREQUENCY_PHRASES:
        if phrase in t:
            return days
    return val


def _maybe_excel_serial(v) -> pd.Timestamp | type(pd.NaT):
    if pd.isna(v):
        return pd.NaT
    if isinstance(v, bool):
        return pd.NaT
    try:
        f = float(v)
    except (TypeError, ValueError):
        return pd.NaT
    if 20000 < f < 60000:
        return pd.to_datetime(f, unit="D", origin="1899-12-30", errors="coerce")
    return pd.NaT


def parse_single_date(v) -> pd.Timestamp | type(pd.NaT):
    if pd.isna(v):
        return pd.NaT
    if isinstance(v, pd.Timestamp):
        return v
    if isinstance(v, datetime):
        return pd.Timestamp(v)
    ex = _maybe_excel_serial(v)
    if pd.notna(ex):
        return ex
    s = str(v).strip()
    if not s:
        return pd.NaT
    for dayfirst in (False, True):
        p = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst)
        if pd.notna(p):
            return p
    return pd.NaT


def clean_schedule_dates(series: pd.Series) -> tuple[pd.Series, list[str]]:
    logs: list[str] = []
    out = []
    for i, v in series.items():
        orig = v
        p = parse_single_date(v)
        if pd.isna(p) and pd.notna(orig) and str(orig).strip():
            logs.append(f"row {i}: could not parse date {orig!r} -> dropped later")
        out.append(p)
    s = pd.Series(out, index=series.index, dtype="datetime64[ns]")
    return s, logs


def force_integer_frequency(series: pd.Series, default: int = 30) -> tuple[pd.Series, list[str]]:
    logs: list[str] = []
    out = []
    for i, v in series.items():
        c = coerce_order_frequency(v)
        if isinstance(c, str):
            n = pd.to_numeric(c, errors="coerce")
            if pd.notna(n) and abs(n - round(n)) < 1e-6:
                c = int(round(n))
                logs.append(f"row {i}: frequency string {v!r} -> int {c}")
            else:
                logs.append(f"row {i}: unknown frequency {v!r} -> default {default}")
                c = default
        elif isinstance(c, float) and pd.notna(c):
            if abs(c - round(c)) < 1e-6:
                c = int(round(c))
            else:
                logs.append(f"row {i}: non-integer float frequency {v} -> rounded {int(round(c))}")
                c = int(round(c))
        elif pd.isna(c):
            logs.append(f"row {i}: missing frequency -> default {default}")
            c = default
        else:
            c = int(c)
        out.append(c)
    return pd.Series(out, index=series.index, dtype="int64"), logs


def adf_collect_text(node) -> list[str]:
    """Pull text leaves from Atlassian Document Format (ADF) for keyword search."""
    out: list[str] = []
    if isinstance(node, dict):
        if node.get("type") == "text" and "text" in node:
            out.append(str(node["text"]))
        for c in node.get("content", []) or []:
            out.extend(adf_collect_text(c))
    elif isinstance(node, list):
        for x in node:
            out.extend(adf_collect_text(x))
    return out


def description_to_search_text(description) -> str:
    if description is None:
        return ""
    if isinstance(description, str):
        return description
    if isinstance(description, dict):
        if description.get("type") == "doc":
            return " ".join(adf_collect_text(description))
        return json.dumps(description)
    return str(description)


def parse_anchor_weekday(text_lower: str) -> int | None:
    """
    Map weekday name near 'anchor' to Mon=0 … Sun=6. Returns None if not found.
    """
    full = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
    short = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")
    # Prefer explicit "anchor ... wednesday" window
    for m in re.finditer(r"anchor[^\n.]{0,120}", text_lower):
        chunk = m.group(0)
        for i, name in enumerate(full):
            if name in chunk:
                return i
        for i, name in enumerate(short):
            if re.search(rf"\b{name}\b", chunk):
                return i
    # "wednesday anchor" / "on wednesdays" near fiscal
    if "fiscal" in text_lower or "91" in text_lower or "182" in text_lower:
        for i, name in enumerate(full):
            if name in text_lower:
                return i
        for i, name in enumerate(short):
            if re.search(rf"\b{name}\b", text_lower):
                return i
    return None


def infer_fiscal_from_jira_text(summary: str, description_plain: str) -> dict:
    """
    Decide whether order_sch_main should set IS_FISCAL (91-day fiscal expansion).
    Writes rationale for humans / skills to audit.
    """
    text = f"{summary}\n{description_plain}".lower()

    negatives = (
        "non-fiscal",
        "non fiscal",
        "no fiscal",
        "not fiscal",
        "without fiscal",
        "gregorian only",
        "calendar month only",
        "do not use fiscal",
        "skip fiscal",
    )
    if any(n in text for n in negatives):
        return {
            "is_fiscal": False,
            "fiscal_anchor_weekday": 0,
            "rationale": "matched explicit non-fiscal / calendar-only phrasing",
            "explicit_non_fiscal": True,
        }

    positives = (
        "fiscal month",
        "fiscal quarter",
        "fiscal calendar",
        "fiscal year",
        "fiscal schedule",
        "fy month",
        "91-day",
        "91 day",
        "91 days",
        "182-day",
        "182 day",
        "182 days",
        "quarterly fiscal",
        "fiscal anchor",
        "generate_91",
        "three fiscal",
        "next 3 fiscal",
        "every two fiscal months",
        "semi-annual fiscal",
        "semiannual fiscal",
    )
    # standalone "fiscal" near order schedule context
    if "fiscal" in text and any(
        k in text for k in ("order schedule", "order_schedule", "schedule", "submission", "upload")
    ):
        hit = True
    else:
        hit = any(p in text for p in positives)

    if not hit:
        return {
            "is_fiscal": False,
            "fiscal_anchor_weekday": 0,
            "rationale": "no fiscal / fiscal-cadence keywords; default IS_FISCAL=False",
            "explicit_non_fiscal": False,
        }

    wd = parse_anchor_weekday(text)
    if wd is None:
        wd = 0
        anchor_note = "default Mon=0 (no weekday found near anchor/fiscal)"
    else:
        anchor_note = f"parsed anchor weekday Mon=0..Sun=6 -> {wd}"

    return {
        "is_fiscal": True,
        "fiscal_anchor_weekday": wd,
        "rationale": f"fiscal keywords matched; {anchor_note}",
        "explicit_non_fiscal": False,
    }


def is_fiscal_frequency_days(n: int) -> bool:
    """True for 91, 182, 273, … (multiples of 91 used for fiscal quarter / half-year cadence)."""
    return n >= 91 and n % 91 == 0


def infer_fiscal_anchor_weekday_from_submission_fiscal_rows(cleaned: pd.DataFrame) -> int | None:
    """
    Mon=0 … Sun=6 from order_schedule_date on rows where order_frequency is a fiscal cadence
    (multiple of 91, at least 91). Single weekday if all agree; else modal weekday.
    """
    if cleaned.empty or "order_frequency" not in cleaned.columns:
        return None
    freq = pd.to_numeric(cleaned["order_frequency"], errors="coerce")
    m = freq.notna() & (freq >= 91) & (freq % 91 == 0)
    if not m.any():
        return None
    dts = pd.to_datetime(cleaned.loc[m, "order_schedule_date"], errors="coerce").dropna()
    if dts.empty:
        return None
    wds = dts.dt.weekday
    if wds.nunique() == 1:
        return int(wds.iloc[0])
    return int(wds.mode().iloc[0])


def infer_fiscal_from_order_frequency_column(
    cleaned: pd.DataFrame,
    *,
    summary: str,
    description_plain: str,
) -> dict | None:
    """
    If any row has order_frequency that is a fiscal cadence (91, 182, …), treat as fiscal merge.
    Anchor weekday: issue text (parse_anchor_weekday) if present, else inferred from those row
    dates, else Mon=0.
    """
    if cleaned.empty or "order_frequency" not in cleaned.columns:
        return None
    freq = pd.to_numeric(cleaned["order_frequency"], errors="coerce")
    if not (freq.notna() & (freq >= 91) & (freq % 91 == 0)).any():
        return None
    text = f"{summary}\n{description_plain}".lower()
    text_wd = parse_anchor_weekday(text)
    data_wd = infer_fiscal_anchor_weekday_from_submission_fiscal_rows(cleaned)
    if text_wd is not None:
        wd = text_wd
        anchor_note = f"anchor weekday from issue text Mon=0..Sun=6 -> {wd}"
    elif data_wd is not None:
        wd = data_wd
        names = "Mon Tue Wed Thu Fri Sat Sun".split()
        anchor_note = (
            f"anchor weekday inferred from fiscal-cadence row dates -> {wd} ({names[wd]})"
        )
    else:
        wd = 0
        anchor_note = "default Mon=0 (no weekday in issue text or parseable fiscal-row dates)"
    return {
        "is_fiscal": True,
        "fiscal_anchor_weekday": wd,
        "rationale": f"order_frequency is fiscal cadence (91/182/… multiple) in Excel; {anchor_note}",
    }


def clean_submission_dataframe(
    df: pd.DataFrame,
    *,
    frequency_default: int = 30,
) -> tuple[pd.DataFrame, list[str]]:
    all_logs: list[str] = []
    work = df.copy()
    work.columns = [str(c).strip() for c in work.columns]

    missing = [c for c in ORDER_COLUMNS if c not in work.columns]
    if missing:
        raise ValueError(f"Submission missing required columns: {missing}. Found: {list(work.columns)}")

    for col in ("order_group_description", "order_review_calendar"):
        if col in work.columns and work[col].dtype == object:
            work[col] = work[col].astype(str).replace({"nan": np.nan, "None": np.nan})
            work[col] = work[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            if col == "order_review_calendar":
                work[col] = work[col].str.replace(",", "", regex=False)

    # destination_code: do not impute; keep blank as blank and preserve literals like NA / Missing.
    if "destination_code" in work.columns:
        dc = work["destination_code"]
        if dc.dtype == object:
            work["destination_code"] = dc.apply(
                lambda x: x.strip()
                if isinstance(x, str)
                else (np.nan if pd.isna(x) else x)
            )
        else:
            work["destination_code"] = dc.where(dc.notna(), np.nan)

    dts, dlogs = clean_schedule_dates(work["order_schedule_date"])
    all_logs.extend(dlogs)
    bad_date = dts.isna() & work["order_group_description"].notna()
    if bad_date.any():
        all_logs.append(f"dropped {bad_date.sum()} row(s) with invalid dates")
    work["_dt"] = dts
    work = work.dropna(subset=["_dt", "order_group_description"])
    work["order_schedule_date"] = work["_dt"].map(format_schedule_date)
    work = work.drop(columns=["_dt"])

    freqs, flogs = force_integer_frequency(work["order_frequency"], default=frequency_default)
    all_logs.extend(flogs)
    work["order_frequency"] = freqs
    work = work[work["order_group_description"].astype(str).str.strip().ne("")]

    out = work[ORDER_COLUMNS].copy()
    return out.reset_index(drop=True), all_logs


def _basic_auth_header(email: str, token: str) -> str:
    raw = f"{email}:{token}".encode()
    return "Basic " + base64.b64encode(raw).decode()


def jira_request_json(url: str, email: str, token: str) -> dict:
    req = urllib.request.Request(url)
    req.add_header("Authorization", _basic_auth_header(email, token))
    req.add_header("Accept", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        raise RuntimeError(f"Jira HTTP {e.code}: {body}") from e


def jira_download(url: str, email: str, token: str) -> bytes:
    req = urllib.request.Request(url)
    req.add_header("Authorization", _basic_auth_header(email, token))
    with urllib.request.urlopen(req, timeout=300) as resp:
        return resp.read()


def fetch_issue_fields(
    base_url: str,
    issue_key: str,
    email: str,
    token: str,
    api_version: str,
    fields: str,
) -> dict:
    base = base_url.rstrip("/")
    url = f"{base}/rest/api/{api_version}/issue/{issue_key}?fields={fields}"
    return jira_request_json(url, email, token)


def list_excel_attachments(
    attachments: list[dict],
    preferred_name: str | None,
) -> list[dict]:
    """Excel attachments on the issue. If `preferred_name` is set, return only that match (one item)."""
    xlsx = [a for a in attachments if a.get("filename", "").lower().endswith((".xlsx", ".xls"))]
    if not xlsx:
        names = [a.get("filename") for a in attachments]
        raise FileNotFoundError(f"No Excel attachment on issue. Attachments: {names}")
    if preferred_name:
        for a in xlsx:
            if a.get("filename") == preferred_name:
                return [a]
        lower = preferred_name.lower()
        for a in xlsx:
            if a.get("filename", "").lower() == lower:
                return [a]
        raise FileNotFoundError(
            f"No attachment named {preferred_name!r}. Available: {[x.get('filename') for x in xlsx]}"
        )
    return sorted(xlsx, key=lambda a: (a.get("filename") or "").lower())


def run_cli() -> None:
    from toolkit_env import bootstrap_toolkit_env, default_issue_key as _default_issue_key

    bootstrap_toolkit_env()

    p = argparse.ArgumentParser(
        description="Download Jira Excel attachment(s), merge rows, and save one cleaned CSV per issue."
    )
    p.add_argument(
        "issue_key",
        nargs="?",
        default=None,
        help="Jira issue key, e.g. PROJ-123 (optional if default_issue_key in config/local.settings.json)",
    )
    p.add_argument(
        "--attachment",
        help="Process only this Excel filename on the issue (omit to download and merge all .xlsx/.xls)",
    )
    p.add_argument(
        "--excel-path",
        type=Path,
        help="Local .xlsx/.xls: skip Jira REST API (use with summary/description from Cursor Jira)",
    )
    p.add_argument(
        "--manual-csv",
        type=Path,
        help="Use a CSV with ORDER_COLUMNS as the cleaned submission (no Excel). With JIRA_* set, "
        "still loads issue summary/description for fiscal inference; otherwise pass --summary / --description-*.",
    )
    p.add_argument(
        "--summary",
        default="",
        help="Issue summary for fiscal inference when using --excel-path",
    )
    p.add_argument(
        "--description-text",
        default="",
        help="Issue description (plain) for fiscal inference when using --excel-path",
    )
    p.add_argument(
        "--description-file",
        type=Path,
        help="Read description from file (UTF-8) for fiscal inference with --excel-path",
    )
    p.add_argument(
        "--frequency-default",
        type=int,
        default=30,
        help="Default order_frequency when value is missing or unparseable (default 30)",
    )
    g = p.add_mutually_exclusive_group()
    g.add_argument("--fiscal", action="store_true", help="Force IS_FISCAL=true in pipeline_meta.json")
    g.add_argument("--no-fiscal", action="store_true", help="Force IS_FISCAL=false")
    args = p.parse_args()

    raw_dir = TOOLKIT_ROOT / "jira_downloads" / "raw"
    cleaned_dir = TOOLKIT_ROOT / "jira_downloads" / "cleaned"
    raw_dir.mkdir(parents=True, exist_ok=True)
    cleaned_dir.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    issue_key = (args.issue_key or _default_issue_key() or "").strip().upper()
    if not issue_key:
        print(
            "Missing issue key: pass as argument or set default_issue_key in config/local.settings.json",
            file=sys.stderr,
        )
        sys.exit(1)

    if args.manual_csv and args.excel_path:
        print("Use only one of --manual-csv and --excel-path.", file=sys.stderr)
        sys.exit(1)

    raw_paths: list[tuple[Path, str]] = []
    summary = ""
    description_plain = ""
    description_dump = ""

    if args.manual_csv:
        mc = args.manual_csv.expanduser().resolve()
        if not mc.is_file():
            print(f"Not a file: {mc}", file=sys.stderr)
            sys.exit(1)
        base_m = os.environ.get("JIRA_BASE_URL")
        email_m = os.environ.get("JIRA_EMAIL")
        token_m = os.environ.get("JIRA_API_TOKEN")
        api_ver_m = os.environ.get("JIRA_API_VERSION", "3")
        if base_m and email_m and token_m:
            issue_data_m = fetch_issue_fields(
                base_m, issue_key, email_m, token_m, api_ver_m, "summary,description"
            )
            fields_m = issue_data_m.get("fields", {}) or {}
            summary = fields_m.get("summary") or ""
            dr = fields_m.get("description")
            description_plain = description_to_search_text(dr)
            if isinstance(dr, dict):
                description_dump = json.dumps(dr, indent=2)
            elif dr is None:
                description_dump = ""
            else:
                description_dump = str(dr)
        else:
            summary = args.summary or ""
            if args.description_file:
                description_plain = args.description_file.expanduser().read_text(encoding="utf-8")
            else:
                description_plain = args.description_text or ""
            description_dump = description_plain
        if not summary.strip() and not description_plain.strip():
            print(
                "Warning: --manual-csv with no Jira text; fiscal inference may miss. "
                "Set JIRA_* or pass --summary / --description-*.",
                file=sys.stderr,
            )
        df_manual = pd.read_csv(mc)
        part_m, logs_m = clean_submission_dataframe(df_manual, frequency_default=args.frequency_default)
        cleaned = part_m.reset_index(drop=True)
        logs = list(logs_m)
        raw_paths.append((mc, mc.name))
        print(f"Using manual CSV {mc} ({len(cleaned)} row(s))")
    elif args.excel_path:
        excel_src = args.excel_path.expanduser().resolve()
        if not excel_src.is_file():
            print(f"Not a file: {excel_src}", file=sys.stderr)
            sys.exit(1)
        summary = args.summary or ""
        if args.description_file:
            description_plain = args.description_file.expanduser().read_text(encoding="utf-8")
        else:
            description_plain = args.description_text or ""
        description_dump = description_plain
        fn_local = excel_src.name
        rp_local = raw_dir / f"{issue_key}_{stamp}_{fn_local}"
        shutil.copy2(excel_src, rp_local)
        print(f"Copied local Excel {excel_src} -> {rp_local}")
        raw_paths.append((rp_local, fn_local))
        if not summary.strip() and not description_plain.strip():
            print(
                "Warning: no --summary or --description; fiscal inference defaults to off. "
                "Pass text from the Jira issue (e.g. from Cursor) for automatic IS_FISCAL.",
                file=sys.stderr,
            )
    else:
        base = os.environ.get("JIRA_BASE_URL")
        email = os.environ.get("JIRA_EMAIL")
        token = os.environ.get("JIRA_API_TOKEN")
        api_ver = os.environ.get("JIRA_API_VERSION", "3")

        if not base or not email or not token:
            print(
                "Set JIRA_BASE_URL, JIRA_EMAIL, and JIRA_API_TOKEN — or use --excel-path / --manual-csv.",
                file=sys.stderr,
            )
            sys.exit(1)

        issue_data = fetch_issue_fields(
            base, issue_key, email, token, api_ver, "attachment,summary,description"
        )
        fields = issue_data.get("fields", {}) or {}
        summary = fields.get("summary") or ""
        description_raw = fields.get("description")
        description_plain = description_to_search_text(description_raw)
        if isinstance(description_raw, dict):
            description_dump = json.dumps(description_raw, indent=2)
        elif description_raw is None:
            description_dump = ""
        else:
            description_dump = str(description_raw)

        atts = fields.get("attachment", []) or []
        try:
            excel_atts = list_excel_attachments(atts, args.attachment)
        except FileNotFoundError as e:
            print(
                f"{e}\n"
                "Attach Excel on the issue, or run with --excel-path / --manual-csv "
                "(build CSV from ticket details when there is no workbook).",
                file=sys.stderr,
            )
            sys.exit(1)
        names = [a["filename"] for a in excel_atts]
        print(f"Excel on issue: {len(excel_atts)} file(s) -> {names}")
        for att in excel_atts:
            fn = att["filename"]
            content_url = att["content"]
            rp = raw_dir / f"{issue_key}_{stamp}_{fn}"
            print(f"Downloading {fn!r} -> {rp}")
            rp.write_bytes(jira_download(content_url, email, token))
            raw_paths.append((rp, fn))

    if not args.manual_csv:
        all_cleaned: list[pd.DataFrame] = []
        all_logs: list[str] = []
        for rp, fn in raw_paths:
            df = read_submission_excel(rp)
            part, part_logs = clean_submission_dataframe(df, frequency_default=args.frequency_default)
            all_cleaned.append(part)
            all_logs.extend(f"{fn}: {line}" for line in part_logs)

        cleaned = pd.concat(all_cleaned, ignore_index=True)
        if len(raw_paths) > 1:
            n_before = len(cleaned)
            cleaned = cleaned.drop_duplicates(subset=ORDER_COLUMNS, keep="first").reset_index(drop=True)
            if len(cleaned) < n_before:
                all_logs.append(f"cross-file dedupe: removed {n_before - len(cleaned)} duplicate row(s)")
        logs = all_logs

    source_files = "; ".join(fn for _, fn in raw_paths)
    raw_saved = "; ".join(str(rp) for rp, _ in raw_paths)

    cleaned_csv = cleaned_dir / f"{issue_key}_cleaned.csv"
    cleaned.to_csv(cleaned_csv, index=False)

    if args.fiscal:
        pipeline_flags = {
            "is_fiscal": True,
            "fiscal_anchor_weekday": 0,
            "rationale": "CLI --fiscal override",
        }
    elif args.no_fiscal:
        pipeline_flags = {
            "is_fiscal": False,
            "fiscal_anchor_weekday": 0,
            "rationale": "CLI --no-fiscal override",
        }
    else:
        text_flags = infer_fiscal_from_jira_text(summary, description_plain)
        if text_flags.get("explicit_non_fiscal"):
            pipeline_flags = text_flags
        else:
            from_freq = infer_fiscal_from_order_frequency_column(
                cleaned,
                summary=summary,
                description_plain=description_plain,
            )
            if from_freq is not None:
                pipeline_flags = from_freq
            else:
                pipeline_flags = text_flags

    pipeline_meta_path = cleaned_dir / f"{issue_key}_pipeline_meta.json"
    pipeline_flags_meta = {
        k: v
        for k, v in pipeline_flags.items()
        if k in ("is_fiscal", "fiscal_anchor_weekday", "rationale")
    }
    pipeline_meta = {
        "issue_key": issue_key,
        "is_fiscal": pipeline_flags_meta["is_fiscal"],
        "fiscal_anchor_weekday": int(pipeline_flags_meta["fiscal_anchor_weekday"]),
        "rationale": pipeline_flags_meta["rationale"],
        "cleaned_csv": str(cleaned_csv.relative_to(TOOLKIT_ROOT)),
        "source_attachments": [fn for _, fn in raw_paths],
    }
    pipeline_meta_path.write_text(json.dumps(pipeline_meta, indent=2), encoding="utf-8")
    print(f"Pipeline flags: {pipeline_meta_path}")
    print(json.dumps(pipeline_meta, indent=2))

    meta_path = cleaned_dir / f"{issue_key}_clean_log_{stamp}.txt"
    meta_lines = [
        f"issue={issue_key}",
        f"source_attachments={source_files}",
        f"raw_saved={raw_saved}",
        f"rows_cleaned={len(cleaned)}",
        "",
        "fixes / notes:",
        *logs,
    ]
    meta_path.write_text("\n".join(meta_lines), encoding="utf-8")

    ctx_path = cleaned_dir / f"{issue_key}_jira_context.txt"
    ctx_path.write_text(
        f"Summary:\n{summary}\n\nDescription (plain):\n{description_plain}\n\nDescription (raw dump):\n{description_dump}\n",
        encoding="utf-8",
    )
    print(f"Jira summary/description: {ctx_path}")

    print(f"Cleaned rows: {len(cleaned)} -> {cleaned_csv}")
    print(
        "After order_sch_main.py merge, this file is overwritten with merge-ready rows "
        "(earliest date per pair if non-fiscal, or fiscal calendar expansion 91/182/…)."
    )
    print(f"Log: {meta_path}")
    if logs:
        print("--- fix log (summary) ---")
        for line in logs[:30]:
            print(line)
        if len(logs) > 30:
            print(f"... and {len(logs) - 30} more (see log file)")


if __name__ == "__main__":
    run_cli()
