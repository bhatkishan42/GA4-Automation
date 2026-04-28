"""
Extract events and parameters from one or many Confluence pages and write them
to an Excel file.

Three input modes (mutually exclusive):

  1) --input-excel <file.xlsx>   Batch mode. The Excel must have columns:
                                     url       (required)
                                     platform  (required)
                                     module    (optional; defaults to page title)
                                 Auth via ATLASSIAN_EMAIL / ATLASSIAN_TOKEN env vars.

  2) --url <page-url>            Single page, fetched live.
                                 Needs --platform and credentials.

  3) --html-file <path.html>     Single page, no auth.
                                 Needs --platform and --module.

Output Excel columns:
    Module | platform | event_name | parameter_name

Existing rows in the output file are preserved; duplicates (matching all 4 columns)
are skipped.

USAGE (batch)
-------------
    pip install requests beautifulsoup4 pandas openpyxl python-dotenv

    # .env file in same folder:
    #   ATLASSIAN_EMAIL=you@company.com
    #   ATLASSIAN_TOKEN=your_api_token

    python confluence_events.py --input-excel input_template.xlsx --excel events.xlsx
"""

import argparse
import os
import re
import sys
from urllib.parse import parse_qs, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.auth import HTTPBasicAuth

# Load ATLASSIAN_EMAIL / ATLASSIAN_TOKEN from a .env file in the current dir, if present.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # .env support is optional; env vars set by the shell still work


# --------------------------------------------------------------------------- #
# Confluence: URL parsing + fetch
# --------------------------------------------------------------------------- #

def parse_page_id(url: str):
    """Return (base_url, page_id). Supports both /pages/<id>/ and ?pageId=<id> URLs."""
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"

    m = re.search(r"/pages/(\d+)", parsed.path)
    if m:
        return base, m.group(1)

    qs = parse_qs(parsed.query)
    if "pageId" in qs:
        return base, qs["pageId"][0]

    raise ValueError(f"Could not extract page ID from URL: {url}")


def fetch_page(url: str, email: str, token: str):
    """Return (page_title, storage_format_html)."""
    base, page_id = parse_page_id(url)
    api = f"{base}/wiki/rest/api/content/{page_id}?expand=body.storage,title"
    r = requests.get(api, auth=HTTPBasicAuth(email, token), timeout=30)
    r.raise_for_status()
    data = r.json()
    return data["title"], data["body"]["storage"]["value"]


# --------------------------------------------------------------------------- #
# HTML parsing
# --------------------------------------------------------------------------- #

EVENT_NAME_ALIASES = {"event", "event name", "eventname", "event_name"}


def clean(node) -> str:
    if node is None:
        return ""
    return re.sub(r"\s+", " ", node.get_text(" ", strip=True)).strip()


def column_index(headers, *needles):
    for i, h in enumerate(headers):
        for n in needles:
            if n in h:
                return i
    return None


def parse_events_table(table):
    """Return list of (event_name, parameter_name) for one table, or [] if not an events table."""
    rows = table.find_all("tr")
    if len(rows) < 2:
        return []

    headers = [clean(c).lower() for c in rows[0].find_all(["th", "td"])]
    name_idx    = column_index(headers, "event parameter name", "parameter name")
    value_idx   = column_index(headers, "value")
    trigger_idx = column_index(headers, "event trigger", "trigger")

    if name_idx is None or value_idx is None:
        return []

    parameters = []
    trigger_rows_left = 0

    for tr in rows[1:]:
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue

        has_trigger_cell = (trigger_idx is not None and trigger_rows_left == 0)
        if has_trigger_cell and trigger_idx < len(cells):
            try:
                trigger_rows_left = int(cells[trigger_idx].get("rowspan", "1"))
            except (TypeError, ValueError):
                trigger_rows_left = 1
        if trigger_rows_left > 0:
            trigger_rows_left -= 1

        def cell_at(col_idx):
            if col_idx is None:
                return ""
            if has_trigger_cell or trigger_idx is None:
                shift = 0
            else:
                shift = -1 if col_idx > trigger_idx else 0
            real = col_idx + shift
            if 0 <= real < len(cells):
                return clean(cells[real])
            return ""

        pname  = cell_at(name_idx)
        pvalue = cell_at(value_idx)
        if pname:
            parameters.append((pname, pvalue))

    if not parameters:
        return []

    event_name = None
    for pname, pvalue in parameters:
        if pname.strip().lower() in EVENT_NAME_ALIASES and pvalue:
            event_name = pvalue
            break
    if not event_name:
        return []

    return [(event_name, pname) for pname, _ in parameters]


def extract_all_events(html: str):
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for table in soup.find_all("table"):
        out.extend(parse_events_table(table))
    return out


# --------------------------------------------------------------------------- #
# Excel I/O
# --------------------------------------------------------------------------- #

REQUIRED_COLS = ["Module", "platform", "event_name", "parameter_name"]


def load_or_create(xlsx_path: str) -> pd.DataFrame:
    if os.path.exists(xlsx_path):
        df = pd.read_excel(xlsx_path)
        for col in REQUIRED_COLS:
            if col not in df.columns:
                df[col] = ""
        other_cols = [c for c in df.columns if c not in REQUIRED_COLS]
        return df[REQUIRED_COLS + other_cols]
    return pd.DataFrame(columns=REQUIRED_COLS)


def append_unique(df: pd.DataFrame, new_rows: list):
    if df.empty:
        existing = set()
    else:
        existing = set(map(tuple, df[REQUIRED_COLS].astype(str).values.tolist()))

    to_add = []
    for row in new_rows:
        key = tuple(str(row[c]) for c in REQUIRED_COLS)
        if key not in existing:
            to_add.append(row)
            existing.add(key)

    if to_add:
        df = pd.concat([df, pd.DataFrame(to_add)], ignore_index=True)
    return df, len(to_add)


# --------------------------------------------------------------------------- #
# Per-page processing
# --------------------------------------------------------------------------- #

def process_url(url: str, platform: str, module_override, email: str, token: str):
    """Return list of output-row dicts for one Confluence URL. Raises on fetch error."""
    title, html = fetch_page(url, email, token)
    module = module_override if module_override else title

    pairs = extract_all_events(html)
    return module, [
        {"Module": module, "platform": platform, "event_name": ev, "parameter_name": pn}
        for ev, pn in pairs
    ]


def process_html_file(path: str, platform: str, module: str):
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()
    pairs = extract_all_events(html)
    return [
        {"Module": module, "platform": platform, "event_name": ev, "parameter_name": pn}
        for ev, pn in pairs
    ]


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--input-excel", help="Excel file with url, platform, [module] columns (batch mode)")
    src.add_argument("--url",         help="Single Confluence page URL")
    src.add_argument("--html-file",   help="Local HTML file with the page's storage-format content")

    ap.add_argument("--excel",    required=True, help="Output Excel file (created if missing, appended to otherwise)")
    ap.add_argument("--platform", help="Platform tag (App / Web). Required with --url and --html-file.")
    ap.add_argument("--module",   help="Module override. Optional with --url, required with --html-file.")
    ap.add_argument("--email",    default=os.environ.get("ATLASSIAN_EMAIL"))
    ap.add_argument("--token",    default=os.environ.get("ATLASSIAN_TOKEN"))
    args = ap.parse_args()

    out_df = load_or_create(args.excel)
    total_added = 0
    total_dupes = 0

    # ---- Batch mode ---------------------------------------------------------
    if args.input_excel:
        if not args.email or not args.token:
            sys.exit("Batch mode needs credentials. Set ATLASSIAN_EMAIL and ATLASSIAN_TOKEN env vars, "
                     "or pass --email and --token. "
                     "(Get a token: https://id.atlassian.com/manage-profile/security/api-tokens)")

        if not os.path.exists(args.input_excel):
            sys.exit(f"Input file not found: {args.input_excel}")

        in_df = pd.read_excel(args.input_excel)
        in_df.columns = [str(c).strip().lower() for c in in_df.columns]

        if "url" not in in_df.columns or "platform" not in in_df.columns:
            sys.exit(f"Input Excel must have 'url' and 'platform' columns. Got: {list(in_df.columns)}")

        has_module_col = "module" in in_df.columns

        print(f"Reading {len(in_df)} row(s) from {args.input_excel}\n")

        for i, row in in_df.iterrows():
            url      = str(row["url"]).strip()
            platform = str(row["platform"]).strip()
            module_override = None
            if has_module_col and pd.notna(row["module"]) and str(row["module"]).strip():
                module_override = str(row["module"]).strip()

            if not url or url.lower() == "nan":
                print(f"[{i+1}] (skipped: empty url)")
                continue
            if not platform or platform.lower() == "nan":
                print(f"[{i+1}] (skipped: empty platform) {url}")
                continue

            print(f"[{i+1}/{len(in_df)}] {url}")
            try:
                module, new_rows = process_url(url, platform, module_override, args.email, args.token)
            except requests.HTTPError as e:
                print(f"        ERROR fetching page: {e}")
                continue
            except Exception as e:
                print(f"        ERROR: {e}")
                continue

            if not new_rows:
                print(f"        Module: {module!r} -- no events tables found")
                continue

            distinct = sorted({r["event_name"] for r in new_rows})
            out_df, added = append_unique(out_df, new_rows)
            dupes = len(new_rows) - added
            total_added += added
            total_dupes += dupes
            print(f"        Module: {module!r} | platform: {platform} | "
                  f"{len(new_rows)} rows ({len(distinct)} events) -> +{added} new, {dupes} dup")

    # ---- Single URL mode ----------------------------------------------------
    elif args.url:
        if not args.platform:
            sys.exit("--url requires --platform")
        if not args.email or not args.token:
            sys.exit("--url needs credentials (ATLASSIAN_EMAIL / ATLASSIAN_TOKEN env vars or --email/--token).")

        print(f"Fetching: {args.url}")
        module, new_rows = process_url(args.url, args.platform, args.module, args.email, args.token)
        print(f"Module:   {module}")
        print(f"Platform: {args.platform}")
        if not new_rows:
            print("No events tables found on this page.")
        else:
            distinct = sorted({r["event_name"] for r in new_rows})
            print(f"Extracted {len(new_rows)} (event, parameter) rows across {len(distinct)} events.")
            out_df, total_added = append_unique(out_df, new_rows)
            total_dupes = len(new_rows) - total_added

    # ---- HTML-file mode -----------------------------------------------------
    else:
        if not args.platform:
            sys.exit("--html-file requires --platform")
        if not args.module:
            sys.exit("--html-file requires --module \"<Module name>\"")

        new_rows = process_html_file(args.html_file, args.platform, args.module)
        if not new_rows:
            print("No events tables found in the HTML file.")
        else:
            distinct = sorted({r["event_name"] for r in new_rows})
            print(f"Module:   {args.module}")
            print(f"Platform: {args.platform}")
            print(f"Extracted {len(new_rows)} (event, parameter) rows across {len(distinct)} events.")
            out_df, total_added = append_unique(out_df, new_rows)
            total_dupes = len(new_rows) - total_added

    # ---- Save ---------------------------------------------------------------
    out_df.to_excel(args.excel, index=False)
    print(f"\nDone. Appended {total_added} new rows; skipped {total_dupes} duplicates.")
    print(f"Saved -> {args.excel}")


if __name__ == "__main__":
    main()