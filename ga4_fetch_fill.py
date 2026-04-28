#!/usr/bin/env python3
"""
GA4 Expected-Value Fetch-and-Fill — Phase 2

Reads an Excel workbook with a 'Main' sheet (any other sheets — Lookup,
Exhaustive list, etc. — are preserved untouched in the output).

Main sheet column sequence:
    Module | platform | event_name | parameter_name |
    required | api_column | expected_value

The platform → GA4 streamName mapping is defined in STREAM_MAPS near the top
of this script. Two environments are supported:
    --env prod      uses prod stream names
    --env nonprod   uses non-prod stream names (DEFAULT)

Each GA4 query is filtered by BOTH eventName AND streamName, so an `App` row
never receives Web data and vice versa.

For every row where:
    required        == yes
    api_column      is filled
    expected_value  is blank or contains <<...>> placeholders

…it queries GA4 for the given event + api_column dimension over the date range,
and fills whatever values it finds (pipe-separated) into the expected_value cell.

  - Rows with required=no are left untouched.
  - Rows with required blank are left untouched and reported.
  - Rows that already have a non-template expected_value are skipped.
  - Output is written to <input>_filled.xlsx (alongside the input file).

USAGE
-----
    pip install pandas openpyxl google-analytics-data google-auth

    python ga4_fetch_fill.py \
        --input  events_batch_fill.xlsx \
        --date   05042026-07042026

Default credentials: ./nonprodv2.json
Default property   : 195772067
Override either with --credentials / --property, or via env vars
GA4_CREDENTIALS / GA4_PROPERTY_ID (loaded from .env if python-dotenv installed).

Optional filters and preview:
    python ga4_fetch_fill.py --input events_batch_fill.xlsx --date 05042026-07042026 \
        --module "Video Tag Listing"
    python ga4_fetch_fill.py --input events_batch_fill.xlsx --date 05042026-07042026 --dry-run
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

GA4_MAX_DIMENSIONS = 8

REQUIRED_INPUT_COLS = ["event_name", "parameter_name", "api_column", "required"]


# ─────────────────────────────────────────────────────────────────────────────
#  STREAM MAPPINGS — pick one with --env prod | --env nonprod
# ─────────────────────────────────────────────────────────────────────────────
#  Maps the 'platform' column in your Main sheet to the GA4 streamName values
#  that should be queried. Stream names DIFFER between prod and non-prod.
#
#  Keys must be lowercase (the script lowercases the platform column before
#  lookup). Values must match GA4 stream names EXACTLY (case- and
#  space-sensitive).
# ─────────────────────────────────────────────────────────────────────────────
STREAM_MAPS = {
    "prod": {
        "app": ["Core Android App", "Core IOS App"],
        "web": ["Web Stream"],
    },
    "nonprod": {
        # NOTE: stream names are case- and space-sensitive. No trailing spaces.
        # Verify each stream's platform via `python discover_streams.py`
        # before changing this dict.
        "app": ["com.fodmltd.OfficialF1.ios", "F1 QA", "F1 Alpha", "F1 Prod"],
        "web": ["web-nonprodf1", "Fantasy non prod"],
    },
}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def parse_date_range(date_str):
    """DDMMYYYY-DDMMYYYY -> (YYYY-MM-DD, YYYY-MM-DD)"""
    try:
        a, b = date_str.strip().split("-")
        start = datetime.strptime(a.strip(), "%d%m%Y").strftime("%Y-%m-%d")
        end   = datetime.strptime(b.strip(), "%d%m%Y").strftime("%Y-%m-%d")
        return start, end
    except Exception:
        sys.exit(f"[ERROR] Invalid date format: {date_str!r}. Expected DDMMYYYY-DDMMYYYY")


def parse_required(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    v = str(val).strip().lower()
    if v in ("yes", "true", "1", "y"):
        return True
    if v in ("no", "false", "0", "n"):
        return False
    return None


def is_valid_dim(val):
    if pd.isna(val):
        return False
    v = str(val).strip()
    return v not in ("", "-", "—", "nan", "N/A", "n/a")


def is_blank(val):
    if pd.isna(val):
        return True
    return str(val).strip() in ("", "-", "—", "nan")


# Templates: anything with a <<...>> placeholder. The literal text BEFORE the
# placeholder anchors the start of the GA4 value; the literal text AFTER it
# anchors the end. So:
#   "<<page>>"             -> match anything
#   "landing_<<page>>"     -> startswith "landing_"
#   "<<page>>_clicked"     -> endswith   "_clicked"
#   "pre_<<x>>_post"       -> startswith "pre_"  AND endswith "_post"
TEMPLATE_PATTERN = re.compile(r"<<.+?>>")


def has_template(val) -> bool:
    """True if the cell contains any <<...>> placeholder anywhere."""
    if pd.isna(val):
        return False
    return bool(TEMPLATE_PATTERN.search(str(val)))


def needs_fill(val) -> bool:
    """A cell needs GA4 fill if it's blank OR contains any <<...>> template."""
    return is_blank(val) or has_template(val)


def expand_chunk(chunk: str, ga4_values: set) -> list:
    """
    Expand one pipe-separated chunk against the GA4 values for this row.

      - Pure literal (no <<...>>)   -> returned as-is.
      - Has <<...>>                  -> the literal text before the FIRST '<<'
                                        is the required prefix; the literal
                                        text after the LAST '>>' is the
                                        required suffix. Anything between
                                        placeholders is treated as wildcard.
    """
    chunk = chunk.strip()
    if not chunk:
        return []
    if "<<" not in chunk or ">>" not in chunk:
        return [chunk]
    first = chunk.find("<<")
    last  = chunk.rfind(">>")
    if last <= first:
        return [chunk]  # malformed -- keep literal
    prefix = chunk[:first]
    suffix = chunk[last + 2:]
    return [v for v in ga4_values
            if v.startswith(prefix) and v.endswith(suffix)]


def expand_cell(raw, ga4_values: set) -> set:
    """Expand an entire expected_value cell against GA4 values."""
    if is_blank(raw):
        return set(ga4_values)
    out = set()
    for chunk in str(raw).split("|"):
        out.update(expand_chunk(chunk, ga4_values))
    return out


def derive_output_path(input_path: str) -> str:
    p = Path(input_path)
    return str(p.with_name(f"{p.stem}_filled{p.suffix or '.xlsx'}"))


# --------------------------------------------------------------------------- #
# Input
# --------------------------------------------------------------------------- #

def read_input(path: str):
    """
    Read input file. Returns (main_df, other_sheets, main_sheet_name).

      - For Excel: looks for a sheet named 'Main' (case-insensitive). Other
        sheets (e.g. 'Lookup') are returned untouched in `other_sheets` so
        they can be carried over to the output unchanged.
      - For CSV: other_sheets is None.

    Expected column sequence in the Main sheet:
        Module | platform | event_name | parameter_name | required | api_column | expected_value
    """
    print(f"[INFO] Reading: {path}")

    if path.lower().endswith(".csv"):
        df = pd.read_csv(path)
        other_sheets = None
        main_sheet_name = "Main"
    else:
        all_sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        if not all_sheets:
            sys.exit(f"[ERROR] No sheets found in {path}")

        main_sheet_name = next(
            (k for k in all_sheets.keys() if k.strip().lower() == "main"),
            None,
        )
        if main_sheet_name is None:
            first = list(all_sheets.keys())[0]
            print(f"[WARN] No 'Main' sheet found. Sheets: {list(all_sheets.keys())}. "
                  f"Falling back to first sheet: '{first}'")
            main_sheet_name = first
        else:
            print(f"[INFO] Using sheet: '{main_sheet_name}' "
                  f"(other sheets kept as-is: {[k for k in all_sheets if k != main_sheet_name] or 'none'})")

        df = all_sheets[main_sheet_name]
        other_sheets = {k: v for k, v in all_sheets.items() if k != main_sheet_name}

    df.columns = df.columns.str.strip().str.lower()

    missing = [c for c in REQUIRED_INPUT_COLS if c not in df.columns]
    if missing:
        sys.exit(f"[ERROR] Main sheet is missing column(s): {missing}. Got: {list(df.columns)}")

    if "expected_value" not in df.columns:
        df["expected_value"] = ""
    # Force object dtype + clean NaN -> "" so we can write filled strings back later.
    # Without this, an all-blank column comes back as float64 from openpyxl and
    # raises LossySetitemError on assignment.
    df["expected_value"] = df["expected_value"].fillna("").astype(object)

    df = df[df["event_name"].notna() & (df["event_name"].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    df["_required_parsed"] = df["required"].apply(parse_required)

    print(f"[INFO] Total rows loaded from Main: {len(df)}")
    return df, other_sheets, main_sheet_name


def apply_module_filter(df: pd.DataFrame, module_arg: str) -> pd.DataFrame:
    if not module_arg or "module" not in df.columns:
        return df
    modules = [m.strip() for m in module_arg.split(",") if m.strip()]
    print(f"[INFO] Module filter(s): {modules}")
    mask = pd.Series([False] * len(df), index=df.index)
    for m in modules:
        sub = df["module"].astype(str).str.lower().str.contains(m.lower(), na=False)
        print(f"         '{m}' -> {int(sub.sum())} row(s)")
        mask = mask | sub
    out = df[mask].reset_index(drop=True)
    print(f"[INFO] After module filter: {len(out)} / {len(df)}")
    return out


# --------------------------------------------------------------------------- #
# GA4 fetch (batched — same pattern as v2.3 validator)
# --------------------------------------------------------------------------- #

def fetch_from_ga4(client, property_id, eligible_df, platform_to_streams,
                   start_date, end_date):
    """
    Run GA4 queries grouped by platform, applying a streamName filter per
    platform so events from one platform never bleed into another's rows.

    Returns:
        actual:          {(platform_lc, event_name): {dim: set(values)}}
        events_found:    set of (platform_lc, event_name) seen in GA4
        invalid_dims:    set of dimensions GA4 rejected
        skipped_platforms: set of platforms with no Lookup mapping (rows
                           for these are reported as event-not-found later)
    """
    from google.analytics.data_v1beta.types import (
        RunReportRequest, Dimension, Metric, DateRange,
        FilterExpression, FilterExpressionList, Filter,
    )

    actual       = {}
    events_found = set()
    invalid_dims = set()
    skipped_platforms = set()

    # Group eligible rows by their platform value (preserved as written)
    grouped = eligible_df.groupby(eligible_df["platform"].astype(str).str.strip())

    for platform_raw, group_df in grouped:
        platform_lc = platform_raw.lower()
        streams = platform_to_streams.get(platform_lc)

        if not streams:
            print(f"[WARN] No streams configured for platform {platform_raw!r} in Lookup sheet "
                  f"-- skipping {len(group_df)} row(s).")
            skipped_platforms.add(platform_raw)
            continue

        events = sorted({str(e).strip() for e in group_df["event_name"]})
        dims   = sorted({
            str(d).strip() for d in group_df["api_column"]
            if is_valid_dim(d)
        })

        if not events or not dims:
            continue

        for ev in events:
            actual[(platform_lc, ev)] = {}

        # Combined filter: eventName IN events AND streamName IN streams
        combined_filter = FilterExpression(
            and_group=FilterExpressionList(expressions=[
                FilterExpression(filter=Filter(
                    field_name="eventName",
                    in_list_filter=Filter.InListFilter(values=events),
                )),
                FilterExpression(filter=Filter(
                    field_name="streamName",
                    in_list_filter=Filter.InListFilter(values=streams),
                )),
            ])
        )

        batches = [dims[i:i + GA4_MAX_DIMENSIONS]
                   for i in range(0, len(dims), GA4_MAX_DIMENSIONS)]

        print(f"\n[INFO] Platform: {platform_raw!r}")
        print(f"       Streams : {streams}")
        print(f"       Events  : {len(events)}")
        print(f"       Dims    : {len(dims)} (in {len(batches)} batch(es))")

        for idx, batch in enumerate(batches, 1):
            print(f"  -> Batch {idx}/{len(batches)}: {batch}")
            try:
                req = RunReportRequest(
                    property=f"properties/{property_id}",
                    dimensions=[Dimension(name="eventName")] + [Dimension(name=d) for d in batch],
                    metrics=[Metric(name="eventCount")],
                    date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
                    dimension_filter=combined_filter,
                    limit=100000,
                )
                resp = client.run_report(req)
                for row in resp.rows:
                    ev = row.dimension_values[0].value.strip()
                    key = (platform_lc, ev)
                    if key not in actual:
                        continue
                    events_found.add(key)
                    for i, d in enumerate(batch):
                        v = row.dimension_values[i + 1].value.strip()
                        if v and v != "(not set)":
                            actual[key].setdefault(d, set()).add(v)
            except Exception as e:
                print(f"  [!] Batch failed: {e} -- retrying individually...")
                for d in batch:
                    try:
                        req = RunReportRequest(
                            property=f"properties/{property_id}",
                            dimensions=[Dimension(name="eventName"), Dimension(name=d)],
                            metrics=[Metric(name="eventCount")],
                            date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
                            dimension_filter=combined_filter,
                            limit=100000,
                        )
                        resp = client.run_report(req)
                        for row in resp.rows:
                            ev = row.dimension_values[0].value.strip()
                            key = (platform_lc, ev)
                            if key not in actual:
                                continue
                            events_found.add(key)
                            v = row.dimension_values[1].value.strip()
                            if v and v != "(not set)":
                                actual[key].setdefault(d, set()).add(v)
                    except Exception:
                        print(f"    [X] Invalid dimension (not in GA4): {d}")
                        invalid_dims.add(d)
            time.sleep(0.3)

    if invalid_dims:
        print(f"[WARN] Dimensions rejected by GA4: {invalid_dims}")

    return actual, events_found, invalid_dims, skipped_platforms


# --------------------------------------------------------------------------- #
# Fill
# --------------------------------------------------------------------------- #

def fill_expected_values(df: pd.DataFrame, actual, events_found, invalid_dims):
    """Mutates df in place: writes pipe-joined values into expected_value where applicable."""
    filled = empty = invalid_dim = event_nf = 0

    for i, row in df.iterrows():
        if row["_required_parsed"] is not True:
            continue
        if not needs_fill(row["expected_value"]):
            continue
        api_col = str(row["api_column"]).strip() if is_valid_dim(row["api_column"]) else ""
        if not api_col:
            continue

        platform_lc = str(row["platform"]).strip().lower()
        ev = str(row["event_name"]).strip()
        key = (platform_lc, ev)

        if api_col in invalid_dims:
            invalid_dim += 1
            continue
        if key not in events_found:
            event_nf += 1
            continue

        ga4_values = actual.get(key, {}).get(api_col, set())
        expanded = expand_cell(row["expected_value"], ga4_values)
        if not expanded:
            empty += 1
            continue

        df.at[i, "expected_value"] = "|".join(sorted(expanded))
        filled += 1

    return filled, empty, invalid_dim, event_nf


# --------------------------------------------------------------------------- #
# Output
# --------------------------------------------------------------------------- #

# Canonical column order for the Main sheet on output.
MAIN_COLUMN_ORDER = [
    "module", "platform", "event_name", "parameter_name",
    "required", "api_column", "expected_value",
]


def write_output(df: pd.DataFrame, output_path: str,
                 other_sheets=None, main_sheet_name: str = "Main"):
    out = df.drop(columns=["_required_parsed"], errors="ignore")

    # Reorder Main columns to the canonical sequence; keep any extras at the end.
    ordered = [c for c in MAIN_COLUMN_ORDER if c in out.columns]
    extras  = [c for c in out.columns if c not in ordered]
    out = out[ordered + extras]

    if not other_sheets:
        out.to_excel(output_path, index=False, sheet_name=main_sheet_name, engine="openpyxl")
    else:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            out.to_excel(writer, sheet_name=main_sheet_name, index=False)
            for name, sheet_df in other_sheets.items():
                sheet_df.to_excel(writer, sheet_name=name, index=False)
        print(f"[INFO] Preserved sheets: {list(other_sheets.keys())}")

    print(f"[INFO] Output saved -> {output_path}")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def parse_args():
    ap = argparse.ArgumentParser(description="Fetch GA4 values and fill expected_value column.")
    ap.add_argument("--input",       required=True, help="Input Excel/CSV (e.g. events_batch_fill.xlsx)")
    ap.add_argument("--date",        required=True, help="Date range DDMMYYYY-DDMMYYYY  e.g. 05042026-07042026")
    ap.add_argument("--module",      help="Optional comma-separated module filter (substring match)")
    ap.add_argument("--property",    default=os.environ.get("GA4_PROPERTY_ID", "195772067"),
                    help="GA4 Property ID (default: 195772067, or set GA4_PROPERTY_ID in .env)")
    ap.add_argument("--credentials", default=os.environ.get("GA4_CREDENTIALS", "nonprodv2.json"),
                    help="Service-account JSON path (default: nonprodv2.json, or set GA4_CREDENTIALS in .env)")
    ap.add_argument("--env",         choices=["prod", "nonprod"], default="nonprod",
                    help="Which stream mapping to use (default: nonprod). See STREAM_MAPS in script.")
    ap.add_argument("--output",      help="Output file path (default: <input>_filled.xlsx)")
    ap.add_argument("--dry-run",     action="store_true", help="Show plan without calling GA4")
    return ap.parse_args()


def main():
    args = parse_args()
    start_date, end_date = parse_date_range(args.date)
    output_path = args.output or derive_output_path(args.input)

    print(f"[INFO] Date range  : {start_date} -> {end_date}")
    print(f"[INFO] Input file  : {args.input}")
    print(f"[INFO] Output file : {output_path}")
    print(f"[INFO] Dry run     : {args.dry_run}")
    print("")

    df, other_sheets, main_sheet_name = read_input(args.input)
    df = apply_module_filter(df, args.module)

    platform_to_streams = STREAM_MAPS[args.env]
    print(f"[INFO] Stream mapping ({args.env}):")
    for plat, streams in platform_to_streams.items():
        print(f"         {plat!r:>8} -> {streams}")

    blank_required = int(df["_required_parsed"].isna().sum())
    skipped_no     = int((df["_required_parsed"] == False).sum())
    if blank_required:
        print(f"[WARN] {blank_required} row(s) have a blank 'required' column -- left untouched.")
    if skipped_no:
        print(f"[INFO] {skipped_no} row(s) marked required=no -- left untouched.")

    eligible_mask = (
        (df["_required_parsed"] == True) &
        (df["api_column"].apply(is_valid_dim)) &
        (df["expected_value"].apply(needs_fill))
    )
    eligible = df[eligible_mask]
    templates = int(((df["_required_parsed"] == True) &
                     df["expected_value"].apply(has_template)).sum())
    already_filled = int(((df["_required_parsed"] == True) &
                          ~df["expected_value"].apply(needs_fill)).sum())
    if templates:
        print(f"[INFO] {templates} required row(s) contain <<...>> templates -- "
              f"GA4 values will be filtered by literal prefix/suffix.")
    if already_filled:
        print(f"[INFO] {already_filled} required row(s) already have expected_value -- skipped.")

    # Per-platform breakdown of what will be queried
    platform_breakdown = {}
    if not eligible.empty:
        for plat_raw, gdf in eligible.groupby(eligible["platform"].astype(str).str.strip()):
            evs  = sorted({str(e).strip() for e in gdf["event_name"]})
            dims = sorted({str(d).strip() for d in gdf["api_column"] if is_valid_dim(d)})
            platform_breakdown[plat_raw] = (evs, dims)

    print(f"[INFO] Rows needing fetch    : {len(eligible)}")
    for plat, (evs, dims) in platform_breakdown.items():
        mapped = "✓" if plat.lower() in platform_to_streams else "✗ (add to STREAM_MAPS in script)"
        print(f"         platform={plat!r:>8}  events={len(evs)}  dims={len(dims)}  {mapped}")

    if args.dry_run:
        print("")
        for plat, (evs, dims) in platform_breakdown.items():
            print(f"[DRY RUN] platform={plat!r}")
            print(f"          streams: {platform_to_streams.get(plat.lower(), '<none — would skip>')}")
            print(f"          events:  {evs}")
            print(f"          dims:    {dims}")
        print("[DRY RUN] No API call made.")
        return

    if not platform_breakdown:
        print("[INFO] Nothing to fetch -- writing input through unchanged.")
        write_output(df, output_path, other_sheets, main_sheet_name)
        return

    if not args.property or not args.credentials:
        sys.exit("[ERROR] Need --property and --credentials (or GA4_PROPERTY_ID / GA4_CREDENTIALS in .env)")

    try:
        from google.oauth2 import service_account
        from google.analytics.data_v1beta import BetaAnalyticsDataClient

        print(f"\n[INFO] Authenticating with: {args.credentials}")
        creds = service_account.Credentials.from_service_account_file(
            args.credentials,
            scopes=["https://www.googleapis.com/auth/analytics.readonly"],
        )
        with open(args.credentials) as f:
            info = json.load(f)
        print(f"[INFO] Service Account : {info.get('client_email', '?')}")
        print(f"[INFO] Property ID     : {args.property}")

        client = BetaAnalyticsDataClient(credentials=creds)
        actual, found, invalid_dims, skipped_platforms = fetch_from_ga4(
            client, args.property, eligible, platform_to_streams, start_date, end_date
        )
    except Exception as e:
        sys.exit(f"[ERROR] GA4 authentication/fetch failed: {e}")

    filled, empty, inv_dim, ev_nf = fill_expected_values(df, actual, found, invalid_dims)
    write_output(df, output_path, other_sheets, main_sheet_name)

    print("")
    print("=" * 60)
    print("  FETCH-AND-FILL SUMMARY")
    print("=" * 60)
    print(f"  Rows filled with GA4 values         : {filled}")
    print(f"  Rows with no values in GA4          : {empty}")
    print(f"  Rows with invalid api_column        : {inv_dim}")
    print(f"  Rows whose (platform,event) missing : {ev_nf}")
    print(f"  Rows already had expected_value     : {already_filled}")
    print(f"  Rows with blank 'required'          : {blank_required}")
    print(f"  Rows marked required=no             : {skipped_no}")
    if skipped_platforms:
        print(f"  Platforms skipped (no mapping)      : {sorted(skipped_platforms)}")
    print(f"  Output -> {output_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()