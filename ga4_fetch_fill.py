#!/usr/bin/env python3
"""
GA4 Expected-Value Fetch-and-Fill — Phase 3 (PROD)

Reads an Excel workbook with a 'Main' sheet (any other sheets — Lookup,
Exhaustive list, etc. — are preserved untouched in the output) and fills the
`ga4_expected_values` column from the GA4 Data API.

Main sheet expected columns (names are case-insensitive; remap below if yours
differ):

    Link | Module | event_name | platform | parameter_name |
    bq_column | api_column | required | rules_expected_values | ga4_expected_values

Output column (filled by this script):
    ga4_expected_values

Reference column (read for prefix/suffix templates, NEVER modified):
    rules_expected_values

The platform → GA4 streamName mapping lives in STREAM_MAPS below. The project
has migrated to PROD, so `--env prod` is now the default. `nonprod` is still
selectable if you ever need to validate against the QA property.

Each GA4 query is filtered by BOTH eventName AND streamName, so an `App` row
never receives Web data and vice versa.

A row is processed when ALL of these are true:
    1. required               == yes
    2. api_column             is in valid GA4 dimension format
                                (e.g. `customEvent:foo`, `eventName`,
                                `pagePath`). Rows with `Not Found`, blank,
                                or malformed entries are skipped automatically.
    3. ga4_expected_values    is blank (so a re-run only fills the missing
                                rows — safe to retry after a crash)

Behaviour notes
---------------
- Output OVERWRITES the input file in place. Use `--output` to write elsewhere.
- A `.tmp` file is used during write, then atomically renamed, so a crash
  during write will never destroy the original.
- api_column values are sent to GA4 EXACTLY AS WRITTEN (whitespace-trimmed
  only). This property registers many dimensions with a `[event_context]`
  suffix in the api_name itself (e.g. `customEvent:actionType[openDeepLink]`)
  — those brackets are part of the registered name, NOT annotation.
- `rules_expected_values` is read for `<<placeholder>>` templates that filter
  GA4 values by literal prefix/suffix. The column itself is NEVER modified.
- Column header names can be remapped via the COL_* constants near the top
  if your sheet uses different headers.

USAGE
-----
    pip install pandas openpyxl google-analytics-data google-auth python-dotenv

    python ga4_fetch_fill.py --input mysheet.xlsx --date 05042026-07042026

Default credentials: ./prod.json   (override with --credentials or GA4_CREDENTIALS)
Default property   : taken from GA4_PROPERTY_ID in .env (override with --property)

Optional flags:
    --module "Video Tag Listing"   # filter to one module (substring match)
    --dry-run                       # print plan, no API calls, no writes
    --env nonprod                   # if you need to validate against QA again
    --output path/to/other.xlsx    # write to a different file instead of in-place
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
# Column name configuration  --  edit here if your sheet uses different headers
# --------------------------------------------------------------------------- #
# All comparisons are done case-insensitively after stripping whitespace, so
# "Api_Column" and "api_column" are treated the same. Update the right-hand
# side if your sheet uses a different header name.

COL_MODULE      = "module"
COL_PLATFORM    = "platform"
COL_EVENT       = "event_name"
COL_PARAM       = "parameter_name"
COL_BQ          = "bq_column"
COL_API         = "api_column"
COL_REQUIRED    = "required"
COL_RULES       = "rules_expected_values"   # read-only, has <<...>> templates
COL_GA4_VALUES  = "ga4_expected_values"     # WRITE TARGET — filled by this script
COL_STATUS      = "ga4_check_status"        # WRITE TARGET — coverage tracking

# --- Value written into COL_STATUS ---
# A non-blank value means "this script attempted to process this row in some
# run". Downstream scripts count non-blank cells for coverage %.
STATUS_CHECKED = "checked"

# Columns that MUST exist in the input (others like Link / Module are optional
# but preserved if present).
REQUIRED_INPUT_COLS = [COL_EVENT, COL_PARAM, COL_API, COL_REQUIRED, COL_PLATFORM]


# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

GA4_MAX_DIMENSIONS = 8


# ─────────────────────────────────────────────────────────────────────────────
#  STREAM MAPPINGS — pick one with --env prod | --env nonprod
# ─────────────────────────────────────────────────────────────────────────────
#  Maps the 'platform' column in your Main sheet to the GA4 streamName values
#  that should be queried. Stream names DIFFER between prod and non-prod.
#
#  Keys must be lowercase (the script lowercases the platform column before
#  lookup). Values must match GA4 stream names EXACTLY (case- and
#  space-sensitive). Run discover_streams.py / verify_creds.py to list the
#  actual stream names in the property before editing this dict.
# ─────────────────────────────────────────────────────────────────────────────
STREAM_MAPS = {
    "prod": {
        # NOTE: the character between "Core" and "iOS" / "Android" is an
        # EN-DASH (–, U+2013), NOT a regular hyphen (-, U+002D). GA4 stream
        # names are character-sensitive — copy/paste from the GA4 Admin UI
        # rather than retyping if you ever edit these.
        "app": ["Formula 1 Core–iOS", "Formula 1 Core–Android"],
        "web": ["Web Stream"],
    },
    "nonprod": {
        "app": ["com.fodmltd.OfficialF1.ios", "F1 QA", "F1 Alpha", "F1 Prod"],
        "web": ["web-nonprodf1", "Fantasy non prod"],
    },
}


# --------------------------------------------------------------------------- #
# Helpers — parsing, validation, template expansion
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


# api_column entries we treat as "no dimension" — skipped silently.
_INVALID_API_LITERALS = {"", "-", "—", "nan", "n/a", "na", "not found", "tbd", "?"}

# Valid GA4 dimension api_names look like one of:
#   eventName                                 (single identifier)
#   pagePath                                  (single identifier)
#   customEvent:actionType                    (scope:param)
#   customEvent:actionType[openDeepLink]      (scope:param[event_context])
#
# This property registers many dimensions WITH a [event_context] suffix as
# part of the api_name itself — those brackets are not annotations, they are
# part of the literal name GA4 expects. So we keep them.
_API_DIM_RE = re.compile(
    r"^[A-Za-z][A-Za-z0-9_]*"           # first identifier (scope or standalone)
    r"(:[A-Za-z][A-Za-z0-9_]*"          # optional :paramName
    r"(\[[^\[\]]+\])?"                  # optional [event_context] suffix
    r")?$"
)


def clean_api_column(val):
    """
    Normalize an api_column cell into the actual GA4 dimension name.

    Returns the cleaned dimension string if it's a valid GA4 dimension,
    or None if the cell is blank, "Not Found", or otherwise malformed.

    Whitespace is trimmed; brackets like [openDeepLink] are PRESERVED as
    they are part of the registered api_name in this property.
    """
    if pd.isna(val):
        return None
    raw = str(val).strip()
    if raw.lower() in _INVALID_API_LITERALS:
        return None
    if not _API_DIM_RE.match(raw):
        return None
    return raw


def is_valid_api_column(val) -> bool:
    """True iff this api_column cell is in a usable GA4 dimension format."""
    return clean_api_column(val) is not None


def is_blank(val):
    if pd.isna(val):
        return True
    return str(val).strip() in ("", "-", "—", "nan")


# Templates: anything with a <<...>> placeholder. The literal text BEFORE the
# first placeholder anchors the start of the GA4 value; the literal text
# AFTER the last placeholder anchors the end. Examples:
#   "<<page>>"           -> match anything
#   "landing_<<page>>"   -> startswith "landing_"
#   "<<page>>_clicked"   -> endswith   "_clicked"
#   "pre_<<x>>_post"     -> startswith "pre_"  AND endswith "_post"
TEMPLATE_PATTERN = re.compile(r"<<.+?>>")


def has_template(val) -> bool:
    """True if the cell contains any <<...>> placeholder anywhere."""
    if pd.isna(val):
        return False
    return bool(TEMPLATE_PATTERN.search(str(val)))


def _expand_template_chunk(chunk: str, ga4_values: set) -> set:
    """
    Apply one <<...>>-bearing chunk's prefix/suffix filter to ga4_values.

    The literal text BEFORE the first '<<' is the required prefix; the
    literal text AFTER the last '>>' is the required suffix. Returns the
    GA4 values that satisfy both anchors.

    Anchorless templates (e.g. '<<x>>', '<<listing page title>>' -- where
    there is no literal text before '<<' or after '>>') are SKIPPED on
    purpose: they'd match every GA4 value and produce a giant
    pipe-separated dump that is rarely what the author meant. If the
    author actually wanted "everything", they should leave the rules
    cell blank.
    """
    first = chunk.find("<<")
    last  = chunk.rfind(">>")
    if last <= first:
        return set()  # malformed brackets, ignore
    prefix = chunk[:first]
    suffix = chunk[last + 2:]
    if not prefix and not suffix:
        # Anchorless -- would match everything. Skip to avoid dumping the
        # entire dimension's value space into one cell.
        return set()
    return {v for v in ga4_values
            if v.startswith(prefix) and v.endswith(suffix)}


def expand_cell(rules_raw, ga4_values: set) -> set:
    """
    Decide what goes into ga4_expected_values for one row.

    The rules column is consulted ONLY for prefix/suffix filtering hints
    via <<...>> placeholders. Anything else -- plain literals, human
    commentary like "Eg: all_videos_clicked archive_videos_clicked
    (only English)" -- is treated as commentary and ignored.

    Comparison/validation between the rules column and what GA4 actually
    returned is intentionally OUT OF SCOPE for this script; that is handled
    downstream. This function's only job is to fetch GA4 values and
    optionally narrow them by an explicit template.

    Returns:
      - blank rules                       -> all GA4 values (no filter)
      - rules has no <<...>> at all       -> all GA4 values (literals
                                              are commentary, not filters)
      - rules has one or more anchored    -> GA4 values matching any of
        templates (e.g. 'landing_<<x>>')    the templates' prefix/suffix
      - rules has only anchorless         -> empty set; row stays blank.
        templates (e.g. '<<x>>')            We skip these on purpose --
                                              dumping every GA4 value into
                                              one cell creates a useless
                                              pipe-blob. If the author
                                              wants "everything", they
                                              should leave rules blank.

    Examples (with GA4 values = {'landing_home', 'home_clicked',
                                  'archive_videos_clicked', 'audio'}):
      ''                              -> {'landing_home','home_clicked',
                                          'archive_videos_clicked','audio'}
      'header'                        -> all of them (literal ignored)
      '<<x>>'                         -> set()    (anchorless: skipped)
      '<<listing page title>>'        -> set()    (anchorless: skipped)
      'landing_<<x>>'                 -> {'landing_home'}
      '<<x>>_clicked'                 -> {'home_clicked',
                                          'archive_videos_clicked'}
      '<<x>>_clicked | Eg: a b (en)'  -> {'home_clicked',
                                          'archive_videos_clicked'}
                                          (commentary chunk ignored)
    """
    if is_blank(rules_raw):
        return set(ga4_values)

    rules_str = str(rules_raw)
    # No template anywhere -> rules cell is pure commentary, return everything.
    if "<<" not in rules_str or ">>" not in rules_str:
        return set(ga4_values)

    # At least one template is present. Apply each template chunk and union
    # the results. Non-template chunks (commentary like "Eg: ...") are
    # silently ignored.
    out = set()
    for chunk in rules_str.split("|"):
        chunk = chunk.strip()
        if "<<" in chunk and ">>" in chunk:
            out.update(_expand_template_chunk(chunk, ga4_values))
    return out


# --------------------------------------------------------------------------- #
# Input
# --------------------------------------------------------------------------- #

def read_input(path: str):
    """
    Read input file. Returns (main_df, other_sheets, main_sheet_name).

      - For Excel: looks for a sheet named 'Main' (case-insensitive). Other
        sheets are kept untouched and written back as-is.
      - For CSV: other_sheets is None.
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
            others = [k for k in all_sheets if k != main_sheet_name]
            print(f"[INFO] Using sheet: '{main_sheet_name}' "
                  f"(other sheets kept as-is: {others or 'none'})")

        df = all_sheets[main_sheet_name]
        other_sheets = {k: v for k, v in all_sheets.items() if k != main_sheet_name}

    df.columns = df.columns.str.strip().str.lower()

    missing = [c for c in REQUIRED_INPUT_COLS if c not in df.columns]
    if missing:
        sys.exit(f"[ERROR] Main sheet is missing column(s): {missing}. "
                 f"Got: {list(df.columns)}")

    # Make sure rules, output, and status columns exist as object-typed columns.
    if COL_RULES not in df.columns:
        print(f"[WARN] '{COL_RULES}' column not found -- creating it blank.")
        df[COL_RULES] = ""
    if COL_GA4_VALUES not in df.columns:
        print(f"[INFO] '{COL_GA4_VALUES}' column not found -- creating it blank.")
        df[COL_GA4_VALUES] = ""
    if COL_STATUS not in df.columns:
        print(f"[INFO] '{COL_STATUS}' column not found -- creating it blank.")
        df[COL_STATUS] = ""

    # Force object dtype so we can write strings back without LossySetitemError.
    df[COL_RULES]      = df[COL_RULES].fillna("").astype(object)
    df[COL_GA4_VALUES] = df[COL_GA4_VALUES].fillna("").astype(object)
    df[COL_STATUS]     = df[COL_STATUS].fillna("").astype(object)

    # Drop fully blank rows (no event_name).
    df = df[df[COL_EVENT].notna() & (df[COL_EVENT].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    df["_required_parsed"] = df[COL_REQUIRED].apply(parse_required)
    # Pre-clean the api column once. None means "skip this row".
    df["_api_clean"] = df[COL_API].apply(clean_api_column)

    print(f"[INFO] Total rows loaded from Main: {len(df)}")
    return df, other_sheets, main_sheet_name


def apply_module_mask(df: pd.DataFrame, module_arg: str) -> pd.Series:
    """
    Return a boolean mask selecting rows that match the module filter.

    IMPORTANT: this returns a mask, not a sliced dataframe. The caller keeps
    the full dataframe and uses the mask only to decide which rows to process
    -- this is what makes module-scoped runs safe to write back in place
    (rows outside the filter are preserved unchanged).

    If `module_arg` is empty/None or there is no module column, returns a
    mask that selects every row (i.e. no filtering).
    """
    if not module_arg or COL_MODULE not in df.columns:
        return pd.Series([True] * len(df), index=df.index)

    modules = [m.strip() for m in module_arg.split(",") if m.strip()]
    print(f"[INFO] Module filter(s): {modules}")
    mask = pd.Series([False] * len(df), index=df.index)
    for m in modules:
        sub = df[COL_MODULE].astype(str).str.lower().str.contains(m.lower(), na=False)
        print(f"         '{m}' -> {int(sub.sum())} row(s)")
        mask = mask | sub
    print(f"[INFO] After module filter: {int(mask.sum())} / {len(df)}  "
          f"(remaining {len(df) - int(mask.sum())} row(s) are preserved as-is)")
    return mask


# --------------------------------------------------------------------------- #
# GA4 fetch (batched — same pattern as the previous script)
# --------------------------------------------------------------------------- #

def fetch_valid_dimensions(client, property_id) -> set:
    """
    Pull every dimension api_name registered in the property via the
    Metadata API. Used for upfront validation so we don't fire runReport
    queries against dimensions GA4 is going to reject.
    """
    from google.analytics.data_v1beta.types import GetMetadataRequest
    req = GetMetadataRequest(name=f"properties/{property_id}/metadata")
    meta = client.get_metadata(req)
    return {d.api_name for d in meta.dimensions}


def suggest_dimension(unknown: str, valid_dims: set):
    """
    Find a likely-correct replacement for an unrecognized dimension.

    Returns (corrected_name, reason) or (None, None) if no good match.

    Tries, in order:
      1. Case-insensitive match — usually a typo we can auto-fix.
      2. Different scope — customEvent <-> customUser <-> customItem.
         Returned as a suggestion, NOT auto-applied (the user has to
         decide whether the data they want is actually item-scoped, etc.).
    """
    lower_map = {d.lower(): d for d in valid_dims}

    # 1. Case-only mismatch
    canonical = lower_map.get(unknown.lower())
    if canonical and canonical != unknown:
        return canonical, "case mismatch"

    # 2. Scope swap
    if ":" in unknown:
        scope, suffix = unknown.split(":", 1)
        for alt_scope in ("customEvent", "customUser", "customItem"):
            if alt_scope == scope:
                continue
            candidate = f"{alt_scope}:{suffix}"
            if candidate in valid_dims:
                return candidate, f"different scope ({scope} -> {alt_scope})"
            alt = lower_map.get(candidate.lower())
            if alt:
                return alt, f"different scope ({scope} -> {alt_scope})"
    return None, None


def validate_and_remap_dimensions(df: pd.DataFrame, valid_dims: set) -> pd.DataFrame:
    """
    Walk every cleaned api_column in the dataframe, validate against the
    property's registered dimensions, and:

      - auto-fix case-only mismatches in the _api_clean column
      - leave scope mismatches untouched but warn loudly with suggestions
      - leave totally unknown dimensions untouched (they'll be skipped at
        fetch time and counted in the summary)

    Returns the (possibly modified) dataframe.
    """
    # NaN passes a `if d` truthiness check (bool(NaN) is True), which then
    # blows up sorted() when it tries to compare a float with strings. Filter
    # explicitly to strings only.
    unique_dims = sorted({d for d in df["_api_clean"] if isinstance(d, str) and d})
    if not unique_dims:
        return df

    auto_fixed   = {}   # original -> canonical (case fix)
    suggestions  = {}   # original -> (suggestion, reason)
    unrecognized = []

    for d in unique_dims:
        if d in valid_dims:
            continue
        sug, reason = suggest_dimension(d, valid_dims)
        if sug and reason == "case mismatch":
            auto_fixed[d] = sug
        elif sug:
            suggestions[d] = (sug, reason)
        else:
            unrecognized.append(d)

    if auto_fixed:
        print(f"\n[INFO] Auto-correcting {len(auto_fixed)} case-only mismatch(es):")
        for orig, fixed in auto_fixed.items():
            print(f"         {orig!r}  ->  {fixed!r}")
        df["_api_clean"] = df["_api_clean"].map(lambda d: auto_fixed.get(d, d))

    if suggestions:
        print(f"\n[WARN] {len(suggestions)} dimension(s) not registered as "
              f"requested -- consider updating your sheet:")
        for orig, (sug, reason) in suggestions.items():
            print(f"         {orig!r}  ->  {sug!r}   ({reason})")
        print(f"       These rows WILL BE SKIPPED until you update the "
              f"{COL_API} cells. We don't auto-remap scope changes because "
              f"event-scoped vs item-scoped data isn't interchangeable.")

    if unrecognized:
        print(f"\n[WARN] {len(unrecognized)} dimension(s) not found in property "
              f"metadata at all:")
        for d in unrecognized:
            print(f"         {d!r}")
        print(f"       These are likely missing from "
              f"Admin -> Custom Definitions in GA4. The relevant rows will "
              f"be skipped in this run.")

    return df


def fetch_from_ga4(client, property_id, eligible_df, platform_to_streams,
                   start_date, end_date):
    """
    Run GA4 queries grouped by platform, applying a streamName filter per
    platform so events from one platform never bleed into another's rows.

    Uses the cleaned api_column (no [annotation], validated format).

    Returns:
        actual:          {(platform_lc, event_name): {dim: set(values)}}
        events_found:    set of (platform_lc, event_name) seen in GA4
        invalid_dims:    set of dimensions GA4 rejected at runtime
        skipped_platforms: set of platforms with no STREAM_MAPS entry
    """
    from google.analytics.data_v1beta.types import (
        RunReportRequest, Dimension, Metric, DateRange,
        FilterExpression, FilterExpressionList, Filter,
    )

    actual       = {}
    events_found = set()
    invalid_dims = set()
    skipped_platforms = set()

    grouped = eligible_df.groupby(eligible_df[COL_PLATFORM].astype(str).str.strip())

    for platform_raw, group_df in grouped:
        platform_lc = platform_raw.lower()
        streams = platform_to_streams.get(platform_lc)

        if not streams:
            print(f"[WARN] No streams configured for platform {platform_raw!r} "
                  f"-- skipping {len(group_df)} row(s).")
            skipped_platforms.add(platform_raw)
            continue

        events = sorted({str(e).strip() for e in group_df[COL_EVENT]})
        # _api_clean is guaranteed non-None for eligible rows. Still filter
        # to strings only so a stray NaN can't crash sorted().
        dims   = sorted({d for d in group_df["_api_clean"] if isinstance(d, str) and d})

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
                        print(f"    [X] Invalid dimension (rejected by GA4): {d}")
                        invalid_dims.add(d)
            time.sleep(0.3)

    if invalid_dims:
        print(f"[WARN] Dimensions rejected by GA4 at runtime: {invalid_dims}")

    return actual, events_found, invalid_dims, skipped_platforms


# --------------------------------------------------------------------------- #
# Fill — writes ONLY into ga4_expected_values, never touches rules_expected_values
# --------------------------------------------------------------------------- #

def _is_anchorless_only(rules_raw) -> bool:
    """
    True if rules contains <<...>> templates AND every template is anchorless
    (no literal text outside the brackets). Used to distinguish 'we
    deliberately skipped this row' from 'we queried and got nothing'.

    Returns False for blank rules and for rules with no templates at all
    (those are valid 'no filter, return everything' cases).
    """
    if is_blank(rules_raw):
        return False
    rules_str = str(rules_raw)
    if "<<" not in rules_str or ">>" not in rules_str:
        return False
    seen_template = False
    for chunk in rules_str.split("|"):
        chunk = chunk.strip()
        if "<<" not in chunk or ">>" not in chunk:
            continue
        first = chunk.find("<<")
        last  = chunk.rfind(">>")
        if last <= first:
            continue
        seen_template = True
        prefix = chunk[:first]
        suffix = chunk[last + 2:]
        if prefix or suffix:
            return False  # found at least one anchored template
    return seen_template  # only True if we saw templates AND all were anchorless


def fill_ga4_values(df: pd.DataFrame, actual, events_found, invalid_dims,
                    scope_mask=None):
    """
    Mutates df in place: writes pipe-joined values into ga4_expected_values
    for every eligible row that doesn't already have a value.

    `scope_mask` (optional) is a boolean Series. When provided, only rows
    where the mask is True are touched -- rows outside the mask are left
    completely untouched.

    Status (`ga4_check_status`) is intentionally NOT written here. It's
    computed in a separate post-processing pass (`mark_checked_status`)
    so that "checked" tracks whether GA4 actually returned data for the
    row, including data from earlier runs that's already in the file.

    Returns counters: (filled, no_data, invalid_dim, event_nf,
                       skipped_no_api, skipped_anchorless). These are for
    the run's console summary -- they are NOT written to the sheet.
    """
    filled = no_data = invalid_dim = event_nf = 0
    skipped_no_api = skipped_anchorless = 0

    for i, row in df.iterrows():
        if scope_mask is not None and not bool(scope_mask.loc[i]):
            continue
        if row["_required_parsed"] is not True:
            continue
        # Retry safety: never overwrite a row that already has a value.
        if not is_blank(row[COL_GA4_VALUES]):
            continue

        api_dim = row["_api_clean"]

        if not isinstance(api_dim, str) or not api_dim:
            skipped_no_api += 1
            continue
        if _is_anchorless_only(row[COL_RULES]):
            skipped_anchorless += 1
            continue

        platform_lc = str(row[COL_PLATFORM]).strip().lower()
        ev = str(row[COL_EVENT]).strip()
        key = (platform_lc, ev)

        if api_dim in invalid_dims:
            invalid_dim += 1
            continue
        if key not in events_found:
            event_nf += 1
            continue

        ga4_values = actual.get(key, {}).get(api_dim, set())
        expanded = expand_cell(row[COL_RULES], ga4_values)
        if not expanded:
            no_data += 1
            continue

        df.at[i, COL_GA4_VALUES] = "|".join(sorted(expanded))
        filled += 1

    return filled, no_data, invalid_dim, event_nf, skipped_no_api, skipped_anchorless


# Parameter-name value that identifies the event-level row (the row that
# represents the event itself rather than one of its parameters). Used by
# the rollup step in mark_checked_status. Compared case-insensitively.
EVENT_LEVEL_PARAM_NAME = "event"


def mark_checked_status(df: pd.DataFrame, scope_mask=None) -> tuple:
    """
    Walk in-scope rows and mark `ga4_check_status` according to the policy:

      1. PARAMETER ROW: marked 'checked' iff its ga4_expected_values is
         non-empty. (i.e. this row's GA4 query actually returned a value.)
      2. EVENT-LEVEL ROW (parameter_name == 'event'): marked 'checked' iff
         AT LEAST ONE parameter row of the same (event_name, platform) is
         itself 'checked' by rule 1.

    Out-of-scope rows are never touched. Existing 'checked' status from
    earlier runs on out-of-scope rows is preserved as-is.

    This function is idempotent inside scope: re-running recomputes
    'checked' from the current state of ga4_expected_values, so the column
    always reflects what's currently true.

    Returns (param_checked, event_checked) -- two counts for the run summary.
    """
    if scope_mask is None:
        scope_mask = pd.Series([True] * len(df), index=df.index)

    has_value = ~df[COL_GA4_VALUES].apply(is_blank)

    # Rule 1: parameter rows with GA4 values -> 'checked'
    param_mask = scope_mask & has_value
    df.loc[param_mask, COL_STATUS] = STATUS_CHECKED
    param_checked = int(param_mask.sum())

    # Rule 2: event-level rollup. For each (event_name, platform) where any
    # parameter row got data, mark the event-level row 'checked'.
    event_checked = 0
    if param_mask.any():
        # Find every (event, platform) that has at least one filled param row.
        events_with_data = set()
        for (ev, pl), _ in df.loc[param_mask].groupby([COL_EVENT, COL_PLATFORM]):
            events_with_data.add((str(ev).strip(), str(pl).strip()))

        # Mark the event-level row(s) for those events.
        param_lower = df[COL_PARAM].astype(str).str.strip().str.lower()
        ev_str = df[COL_EVENT].astype(str).str.strip()
        pl_str = df[COL_PLATFORM].astype(str).str.strip()
        for ev, pl in events_with_data:
            event_row_mask = (
                scope_mask
                & (ev_str == ev)
                & (pl_str == pl)
                & (param_lower == EVENT_LEVEL_PARAM_NAME)
            )
            n = int(event_row_mask.sum())
            if n:
                df.loc[event_row_mask, COL_STATUS] = STATUS_CHECKED
                event_checked += n

    return param_checked, event_checked


# --------------------------------------------------------------------------- #
# Output — atomic, in-place by default
# --------------------------------------------------------------------------- #

# Canonical column order for the Main sheet on output. Any extra columns
# (e.g. Link, comments) are appended after these in their original order.
MAIN_COLUMN_ORDER = [
    "link", COL_MODULE, COL_EVENT, COL_PLATFORM, COL_PARAM,
    COL_BQ, COL_API, COL_REQUIRED, COL_RULES, COL_GA4_VALUES, COL_STATUS,
]


def write_output_atomic(df: pd.DataFrame, output_path: str,
                        other_sheets=None, main_sheet_name: str = "Main"):
    """
    Write to a `.tmp` file, then atomically rename over `output_path`.

    Two failure modes are handled differently:
      - Write itself fails  -> the .tmp file is partial/garbage, delete it.
      - Rename fails        -> the .tmp file holds the user's good data;
                                KEEP IT and tell them how to recover. This
                                is the common case on Windows when the
                                target file is open in Excel.

    Output format is chosen from the output_path extension:
      - .csv        -> CSV (other_sheets are ignored with a warning)
      - .xlsx/.xlsm -> Excel workbook (other_sheets preserved as-is)
    """
    out = df.drop(columns=["_required_parsed", "_api_clean"], errors="ignore")

    ordered = [c for c in MAIN_COLUMN_ORDER if c in out.columns]
    extras  = [c for c in out.columns if c not in ordered]
    out = out[ordered + extras]

    is_csv = output_path.lower().endswith(".csv")
    # Insert ".tmp" BEFORE the extension so the temp file has a valid
    # extension that pandas writers will accept. e.g.
    #   data_check.xlsx -> data_check.tmp.xlsx
    #   data_check.csv  -> data_check.tmp.csv
    base, ext = os.path.splitext(output_path)
    tmp_path = base + ".tmp" + ext

    # --- Step 1: write to tmp file ---
    try:
        if is_csv:
            if other_sheets:
                print(f"[WARN] Output is CSV -- {len(other_sheets)} other sheet(s) "
                      f"will NOT be preserved: {list(other_sheets.keys())}")
            out.to_csv(tmp_path, index=False)
        elif not other_sheets:
            out.to_excel(tmp_path, index=False, sheet_name=main_sheet_name,
                         engine="openpyxl")
        else:
            with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                out.to_excel(writer, sheet_name=main_sheet_name, index=False)
                for name, sheet_df in other_sheets.items():
                    sheet_df.to_excel(writer, sheet_name=name, index=False)
            print(f"[INFO] Preserved sheets: {list(other_sheets.keys())}")
    except Exception:
        # Write itself failed -- clean up the partial tmp file.
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise

    # --- Step 2: atomic rename ---
    # If this fails, the tmp file holds the user's good data. Don't delete it.
    try:
        os.replace(tmp_path, output_path)
        print(f"[INFO] Output saved -> {output_path}")
    except PermissionError as e:
        print(f"\n[ERROR] Could not overwrite {output_path!r}: {e}")
        print( "        On Windows, this usually means the file is open in")
        print( "        Excel or another program.")
        print( "")
        print( "        ** Your filled data is preserved in: **")
        print(f"            {tmp_path}")
        print( "")
        print(f"        To recover, close the program holding {output_path!r}")
        print( "        open and either:")
        print(f"          - delete {os.path.basename(output_path)} and rename "
              f"{os.path.basename(tmp_path)} -> {os.path.basename(output_path)}")
        print( "          - OR re-run the script "
              "(it skips rows already filled, so re-runs are cheap)")
        raise


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def parse_args():
    ap = argparse.ArgumentParser(
        description="Fetch GA4 values and fill the ga4_expected_values column "
                    "(in place by default)."
    )
    ap.add_argument("--input",       default=r"C:\Users\kisbhat\Downloads\GA4 Automation\data_check.csv",
                    help="Input Excel/CSV file (will be overwritten in place "
                         "unless --output is given). Default points to "
                         "data_check.csv in the GA4 Automation folder.")
    ap.add_argument("--date",        required=True,
                    help="Date range DDMMYYYY-DDMMYYYY  e.g. 05042026-07042026")
    ap.add_argument("--module",      help="Optional comma-separated module filter (substring match)")
    ap.add_argument("--property",    default=os.environ.get("GA4_PROPERTY_ID")
                                          or os.environ.get("GA_PROPERTY_ID"),
                    help="GA4 Property ID (default: GA4_PROPERTY_ID / GA_PROPERTY_ID from .env)")
    ap.add_argument("--credentials", default=os.environ.get("GA4_CREDENTIALS")
                                          or os.environ.get("GA_CREDENTIALS_PATH")
                                          or "prod.json",
                    help="Service-account JSON path (default: prod.json, "
                         "or GA4_CREDENTIALS / GA_CREDENTIALS_PATH from .env)")
    ap.add_argument("--env",         choices=["prod", "nonprod"], default="prod",
                    help="Which stream mapping to use (default: prod). See STREAM_MAPS in script.")
    ap.add_argument("--output",      help="Output file path (default: overwrite --input)")
    ap.add_argument("--dry-run",     action="store_true", help="Show plan without calling GA4")
    return ap.parse_args()


def main():
    args = parse_args()
    start_date, end_date = parse_date_range(args.date)
    output_path = args.output or args.input  # default: overwrite input

    print(f"[INFO] Date range  : {start_date} -> {end_date}")
    print(f"[INFO] Input file  : {args.input}")
    print(f"[INFO] Output file : {output_path}"
          f"{'   (in-place overwrite)' if output_path == args.input else ''}")
    print(f"[INFO] Environment : {args.env}")
    print(f"[INFO] Dry run     : {args.dry_run}")
    print("")

    df, other_sheets, main_sheet_name = read_input(args.input)
    # Module filter is now a MASK, not a slice. The full df is preserved so
    # rows outside the filter are written back unchanged.
    module_mask = apply_module_mask(df, args.module)

    platform_to_streams = STREAM_MAPS[args.env]
    print(f"[INFO] Stream mapping ({args.env}):")
    for plat, streams in platform_to_streams.items():
        print(f"         {plat!r:>8} -> {streams}")

    # --- Pre-flight row counts (scoped to the module filter) ---
    in_scope = module_mask
    blank_required = int((in_scope & df["_required_parsed"].isna()).sum())
    skipped_no     = int((in_scope & (df["_required_parsed"] == False)).sum())
    if blank_required:
        print(f"[WARN] {blank_required} row(s) have a blank '{COL_REQUIRED}' "
              f"column -- left untouched.")
    if skipped_no:
        print(f"[INFO] {skipped_no} row(s) marked required=no -- left untouched.")

    # Rows skipped because api_column is blank / Not Found / malformed.
    invalid_api_mask = in_scope & (df["_required_parsed"] == True) & df["_api_clean"].isna()
    invalid_api_count = int(invalid_api_mask.sum())
    if invalid_api_count:
        print(f"[INFO] {invalid_api_count} required row(s) have an unusable "
              f"'{COL_API}' (blank / 'Not Found' / malformed) -- skipped.")

    # Rows already filled in ga4_expected_values are skipped (retry-safe).
    already_filled = int((in_scope &
                          (df["_required_parsed"] == True) &
                          ~df[COL_GA4_VALUES].apply(is_blank)).sum())
    if already_filled:
        print(f"[INFO] {already_filled} required row(s) already have "
              f"'{COL_GA4_VALUES}' set -- skipped (retry-safe).")

    # Eligible = in scope AND required AND valid api AND ga4_expected_values blank.
    eligible_mask = (
        in_scope &
        (df["_required_parsed"] == True) &
        df["_api_clean"].notna() &
        df[COL_GA4_VALUES].apply(is_blank)
    )
    eligible = df[eligible_mask]

    templates = int((in_scope &
                     (df["_required_parsed"] == True) &
                     df[COL_RULES].apply(has_template)).sum())
    if templates:
        print(f"[INFO] {templates} required row(s) have <<...>> templates in "
              f"'{COL_RULES}' -- GA4 values will be filtered by literal "
              f"prefix/suffix.")

    # Per-platform breakdown of what will be queried.
    platform_breakdown = {}
    if not eligible.empty:
        for plat_raw, gdf in eligible.groupby(eligible[COL_PLATFORM].astype(str).str.strip()):
            evs  = sorted({str(e).strip() for e in gdf[COL_EVENT]})
            dims = sorted({d for d in gdf["_api_clean"] if isinstance(d, str) and d})
            platform_breakdown[plat_raw] = (evs, dims)

    print(f"[INFO] Rows needing fetch    : {len(eligible)}")
    for plat, (evs, dims) in platform_breakdown.items():
        mapped = "✓" if plat.lower() in platform_to_streams else "✗ (add to STREAM_MAPS)"
        print(f"         platform={plat!r:>8}  events={len(evs)}  dims={len(dims)}  {mapped}")

    if args.dry_run:
        print("")
        for plat, (evs, dims) in platform_breakdown.items():
            print(f"[DRY RUN] platform={plat!r}")
            print(f"          streams: {platform_to_streams.get(plat.lower(), '<none — would skip>')}")
            print(f"          events:  {evs}")
            print(f"          dims:    {dims}")
        print("[DRY RUN] No API call made. No file written.")
        return

    if not platform_breakdown:
        print("[INFO] Nothing to fetch -- writing input through unchanged.")
        write_output_atomic(df, output_path, other_sheets, main_sheet_name)
        return

    if not args.property or not args.credentials:
        sys.exit("[ERROR] Need --property and --credentials "
                 "(or GA4_PROPERTY_ID / GA4_CREDENTIALS in .env)")

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

        # --- Pre-flight dimension validation ---
        # Pull metadata once and check every api_column against the property's
        # registered dimensions. Auto-fixes case typos, surfaces scope
        # mismatches, and reports anything that simply isn't registered.
        print(f"\n[INFO] Validating api_column values against property metadata...")
        valid_dims = fetch_valid_dimensions(client, args.property)
        print(f"[INFO] Property has {len(valid_dims)} registered dimension(s).")
        df = validate_and_remap_dimensions(df, valid_dims)

        # Re-derive the eligible slice in case validation auto-corrected some
        # api columns (case fixes change the _api_clean values).
        eligible_mask = (
            in_scope &
            (df["_required_parsed"] == True) &
            df["_api_clean"].notna() &
            df[COL_GA4_VALUES].apply(is_blank)
        )
        eligible = df[eligible_mask]

        actual, found, invalid_dims, skipped_platforms = fetch_from_ga4(
            client, args.property, eligible, platform_to_streams,
            start_date, end_date,
        )
    except Exception as e:
        sys.exit(f"[ERROR] GA4 authentication/fetch failed: {e}")

    filled, no_data, inv_dim, ev_nf, skp_no_api, skp_anchorless = fill_ga4_values(
        df, actual, found, invalid_dims, scope_mask=in_scope,
    )

    # Compute / refresh the ga4_check_status column based on what's now
    # in ga4_expected_values (including data from prior runs already on disk).
    param_checked, event_checked = mark_checked_status(df, scope_mask=in_scope)

    write_output_atomic(df, output_path, other_sheets, main_sheet_name)

    attempted = filled + no_data + inv_dim + ev_nf + skp_no_api + skp_anchorless

    print("")
    print("=" * 60)
    print("  FETCH-AND-FILL SUMMARY")
    print("=" * 60)
    print(f"  Rows attempted in this run          : {attempted}")
    print(f"    -> filled with GA4 values         : {filled}")
    print(f"    -> queried but no data            : {no_data}")
    print(f"    -> event not found in date range  : {ev_nf}")
    print(f"    -> dimension rejected by GA4      : {inv_dim}")
    print(f"    -> skipped (api_column unusable)  : {skp_no_api}")
    print(f"    -> skipped (anchorless template)  : {skp_anchorless}")
    print(f"  Rows already had ga4_expected_values: {already_filled}")
    print(f"  Rows with unusable api_column       : {invalid_api_count}")
    print(f"  Rows with blank 'required'          : {blank_required}")
    print(f"  Rows marked required=no             : {skipped_no}")
    if skipped_platforms:
        print(f"  Platforms skipped (no mapping)      : {sorted(skipped_platforms)}")
    print(f"  Output -> {output_path}")
    print("=" * 60)
    print(f"  ga4_check_status (in-scope) :")
    print(f"    parameter rows with GA4 data     : {param_checked}")
    print(f"    event-level rows rolled up       : {event_checked}")
    print(f"  ('checked' = parameter row got GA4 data,")
    print(f"   OR event row has at least one such parameter)")
    print("=" * 60)


if __name__ == "__main__":
    main()