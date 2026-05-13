#!/usr/bin/env python3
"""
GA4 Expected-Value Fetch-and-Fill — Phase 3 (PROD)

Reads an Excel workbook with a 'Main' sheet (any other sheets — Lookup,
Exhaustive list, etc. — are preserved untouched in the output) and fills the
`ga4_actual_value` column from the GA4 Data API.

Main sheet expected columns (names are case-insensitive; remap below if yours
differ):

    Link | Module | event_name | platform | parameter_name |
    bq_column | api_column | required | pageType_check |
    rules_expected_values | ga4_actual_value | ga4_check_status

Output column (filled by this script):
    ga4_actual_value
    ga4_check_status

Reference column (read for prefix/suffix templates, NEVER modified):
    rules_expected_values

The platform → GA4 streamName mapping lives in STREAM_MAPS below. The project
has migrated to PROD, so `--env prod` is now the default.

Each GA4 query is filtered by BOTH eventName AND streamName, so an `App` row
never receives Web data and vice versa. Optionally a third filter on a
pageType dimension is added — see "pageType filter" below.

DEFAULT-MODE eligibility (a row is processed when ALL are true):
    1. required               == yes
    2. api_column             is in valid GA4 dimension format
                                (e.g. `customEvent:foo`, `eventName`,
                                `pagePath`). Rows with `Not Found`, blank,
                                or malformed entries are skipped automatically.
    3. ga4_actual_value       is blank (so a re-run only fills the missing
                                rows — safe to retry after a crash)
    4. rules_expected_values  is NOT made up entirely of anchorless templates
                                (e.g. `<<x>>`, `<<page title>>`). Anchorless
                                templates would dump every GA4 value into one
                                cell — useless in default mode where the
                                point is to validate against a known set.
                                Use --ignore-required to fetch them anyway.

pageType filter (applies in ALL modes — default and --ignore-required)
----------------------------------------------------------------------
A column `pageType_check` (yes/no) lets you scope a query to one specific
page version of an event. When ANY row in an (event, platform) group has
`pageType_check=yes`, the entire group is queried with an extra filter
clause:

    customEvent:pageType[...] IN [<values from the pageType row's
                                   rules_expected_values, split on '|'>]

The "pageType row" in the group is identified by `api_column` starting
with `customEvent:pageType` (case-insensitive). The filter value comes
from THAT row's rules_expected_values.

If pageType_check=yes but the pageType row is missing, blank, or contains
a `<<...>>` template, the WHOLE GROUP is skipped (no fetch, no writes —
fix the sheet and re-run).

Why: many (event, parameter) pairs return values aggregated across many
pages. Without page-scoping you get a giant pipe-blob. The pageType
filter narrows it down to one page version so the result is meaningful.

--ignore-required MODE
----------------------
Skips the `required` column as an eligibility gate, AND treats anchorless
`<<...>>` templates as "no filter" instead of skipping them. The point is
to fetch ALL available GA4 values so you can review them and update
`rules_expected_values` afterwards.

What it does:
  - Eligibility ignores the `required` column entirely (yes/no/blank rows
    are all considered).
  - Anchorless templates (`<<x>>`, `<<page title>>`) NO LONGER skip the
    row -- they're treated like a blank rules cell, returning every GA4
    value.
  - Anchored templates (`landing_<<x>>`, `<<x>>_clicked`) still filter
    by their literal prefix/suffix as in default mode.
  - pageType filter still applies the same way.
  - Output goes to a SEPARATE FILE: `ignore_required.<ext>` in the
    input's directory (so the curated default-mode file isn't touched).
    Override with --output.
  - Does NOT mutate the `required` column. Auto-classification (e.g.
    "rows with >N values become required=no") was removed -- you'll
    update `required` manually after reviewing the data. Threshold-based
    auto-classification will be revisited later.

Behaviour notes
---------------
- Output OVERWRITES the target file in place (input in default mode,
  ignore_required.<ext> in --ignore-required mode). Use `--output` to
  write elsewhere.
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

    # Default mode — fill ga4_actual_value for required=yes rows
    # (writes back over --input):
    python ga4_fetch_fill.py --input mysheet.xlsx --date 05042026-07042026

    # Default mode, narrowed to one platform / module:
    python ga4_fetch_fill.py --input mysheet.xlsx --date 05042026-07042026 \
        --platform App --module "Video Tag Listing"

    # Ignore-required mode — fetch everything ignoring `required` and
    # anchorless-template skips, write to ignore_required.<ext>:
    python ga4_fetch_fill.py --input mysheet.xlsx --date 05042026-07042026 \
        --ignore-required --platform App

    # Dry run — print the plan, no API calls, no writes:
    python ga4_fetch_fill.py --input mysheet.xlsx --date 05042026-07042026 \
        --dry-run

Default credentials: ./prod.json   (override with --credentials or GA4_CREDENTIALS)
Default property   : taken from GA4_PROPERTY_ID in .env (override with --property)

Optional flags:
    --module "Video Tag Listing"   # filter to one module (substring match)
    --platform App                 # filter to one platform (exact match)
    --ignore-required              # ignore required + anchorless skip;
                                   # writes to ignore_required.<ext>
    --dry-run                      # print plan, no API calls, no writes
    --env nonprod                  # if you need to validate against QA again
    --output path/to/other.xlsx   # write to a different file instead of in-place
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
COL_PAGETYPE_CK = "pagetype_check"             # NEW: yes/no, optional column
COL_RULES       = "rules_expected_values"      # read-only, has <<...>> templates
COL_CONFLUENCE  = "confluence values"          # read-only, user reference notes
COL_GA4_VALUES  = "ga4_actual_value"           # WRITE TARGET — filled by this script
COL_STATUS      = "ga4_check_status"           # WRITE TARGET — coverage tracking

# --- Value written into COL_STATUS ---
# A non-blank value means "this script attempted to process this row in some
# run". Downstream scripts count non-blank cells for coverage %.
STATUS_CHECKED = "checked"

# Default sheet name to look for inside an Excel workbook. Override per-run
# with --sheet. Match is case-insensitive (whitespace and case differences
# are tolerated). If no sheet by this name exists, the script falls back
# to the first sheet with a warning.
DEFAULT_SHEET_NAME = "API Confluence (Final Version)"

# Excel hard limit is 32767 chars per cell. Pandas/openpyxl will silently
# truncate values longer than this and emit a UserWarning. Truncating mid-
# value is dangerous because the last value would be cut in half and look
# like a real (but mangled) data point. We pre-truncate at value boundaries
# and append a marker so it's obvious the cell was capped.
EXCEL_MAX_CELL = 32_767
TRUNCATION_MARKER = "...[TRUNCATED]"


def cap_for_excel(joined: str, values: list) -> str:
    """
    Return `joined` if it fits in an Excel cell. Otherwise drop trailing
    values one at a time (preserving full pipe-separated entries) until the
    string + truncation marker fits. Returns at minimum the marker alone if
    even one value is too long for Excel.

    `values` is the already-sorted list of GA4 values that produced
    `joined`; it lets us shrink at value boundaries instead of mid-string.
    """
    if len(joined) <= EXCEL_MAX_CELL:
        return joined
    # Reserve room for the marker and a leading separator.
    budget = EXCEL_MAX_CELL - len(TRUNCATION_MARKER) - 1
    out = []
    used = 0
    for v in values:
        # +1 for the pipe separator we'd add before this value.
        cost = len(v) + (1 if out else 0)
        if used + cost > budget:
            break
        out.append(v)
        used += cost
    if not out:
        # Single value alone exceeds the cell limit — just emit the marker.
        return TRUNCATION_MARKER
    return "|".join(out) + "|" + TRUNCATION_MARKER

# Columns that MUST exist in the input (others like Link / Module are optional
# but preserved if present). pageType_check is NOT required here -- it's
# treated as optional and absence is interpreted as "no row needs pageType
# scoping" so older sheets keep working.
REQUIRED_INPUT_COLS = [COL_EVENT, COL_PARAM, COL_API, COL_REQUIRED, COL_PLATFORM]


# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

# GA4 caps the number of dimensions per RunReport request. The cap as
# observed in this property is 9 dimensions per request, but GA4 counts
# both the grouping dimension (eventName, which we always add) AND the
# dimensions referenced in the filter (eventName + streamName) toward
# that total. Net effect: usable user-dims per batch = 7. If you see a
# 400 "Requests are limited to N dimensions in a nested request" error,
# lower this further.
#
# When a pageType filter is added (pageType_check=yes groups), it consumes
# one more slot in the filter's dim count, so usable user-dims drops by 1.
# That adjustment is made at the bucket level in fetch_from_ga4().
GA4_MAX_DIMENSIONS = 7


# ─────────────────────────────────────────────────────────────────────────────
#  ITEM-SCOPED DIMENSIONS — separate query path
# ─────────────────────────────────────────────────────────────────────────────
#  GA4 has two scopes that matter here: event-scoped (the default) and
#  item-scoped (used for ecommerce items in events like purchase, view_item,
#  add_to_cart, etc.). The two scopes are NOT mixable in a single RunReport
#  request -- pairing an item dimension with the event-scoped `eventCount`
#  metric returns a 400 with the message:
#
#      "Please remove eventCount to make the request compatible..."
#
#  GA4 docs: item dimensions only join with item metrics (`itemsViewed`,
#  `itemQuantity`, `itemRevenue`, etc.). So the script issues a separate
#  query path for item dims using `itemsViewed` as the metric. The same
#  eventName + streamName + pageType filter clauses still apply, so we
#  retain proper scoping by event and platform.
#
#  Identification is by exact api_name (lowercased): item dimensions in
#  GA4 always have a fixed set of names, none of which use the
#  `customEvent:` / `customUser:` / `customItem:` prefix. They look like
#  bare identifiers (`itemId`, `itemName`, etc.). The set below covers the
#  ones registered in this property; if GA4 introduces new ones later,
#  add them here. `customEvent:price` is included even though it's
#  prefixed -- GA4 routes it through the item-scoped path in this property
#  because it's logged on the items array of ecommerce events.
# ─────────────────────────────────────────────────────────────────────────────
ITEM_SCOPED_DIMS = {
    "itemid",
    "itemname",
    "itembrand",
    "itemvariant",
    "itemcategory",
    "itemcategory2",
    "itemcategory3",
    "itemcategory4",
    "itemcategory5",
    "itemlistid",
    "itemlistname",
    "itemaffiliation",
    "itempromotionid",
    "itempromotionname",
    "itempromotioncreativename",
    "itempromotioncreativeslot",
    "shippingtier",
    "customevent:price",
}


def is_item_scoped_dim(api_name: str) -> bool:
    """
    True if `api_name` is a GA4 item-scoped dimension that must be queried
    with an item metric (itemsViewed/itemQuantity), not with eventCount.

    Comparison is case-insensitive; the canonical api_name from the sheet
    is preserved (case-only typos are auto-fixed earlier by
    validate_and_remap_dimensions).
    """
    if not isinstance(api_name, str):
        return False
    return api_name.strip().lower() in ITEM_SCOPED_DIMS


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


# ─────────────────────────────────────────────────────────────────────────────
#  MODULE-BASED EXTRA STREAMS — additive overrides on top of STREAM_MAPS
# ─────────────────────────────────────────────────────────────────────────────
#  Some modules fire events on additional GA4 streams beyond the platform's
#  normal ones. For example, Fantasy-related events fire on both the normal
#  App/Web streams AND on the "Fantasy Prod and Predict" stream, so a query
#  for Fantasy module rows must include all three streams in its filter or
#  data on Fantasy Prod and Predict gets missed.
#
#  Keys: lowercase substring to match against the module column (case-
#        insensitive). If a row's module contains the key as a substring,
#        the corresponding streams are added to that row's stream filter.
#  Values: dict keyed by env ('prod', 'nonprod') -> list of additional
#        stream names. Stream names must match GA4 EXACTLY.
#
#  Multiple keys can match a single module; all matching streams are unioned.
#  The base STREAM_MAPS streams for the row's platform are ALWAYS included,
#  these are additive only.
#
#  To leave a module unaffected in a given env, omit that env from the
#  inner dict (don't put an empty list -- that'd be a deliberate "no extra
#  streams here" override, same as omission, but the convention is omit).
# ─────────────────────────────────────────────────────────────────────────────
MODULE_EXTRA_STREAMS = {
    # Fantasy events fire on the Fantasy stream in addition to the normal
    # App/Web streams. Applies to BOTH platforms. Nonprod left out because
    # the equivalent nonprod stream name isn't confirmed -- if you need
    # this in nonprod, add e.g. {"nonprod": ["Fantasy non prod"]}.
    "fantasy analytics (existing datalayers)": {
        "prod": ["Fantasy Prod and Predict"],
    },
}


def get_extra_streams_for_module(module_value: str, env: str) -> list:
    """
    Return any additional GA4 stream names that should be unioned into the
    base platform streams for rows in the given module. Returns [] when no
    overrides match.

    Matching is case-insensitive substring -- so "Fantasy Analytics
    (Existing Datalayers) - App" also matches the "fantasy analytics
    (existing datalayers)" key. Multiple keys can match; results are
    unioned and de-duplicated, preserving first-seen order.
    """
    if not isinstance(module_value, str) or not module_value.strip():
        return []
    mod_lc = module_value.strip().lower()
    seen = set()
    out = []
    for substr, env_map in MODULE_EXTRA_STREAMS.items():
        if substr in mod_lc:
            for s in env_map.get(env, []):
                if s not in seen:
                    seen.add(s)
                    out.append(s)
    return out


# --------------------------------------------------------------------------- #
# Helpers — parsing, validation, template expansion
# --------------------------------------------------------------------------- #

# --------------------------------------------------------------------------- #
# Event-name placeholders — `<<...>>` patterns inside event_name
# --------------------------------------------------------------------------- #
# Some sheet rows use placeholder-bearing event names like
# `<<listingname>>_article_listing_click` to represent a family of real GA4
# events (`latest_article_listing_click`, `archive_article_listing_click`,
# etc.) without enumerating every variant. Without special handling these
# would be sent to GA4 as exact-match strings via InListFilter and never
# match anything. Instead we convert them to GA4 string-match filters
# (BEGINS_WITH / ENDS_WITH / FULL_REGEXP) so the script picks up every
# real event the placeholder covers.
#
# The classification produces one of three "shapes":
#
#   ('exact',     name)              -- no placeholder; use InListFilter
#   ('begins',    prefix)            -- one placeholder, suffix only is literal
#                                       (e.g. '<<x>>_clicked' -> ENDS_WITH '_clicked')
#                                       Yes, the bucket is named 'begins' but
#                                       the GA4 op is ENDS_WITH; see below.
#   ('ends',      suffix)            -- one placeholder, prefix only is literal
#                                       (e.g. 'pre_<<x>>' -> BEGINS_WITH 'pre_')
#   ('regex',     compiled_pattern)  -- literal-placeholder-literal, multiple
#                                       placeholders, etc.; needs FULL_REGEXP
#   ('skip',      reason)            -- pure anchorless ('<<x>>') or malformed
#
# Naming note: 'begins'/'ends' refer to *which side has the placeholder*,
# not the GA4 op. A placeholder at the beginning means the literal is at
# the end -> ENDS_WITH on GA4's side. We'll translate at filter-build time.

# Match any <<...>> placeholder with non-empty contents.
_EVENTNAME_PLACEHOLDER_RE = re.compile(r"<<[^<>]*>>")


def classify_event_name(name: str):
    """
    Classify an event_name cell. Returns (kind, payload) where kind is one
    of 'exact' | 'begins' | 'ends' | 'regex' | 'skip' and payload is:

      'exact'  -> the literal name (string)
      'begins' -> the literal prefix (placeholder is at the END)
      'ends'   -> the literal suffix (placeholder is at the START)
      'regex'  -> a compiled re.Pattern with ^...$ anchors
      'skip'   -> a human-readable reason string

    See module-level comment for naming rationale.
    """
    if name is None:
        return ('skip', 'event_name is None')
    s = str(name).strip()
    if not s:
        return ('skip', 'event_name is blank')

    # Fast path: no placeholder -> exact match (the common case).
    if "<<" not in s:
        return ('exact', s)

    # Find every <<...>> span. If the regex finds nothing but '<<' is in
    # the string, the brackets are malformed (e.g. unclosed '<<foo' or
    # nested). Skip with a clear reason.
    matches = list(_EVENTNAME_PLACEHOLDER_RE.finditer(s))
    if not matches:
        return ('skip', f"malformed placeholder brackets in event_name: {s!r}")

    # Pure anchorless: the whole string is one placeholder, no literal text.
    # Would match every event -- almost certainly a sheet error, skip.
    if (len(matches) == 1
            and matches[0].start() == 0
            and matches[0].end() == len(s)):
        return ('skip',
                f"event_name is a single bare placeholder ({s!r}) -- "
                f"would match every event in the property")

    # Single placeholder: collapses to BEGINS_WITH or ENDS_WITH depending
    # on which side has the literal. These are cheaper for GA4 than regex.
    if len(matches) == 1:
        m = matches[0]
        prefix = s[:m.start()]
        suffix = s[m.end():]
        if prefix and not suffix:
            # 'race_<<x>>' -> placeholder at end -> BEGINS_WITH prefix
            return ('ends', prefix)   # 'ends' = literal at start; see comment above
        if suffix and not prefix:
            # '<<x>>_clicked' -> placeholder at start -> ENDS_WITH suffix
            return ('begins', suffix) # 'begins' = literal at end
        # Both sides have literal -> needs regex anyway. Fall through.

    # General case: build a regex. Replace each placeholder with `.*` and
    # escape every literal segment. Anchor with ^...$ so we match the whole
    # event name, not a substring.
    parts = []
    last_end = 0
    for m in matches:
        parts.append(re.escape(s[last_end:m.start()]))
        parts.append(".*")
        last_end = m.end()
    parts.append(re.escape(s[last_end:]))
    try:
        pattern = re.compile("^" + "".join(parts) + "$")
    except re.error as e:
        return ('skip', f"could not build regex for event_name {s!r}: {e}")
    return ('regex', pattern)


def parse_date_range(date_str):
    """DDMMYYYY-DDMMYYYY -> (YYYY-MM-DD, YYYY-MM-DD)"""
    try:
        a, b = date_str.strip().split("-")
        start = datetime.strptime(a.strip(), "%d%m%Y").strftime("%Y-%m-%d")
        end   = datetime.strptime(b.strip(), "%d%m%Y").strftime("%Y-%m-%d")
        return start, end
    except Exception:
        sys.exit(f"[ERROR] Invalid date format: {date_str!r}. Expected DDMMYYYY-DDMMYYYY")


def parse_yes_no(val):
    """
    Parse a yes/no/true/false/1/0/y/n cell.
    Returns True / False / None (None for blank or unrecognized).
    """
    if pd.isna(val) or str(val).strip() == "":
        return None
    v = str(val).strip().lower()
    if v in ("yes", "true", "1", "y"):
        return True
    if v in ("no", "false", "0", "n"):
        return False
    return None


# parse_required is the same operation as parse_yes_no -- kept as an alias
# for backward-compatibility with anything that might import it.
parse_required = parse_yes_no


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


def _expand_template_chunk(chunk: str, ga4_values: set,
                           treat_anchorless_as_all: bool = False) -> set:
    """
    Apply one <<...>>-bearing chunk's prefix/suffix filter to ga4_values.

    The literal text BEFORE the first '<<' is the required prefix; the
    literal text AFTER the last '>>' is the required suffix. Returns the
    GA4 values that satisfy both anchors.

    Anchorless templates (e.g. '<<x>>', '<<listing page title>>' -- where
    there is no literal text before '<<' or after '>>') are SKIPPED on
    purpose in default mode: they'd match every GA4 value and produce a
    giant pipe-separated dump that is rarely what the author meant. If
    the author actually wanted "everything", they should leave the rules
    cell blank.

    In --ignore-required mode, the goal is exactly to dump everything so
    the user can review and write proper rules afterwards. Pass
    `treat_anchorless_as_all=True` to make anchorless templates return
    every GA4 value instead of being skipped.
    """
    first = chunk.find("<<")
    last  = chunk.rfind(">>")
    if last <= first:
        return set()  # malformed brackets, ignore
    prefix = chunk[:first]
    suffix = chunk[last + 2:]
    if not prefix and not suffix:
        # Anchorless. Default: skip (would dump everything). In
        # ignore-required mode: return everything (that's the point).
        if treat_anchorless_as_all:
            return set(ga4_values)
        return set()
    return {v for v in ga4_values
            if v.startswith(prefix) and v.endswith(suffix)}


def expand_cell(rules_raw, ga4_values: set,
                treat_anchorless_as_all: bool = False) -> set:
    """
    Decide what goes into ga4_actual_value for one row.

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
            out.update(_expand_template_chunk(
                chunk, ga4_values,
                treat_anchorless_as_all=treat_anchorless_as_all,
            ))
    return out


# --------------------------------------------------------------------------- #
# pageType filter — per (event, platform) group
# --------------------------------------------------------------------------- #

def get_pagetype_filter(group_df: pd.DataFrame):
    """
    Determine the pageType filter for one (event, platform) group.

    A group needs a pageType filter when ANY row in the group has
    pageType_check == yes. The filter dimension is the api_column of the
    row whose api_column starts (case-insensitively) with
    'customEvent:pageType'. The filter values come from THAT row's
    rules_expected_values, split on '|'.

    Returns one of:
      ('none', None,  None)      -- no rows have pageType_check=yes; no
                                    filter applied; default behaviour
      ('ok',   dim,   values)    -- filter usable; pass dim+values to GA4
      ('skip', reason, None)     -- pageType_check=yes but filter value
                                    can't be derived; the WHOLE group must
                                    be skipped (no fetch, no writes); fix
                                    the sheet and re-run

    Skip reasons covered:
      - no row in the group has api_column starting with customEvent:pageType
      - the pageType row's rules_expected_values is blank
      - the pageType row's rules_expected_values contains a <<...>> template
        (mixed templates+literals also count -- ambiguous, skip to be safe)
    """
    if COL_PAGETYPE_CK not in group_df.columns:
        return ('none', None, None)

    flags = group_df[COL_PAGETYPE_CK].apply(parse_yes_no)
    if not (flags == True).any():
        return ('none', None, None)

    # Identify the pageType row by api_column prefix (case-insensitive).
    def _is_pt_dim(d):
        return isinstance(d, str) and d.lower().startswith("customevent:pagetype")

    pt_mask = group_df["_api_clean"].apply(_is_pt_dim)
    pt_rows = group_df[pt_mask]

    if pt_rows.empty:
        return ('skip',
                'no row with api_column starting with customEvent:pageType',
                None)

    # If multiple, take the first -- shouldn't happen in practice and we
    # don't want to silently combine values from different bracketed
    # variants like [page_view] vs [screen_view].
    pt_row = pt_rows.iloc[0]
    pt_dim = pt_row["_api_clean"]
    pt_rules = pt_row[COL_RULES]

    if is_blank(pt_rules):
        return ('skip',
                "pageType row's rules_expected_values is blank",
                None)

    pt_str = str(pt_rules).strip()
    if "<<" in pt_str or ">>" in pt_str:
        return ('skip',
                "pageType row's rules_expected_values contains a <<...>> template",
                None)

    values = [v.strip() for v in pt_str.split("|") if v.strip()]
    if not values:
        return ('skip',
                "pageType row's rules_expected_values has no usable values",
                None)

    return ('ok', pt_dim, values)


# --------------------------------------------------------------------------- #
# Input
# --------------------------------------------------------------------------- #

def read_input(path: str, sheet_name: str = DEFAULT_SHEET_NAME):
    """
    Read input file. Returns (main_df, other_sheets, main_sheet_name).

      - For Excel: looks for a sheet whose name matches `sheet_name`
        case-insensitively (whitespace stripped). Other sheets in the
        workbook are kept untouched and written back as-is on output.
        Falls back to the first sheet with a warning if no match.
      - For CSV: `sheet_name` is ignored; other_sheets is None.
    """
    print(f"[INFO] Reading: {path}")

    if path.lower().endswith(".csv"):
        # Try UTF-8 first (the modern default). If that blows up on a byte
        # like 0x97 it's almost certainly an Excel-saved CSV using the
        # Windows codepage cp1252 (curly quotes, en-dashes, em-dashes,
        # smart apostrophes -- all of which map to 0x91-0x97). Fall back
        # rather than making the user re-save the file. Order matters:
        # cp1252 is a strict superset of latin-1 in the bytes that show up
        # in practice, but it's the encoding Excel actually writes, so we
        # try it before latin-1.
        try:
            df = pd.read_csv(path, encoding="utf-8")
        except UnicodeDecodeError as e:
            print(f"[WARN] {os.path.basename(path)} is not valid UTF-8 "
                  f"(byte {hex(e.object[e.start])} at position {e.start}). "
                  f"Retrying with cp1252 -- this is what Excel writes when "
                  f"you save as plain 'CSV (Comma delimited)' on Windows.")
            try:
                df = pd.read_csv(path, encoding="cp1252")
            except UnicodeDecodeError:
                print(f"[WARN] cp1252 also failed -- falling back to latin-1 "
                      f"(this won't fail but may garble some characters).")
                df = pd.read_csv(path, encoding="latin-1")
        other_sheets = None
        main_sheet_name = sheet_name
    else:
        all_sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        if not all_sheets:
            sys.exit(f"[ERROR] No sheets found in {path}")

        # Case-insensitive, whitespace-tolerant sheet lookup. This lets
        # 'API Confluence (Final Version)' match 'api confluence (final version)'
        # or similar minor capitalization differences.
        target_norm = sheet_name.strip().lower()
        main_sheet_name = next(
            (k for k in all_sheets.keys() if k.strip().lower() == target_norm),
            None,
        )
        if main_sheet_name is None:
            first = list(all_sheets.keys())[0]
            print(f"[WARN] No {sheet_name!r} sheet found. "
                  f"Sheets: {list(all_sheets.keys())}. "
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
        sys.exit(f"[ERROR] Input sheet is missing column(s): {missing}. "
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

    # pageType_check is OPTIONAL. If the column is missing we treat every
    # row as pageType_check=no (no filtering). This keeps older sheets
    # working without changes.
    if COL_PAGETYPE_CK not in df.columns:
        print(f"[INFO] '{COL_PAGETYPE_CK}' column not found -- "
              f"no pageType filtering will be applied.")

    # Force object dtype so we can write strings back without LossySetitemError.
    df[COL_RULES]      = df[COL_RULES].fillna("").astype(object)
    df[COL_GA4_VALUES] = df[COL_GA4_VALUES].fillna("").astype(object)
    df[COL_STATUS]     = df[COL_STATUS].fillna("").astype(object)
    df[COL_REQUIRED]   = df[COL_REQUIRED].astype(object)

    # Drop fully blank rows (no event_name).
    df = df[df[COL_EVENT].notna() & (df[COL_EVENT].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    df["_required_parsed"] = df[COL_REQUIRED].apply(parse_yes_no)
    # Pre-clean the api column once. None means "skip this row".
    df["_api_clean"] = df[COL_API].apply(clean_api_column)
    # Pre-classify the event_name once. Most rows are 'exact'; placeholder
    # rows (`<<...>>`) become 'begins'/'ends'/'regex' and route through the
    # GA4 string-match filter path instead of the InListFilter exact-match
    # path. Stored as a tuple (kind, payload).
    df["_event_class"] = df[COL_EVENT].apply(classify_event_name)

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


def apply_platform_mask(df: pd.DataFrame, platform_arg: str) -> pd.Series:
    """
    Return a boolean mask selecting rows that match the platform filter.

    Like apply_module_mask, this is a MASK (not a slice) so out-of-scope
    rows survive into the output untouched.

    Match is case-insensitive and EXACT (not substring) -- platform values
    are short tokens like 'App'/'Web' and substring matching would be
    surprising. Comma-separate to allow multiple, e.g. --platform App,Web.
    """
    if not platform_arg:
        return pd.Series([True] * len(df), index=df.index)

    platforms = [p.strip().lower() for p in platform_arg.split(",") if p.strip()]
    print(f"[INFO] Platform filter(s): {platforms}")
    plat_lower = df[COL_PLATFORM].astype(str).str.strip().str.lower()
    mask = plat_lower.isin(platforms)
    print(f"[INFO] After platform filter: {int(mask.sum())} / {len(df)}  "
          f"(remaining {len(df) - int(mask.sum())} row(s) are preserved as-is)")
    return mask


# --------------------------------------------------------------------------- #
# Query bucket builder — groups rows by (platform, pageType filter, event-shape)
# --------------------------------------------------------------------------- #

def _build_query_buckets(eligible_df: pd.DataFrame, platform_to_streams: dict,
                         env: str = "prod"):
    """
    Group eligible rows into query buckets. Each bucket can share a single
    (batched) RunReport because all its rows share the same platform AND
    the same dimension_filter.

    There are two flavours of bucket:

      EXACT bucket: many real events grouped together via InListFilter on
        eventName. Cheap -- one query covers many events. Used for any
        (event, platform) group whose event_name is a literal string with
        no <<...>> placeholder.

      PATTERN bucket: ONE sheet-row event_name (with placeholder) that we
        translate into a GA4 string-match filter (BEGINS_WITH, ENDS_WITH,
        or FULL_REGEXP). Each placeholder gets its own bucket because each
        has a different filter clause -- they can't share a request.
        At fetch time, GA4 returns the real matched event names in its
        rows; the fetcher maps those back to the original placeholder
        sheet-event so the values are unioned into the right cell.

    Bucket key when grouping shares: (platform, pageType_signature).
    Events with the same pageType filter on the same platform share an
    EXACT bucket if literal, or each get their own PATTERN bucket if
    placeholder-bearing.

    Events whose (event, platform) group has pageType_check=yes but lacks
    a usable filter value are NOT put in any bucket -- they go into the
    skipped_groups list and the caller treats their rows as "skipped due
    to bad pageType filter" (no fetch, no fill). Same for events with a
    'skip' classification (e.g. pure-anchorless event_name).

    Returns:
      buckets:           list of dicts. Common fields:
                            'platform_raw', 'platform_lc', 'streams',
                            'dims', 'pt_dim', 'pt_values'
                         EXACT-bucket-only:
                            'kind': 'exact'
                            'events': sorted list of literal event names
                         PATTERN-bucket-only:
                            'kind': 'pattern'
                            'sheet_event': original placeholder string from
                                           the sheet (used to route results
                                           back to the right rows)
                            'match_kind': 'begins' | 'ends' | 'regex'
                            'match_payload': string for begins/ends, or
                                             a compiled regex pattern
      skipped_groups:    list of (event_name, platform_raw, module_key, reason)
      skipped_platforms: set of platforms with no STREAM_MAPS entry
    """
    buckets = []
    skipped_groups = []
    skipped_platforms = set()

    # Some sheets don't have a module column at all (older versions). We
    # still need to group by *something* to keep the per-event-context
    # separation working. Use an empty string '' as the module value when
    # missing — that becomes the single bucket key for every row.
    has_module = COL_MODULE in eligible_df.columns
    if has_module:
        eligible_df = eligible_df.copy()
        eligible_df["_module_key"] = eligible_df[COL_MODULE].fillna("").astype(str).str.strip()
    else:
        eligible_df = eligible_df.copy()
        eligible_df["_module_key"] = ""

    for platform_raw, plat_df in eligible_df.groupby(
            eligible_df[COL_PLATFORM].astype(str).str.strip()):
        platform_lc = platform_raw.lower()
        streams = platform_to_streams.get(platform_lc)
        if not streams:
            skipped_platforms.add(platform_raw)
            continue

        # Per-platform aggregation:
        # - exact_buckets: keyed by (module, pageType-filter signature);
        #   multiple literal events in the same module can share these.
        # - pattern_buckets: list (one per placeholder event per module);
        #   never merged.
        # The module key was added because the SAME (event, platform) pair
        # can appear in multiple modules with DIFFERENT page contexts
        # (e.g. previous_screen_click on App in module 'Video Tag' means
        # the user is leaving a video page, in 'Driver Details' it means
        # leaving a driver page). Without the module split, all those rows
        # would land in one bucket, the first module's pageType row would
        # win the filter, and every other module's rows would get the
        # wrong data. Splitting by module makes each context its own
        # bucket with its own filter, so values returned by GA4 always
        # belong to the right rows.
        exact_buckets = {}
        pattern_buckets = []

        # Group by (module, event). Each (module, event) cell of the sheet
        # gets its own context-specific bucket.
        for (module_key, ev_name), ev_df in plat_df.groupby(
                [plat_df["_module_key"],
                 plat_df[COL_EVENT].astype(str).str.strip()]):
            status, pt_dim, pt_values = get_pagetype_filter(ev_df)
            if status == 'skip':
                # `pt_dim` is the human-readable reason in this branch.
                # Tuple shape: (event_name, platform_raw, module_key, reason)
                # The module_key is needed so skipped_group_keys in
                # fetch_from_ga4 can produce 3-tuple keys matching the
                # fill function's per-row lookup.
                skipped_groups.append((ev_name, platform_raw, module_key, pt_dim))
                continue

            # All rows in this (module, event, platform) group share the
            # same event_name string -> same classification. Take row 0.
            ev_class = ev_df["_event_class"].iloc[0]
            ev_kind, ev_payload = ev_class

            if ev_kind == 'skip':
                skipped_groups.append((ev_name, platform_raw, module_key,
                                       f"event_name skipped: {ev_payload}"))
                continue

            ev_dims = {d for d in ev_df["_api_clean"]
                       if isinstance(d, str) and d}
            sig = None if status == 'none' else (pt_dim, tuple(pt_values))
            pt_dim_v    = pt_dim    if status == 'ok' else None
            pt_values_v = list(pt_values) if status == 'ok' else None

            if ev_kind == 'exact':
                # Module-scoped share: within ONE module, literal events
                # with the same pageType filter can share a bucket. Across
                # modules, never -- each module's context is independent.
                bucket_key = (module_key, sig)
                b = exact_buckets.setdefault(bucket_key, {
                    'module_key': module_key,
                    'events': set(),
                    'dims': set(),
                    'pt_dim': pt_dim_v,
                    'pt_values': pt_values_v,
                })
                b['events'].add(ev_payload)
                b['dims'] |= ev_dims
            else:
                # Pattern (begins / ends / regex). One bucket per
                # (module, placeholder event).
                # Compute final streams for this bucket: base platform
                # streams + any module-specific extras (e.g. Fantasy rows
                # need the Fantasy stream added on top of App/Web). See
                # MODULE_EXTRA_STREAMS for the config.
                bucket_streams = list(streams)
                for s in get_extra_streams_for_module(module_key, env):
                    if s not in bucket_streams:
                        bucket_streams.append(s)
                pattern_buckets.append({
                    'platform_raw': platform_raw,
                    'platform_lc': platform_lc,
                    'streams': bucket_streams,
                    'kind': 'pattern',
                    'module_key': module_key,
                    'sheet_event': ev_name,
                    'match_kind': ev_kind,
                    'match_payload': ev_payload,
                    'dims': sorted(ev_dims),
                    'pt_dim': pt_dim_v,
                    'pt_values': pt_values_v,
                })

        for bucket_key, info in exact_buckets.items():
            # Same module-extras logic as for pattern buckets above.
            bucket_streams = list(streams)
            for s in get_extra_streams_for_module(info['module_key'], env):
                if s not in bucket_streams:
                    bucket_streams.append(s)
            buckets.append({
                'platform_raw': platform_raw,
                'platform_lc': platform_lc,
                'streams': bucket_streams,
                'kind': 'exact',
                'module_key': info['module_key'],
                'events': sorted(info['events']),
                'dims': sorted(info['dims']),
                'pt_dim': info['pt_dim'],
                'pt_values': info['pt_values'],
            })
        buckets.extend(pattern_buckets)

    return buckets, skipped_groups, skipped_platforms


# --------------------------------------------------------------------------- #
# GA4 fetch
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
                   start_date, end_date, env: str = "prod"):
    """
    Run GA4 queries grouped into buckets that share a (platform, pageType
    filter). Within each bucket the request is filtered by:

        eventName IN events
        AND streamName IN streams
        AND <pageType dim> IN <pageType values>     (only if the bucket
                                                      has a pageType filter)

    Uses the cleaned api_column (validated format).

    Returns:
        actual:                {(platform_lc, module_key, sheet_event): {dim: set(values)}}
                               where sheet_event is the event_name string AS
                               WRITTEN IN THE SHEET. For exact buckets that's
                               the literal name (matches the GA4-returned
                               name 1:1). For pattern buckets that's the
                               original placeholder string from the sheet
                               (e.g. '<<listingname>>_article_listing_click')
                               and values from every real GA4 event matching
                               the pattern are unioned under that key.
        events_found:          set of (platform_lc, module_key, sheet_event) seen in GA4
                               (using the same sheet_event-keyed identity)
        invalid_dims:          set of dimensions GA4 rejected at runtime
        skipped_platforms:     set of platforms with no STREAM_MAPS entry
        skipped_group_keys:    set of (platform_lc, module_key, sheet_event) skipped at
                               the bucket-builder stage. Causes: bad pageType
                               filter value, or event_name classified as
                               'skip' (e.g. pure-anchorless `<<x>>`).
    """
    from google.analytics.data_v1beta.types import (
        RunReportRequest, Dimension, Metric, DateRange,
        FilterExpression, FilterExpressionList, Filter,
    )

    actual       = {}
    events_found = set()
    invalid_dims = set()

    buckets, skipped_groups, skipped_platforms = _build_query_buckets(
        eligible_df, platform_to_streams, env=env)

    for plat in skipped_platforms:
        n = (eligible_df[COL_PLATFORM].astype(str).str.strip() == plat).sum()
        print(f"[WARN] No streams configured for platform {plat!r} "
              f"-- skipping {int(n)} row(s).")

    if skipped_groups:
        print(f"\n[WARN] {len(skipped_groups)} (event, platform, module) group(s) "
              f"skipped (bad pageType filter, malformed event_name, etc.):")
        for ev, plat, module_key, reason in skipped_groups[:10]:
            module_disp = f"  [module={module_key!r}]" if module_key else ""
            print(f"         {plat!r:>8} / {ev!r}{module_disp}  --  {reason}")
        if len(skipped_groups) > 10:
            print(f"         ... and {len(skipped_groups) - 10} more")
        print(f"       Fix the relevant cell(s) and re-run.")

    # Set of (platform_lc, module_key, sheet_event) keys that were skipped
    # at the bucket-builder stage. Two reasons feed in here:
    #   - bad/missing pageType filter value (when pageType_check=yes)
    #   - event_name classified as 'skip' (e.g. pure-anchorless `<<x>>`)
    # The fill stage uses this to flag affected rows distinctly from
    # "event not found in date range".
    skipped_group_keys = {(p.lower(), mk, e) for e, p, mk, _ in skipped_groups}

    # Helper: build the eventName filter clause for a bucket. For exact
    # buckets this is an InListFilter; for pattern buckets it's a
    # StringFilter with BEGINS_WITH / ENDS_WITH / FULL_REGEXP.
    def _eventname_filter_clause(bucket):
        if bucket['kind'] == 'exact':
            return FilterExpression(filter=Filter(
                field_name="eventName",
                in_list_filter=Filter.InListFilter(values=bucket['events']),
            ))
        # Pattern bucket. Note the begins/ends naming swap (see
        # classify_event_name docstring): kind 'begins' means the placeholder
        # is at the START of the sheet event_name, so the LITERAL is the
        # suffix -> we ask GA4 for ENDS_WITH that literal. And vice versa.
        mk = bucket['match_kind']
        mp = bucket['match_payload']
        if mk == 'begins':       # literal at end -> ENDS_WITH
            return FilterExpression(filter=Filter(
                field_name="eventName",
                string_filter=Filter.StringFilter(
                    match_type=Filter.StringFilter.MatchType.ENDS_WITH,
                    value=mp,
                ),
            ))
        if mk == 'ends':         # literal at start -> BEGINS_WITH
            return FilterExpression(filter=Filter(
                field_name="eventName",
                string_filter=Filter.StringFilter(
                    match_type=Filter.StringFilter.MatchType.BEGINS_WITH,
                    value=mp,
                ),
            ))
        # 'regex' -> FULL_REGEXP. mp is a compiled re.Pattern; GA4 wants
        # the source string. Our pattern is already ^...$ anchored, which
        # FULL_REGEXP wants implicitly anyway -- the explicit anchors don't
        # hurt and make intent clearer in the request.
        return FilterExpression(filter=Filter(
            field_name="eventName",
            string_filter=Filter.StringFilter(
                match_type=Filter.StringFilter.MatchType.FULL_REGEXP,
                value=mp.pattern,
            ),
        ))

    # Helper: given a real event name returned by GA4 for this bucket,
    # return the sheet_event key under which it should be stored. For
    # exact buckets that's the event itself (only if it was one of the
    # bucket's events -- defensive). For pattern buckets it's always the
    # bucket's sheet_event placeholder.
    def _ga4_event_to_sheet_event(bucket, ga4_event):
        if bucket['kind'] == 'exact':
            return ga4_event if ga4_event in bucket['_events_set'] else None
        # Pattern bucket: GA4's filter already constrained the result to
        # matching events, so any name we get back belongs to this bucket.
        # Defensively re-verify with the local pattern in case GA4's regex
        # interpretation diverges (e.g. case sensitivity edge cases).
        mk = bucket['match_kind']
        mp = bucket['match_payload']
        if mk == 'begins' and ga4_event.endswith(mp):
            return bucket['sheet_event']
        if mk == 'ends'   and ga4_event.startswith(mp):
            return bucket['sheet_event']
        if mk == 'regex'  and mp.match(ga4_event):
            return bucket['sheet_event']
        return None

    for bucket in buckets:
        platform_raw = bucket['platform_raw']
        platform_lc  = bucket['platform_lc']
        streams      = bucket['streams']
        all_dims     = bucket['dims']
        pt_dim       = bucket['pt_dim']
        pt_values    = bucket['pt_values']
        # Module identity for this bucket. All sheet-event keys created
        # from this bucket include the module_key, so values fetched here
        # never leak into rows belonging to a different module's bucket
        # (even when the event_name is identical). See bucket builder for
        # why this matters.
        module_key   = bucket.get('module_key', '')

        # Pre-compute a lookup set for exact-bucket event-name validation
        # (avoid repeated `in list` checks per row).
        if bucket['kind'] == 'exact':
            bucket['_events_set'] = set(bucket['events'])

        # eventName special-case: it's the grouping dim already, so it
        # can't appear in the per-batch dim list (GA4 rejects duplicates).
        # Pull it out and synthesize values directly from each event.
        request_eventname = "eventName" in all_dims
        non_event_name_dims = [d for d in all_dims if d != "eventName"]

        # Split into event-scoped (default path) and item-scoped (separate
        # path with itemsViewed metric -- GA4 rejects item dims paired
        # with eventCount). See ITEM_SCOPED_DIMS at the top of this file
        # for the recognition rules.
        dims = [d for d in non_event_name_dims if not is_item_scoped_dim(d)]
        item_dims = [d for d in non_event_name_dims if is_item_scoped_dim(d)]

        # Identify all sheet-event keys this bucket initializes / fills.
        if bucket['kind'] == 'exact':
            sheet_events_for_bucket = bucket['events']
        else:
            sheet_events_for_bucket = [bucket['sheet_event']]

        if not sheet_events_for_bucket:
            continue
        if not dims and not item_dims and not request_eventname:
            continue

        # Initialize actual for these sheet-event keys.
        for se in sheet_events_for_bucket:
            actual.setdefault((platform_lc, module_key, se), {})
            if request_eventname:
                # For exact buckets, the eventName value IS the sheet event.
                # For pattern buckets we don't pre-seed -- we collect the
                # real event names from GA4 results below and union them.
                if bucket['kind'] == 'exact':
                    actual[(platform_lc, module_key, se)]["eventName"] = {se}

        # Build combined filter clauses.
        filter_clauses = [
            _eventname_filter_clause(bucket),
            FilterExpression(filter=Filter(
                field_name="streamName",
                in_list_filter=Filter.InListFilter(values=streams),
            )),
        ]
        if pt_dim and pt_values:
            filter_clauses.append(FilterExpression(filter=Filter(
                field_name=pt_dim,
                in_list_filter=Filter.InListFilter(values=list(pt_values)),
            )))
        combined_filter = FilterExpression(
            and_group=FilterExpressionList(expressions=filter_clauses))

        # Each filter dim consumes one slot from the per-request dim cap
        # (because GA4 counts filter-referenced dims toward the total).
        # Default usable user-dims: 7 (eventName + streamName already cost
        # 2). With a pageType filter, drop to 6.
        max_dims = GA4_MAX_DIMENSIONS - (1 if pt_dim else 0)

        # Header for this bucket's logs.
        if bucket['kind'] == 'exact':
            bucket_label = (f"platform={platform_raw!r}   "
                            f"events={len(bucket['events'])} (exact)")
        else:
            mk = bucket['match_kind']
            mp = bucket['match_payload']
            mp_str = mp.pattern if mk == 'regex' else mp
            ga4_op = {'begins': 'ENDS_WITH', 'ends': 'BEGINS_WITH',
                      'regex': 'FULL_REGEXP'}[mk]
            bucket_label = (f"platform={platform_raw!r}   "
                            f"event=PATTERN {bucket['sheet_event']!r}  "
                            f"({ga4_op} {mp_str!r})")
        if pt_dim:
            bucket_label += f"   pageType filter: {pt_dim} IN {list(pt_values)}"

        print(f"\n[INFO] {bucket_label}")
        print(f"       Streams : {streams}")

        # eventName-only probe: when there are no event-scoped dims to
        # query, we still need GA4 to confirm which events actually exist
        # in the date range (and for pattern buckets, to discover WHICH
        # events match). Item-scoped dims aren't queried here -- they
        # have their own loop below.
        if request_eventname and not dims:
            print(f"       (eventName-only probe)")
            try:
                probe = RunReportRequest(
                    property=f"properties/{property_id}",
                    dimensions=[Dimension(name="eventName")],
                    metrics=[Metric(name="eventCount")],
                    date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
                    dimension_filter=combined_filter,
                    limit=100000,
                )
                resp = client.run_report(probe)
                for row in resp.rows:
                    ga4_ev = row.dimension_values[0].value.strip()
                    se = _ga4_event_to_sheet_event(bucket, ga4_ev)
                    if se is None:
                        continue
                    key = (platform_lc, module_key, se)
                    events_found.add(key)
                    # Pattern buckets that requested eventName as a dim:
                    # collect the matched real event names into the cell.
                    if bucket['kind'] == 'pattern':
                        actual[key].setdefault("eventName", set()).add(ga4_ev)
            except Exception as e:
                print(f"  [!] eventName probe failed: {e}")
            time.sleep(0.3)
            # NOTE: do NOT `continue` here -- fall through so the
            # item-scoped block below still runs if item_dims is non-empty.

        if dims:
            batches = [dims[i:i + max_dims] for i in range(0, len(dims), max_dims)]
            print(f"       Dims    : {len(dims)} (in {len(batches)} batch(es), "
                  f"max {max_dims}/batch)")
        else:
            batches = []

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
                    ga4_ev = row.dimension_values[0].value.strip()
                    se = _ga4_event_to_sheet_event(bucket, ga4_ev)
                    if se is None:
                        continue
                    key = (platform_lc, module_key, se)
                    events_found.add(key)
                    # Pattern buckets accumulate the matched real event
                    # names under the eventName dim (when requested).
                    if bucket['kind'] == 'pattern' and request_eventname:
                        actual[key].setdefault("eventName", set()).add(ga4_ev)
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
                            ga4_ev = row.dimension_values[0].value.strip()
                            se = _ga4_event_to_sheet_event(bucket, ga4_ev)
                            if se is None:
                                continue
                            key = (platform_lc, module_key, se)
                            events_found.add(key)
                            if bucket['kind'] == 'pattern' and request_eventname:
                                actual[key].setdefault("eventName", set()).add(ga4_ev)
                            v = row.dimension_values[1].value.strip()
                            if v and v != "(not set)":
                                actual[key].setdefault(d, set()).add(v)
                    except Exception:
                        print(f"    [X] Invalid dimension (rejected by GA4): {d}")
                        invalid_dims.add(d)
            time.sleep(0.3)

        # ----------------------------------------------------------------- #
        # Item-scoped query path
        # ----------------------------------------------------------------- #
        # Item dimensions (itemId, itemName, itemCategory*, etc.) are not
        # compatible with eventCount -- GA4 returns 400 "remove eventCount
        # to make the request compatible". We re-query just those dims
        # using `itemsViewed` as the metric, which IS item-scoped and
        # joins cleanly with item dims.
        #
        # The same eventName + streamName + pageType filter applies, so
        # data stays scoped to the right events on the right platform.
        # eventName is the grouping dim (same as the event-scoped path),
        # which lets us route each result row back to its sheet_event key.
        #
        # itemsViewed is logged on impression-style ecommerce events
        # (view_item, view_item_list, view_promotion, etc.). For other
        # ecommerce events (purchase, add_to_cart, etc.) we may also try
        # itemQuantity if the first metric returns nothing, but starting
        # with itemsViewed is the closest analog to eventCount.
        if item_dims:
            item_batches = [item_dims[i:i + max_dims]
                            for i in range(0, len(item_dims), max_dims)]
            print(f"       Item dims: {len(item_dims)} (in "
                  f"{len(item_batches)} batch(es), via itemsViewed metric)")

            for idx, batch in enumerate(item_batches, 1):
                print(f"  -> Item batch {idx}/{len(item_batches)}: {batch}")
                try:
                    req = RunReportRequest(
                        property=f"properties/{property_id}",
                        dimensions=[Dimension(name="eventName")]
                                   + [Dimension(name=d) for d in batch],
                        metrics=[Metric(name="itemsViewed")],
                        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
                        dimension_filter=combined_filter,
                        limit=100000,
                    )
                    resp = client.run_report(req)
                    for row in resp.rows:
                        ga4_ev = row.dimension_values[0].value.strip()
                        se = _ga4_event_to_sheet_event(bucket, ga4_ev)
                        if se is None:
                            continue
                        key = (platform_lc, module_key, se)
                        events_found.add(key)
                        if bucket['kind'] == 'pattern' and request_eventname:
                            actual[key].setdefault("eventName", set()).add(ga4_ev)
                        for i, d in enumerate(batch):
                            v = row.dimension_values[i + 1].value.strip()
                            if v and v != "(not set)":
                                actual[key].setdefault(d, set()).add(v)
                except Exception as e:
                    # An entire item-batch can fail for two reasons we care
                    # about: (1) a single dim in the batch is unregistered
                    # in this property, or (2) itemsViewed itself isn't
                    # populated for the matched events (rare). Retry each
                    # dim individually so one bad dim doesn't poison the
                    # batch -- same retry pattern as the event-scoped path.
                    print(f"  [!] Item batch failed: {e} -- retrying individually...")
                    for d in batch:
                        try:
                            req = RunReportRequest(
                                property=f"properties/{property_id}",
                                dimensions=[Dimension(name="eventName"),
                                            Dimension(name=d)],
                                metrics=[Metric(name="itemsViewed")],
                                date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
                                dimension_filter=combined_filter,
                                limit=100000,
                            )
                            resp = client.run_report(req)
                            for row in resp.rows:
                                ga4_ev = row.dimension_values[0].value.strip()
                                se = _ga4_event_to_sheet_event(bucket, ga4_ev)
                                if se is None:
                                    continue
                                key = (platform_lc, module_key, se)
                                events_found.add(key)
                                if bucket['kind'] == 'pattern' and request_eventname:
                                    actual[key].setdefault("eventName", set()).add(ga4_ev)
                                v = row.dimension_values[1].value.strip()
                                if v and v != "(not set)":
                                    actual[key].setdefault(d, set()).add(v)
                        except Exception:
                            print(f"    [X] Invalid item dimension (rejected by GA4): {d}")
                            invalid_dims.add(d)
                time.sleep(0.3)

    if invalid_dims:
        print(f"[WARN] Dimensions rejected by GA4 at runtime: {invalid_dims}")

    return actual, events_found, invalid_dims, skipped_platforms, skipped_group_keys


# --------------------------------------------------------------------------- #
# Fill — writes ONLY into ga4_actual_value, never touches rules_expected_values
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


def _row_is_processable(row, scope_mask, idx, ignore_required: bool):
    """
    Quick pre-filter shared by fill / classify / report. Returns True if the
    row is in scope, has a usable api_column, and isn't already filled.

    `ignore_required` controls TWO things:
      1. Whether required=yes is enforced (default mode does, ignore-required
         mode skips this check).
      2. Whether anchorless-only template rows are skipped. Default mode
         skips them (they'd dump every value into one cell). Ignore-required
         mode KEEPS them (dumping every value is the whole point).
    """
    if scope_mask is not None and not bool(scope_mask.loc[idx]):
        return False
    if not ignore_required and row["_required_parsed"] is not True:
        return False
    if not is_blank(row[COL_GA4_VALUES]):
        return False
    api_dim = row["_api_clean"]
    if not isinstance(api_dim, str) or not api_dim:
        return False
    # Anchorless-only check is a default-mode-only guard. In ignore-required
    # mode we want exactly the dump that the guard prevents.
    if not ignore_required and _is_anchorless_only(row[COL_RULES]):
        return False
    return True


def fill_ga4_values(df: pd.DataFrame, actual, events_found, invalid_dims,
                    skipped_group_keys=None, scope_mask=None,
                    ignore_required: bool = False):
    """
    Mutates df in place: writes pipe-joined values into ga4_actual_value
    for every eligible row that doesn't already have a value.

    `scope_mask` (optional) is a boolean Series. When provided, only rows
    where the mask is True are touched -- rows outside the mask are left
    completely untouched.

    `ignore_required=True` (set by --ignore-required) does TWO things:
      1. Bypasses the required=yes eligibility gate (rows with required=no
         and required=blank are also processed).
      2. Treats anchorless `<<...>>` templates as "no filter" instead of
         skipping them. The default-mode skip exists because anchorless
         dumps every value into one cell, which is rarely useful for
         validation. In ignore-required mode that dump IS the goal --
         you want to see every value so you can write proper rules.

    `skipped_group_keys` is the set of (platform_lc, module_key, sheet_event) keys
    that the bucket-builder/fetcher refused to query. Two reasons cluster
    here -- bad pageType filter values, and skip-classified event_names
    (e.g. pure-anchorless `<<x>>`). Affected rows are counted under
    skipped_group rather than as event-not-found, since the underlying
    cause is a sheet-level config issue, not missing GA4 data.

    Status (`ga4_check_status`) is intentionally NOT written here. It's
    computed in a separate post-processing pass (`mark_checked_status`)
    so that "checked" tracks whether GA4 actually returned data for the
    row, including data from earlier runs that's already in the file.

    Returns counters:
        (filled, no_data, invalid_dim, event_nf,
         skipped_no_api, skipped_anchorless, skipped_group, truncated)

    Note: skipped_anchorless is always 0 when ignore_required=True
    (anchorless rows are processed in that mode, not skipped).
    """
    if skipped_group_keys is None:
        skipped_group_keys = set()

    filled = no_data = invalid_dim = event_nf = 0
    skipped_no_api = skipped_anchorless = skipped_group = 0
    truncated = 0

    for i, row in df.iterrows():
        # Per-row in-scope checks (mask, required, already-filled).
        if scope_mask is not None and not bool(scope_mask.loc[i]):
            continue
        if not ignore_required and row["_required_parsed"] is not True:
            continue
        # Retry safety: never overwrite a row that already has a value.
        if not is_blank(row[COL_GA4_VALUES]):
            continue

        api_dim = row["_api_clean"]
        if not isinstance(api_dim, str) or not api_dim:
            skipped_no_api += 1
            continue
        # Anchorless-only guard is default-mode-only. In ignore-required
        # mode, anchorless rows are exactly what we want to fill.
        if not ignore_required and _is_anchorless_only(row[COL_RULES]):
            skipped_anchorless += 1
            continue

        platform_lc = str(row[COL_PLATFORM]).strip().lower()
        ev = str(row[COL_EVENT]).strip()
        # Pull the module value the same way the bucket builder did, so
        # the lookup key matches exactly. Missing/blank module column ->
        # empty string '' (same default as the builder).
        if COL_MODULE in row.index:
            module_key = str(row[COL_MODULE]).strip() if not pd.isna(row[COL_MODULE]) else ""
        else:
            module_key = ""
        key = (platform_lc, module_key, ev)

        # Group-level skip (bad pageType filter or unusable event_name)
        # takes precedence over event-not-found, because the underlying
        # reason for "no data" is a sheet-level config issue.
        if key in skipped_group_keys:
            skipped_group += 1
            continue
        if api_dim in invalid_dims:
            invalid_dim += 1
            continue
        if key not in events_found:
            event_nf += 1
            continue

        ga4_values = actual.get(key, {}).get(api_dim, set())
        # In ignore-required mode, anchorless templates expand to "all
        # values" instead of returning empty -- this is the visible
        # behaviour change the user sees in their output.
        expanded = expand_cell(
            row[COL_RULES], ga4_values,
            treat_anchorless_as_all=ignore_required,
        )
        if not expanded:
            no_data += 1
            continue

        sorted_values = sorted(expanded)
        joined = "|".join(sorted_values)
        capped = cap_for_excel(joined, sorted_values)
        if capped != joined:
            truncated += 1
        df.at[i, COL_GA4_VALUES] = capped
        filled += 1

    return (filled, no_data, invalid_dim, event_nf,
            skipped_no_api, skipped_anchorless, skipped_group, truncated)


# Parameter-name value that identifies the event-level row (the row that
# represents the event itself rather than one of its parameters). Used by
# the rollup step in mark_checked_status. Compared case-insensitively.
EVENT_LEVEL_PARAM_NAME = "event"


def mark_checked_status(df: pd.DataFrame, scope_mask=None) -> tuple:
    """
    Walk in-scope rows and mark `ga4_check_status` according to the policy:

      1. PARAMETER ROW: marked 'checked' iff its ga4_actual_value is
         non-empty. (i.e. this row's GA4 query actually returned a value.)
      2. EVENT-LEVEL ROW (parameter_name == 'event'): marked 'checked' iff
         AT LEAST ONE parameter row of the same (event_name, platform) is
         itself 'checked' by rule 1.

    Out-of-scope rows are never touched. Existing 'checked' status from
    earlier runs on out-of-scope rows is preserved as-is.

    This function is idempotent inside scope: re-running recomputes
    'checked' from the current state of ga4_actual_value, so the column
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
# COL_CONFLUENCE ('confluence values') is the user's own reference column,
# preserved untouched on writeback -- the script never reads or modifies it.
MAIN_COLUMN_ORDER = [
    "link", COL_MODULE, COL_EVENT, COL_PLATFORM, COL_PARAM,
    COL_BQ, COL_API, COL_REQUIRED, COL_PAGETYPE_CK,
    COL_RULES, COL_CONFLUENCE, COL_GA4_VALUES, COL_STATUS,
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
        description="Fetch GA4 values and fill the ga4_actual_value column. "
                    "Default mode writes back to --input. --ignore-required "
                    "writes to a separate ignore_required.<ext> file."
    )
    ap.add_argument("--input",       default=r"C:\Users\kisbhat\Downloads\GA4 Automation\data_check.csv",
                    help="Input Excel/CSV file. In default mode this is "
                         "overwritten in place (unless --output is given). "
                         "In --ignore-required mode it is read but never "
                         "modified -- output goes to ignore_required.<ext> "
                         "in the same directory. Default points to "
                         "data_check.csv in the GA4 Automation folder.")
    ap.add_argument("--date",        required=True,
                    help="Date range DDMMYYYY-DDMMYYYY  e.g. 05042026-07042026")
    ap.add_argument("--module",      help="Optional comma-separated module filter "
                                          "(substring match)")
    ap.add_argument("--platform",    help="Optional comma-separated platform filter "
                                          "(exact match, case-insensitive). "
                                          "E.g. 'App' or 'App,Web'.")
    ap.add_argument("--property",    default=os.environ.get("GA4_PROPERTY_ID")
                                          or os.environ.get("GA_PROPERTY_ID"),
                    help="GA4 Property ID (default: GA4_PROPERTY_ID / "
                         "GA_PROPERTY_ID from .env)")
    ap.add_argument("--credentials", default=os.environ.get("GA4_CREDENTIALS")
                                          or os.environ.get("GA_CREDENTIALS_PATH")
                                          or "prod.json",
                    help="Service-account JSON path (default: prod.json, "
                         "or GA4_CREDENTIALS / GA_CREDENTIALS_PATH from .env)")
    ap.add_argument("--env",         choices=["prod", "nonprod"], default="prod",
                    help="Which stream mapping to use (default: prod). "
                         "See STREAM_MAPS in script.")
    ap.add_argument("--output",      help="Output file path. "
                                          "Default mode default: overwrite --input. "
                                          "Ignore-required mode default: "
                                          "ignore_required.<ext> in input's directory.")
    ap.add_argument("--sheet",       default=DEFAULT_SHEET_NAME,
                    help=f"Excel sheet name to read (case-insensitive, "
                         f"whitespace-tolerant). Default: "
                         f"{DEFAULT_SHEET_NAME!r}. Ignored for CSV inputs.")
    ap.add_argument("--dry-run",     action="store_true",
                    help="Show plan without calling GA4 or writing anything.")

    # ---- Ignore-required mode ----
    # Renamed from --triage. Two behavioural changes vs default mode:
    #   1. The `required` column is ignored as an eligibility gate.
    #   2. Anchorless `<<...>>` templates no longer skip the row -- they
    #      expand to "all GA4 values" so you can see the full data and
    #      write proper rules afterwards.
    # pageType filtering still applies. Output goes to a separate file
    # so the curated default-mode data is never touched.
    ap.add_argument("--ignore-required", action="store_true",
                    help="Fetch GA4 values ignoring both the 'required' column "
                         "and the default skip on anchorless `<<...>>` templates. "
                         "pageType filter still applies. Output is written to "
                         "ignore_required.<ext> in the input's directory (or "
                         "to --output if given). The `required` column is NOT "
                         "modified -- update it manually after reviewing the "
                         "fetched data.")
    return ap.parse_args()


def main():
    args = parse_args()

    # ---- Mode selection ----
    # Two modes now (the old --triage / --triage-report split has been
    # collapsed into a single --ignore-required mode; auto-classification
    # of `required` was removed and is deferred for later).
    mode = "ignore_required" if args.ignore_required else "default"

    start_date, end_date = parse_date_range(args.date)

    # ---- Output path resolution ----
    # Default mode: overwrite --input unless --output overrides.
    # Ignore-required mode: write to ignore_required.<ext> in input's
    # directory unless --output overrides. Never overwrites the input,
    # so the curated default-mode file stays clean.
    if args.output:
        output_path = args.output
    elif mode == "ignore_required":
        input_dir = os.path.dirname(os.path.abspath(args.input))
        ext = os.path.splitext(args.input)[1] or ".xlsx"
        output_path = os.path.join(input_dir, f"ignore_required{ext}")
    else:
        output_path = args.input

    # Friendly tags for the output-line summary.
    if output_path == args.input:
        out_tag = "   (in-place overwrite)"
    elif mode == "ignore_required":
        out_tag = "   (ignore-required mode -- input file not modified)"
    else:
        out_tag = ""

    print(f"[INFO] Mode        : {mode}")
    print(f"[INFO] Date range  : {start_date} -> {end_date}")
    print(f"[INFO] Input file  : {args.input}")
    print(f"[INFO] Output file : {output_path}{out_tag}")
    print(f"[INFO] Environment : {args.env}")
    print(f"[INFO] Dry run     : {args.dry_run}")
    print("")

    df, other_sheets, main_sheet_name = read_input(args.input, sheet_name=args.sheet)

    # Build scope mask = module filter AND platform filter.
    module_mask   = apply_module_mask(df, args.module)
    platform_mask = apply_platform_mask(df, args.platform)
    in_scope      = module_mask & platform_mask
    if args.module or args.platform:
        print(f"[INFO] Combined scope: {int(in_scope.sum())} / {len(df)} row(s)")

    platform_to_streams = STREAM_MAPS[args.env]
    print(f"[INFO] Stream mapping ({args.env}):")
    for plat, streams in platform_to_streams.items():
        print(f"         {plat!r:>8} -> {streams}")

    # --- Pre-flight row counts (scoped to module + platform filters) ---
    blank_required = int((in_scope & df["_required_parsed"].isna()).sum())
    skipped_no     = int((in_scope & (df["_required_parsed"] == False)).sum())

    if mode == "default":
        if blank_required:
            print(f"[WARN] {blank_required} row(s) have a blank '{COL_REQUIRED}' "
                  f"column -- left untouched.")
        if skipped_no:
            print(f"[INFO] {skipped_no} row(s) marked required=no -- left untouched.")
    else:
        # ignore_required mode: required column not used as a gate.
        print(f"[INFO] ignore-required mode: '{COL_REQUIRED}' column ignored, "
              f"anchorless `<<...>>` templates expand to all GA4 values")
        print(f"         in-scope rows with required=yes  : "
              f"{int((in_scope & (df['_required_parsed'] == True)).sum())}")
        print(f"         in-scope rows with required=no   : {skipped_no}")
        print(f"         in-scope rows with required blank: {blank_required}")

    # Rows skipped because api_column is blank / Not Found / malformed.
    if mode == "default":
        invalid_api_mask = in_scope & (df["_required_parsed"] == True) & df["_api_clean"].isna()
    else:
        invalid_api_mask = in_scope & df["_api_clean"].isna()
    invalid_api_count = int(invalid_api_mask.sum())
    if invalid_api_count:
        print(f"[INFO] {invalid_api_count} in-scope row(s) have an unusable "
              f"'{COL_API}' (blank / 'Not Found' / malformed) -- skipped.")

    # Rows already filled in ga4_actual_value are skipped (retry-safe).
    if mode == "default":
        already_filled = int((in_scope &
                              (df["_required_parsed"] == True) &
                              ~df[COL_GA4_VALUES].apply(is_blank)).sum())
    else:
        already_filled = int((in_scope &
                              ~df[COL_GA4_VALUES].apply(is_blank)).sum())
    if already_filled:
        print(f"[INFO] {already_filled} in-scope row(s) already have "
              f"'{COL_GA4_VALUES}' set -- skipped (retry-safe).")

    # Mode-aware eligibility for fetch.
    if mode == "default":
        eligible_mask = (
            in_scope &
            (df["_required_parsed"] == True) &
            df["_api_clean"].notna() &
            df[COL_GA4_VALUES].apply(is_blank)
        )
    else:  # ignore_required
        eligible_mask = (
            in_scope &
            df["_api_clean"].notna() &
            df[COL_GA4_VALUES].apply(is_blank)
        )
    eligible = df[eligible_mask]

    # Templates note (default mode only).
    if mode == "default":
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

    # Pre-compute pageType filter outcome for the dry-run plan / heads-up.
    # (fetch_from_ga4 will recompute this, but showing it here lets you see
    # the breakdown without making any API calls.)
    if not eligible.empty:
        _, pre_skip_groups, _ = _build_query_buckets(
            eligible, platform_to_streams, env=args.env)
        if pre_skip_groups:
            print(f"[WARN] {len(pre_skip_groups)} (event, platform, module) group(s) "
                  f"will be SKIPPED due to bad pageType filter:")
            for ev, plat, module_key, reason in pre_skip_groups[:10]:
                module_disp = f"  [module={module_key!r}]" if module_key else ""
                print(f"         {plat!r:>8} / {ev!r}{module_disp}  --  {reason}")
            if len(pre_skip_groups) > 10:
                print(f"         ... and {len(pre_skip_groups) - 10} more")

    if args.dry_run:
        print("")
        for plat, (evs, dims) in platform_breakdown.items():
            print(f"[DRY RUN] platform={plat!r}")
            print(f"          streams: {platform_to_streams.get(plat.lower(), '<none — would skip>')}")
            print(f"          events:  {evs}")
            print(f"          dims:    {dims}")
        if mode == "ignore_required":
            print(f"[DRY RUN] Mode is --ignore-required: would fetch with "
                  f"`required` and anchorless-template skip ignored, write to "
                  f"{output_path!r}.")
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
        if mode == "default":
            eligible_mask = (
                in_scope &
                (df["_required_parsed"] == True) &
                df["_api_clean"].notna() &
                df[COL_GA4_VALUES].apply(is_blank)
            )
        else:
            eligible_mask = (
                in_scope &
                df["_api_clean"].notna() &
                df[COL_GA4_VALUES].apply(is_blank)
            )
        eligible = df[eligible_mask]

        actual, found, invalid_dims, skipped_platforms, skipped_group_keys = \
            fetch_from_ga4(client, args.property, eligible, platform_to_streams,
                           start_date, end_date, env=args.env)
    except Exception as e:
        sys.exit(f"[ERROR] GA4 authentication/fetch failed: {e}")

    # Fill ga4_actual_value. ignore_required=True flips both the required
    # eligibility gate and the anchorless-template skip (see fill_ga4_values).
    ignore_required = (mode == "ignore_required")
    (filled, no_data, inv_dim, ev_nf,
     skp_no_api, skp_anchorless, skp_group, truncated) = fill_ga4_values(
        df, actual, found, invalid_dims,
        skipped_group_keys=skipped_group_keys,
        scope_mask=in_scope,
        ignore_required=ignore_required,
    )

    # Compute / refresh the ga4_check_status column based on what's now
    # in ga4_actual_value (including data from prior runs already on disk).
    param_checked, event_checked = mark_checked_status(df, scope_mask=in_scope)

    write_output_atomic(df, output_path, other_sheets, main_sheet_name)

    attempted = filled + no_data + inv_dim + ev_nf + skp_no_api + skp_anchorless + skp_group

    print("")
    print("=" * 60)
    print(f"  FETCH-AND-FILL SUMMARY  ({mode} mode)")
    print("=" * 60)
    print(f"  Rows attempted in this run                 : {attempted}")
    print(f"    -> filled with GA4 values                : {filled}")
    print(f"    -> queried but no data                   : {no_data}")
    print(f"    -> event not found in date range         : {ev_nf}")
    print(f"    -> dimension rejected by GA4             : {inv_dim}")
    print(f"    -> skipped (api_column unusable)         : {skp_no_api}")
    if mode == "default":
        # Anchorless-skip count is only meaningful in default mode -- in
        # ignore-required mode it's always 0 by design.
        print(f"    -> skipped (anchorless rules)            : {skp_anchorless}")
    # Group-level skip covers two reasons: bad pageType filter spec, and
    # event_name patterns we couldn't translate (e.g. fully-anchorless
    # `<<x>>`). Detailed reasons were logged earlier in the run.
    print(f"    -> skipped (bad pageType / event_name)   : {skp_group}")
    if truncated:
        print(f"  Rows truncated to Excel cell limit  : {truncated}")
        print(f"    (cells over {EXCEL_MAX_CELL} chars are capped at value")
        print(f"     boundaries and end with {TRUNCATION_MARKER!r})")
    print(f"  Rows already had ga4_actual_value   : {already_filled}")
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