#!/usr/bin/env python3
"""
List every dimension registered in a GA4 property and dump to CSV.

Useful for figuring out the correct `api_column` values to put in your
QA tracker — especially custom dimensions (customEvent:* / customItem:*
/ customUser:*).

Run:
    python discover_dimensions.py

Outputs:
    ga4_dimensions.csv   — full list, sortable in Excel
    Console               — custom dimensions printed for quick scan

Defaults to property 195772067 with credentials nonprodv2.json.
Override with --property and --credentials.
"""

import argparse
import csv
import os
import sys


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--property",    default=os.environ.get("GA4_PROPERTY_ID", "195772067"))
    ap.add_argument("--credentials", default=os.environ.get("GA4_CREDENTIALS", "nonprodv2.json"))
    ap.add_argument("--output",      default="ga4_dimensions.csv")
    args = ap.parse_args()

    try:
        from google.oauth2 import service_account
        from google.analytics.data_v1beta import BetaAnalyticsDataClient
    except ImportError:
        sys.exit("[ERROR] Install dependencies first: "
                 "pip install google-analytics-data google-auth")

    print(f"[INFO] Property    : {args.property}")
    print(f"[INFO] Credentials : {args.credentials}")

    creds = service_account.Credentials.from_service_account_file(
        args.credentials,
        scopes=["https://www.googleapis.com/auth/analytics.readonly"],
    )
    client = BetaAnalyticsDataClient(credentials=creds)

    metadata = client.get_metadata(name=f"properties/{args.property}/metadata")
    dims = list(metadata.dimensions)
    print(f"[INFO] Total dimensions: {len(dims)}")

    # Classify
    custom_event = []
    custom_item  = []
    custom_user  = []
    standard     = []
    for d in dims:
        api_name = d.api_name
        if api_name.startswith("customEvent:"):
            custom_event.append(d)
        elif api_name.startswith("customItem:"):
            custom_item.append(d)
        elif api_name.startswith("customUser:"):
            custom_user.append(d)
        else:
            standard.append(d)

    # Console summary -- focus on custom dims since those are usually
    # what you map in api_column
    def print_group(title, group):
        if not group:
            return
        print(f"\n{title} ({len(group)}):")
        for d in sorted(group, key=lambda x: x.api_name.lower()):
            ui = d.ui_name or ""
            print(f"  {d.api_name:<55s}  {ui}")

    print_group("=== customEvent: (event-scoped — works with eventCount) ===", custom_event)
    print_group("=== customItem:  (item-scoped — needs itemCount/itemRevenue, NOT eventCount) ===", custom_item)
    print_group("=== customUser:  (user-scoped) ===", custom_user)
    print(f"\n=== standard dimensions: {len(standard)} (see CSV for full list)")

    # Write everything to CSV
    with open(args.output, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["api_name", "ui_name", "category", "scope", "description"])
        for d in sorted(dims, key=lambda x: x.api_name.lower()):
            scope = (
                "customEvent" if d.api_name.startswith("customEvent:") else
                "customItem"  if d.api_name.startswith("customItem:")  else
                "customUser"  if d.api_name.startswith("customUser:")  else
                "standard"
            )
            w.writerow([d.api_name, d.ui_name, d.category, scope, d.description])
    print(f"\n[INFO] Wrote {len(dims)} rows -> {args.output}")
    print("       Open this in Excel and search/filter to find the api_name "
          "you should use in your QA tracker's `api_column` column.")


if __name__ == "__main__":
    main()
