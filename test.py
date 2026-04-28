#!/usr/bin/env python3
"""
List every GA4 stream with its platform (iOS / Android / Web) and a sample
of recent event names. Use this to figure out which streams to put in the
STREAM_MAPS dict in ga4_fetch_fill.py.

Run:
    python discover_streams.py
    python discover_streams.py --days 30        # widen the window
    python discover_streams.py --show-events    # also list top events per stream

Defaults: property 195772067, credentials nonprodv2.json.
"""

import argparse
import os
import sys
from collections import defaultdict


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--property",    default=os.environ.get("GA4_PROPERTY_ID", "195772067"))
    ap.add_argument("--credentials", default=os.environ.get("GA4_CREDENTIALS", "nonprodv2.json"))
    ap.add_argument("--days",        type=int, default=14, help="Look-back window (default 14)")
    ap.add_argument("--show-events", action="store_true",
                    help="Also list top event names per stream")
    args = ap.parse_args()

    try:
        from google.oauth2 import service_account
        from google.analytics.data_v1beta import BetaAnalyticsDataClient
        from google.analytics.data_v1beta.types import (
            RunReportRequest, Dimension, Metric, DateRange,
        )
    except ImportError:
        sys.exit("[ERROR] pip install google-analytics-data google-auth")

    print(f"[INFO] Property : {args.property}")
    print(f"[INFO] Window   : last {args.days} days\n")

    creds = service_account.Credentials.from_service_account_file(
        args.credentials,
        scopes=["https://www.googleapis.com/auth/analytics.readonly"],
    )
    client = BetaAnalyticsDataClient(credentials=creds)

    # ---- Pass 1: streamName + platform + streamId ---------------------------
    resp = client.run_report(RunReportRequest(
        property=f"properties/{args.property}",
        dimensions=[
            Dimension(name="streamName"),
            Dimension(name="platform"),
            Dimension(name="streamId"),
        ],
        metrics=[Metric(name="eventCount")],
        date_ranges=[DateRange(start_date=f"{args.days}daysAgo", end_date="today")],
        limit=200,
    ))

    rows = []
    for r in resp.rows:
        rows.append({
            "stream":   r.dimension_values[0].value,
            "platform": r.dimension_values[1].value,
            "stream_id": r.dimension_values[2].value,
            "events":   int(r.metric_values[0].value),
        })

    if not rows:
        print(f"[WARN] No data in the last {args.days} days. Try --days 30 or 60.")
        return

    rows.sort(key=lambda x: x["events"], reverse=True)
    print(f"{'Stream Name':<35s} {'Platform':<12s} {'Stream ID':<14s} {'Events':>10s}")
    print("-" * 75)
    for r in rows:
        print(f"{r['stream']:<35s} {r['platform']:<12s} {r['stream_id']:<14s} {r['events']:>10,}")

    # ---- Auto-suggest STREAM_MAPS entry -------------------------------------
    by_platform = defaultdict(list)
    for r in rows:
        plat = r["platform"].lower()
        if plat in ("android", "ios"):
            by_platform["app"].append(r["stream"])
        elif plat == "web":
            by_platform["web"].append(r["stream"])
        else:
            by_platform[plat].append(r["stream"])

    print("\n--- Suggested STREAM_MAPS entry (review before using!) ---")
    print("STREAM_MAPS = {")
    print('    "nonprod": {')
    for plat, streams in by_platform.items():
        formatted = ", ".join(f'"{s}"' for s in streams)
        print(f'        "{plat}": [{formatted}],')
    print("    },")
    print("}")
    print("\n[!] You probably DON'T want every stream above. Common things to drop:")
    print("    * prod streams that leaked into the non-prod property")
    print("    * unrelated apps (e.g. Fantasy when you're testing the main F1 app)")
    print("    * old/deprecated streams")

    # ---- Optional: top events per stream ------------------------------------
    if args.show_events:
        print("\n--- Top events per stream ---")
        resp2 = client.run_report(RunReportRequest(
            property=f"properties/{args.property}",
            dimensions=[
                Dimension(name="streamName"),
                Dimension(name="eventName"),
            ],
            metrics=[Metric(name="eventCount")],
            date_ranges=[DateRange(start_date=f"{args.days}daysAgo", end_date="today")],
            limit=2000,
        ))
        per_stream = defaultdict(list)
        for r in resp2.rows:
            per_stream[r.dimension_values[0].value].append(
                (r.dimension_values[1].value, int(r.metric_values[0].value))
            )
        for stream, evs in per_stream.items():
            evs.sort(key=lambda x: x[1], reverse=True)
            print(f"\n  [{stream}]  top {min(10, len(evs))} of {len(evs)} events:")
            for ev, n in evs[:10]:
                print(f"      {ev:<45s} {n:>8,}")


if __name__ == "__main__":
    main()