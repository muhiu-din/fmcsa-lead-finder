"""
FMCSA Active/Interstate Carrier Lead Finder
--------------------------------------------
Scans a range of USDOT numbers against the official, free FMCSA QCMobile API
(https://mobile.fmcsa.dot.gov) and saves every carrier that is:
  - allowedToOperate == "Y"   (active operating status)
  - carrier operation classification == Interstate

...into an Excel file you can use as a cold-call list.

WHY THIS APPROACH (instead of scraping the SAFER website):
  - SAFER's HTML page isn't meant for automated querying and can block/rate-limit you.
  - MC/MX numbers were retired by FMCSA on Oct 1, 2025 — new authority is now
    issued under the USDOT number, so scanning MC numbers would miss the newest
    (and most cold-call-worthy) carriers. USDOT numbers are the correct target now.
  - The QCMobile API is FMCSA's own public, documented, JSON API for exactly
    this purpose — no HTML parsing, no guessing whether a "hit" means active.

SETUP (one-time):
  1. pip install requests openpyxl
  2. Get a free WebKey: https://mobile.fmcsa.dot.gov/QCDevsite -> log in with
     Login.gov -> My WebKeys -> Get a new WebKey
  3. Paste it into WEBKEY below (or pass --webkey on the command line)

USAGE:
  # Recommended: scan by current USDOT number
  python fmcsa_lead_finder.py --webkey YOUR_KEY --mode dot --start 4000000 --count 200

  # Legacy: scan by MC/MX docket number (only reaches carriers registered
  # before MC numbers were retired on Oct 1, 2025 — use --mode dot for fresh leads)
  python fmcsa_lead_finder.py --webkey YOUR_KEY --mode mc --start 900000 --count 200

  --mode    "dot" (default, recommended) or "mc"
  --start   number to begin scanning from (DOT number, or MC/MX docket number if --mode mc)
  --count   how many numbers to check (NOT how many leads you'll get —
            only a fraction will be active+interstate, so scan generously)
  --output  output .xlsx filename (default: fmcsa_leads.xlsx)
  --delay   seconds to wait between API calls (default: 0.3 — be polite to
            FMCSA's free public API, don't hammer it)

The script saves progress to the Excel file every 25 hits, so if it's
interrupted you still keep what it found so far.
"""

import argparse
import sys
import time
from datetime import datetime

import requests
from openpyxl import Workbook

BASE_URL = "https://mobile.fmcsa.dot.gov/qc/services/carriers/{dot}"
DOCKET_URL = "https://mobile.fmcsa.dot.gov/qc/services/carriers/docket-number/{docket}"

COLUMNS = [
    "DOT Number",
    "MC Number(s)",
    "Legal Name",
    "DBA Name",
    "Phone",
    "City",
    "State",
    "Street",
    "Operation Classification",
    "Allowed To Operate",
]


def _handle_response(resp, label):
    """Shared status-code handling. Returns parsed JSON dict, or None to skip."""
    if resp.status_code in (401, 403):
        print("\nFATAL: FMCSA rejected the WebKey (401/403). Double-check the "
              "key you got from mobile.fmcsa.dot.gov and try again.")
        sys.exit(1)

    if resp.status_code == 429:
        print(f"  [{label}] rate-limited (429) — backing off 5s")
        time.sleep(5)
        return None

    if resp.status_code != 200:
        # Most commonly: no carrier registered under this number.
        return None

    try:
        return resp.json()
    except ValueError:
        return None


def _filter_carrier(carrier: dict, fallback_dot=None):
    """Given one raw carrier dict from the API, return our row dict if it
    passes the active+interstate filter, else None."""
    if not carrier or not isinstance(carrier, dict):
        return None

    allowed = str(carrier.get("allowedToOperate", "")).strip().upper()
    if allowed != "Y":
        return None

    op = carrier.get("carrierOperation", {})
    op_desc = ""
    if isinstance(op, dict):
        op_desc = str(op.get("carrierOperationDesc", op.get("description", ""))).strip()
    elif isinstance(op, str):
        op_desc = op.strip()

    if "inter" not in op_desc.lower():
        return None

    return {
        "DOT Number": carrier.get("dotNumber", fallback_dot),
        "MC Number(s)": carrier.get("mcNumber", carrier.get("docketNumber", "")),
        "Legal Name": carrier.get("legalName", ""),
        "DBA Name": carrier.get("dbaName", ""),
        "Phone": carrier.get("telephone", carrier.get("phoneNumber", "")),
        "City": carrier.get("phyCity", ""),
        "State": carrier.get("phyState", ""),
        "Street": carrier.get("phyStreet", ""),
        "Operation Classification": op_desc,
        "Allowed To Operate": allowed,
    }


def fetch_carrier(dot_number: int, webkey: str, session: requests.Session):
    """Query one DOT number. Returns a dict of carrier fields, or None if
    no record exists / not active+interstate / request failed."""
    url = BASE_URL.format(dot=dot_number)
    try:
        resp = session.get(url, params={"webKey": webkey}, timeout=15)
    except requests.RequestException as e:
        print(f"  [DOT {dot_number}] network error: {e} — skipping")
        return None

    data = _handle_response(resp, f"DOT {dot_number}")
    if data is None:
        return None

    carrier = None
    if isinstance(data, dict):
        content = data.get("content", data)
        if isinstance(content, dict):
            carrier = content.get("carrier", content)

    return _filter_carrier(carrier, fallback_dot=dot_number)


def fetch_carriers_by_docket(docket_number: int, webkey: str, session: requests.Session):
    """Query one MC/MX/docket number. A docket number can map to more than
    one carrier record historically, so this returns a LIST of row dicts
    (usually 0 or 1 item)."""
    url = DOCKET_URL.format(docket=docket_number)
    try:
        resp = session.get(url, params={"webKey": webkey}, timeout=15)
    except requests.RequestException as e:
        print(f"  [MC {docket_number}] network error: {e} — skipping")
        return []

    data = _handle_response(resp, f"MC {docket_number}")
    if data is None:
        return []

    content = data.get("content", data) if isinstance(data, dict) else data

    # This endpoint can come back as a single object or a list of objects
    # depending on how many carriers are tied to the docket number.
    raw_list = content if isinstance(content, list) else [content]

    results = []
    for item in raw_list:
        carrier = item.get("carrier", item) if isinstance(item, dict) else None
        row = _filter_carrier(carrier)
        if row:
            if not row["MC Number(s)"]:
                row["MC Number(s)"] = docket_number
            results.append(row)
    return results


def main():
    parser = argparse.ArgumentParser(description="Find active/interstate FMCSA carrier leads")
    parser.add_argument("--webkey", required=True, help="FMCSA QCMobile WebKey")
    parser.add_argument("--mode", choices=["dot", "mc"], default="dot",
                         help="'dot' scans USDOT numbers (recommended, current). "
                              "'mc' scans legacy MC/MX docket numbers.")
    parser.add_argument("--start", type=int, required=True,
                         help="Number to start scanning from (DOT number or MC/MX docket number, per --mode)")
    parser.add_argument("--count", type=int, default=200, help="How many numbers to check")
    parser.add_argument("--output", default="fmcsa_leads.xlsx", help="Output Excel filename")
    parser.add_argument("--delay", type=float, default=0.3, help="Seconds between requests")
    args = parser.parse_args()

    if args.mode == "mc":
        print("NOTE: MC/MX docket numbers were retired as of Oct 1, 2025 — new "
              "carrier authority no longer gets one. This mode only reaches "
              "carriers that registered before that date. For the freshest "
              "leads, prefer --mode dot.\n")

    wb = Workbook()
    ws = wb.active
    ws.title = "Active Interstate Carriers"
    ws.append(COLUMNS)

    session = requests.Session()
    hits = 0
    checked = 0
    start_time = datetime.now()

    label = "DOT" if args.mode == "dot" else "MC/MX"
    print(f"Scanning {label} {args.start} .. {args.start + args.count - 1} "
          f"(delay={args.delay}s/call)\n")

    for number in range(args.start, args.start + args.count):
        checked += 1

        if args.mode == "dot":
            result = fetch_carrier(number, args.webkey, session)
            results = [result] if result else []
        else:
            results = fetch_carriers_by_docket(number, args.webkey, session)

        for result in results:
            hits += 1
            ws.append([result[col] for col in COLUMNS])
            print(f"  [{checked}/{args.count}] MATCH: {result['Legal Name']} "
                  f"({result['City']}, {result['State']}) — {label} {number}")

        if hits and hits % 25 == 0:
            wb.save(args.output)

        if checked % 50 == 0:
            elapsed = (datetime.now() - start_time).total_seconds()
            print(f"  ...progress: {checked}/{args.count} checked, "
                  f"{hits} leads so far, {elapsed:.0f}s elapsed")

        time.sleep(args.delay)

    wb.save(args.output)
    print(f"\nDone. Checked {checked} DOT numbers, found {hits} active/interstate "
          f"carriers. Saved to {args.output}")


if __name__ == "__main__":
    main()


#-------------------Command to start script-------------------

#     # Recommended — current, active carriers
# python fmcsa_lead_finder.py --webkey YOUR_KEY --mode dot --start 1599945 --count 200

# # Legacy MC/MX docket number scan
# python fmcsa_lead_finder.py --webkey YOUR_KEY --mode mc --start 1599945 --count 200