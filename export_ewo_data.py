#!/usr/bin/env python3
"""
Build data.json for the EWO dashboard.

Takes either a local .xlsx path or a OneDrive "Anyone with the link" sharing
link. With a link it downloads the workbook the way a browser does, which is
the only route that still works: OneDrive's old anonymous API now answers
"unauthenticated", so the sharing link has to be redeemed into a session
first and the file fetched with those cookies.

A browser on another website cannot do this, because the download endpoint
sends no cross-origin headers. That is why this runs on a server, in GitHub
Actions or on a PC, rather than inside the page.

Only the columns the dashboard actually uses are written out, so data.json
stays small even when the workbook is large.

Usage
-----
    python export_ewo_data.py "https://1drv.ms/x/c/..../IQ....?e=xxxx"
    python export_ewo_data.py "C:\\path\\to\\EWO.xlsx" -o data.json

Requires openpyxl:  python -m pip install openpyxl
"""

import argparse
import datetime as dt
import http.cookiejar
import io
import json
import os
import re
import sys
import urllib.parse
import urllib.request

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed. Run:  python -m pip install openpyxl")

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")

SHEET_MAIN, SHEET_FABRIC, SHEET_SEW = "EWO", "Fsum", "Ssum"

# Exactly what index.html reads. Anything else is left out of data.json.
WANTED_COLUMNS = [
    "EWO", "Status", "Buyer", "Team", "Execution Unit",
    "Order Qty", "Fab Qty", "Order Bank Date", "Latest note", "Dept",
]
HEADER_MUST_CONTAIN = ["status", "buyer"]


# --------------------------------------------------------------------------
# Downloading a shared workbook
# --------------------------------------------------------------------------
class _RecordingRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Keeps every URL in the redirect chain, since the download address has
    to be assembled from parts that appear in the middle of it."""

    def __init__(self):
        self.chain = []

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        self.chain.append(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def download_shared_workbook(share_url):
    jar = http.cookiejar.CookieJar()
    recorder = _RecordingRedirectHandler()
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(jar), recorder)
    opener.addheaders = [("User-Agent", UA)]

    # Step 1: follow the sharing link so the anonymous grant becomes a cookie.
    try:
        with opener.open(share_url, timeout=90) as resp:
            final_url = resp.geturl()
            resp.read(2048)
    except Exception as exc:
        raise SystemExit("Could not open the sharing link: %s" % exc)

    chain = recorder.chain + [final_url]
    joined = " ".join(chain)

    # Step 2: dig the document id and the drive path out of the chain.
    doc = re.search(r"sourcedoc=%7B([0-9a-fA-F-]{36})%7D", joined) \
        or re.search(r"sourcedoc=\{([0-9a-fA-F-]{36})\}", joined) \
        or re.search(r"resid=[0-9A-Fa-f]+!s([0-9a-fA-F]{32})", joined)
    personal = re.search(r"(https://[^/]+/personal/[0-9a-zA-Z]+)/", joined)

    if not doc or not personal:
        raise SystemExit(
            "Could not work out the download address from that link.\n"
            "Make sure it is an 'Anyone with the link' share, and that opening it\n"
            "in a private browser window shows the file without asking to sign in.")

    unique_id = doc.group(1)
    if len(unique_id) == 32:  # resid form has no dashes
        unique_id = "%s-%s-%s-%s-%s" % (unique_id[:8], unique_id[8:12],
                                        unique_id[12:16], unique_id[16:20],
                                        unique_id[20:])

    download_url = "%s/_layouts/15/download.aspx?UniqueId=%s" % (
        personal.group(1), unique_id)

    # Step 3: fetch the bytes with the session established in step 1.
    try:
        with opener.open(download_url, timeout=300) as resp:
            ctype = resp.headers.get("Content-Type", "")
            data = resp.read()
    except Exception as exc:
        raise SystemExit("Download failed: %s" % exc)

    if not data.startswith(b"PK"):
        raise SystemExit(
            "The download returned %s rather than a workbook (%d bytes).\n"
            "The sharing link is probably not set to 'Anyone with the link'."
            % (ctype or "unknown content", len(data)))

    print("  downloaded %.1f MB" % (len(data) / 1048576.0))
    return data


# --------------------------------------------------------------------------
# Reading the workbook
# --------------------------------------------------------------------------
def cell_value(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0, 0) else v.isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, dt.time):
        # A bare time where a date belongs carries no information. Excel stores
        # an empty date cell this way, so drop it instead of writing "00:00:00".
        return None
    if isinstance(v, dt.timedelta):
        return v.total_seconds() / 86400.0
    if isinstance(v, (int, float, bool, str)):
        return v
    return str(v)


def find_header_row(rows, must_contain, limit=10):
    for i, row in enumerate(rows[:limit]):
        lowered = [str(c).strip().lower() for c in row if c is not None]
        if all(m in lowered for m in must_contain):
            return i
    return -1


def extract_main_sheet(ws):
    """Header row, then the data rows, cut down to the wanted columns."""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        raise SystemExit('Sheet "%s" is empty.' % SHEET_MAIN)

    header_idx = find_header_row(rows, HEADER_MUST_CONTAIN)
    if header_idx == -1:
        raise SystemExit(
            'No header row containing "Status" and "Buyer" was found in the first\n'
            '10 rows of sheet "%s".' % SHEET_MAIN)

    header = rows[header_idx]
    lookup = {}
    for i, name in enumerate(header):
        if name is not None:
            lookup.setdefault(str(name).strip().lower(), i)

    keep, out_header, missing = [], [], []
    for want in WANTED_COLUMNS:
        idx = lookup.get(want.lower())
        if idx is None:
            missing.append(want)
            continue
        keep.append(idx)
        out_header.append(want)

    if missing:
        print("  note: columns not in the sheet, skipped: %s" % ", ".join(missing),
              file=sys.stderr)

    id_pos = out_header.index("EWO") if "EWO" in out_header else 0
    out = [out_header]
    for row in rows[header_idx + 1:]:
        picked = [cell_value(row[i]) if i < len(row) else None for i in keep]
        if picked[id_pos] in (None, ""):
            continue
        out.append(picked)
    return out


def extract_monthly_sheet(ws):
    """First two columns only. The page ignores rows whose second cell is not
    a number, so the title and header rows fall away by themselves."""
    out = []
    for row in ws.iter_rows(values_only=True):
        pair = [cell_value(row[0] if len(row) > 0 else None),
                cell_value(row[1] if len(row) > 1 else None)]
        if pair[0] is None and pair[1] is None:
            continue
        out.append(pair)
    return out


def main():
    ap = argparse.ArgumentParser(description="Export EWO workbook sheets to data.json")
    ap.add_argument("source", help="a .xlsx path, or a OneDrive sharing link")
    ap.add_argument("-o", "--out", default=None,
                    help="output path (default: data.json beside this script)")
    args = ap.parse_args()

    out = args.out or os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")

    if args.source.lower().startswith(("http://", "https://")):
        print("Downloading from the sharing link")
        blob = io.BytesIO(download_shared_workbook(args.source))
    else:
        src = os.path.abspath(args.source)
        if not os.path.isfile(src):
            sys.exit("Workbook not found: %s" % src)
        print("Reading  %s" % src)
        blob = src

    wb = openpyxl.load_workbook(blob, data_only=True, read_only=True)
    by_lower = {n.lower(): n for n in wb.sheetnames}
    payload = {}

    main_name = by_lower.get(SHEET_MAIN.lower())
    if not main_name:
        sys.exit('No sheet named "%s". Sheets present: %s'
                 % (SHEET_MAIN, ", ".join(wb.sheetnames)))
    payload[SHEET_MAIN] = extract_main_sheet(wb[main_name])
    print("  %-5s %d data rows x %d columns"
          % (SHEET_MAIN, len(payload[SHEET_MAIN]) - 1, len(payload[SHEET_MAIN][0])))

    for label in (SHEET_FABRIC, SHEET_SEW):
        real = by_lower.get(label.lower())
        if not real:
            print("  note: sheet '%s' not found, chart will be empty" % label, file=sys.stderr)
            payload[label] = []
            continue
        payload[label] = extract_monthly_sheet(wb[real])
        print("  %-5s %d rows" % (label, len(payload[label])))

    wb.close()

    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(payload, fh, ensure_ascii=False, separators=(",", ":"))

    print("Wrote    %s  (%.1f KB)" % (out, os.path.getsize(out) / 1024.0))


if __name__ == "__main__":
    main()
