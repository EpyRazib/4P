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

Every named column in the EWO sheet is published, so the dashboard can
slice the data any way it likes.

Usage
-----
    python export_ewo_data.py --folder "https://1drv.ms/f/c/..../Ig....?e=xxxx"
    python export_ewo_data.py "https://1drv.ms/x/c/..../IQ....?e=xxxx"
    python export_ewo_data.py "EWO.xlsx" --delivery "Fabric Delivery.xlsx" -o data.json

--folder is the easiest: share the OneDrive folder that holds both workbooks
as "Anyone with the link" and the script finds them by name, newest first.

The optional --delivery workbook adds the Delivery, Locked Fabric, RFD and
LockRepo sheets, which drive the fabric delivery tab.

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

# Every column that carries a header name is published. Unnamed spacer columns
# are dropped, and a repeated name keeps its first occurrence so the dashboard
# never has to guess which of two identical headers it is looking at.
HEADER_MUST_CONTAIN = ["status", "buyer"]

# Columns the dashboard cannot work without. Their absence is worth shouting
# about, because every chart and filter is built on them.
REQUIRED_COLUMNS = ["EWO", "Status", "Buyer"]


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
# Downloading every workbook in a shared folder
# --------------------------------------------------------------------------
def _session_opener():
    jar = http.cookiejar.CookieJar()
    recorder = _RecordingRedirectHandler()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar), recorder)
    opener.addheaders = [("User-Agent", UA)]
    return opener, recorder


def list_shared_folder(folder_url):
    """Redeem an 'Anyone with the link' folder share and list the files in it.
    Returns (opener, cid_base, [{name, id, size, modified}])."""
    opener, recorder = _session_opener()
    try:
        with opener.open(folder_url, timeout=90) as resp:
            final_url = resp.geturl()
            resp.read(2048)
    except Exception as exc:
        raise SystemExit("Could not open the folder link: %s" % exc)

    chain = " ".join(recorder.chain + [final_url])
    # For a folder the path only ever appears URL-encoded in the query string,
    # e.g. spopath=%2Fpersonal%2F<cid>%2FDocuments%2FDashboard, and the host
    # is whichever onedrive.live.com address the chain settled on.
    spopath = re.search(r"spopath=([^&\s]+)", chain)
    host = re.search(r"https://([^/\s]+)/", final_url)
    if not spopath or not host:
        raise SystemExit(
            "Could not work out the folder's address from that link.\n"
            "Make sure it is an 'Anyone with the link' share of a FOLDER, and that\n"
            "opening it in a private window shows the files without a sign-in.")

    folder_path = urllib.parse.unquote(spopath.group(1))          # /personal/<cid>/Documents/...
    cid = re.match(r"/personal/([0-9a-zA-Z]+)/", folder_path)
    if not cid:
        raise SystemExit("Unexpected folder path in the link: %s" % folder_path)
    base = "https://%s/personal/%s" % (host.group(1), cid.group(1))
    api = ("%s/_api/web/GetFolderByServerRelativePath(decodedurl='%s')/Files"
           "?$select=Name,UniqueId,Length,TimeLastModified"
           % (base, urllib.parse.quote(folder_path, safe="/")))
    req = urllib.request.Request(api, headers={"Accept": "application/json;odata=verbose"})
    try:
        with opener.open(req, timeout=90) as resp:
            data = json.loads(resp.read().decode("utf-8", "replace"))
    except Exception as exc:
        raise SystemExit("The folder opened but its file list could not be read: %s" % exc)

    files = []
    for f in data.get("d", {}).get("results", []):
        files.append({"name": f.get("Name", ""), "id": f.get("UniqueId", ""),
                      "size": int(f.get("Length") or 0), "modified": f.get("TimeLastModified", "")})
    return opener, base, files


def download_by_id(opener, base, unique_id, label):
    url = "%s/_layouts/15/download.aspx?UniqueId=%s" % (base, unique_id)
    try:
        with opener.open(url, timeout=300) as resp:
            data = resp.read()
    except Exception as exc:
        raise SystemExit("Download of %s failed: %s" % (label, exc))
    if not data.startswith(b"PK"):
        raise SystemExit("%s did not come back as a workbook (%d bytes)." % (label, len(data)))
    print("  downloaded %s  %.1f MB" % (label, len(data) / 1048576.0))
    return data


def pick_from_folder(folder_url, ewo_pattern, delivery_pattern):
    """Find the two workbooks by name inside the shared folder. Returns
    (ewo_bytes, delivery_bytes_or_None)."""
    opener, base, files = list_shared_folder(folder_url)
    xlsx = [f for f in files if f["name"].lower().endswith((".xlsx", ".xlsm"))]
    print("  folder holds %d workbook(s): %s" % (len(xlsx), ", ".join(f["name"] for f in xlsx)))

    def find(pattern, what):
        hits = [f for f in xlsx if re.search(pattern, f["name"], re.I)]
        if not hits:
            return None
        # newest first, in case an older copy is still lying in the folder
        hits.sort(key=lambda f: f["modified"], reverse=True)
        if len(hits) > 1:
            print("  note: %d files match the %s pattern, using the newest: %s"
                  % (len(hits), what, hits[0]["name"]), file=sys.stderr)
        return hits[0]

    ewo = find(ewo_pattern, "EWO")
    if not ewo:
        raise SystemExit("No workbook in the folder matches the EWO pattern %r.\n"
                         "Files present: %s" % (ewo_pattern, ", ".join(f["name"] for f in xlsx)))
    dlv = find(delivery_pattern, "delivery")
    if not dlv:
        print("  note: no workbook matches the delivery pattern %r, the delivery tab will be empty"
              % delivery_pattern, file=sys.stderr)

    ewo_bytes = download_by_id(opener, base, ewo["id"], ewo["name"])
    dlv_bytes = download_by_id(opener, base, dlv["id"], dlv["name"]) if dlv else None
    return ewo_bytes, dlv_bytes


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

    keep, out_header, seen = [], [], set()
    for i, name in enumerate(header):
        if name is None:
            continue
        label = str(name).strip()
        if not label or label.lower() in seen:
            continue
        seen.add(label.lower())
        keep.append(i)
        out_header.append(label)

    absent = [c for c in REQUIRED_COLUMNS if c.lower() not in seen]
    if absent:
        raise SystemExit("These columns are missing from the sheet and the dashboard "
                         "cannot run without them: %s" % ", ".join(absent))

    id_pos = out_header.index("EWO")
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



# --------------------------------------------------------------------------
# The fabric delivery workbook
# --------------------------------------------------------------------------
# Nine columns cover every chart on the delivery tab. BookingNo is what ties a
# delivery line to its row in the lock plan. Fabric construction and USD values
# were dropped: they added weight and nothing the dashboard shows.
DELIVERY_COLUMNS = ["DeliveryDate", "ExportOrderNo", "BookingNo", "Buyer", "DeliveryType",
                    "BeneficieryUnit", "ExecutionUnit", "FabricType", "DeliveryQtyKg"]


def norm(h):
    return re.sub(r"\s+", " ", str(h)).strip().lower() if h is not None else ""


def extract_delivery(ws):
    """Every delivery line, trimmed to the columns the dashboard uses."""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:10])
               if "deliverydate" in [norm(c) for c in r] and "exportorderno" in [norm(c) for c in r]), -1)
    if hi == -1:
        print("  note: Delivery sheet has no DeliveryDate/ExportOrderNo header, skipped", file=sys.stderr)
        return []
    lookup = {norm(c): i for i, c in enumerate(rows[hi]) if c is not None}
    keep = [(name, lookup[norm(name)]) for name in DELIVERY_COLUMNS if norm(name) in lookup]
    out = [[k for k, _ in keep]]
    di = lookup.get("deliverydate")
    for r in rows[hi + 1:]:
        if di is None or di >= len(r) or not isinstance(r[di], (dt.date, dt.datetime)):
            continue
        vals = [cell_value(r[i]) if i < len(r) else None for _, i in keep]
        # kilograms to one decimal place, which is all the source carries anyway
        out.append([round(v, 1) if isinstance(v, float) else v for v in vals])
    return out


def extract_lock_plan(ws):
    """One row per booking in the month's lock plan. The daily grid to the
    right is not exported: the plan spreads evenly across the lock window, a
    fact verified against the sheet's own 'Required Till Date' total, so the
    dashboard rebuilds any day's figure from start, end and quantity."""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:10])
               if "ewo" in [norm(c) for c in r] and "booking number" in [norm(c) for c in r]), -1)
    if hi == -1:
        print("  note: Locked Fabric sheet header not found, skipped", file=sys.stderr)
        return []
    L = {norm(c): i for i, c in enumerate(rows[hi]) if c is not None}

    def col(*names):
        for n in names:
            if norm(n) in L:
                return L[norm(n)]
        return None

    spec = [("Booking", col("Booking Number")), ("EWO", col("EWO")), ("MMTeam", col("MM Team")),
            ("Buyer", col("Buyer Name")), ("BuyerTeam", col("Buyer Team")),
            ("Unit", col("Unit")), ("PlanKg", col("Fabric Delivery Balance")),
            ("LockStart", col("Adjusted Lock Fabric Start")), ("LockEnd", col("Adjusted Lock Fabric End")),
            ("ReqTillDate", col("Required Till Date")), ("Delivered", col("Delivered")),
            ("RFD", col("RFD"))]
    spec = [(n, i) for n, i in spec if i is not None]
    out = [[n for n, _ in spec]]
    ei = col("EWO")
    for r in rows[hi + 1:]:
        e = r[ei] if ei is not None and ei < len(r) else None
        if not (isinstance(e, (int, float)) and 100000 < e < 999999):
            continue
        out.append([cell_value(r[i]) if i < len(r) else None for _, i in spec])
    return out


def extract_rfd(ws):
    """Fabric that is ready for delivery, one row per booking."""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hi = next((i for i, r in enumerate(rows[:6])
               if "ewo" in [norm(c) for c in r] and "booking number" in [norm(c) for c in r]), -1)
    if hi == -1:
        return []
    L = {norm(c): i for i, c in enumerate(rows[hi]) if c is not None}

    def col(prefix):
        for k, i in L.items():
            if k.startswith(norm(prefix)):
                return i
        return None

    spec = [("EWO", col("EWO")), ("Booking", col("Booking Number")), ("Buyer", col("Buyer")),
            ("FRStart", col("FR Started")), ("ActualStart", col("Actual Start")),
            ("PIStatus", col("Fabric PI")), ("Unit", col("Received Unit")), ("Kg", col("Total Qty"))]
    spec = [(n, i) for n, i in spec if i is not None]
    out = [[n for n, _ in spec]]
    ei = col("EWO")
    for r in rows[hi + 1:]:
        e = r[ei] if ei is not None and ei < len(r) else None
        if not (isinstance(e, (int, float)) and 100000 < e < 999999):
            continue
        out.append([cell_value(r[i]) if i < len(r) else None for _, i in spec])
    return out


def extract_lock_report_meta(ws):
    """The few report-level inputs that cannot be derived from the rows: the
    reporting window, the month target, and the working-day counts."""
    rows = [list(r) for r in ws.iter_rows(values_only=True, max_row=12)]
    meta = {}
    for r in rows:
        for i, c in enumerate(r):
            t = norm(c)
            nxt = next((x for x in r[i + 1:i + 4] if isinstance(x, (int, float))), None)
            if t.startswith("monthly target") and nxt is not None:
                meta["monthTarget"] = nxt
            if t == "day" and nxt is not None:
                meta["daysInMonth"] = nxt
            if t == "actual" and nxt is not None and "daysDone" not in meta:
                meta["daysDone"] = nxt
            if t == "remain" and nxt is not None:
                meta["daysRemain"] = nxt
            # Two percentages whose formula lives only in the sheet; shown as published.
            if t == "fabric start %" and nxt is not None:
                meta["fabricStartPct"] = nxt
            if t == "fabric end %" and nxt is not None:
                meta["fabricEndPct"] = nxt
    dates = [c for r in rows[:2] for c in r if isinstance(c, (dt.date, dt.datetime))]
    if len(dates) >= 2:
        meta["reportFrom"] = cell_value(dates[0])
        meta["reportTo"] = cell_value(dates[1])
    return meta


def main():
    ap = argparse.ArgumentParser(description="Export EWO workbook sheets to data.json")
    ap.add_argument("source", nargs="?", default=None,
                    help="EWO workbook: a .xlsx path, or a OneDrive sharing link")
    ap.add_argument("--delivery", default=None,
                    help="Fabric Delivery workbook: a .xlsx path, or a OneDrive sharing link")
    ap.add_argument("--folder", default=None,
                    help="an 'Anyone with the link' OneDrive FOLDER share holding both workbooks; "
                         "replaces source and --delivery")
    ap.add_argument("--ewo-pattern", default=r"EWO.*Life.*Cycle",
                    help="regex that picks the EWO workbook by file name inside --folder")
    ap.add_argument("--delivery-pattern", default=r"Fabric.*Delivery",
                    help="regex that picks the delivery workbook by file name inside --folder")
    ap.add_argument("-o", "--out", default=None,
                    help="output path (default: data.json beside this script)")
    args = ap.parse_args()

    out = args.out or os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")

    folder_delivery_bytes = None
    if args.folder:
        print("Listing the shared folder")
        ewo_bytes, folder_delivery_bytes = pick_from_folder(args.folder, args.ewo_pattern, args.delivery_pattern)
        blob = io.BytesIO(ewo_bytes)
    elif not args.source:
        sys.exit("Give either a source workbook or --folder.")
    elif args.source.lower().startswith(("http://", "https://")):
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

    # ---- the second workbook, when given ----
    if args.delivery or folder_delivery_bytes:
        if folder_delivery_bytes:
            blob2 = io.BytesIO(folder_delivery_bytes)
        elif args.delivery.lower().startswith(("http://", "https://")):
            print("Downloading the delivery workbook from its sharing link")
            blob2 = io.BytesIO(download_shared_workbook(args.delivery))
        else:
            src2 = os.path.abspath(args.delivery)
            if not os.path.isfile(src2):
                sys.exit("Delivery workbook not found: %s" % src2)
            print("Reading  %s" % src2)
            blob2 = src2
        wb2 = openpyxl.load_workbook(blob2, data_only=True, read_only=True)
        low2 = {n.lower(): n for n in wb2.sheetnames}

        def sheet2(name):
            real = low2.get(name.lower())
            return wb2[real] if real else None

        ws_d, ws_l, ws_r, ws_m = sheet2("Delivery"), sheet2("Locked Fabric"), sheet2("RFD"), sheet2("LockRepo")
        payload["Delivery"] = extract_delivery(ws_d) if ws_d else []
        payload["LockPlan"] = extract_lock_plan(ws_l) if ws_l else []
        payload["RFD"] = extract_rfd(ws_r) if ws_r else []
        payload["LockMeta"] = extract_lock_report_meta(ws_m) if ws_m else {}
        for k in ("Delivery", "LockPlan", "RFD"):
            print("  %-9s %d rows" % (k, max(0, len(payload[k]) - 1)))
        print("  LockMeta  %s" % payload["LockMeta"])
        wb2.close()

    # The dashboard shows how fresh the data is, and skips re-rendering when
    # this value has not moved since its last poll.
    payload["generated"] = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(payload, fh, ensure_ascii=False, separators=(",", ":"))

    print("Wrote    %s  (%.1f KB)" % (out, os.path.getsize(out) / 1024.0))


if __name__ == "__main__":
    main()
