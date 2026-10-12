#!/usr/bin/env python3
"""
Japan Power Sector LNG Inventory — PDF → Excel converter
Data source: METI (Agency for Natural Resources and Energy)
https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/

Weekly PDF is updated every Tuesday. Each run:
  1. Downloads the latest PDF into this folder (denryoku_LNG_stock.pdf).
  2. Parses the text table (weekly readings + month-end reference values).
  3. Extracts the bottom "参考" line chart as a PNG (it is a raster image inside
     the PDF, so its data labels cannot be read as text).
  4. Persists everything to lng_data.json (accumulates week-by-week).
  5. Writes japan_lng_inventory.xlsx with two tabs:
       • "Monthly (month-end)" – the bottom chart digitized as a table
         (2024 / 5-yr average / 2025 / 2026), with the source chart image embedded.
       • "Weekly (2026)"       – this year's weekly readings, parsed from the PDF text.
  6. Also refreshes the japan_lng_inventory.png line chart (bonus).

Usage
-----
  # Normal weekly run:
  python3 japan_lng_chart.py

  # Pass a manually-downloaded PDF (no WAF issues):
  python3 japan_lng_chart.py --pdf /path/to/denryoku_LNG_stock.pdf

  # Cache your browser's aws-waf-token for automated live downloads:
  python3 japan_lng_chart.py --set-token "<paste value from browser DevTools>"

  # First-time / historical backfill from Wayback Machine archive:
  python3 japan_lng_chart.py --bootstrap

Setup (first time only)
-----------------------
  pip install pdfplumber matplotlib numpy requests openpyxl

Notes
-----
  METI's AWS WAF protects browser navigation, but direct HTTP GET requests
  for the PDF succeed without any special authentication. The --set-token flag
  is a fallback in case METI tightens WAF rules in the future.

  The bottom chart is a flat image in the PDF, so the four-series monthly table
  (especially the 2024 line) is DIGITIZED, not text-parsed. The 2025 and 5-yr
  average columns self-refine from each edition's text table; the source chart
  image is embedded in the workbook so every value can be visually verified.
"""

import argparse
import calendar
import io
import json
import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pdfplumber
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as mticker
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

# ── Constants ────────────────────────────────────────────────────────────────

PDF_URL = (
    "https://www.enecho.meti.go.jp"
    "/category/electricity_and_gas/electricity_measures/pdf/denryoku_LNG_stock.pdf"
)
REFERER = (
    "https://www.enecho.meti.go.jp"
    "/category/electricity_and_gas/electricity_measures/"
)
CDX_API = (
    "https://web.archive.org/cdx/search/cdx"
    "?url=www.enecho.meti.go.jp/category/electricity_and_gas"
    "/electricity_measures/pdf/denryoku_LNG_stock.pdf"
    "&output=json&fl=timestamp,statuscode&filter=statuscode:200&limit=300"
)
WAYBACK_TPL = "https://web.archive.org/web/{ts}if_/" + PDF_URL

# National Diet Library "WARP" archive (warp.ndl.go.jp). Unlike the Internet
# Archive, WARP covers the Sep-2025 → Apr-2026 gap for the METI PDF. It has no
# CDX API, but its viewer resolves /web/{target_ts}/{url} to the *nearest* real
# capture, exposed as an iframe src of the form  /{colDate}/{captureTs}/{url}.
WARP_BASE   = "https://warp.ndl.go.jp"
WARP_URL    = PDF_URL.split("://", 1)[1]          # host+path, no scheme
WARP_VIEWER = WARP_BASE + "/web/{ts}/" + WARP_URL
WARP_START  = "20250901"                          # earliest date to probe

SCRIPT_DIR  = Path(__file__).parent
DATA_FILE   = SCRIPT_DIR / "lng_data.json"
CHART_FILE  = SCRIPT_DIR / "japan_lng_inventory.png"       # matplotlib line chart
PDF_FILE    = SCRIPT_DIR / "denryoku_LNG_stock.pdf"        # downloaded source PDF
PDFCHART_FILE = SCRIPT_DIR / "denryoku_LNG_chart.png"      # chart image lifted from PDF
XLSX_FILE   = SCRIPT_DIR / "japan_lng_inventory.xlsx"      # Excel workbook output
TOKEN_FILE  = SCRIPT_DIR / ".meti_waf_token"   # cached aws-waf-token value

_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Dummy year used so all three series align on the same seasonal x-axis.
_BASE_YEAR = 2000
# Discard implausibly low values — page-number artifacts from PDF text extraction.
_MIN_VALID = 50.0


# ── Date helpers ─────────────────────────────────────────────────────────────

def _last_day(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def _calendar_year(data_month: int, pub_year: int, pub_month: int) -> int:
    """
    Infer the calendar year of a data point given only its month number.

    The METI PDF always shows a rolling window ending at the publication date,
    so any data month *greater* than the publication month must belong to the
    previous calendar year.
    """
    return pub_year if data_month <= pub_month else pub_year - 1


# ── Parse ────────────────────────────────────────────────────────────────────

def parse_pdf(pdf_bytes: bytes) -> tuple[int, int, int, dict]:
    """
    Extract data from the METI LNG stock PDF.

    Returns
    -------
    pub_year     : int  – calendar year of publication
    avg_start    : int  – first year of the 5-yr average range
    avg_end      : int  – last year of the 5-yr average range
    data         : dict – {
        "weekly"      : {iso_date: value_in_man_ton},
        "monthly_prev": {iso_date: value_in_man_ton},   # previous-year month-end
        "monthly_avg" : {pub_year_str: {month_str: value}},
    }

    Year attribution
    ----------------
    The PDF window can span two calendar years (e.g. a Feb 2023 edition may
    show November/December 2022 data alongside January 2023 data).
    Rule: data_month > pub_month  →  calendar year = pub_year − 1
          data_month ≤ pub_month  →  calendar year = pub_year
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = pdf.pages[0].extract_text() or ""

    # Publication date from 令和N年M月D日  (Reiwa era: year 1 = 2019)
    dm = re.search(r"令\s*和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", text)
    if dm:
        pub_year  = 2018 + int(dm.group(1))
        pub_month = int(dm.group(2))
    else:
        today     = date.today()
        pub_year  = today.year
        pub_month = today.month

    # 5-yr average footnote. Wording varies by edition:
    #   "2021～2025年平均"   (calendar-year window)
    #   "2020～2024 年度平均" (fiscal-year window, note the space and 年度)
    am = re.search(r"(\d{4})\s*[～~〜]\s*(\d{4})\s*年\s*度?\s*平均", text)
    avg_start = int(am.group(1)) if am else pub_year - 5
    avg_end   = int(am.group(2)) if am else pub_year - 1

    weekly:       dict[str, float] = {}
    monthly_prev: dict[str, float] = {}
    monthly_avg:  dict[str, dict]  = {}

    for line in text.splitlines():
        # ── Weekly entry: "5/10 時点 212" ──
        wm = re.search(r"(\d{1,2})/(\d{1,2})\s*時点\s+(\d+)", line)
        if wm:
            mon, day = int(wm.group(1)), int(wm.group(2))
            val = float(wm.group(3))
            if val >= _MIN_VALID:
                yr = _calendar_year(mon, pub_year, pub_month)
                try:
                    weekly[date(yr, mon, day).isoformat()] = val
                except ValueError:
                    pass  # invalid date (e.g. Feb 30)

        # ── Month-end reference: "5 月末 232 217" ──
        mm = re.search(r"(\d{1,2})\s*月末\s+(\d+)\s+(\d+)", line)
        if mm:
            mon       = int(mm.group(1))
            prev_val  = float(mm.group(2))
            avg_val   = float(mm.group(3))

            # Previous-year month-end date
            prev_yr   = (pub_year - 1) if mon <= pub_month else (pub_year - 2)
            prev_date = _last_day(prev_yr, mon).isoformat()
            monthly_prev[prev_date] = prev_val

            # 5-yr average stored under range key so multiple PDFs accumulate correctly
            avg_key = f"{avg_start}-{avg_end}"  # e.g. "2021-2025"
            monthly_avg.setdefault(avg_key, {})[str(mon)] = avg_val

    return pub_year, avg_start, avg_end, {
        "weekly":       weekly,
        "monthly_prev": monthly_prev,
        "monthly_avg":  monthly_avg,
    }


# ── Download (live) ──────────────────────────────────────────────────────────

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    ),
    "Referer": REFERER,
    "Accept": "application/pdf,*/*",
    "Accept-Language": "en-US,en;q=0.9",
}


def _load_waf_token() -> str | None:
    """Return cached aws-waf-token, or None if not stored."""
    if TOKEN_FILE.exists():
        return TOKEN_FILE.read_text(encoding="utf-8").strip() or None
    return None


def _try_direct_download() -> bytes | None:
    """Attempt download via requests, optionally with cached WAF token cookie."""
    token = _load_waf_token()
    cookies = {"aws-waf-token": token} if token else {}
    label = "with cached WAF token" if token else "without WAF token"
    print(f"  Trying direct download ({label}) ...", end=" ", flush=True)
    try:
        r = requests.get(PDF_URL, headers=_HEADERS, cookies=cookies, timeout=30)
        if r.status_code == 200 and len(r.content) > 5_000:
            print(f"OK ({len(r.content):,} bytes)")
            return r.content
        print(f"HTTP {r.status_code} — skipped")
    except Exception as e:
        print(f"error: {e}")
    return None


def _try_wayback_latest() -> tuple[bytes, str] | tuple[None, None]:
    """Fetch the most recent Wayback Machine snapshot of the METI PDF."""
    print("  Trying Wayback Machine (most recent snapshot) ...", end=" ", flush=True)
    try:
        r = requests.get(CDX_API, timeout=30)
        r.raise_for_status()
        rows = r.json()
        timestamps = [row[0] for row in rows[1:]]
        if not timestamps:
            print("no snapshots found")
            return None, None
        ts = timestamps[-1]
        pdf = fetch_wayback_pdf(ts)
        if pdf:
            snap_date = f"{ts[0:4]}-{ts[4:6]}-{ts[6:8]}"
            print(f"OK ({len(pdf):,} bytes, snapshot {snap_date})")
            return pdf, snap_date
        print("download failed")
    except Exception as e:
        print(f"error: {e}")
    return None, None


def download_pdf_live() -> tuple[bytes, str]:
    """
    Download the current METI PDF.

    Tries in order:
      1. Direct GET with cached aws-waf-token (fast, requires prior --set-token).
      2. Direct GET without a token (works if WAF is temporarily permissive).
      3. Most-recent NDL WARP capture (covers the Internet Archive gap).
      4. Most-recent Wayback Machine snapshot (may be a few weeks behind).

    Returns (pdf_bytes, source_description).
    If all strategies fail, raises RuntimeError with instructions.
    """
    # Strategy 1 & 2: direct requests (token or no token)
    pdf = _try_direct_download()
    if pdf:
        return pdf, "live"

    # Strategy 3: NDL WARP archive (most recent capture)
    print("  Trying NDL WARP archive (most recent capture) ...", end=" ", flush=True)
    cap = _warp_nearest_capture(date.today().strftime("%Y%m%d"))
    if cap:
        pdf = fetch_warp_pdf(*cap)
        if pdf:
            snap = f"{cap[1][0:4]}-{cap[1][4:6]}-{cap[1][6:8]}"
            print(f"OK ({len(pdf):,} bytes, capture {snap})")
            print(f"\n  [INFO] Using WARP capture from {snap} (see --set-token for live data).")
            return pdf, f"warp:{snap}"
    print("unavailable")

    # Strategy 4: Wayback Machine fallback
    pdf, snap_date = _try_wayback_latest()
    if pdf:
        print(f"\n  [INFO] Using Wayback Machine snapshot from {snap_date}.")
        print("  For live data, copy your aws-waf-token from the browser:")
        print(f"    1. Visit {REFERER} in Chrome/Safari")
        print("    2. Open DevTools > Application > Cookies > www.enecho.meti.go.jp")
        print("    3. Copy the value of 'aws-waf-token'")
        print("    4. Run:  python3 japan_lng_chart.py --set-token '<paste here>'\n")
        return pdf, f"wayback:{snap_date}"

    raise RuntimeError(
        "All download strategies failed.\n\n"
        "Manual option:\n"
        f"  1. Download the PDF from your browser:\n     {PDF_URL}\n"
        "  2. Run:  python3 japan_lng_chart.py --pdf /path/to/denryoku_LNG_stock.pdf\n\n"
        "Token option (enables automated live downloads):\n"
        f"  1. Visit {REFERER} in Chrome/Safari\n"
        "  2. Open DevTools > Application > Cookies > www.enecho.meti.go.jp\n"
        "  3. Copy the value of 'aws-waf-token'\n"
        "  4. Run:  python3 japan_lng_chart.py --set-token '<paste here>'"
    )


# ── Download (Wayback Machine) ───────────────────────────────────────────────

def fetch_wayback_timestamps() -> list[str]:
    """Return all Wayback Machine snapshot timestamps for the METI PDF."""
    r = requests.get(CDX_API, timeout=30)
    r.raise_for_status()
    rows = r.json()
    # First row is the header ["timestamp","statuscode"]
    return [row[0] for row in rows[1:]]


def fetch_wayback_pdf(timestamp: str) -> bytes | None:
    """Download one archived PDF snapshot from the Wayback Machine."""
    url = WAYBACK_TPL.format(ts=timestamp)
    try:
        r = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
        return r.content if r.status_code == 200 and len(r.content) > 5_000 else None
    except Exception:
        return None


def bootstrap_from_wayback(stored: dict) -> dict:
    """
    Download every archived PDF from the Wayback Machine and merge into store.
    Call once (with --bootstrap) to populate historical data.
    """
    print("Querying Wayback Machine CDX API ...")
    try:
        timestamps = fetch_wayback_timestamps()
    except Exception as e:
        print(f"  CDX query failed: {e}")
        return stored

    print(f"  Found {len(timestamps)} snapshots: {timestamps}")

    for ts in timestamps:
        print(f"\n  Fetching {ts} ...", end=" ", flush=True)
        pdf_bytes = fetch_wayback_pdf(ts)
        if not pdf_bytes:
            print("FAILED or too small — skipped")
            continue
        print(f"{len(pdf_bytes):,} bytes")

        try:
            pub_year, avg_start, avg_end, parsed = parse_pdf(pdf_bytes)
            print(
                f"    Published {pub_year} | "
                f"weekly={len(parsed['weekly'])} pts | "
                f"prev_year dates={sorted(parsed['monthly_prev'])}"
            )
            stored = merge(stored, parsed)
        except Exception as e:
            print(f"    Parse error: {e}")

        time.sleep(1.0)  # be polite to archive.org

    return stored


# ── Download (NDL WARP archive) ──────────────────────────────────────────────

_WARP_HDRS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}
_WARP_IFRAME_RE = re.compile(r'src="/(\d{8})/(\d{14})/' + re.escape(WARP_URL))


def _warp_get(url: str, timeout: int = 40) -> requests.Response | None:
    """GET with polite retry/back-off — WARP rate-limits bursts with HTTP 403."""
    for attempt in range(3):
        try:
            r = requests.get(url, timeout=timeout, headers=_WARP_HDRS)
            if r.status_code == 403:          # rate-limited — back off and retry
                time.sleep(5 * (attempt + 1))
                continue
            return r
        except Exception:
            time.sleep(2 * (attempt + 1))
    return None


def _warp_nearest_capture(target_ts: str) -> tuple[str, str] | None:
    """
    Resolve a target timestamp (YYYYMMDDhhmmss or YYYYMMDD) to WARP's nearest
    real capture, returned as (collection_date, capture_ts). None if unavailable.
    """
    ts = (target_ts + "000000")[:14]
    r = _warp_get(WARP_VIEWER.format(ts=ts))
    if r is None or r.status_code != 200:
        return None
    m = _WARP_IFRAME_RE.search(r.text)
    return (m.group(1), m.group(2)) if m else None


def warp_captures(start: str = WARP_START, step_days: int = 30) -> list[tuple[str, str]]:
    """
    Enumerate distinct WARP captures of the METI PDF by probing target dates
    from `start` to today in `step_days` increments (WARP has no CDX API).
    WARP captures the PDF ~monthly, so a monthly cadence with polite delays
    finds them all without tripping the rate limiter. Returns sorted
    (collection_date, capture_ts) pairs.
    """
    start_d = date(int(start[:4]), int(start[4:6]), int(start[6:8]))
    today = date.today()
    seen: dict[str, str] = {}   # capture_ts -> collection_date
    d = start_d
    while d <= today:
        cap = _warp_nearest_capture(d.strftime("%Y%m%d"))
        if cap:
            seen.setdefault(cap[1], cap[0])
        d += timedelta(days=step_days)
        time.sleep(3.0)         # be polite to warp.ndl.go.jp (avoids 403 rate-limits)
    return sorted((col, ts) for ts, col in seen.items())


def fetch_warp_pdf(collection_date: str, capture_ts: str) -> bytes | None:
    """Download one archived PDF from WARP: /{collection_date}/{capture_ts}/{url}."""
    url = f"{WARP_BASE}/{collection_date}/{capture_ts}/{WARP_URL}"
    r = _warp_get(url, timeout=60)
    if r is None:
        return None
    ok = (r.status_code == 200
          and r.content[:5] == b"%PDF-"
          and len(r.content) > 5_000)
    return r.content if ok else None


def bootstrap_from_warp(stored: dict) -> dict:
    """
    Enumerate and download WARP captures of the METI PDF, then merge.
    This is the archive backfill for the Internet Archive's Sep-2025→Apr-2026 gap.

    WARP rate-limits bursts, so this is designed to be resumable: the enumerated
    capture list and the set of already-ingested captures are cached in the store.
    Each run fetches only captures not yet ingested (spaced out); if the rate
    limiter blocks part-way, simply run --bootstrap again later to finish.
    """
    cache = stored.setdefault("warp", {"captures": [], "ingested": []})

    # Probe forward from the newest known capture (cheap on repeat runs); a first
    # run with an empty cache probes the full range from WARP_START once.
    if cache["captures"]:
        newest = max(c[1][:8] for c in cache["captures"])
        start = max(WARP_START, newest)
    else:
        start = WARP_START

    print(f"Enumerating NDL WARP captures from {start} (no CDX API — probing by date) ...")
    try:
        caps = warp_captures(start=start)
    except Exception as e:
        print(f"  WARP enumeration failed: {e}")
        caps = []

    # Merge freshly-found captures into the cached list (keep any earlier finds).
    known = {tuple(c) for c in cache["captures"]} | {tuple(c) for c in caps}
    cache["captures"] = sorted([list(c) for c in known], key=lambda c: c[1])
    if not cache["captures"]:
        print("  No WARP captures found (site may be rate-limiting — retry later).")
        return stored
    print(f"  {len(cache['captures'])} known captures: "
          f"{[c[1][:8] for c in cache['captures']]}")

    ingested = set(cache["ingested"])
    todo = [c for c in cache["captures"] if c[1] not in ingested]
    if not todo:
        print("  All known captures already ingested.")
        return stored
    print(f"  {len(todo)} new capture(s) to fetch.")

    for col_date, cap_ts in todo:
        snap = f"{cap_ts[0:4]}-{cap_ts[4:6]}-{cap_ts[6:8]}"
        print(f"\n  Fetching WARP {snap} ...", end=" ", flush=True)
        pdf_bytes = fetch_warp_pdf(col_date, cap_ts)
        if not pdf_bytes:
            print("blocked/failed — will retry on next --bootstrap run")
            continue
        print(f"{len(pdf_bytes):,} bytes")
        try:
            pub_year, avg_start, avg_end, parsed = parse_pdf(pdf_bytes)
            print(f"    Published {pub_year} | avg {avg_start}-{avg_end} | "
                  f"weekly={len(parsed['weekly'])} pts")
            stored = merge(stored, parsed)
            ingested.add(cap_ts)
            cache["ingested"] = sorted(ingested)
            save_data(stored)          # persist progress after each capture
        except Exception as e:
            print(f"    Parse error: {e}")
        time.sleep(3.0)

    return stored


# ── Data persistence ─────────────────────────────────────────────────────────

def load_data() -> dict:
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    return {"weekly": {}, "monthly_prev": {}, "monthly_avg": {}}


def save_data(d: dict) -> None:
    DATA_FILE.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")


def merge(stored: dict, parsed: dict) -> dict:
    """
    Merge freshly parsed data into the accumulated store.

    Storage layout (all keys are strings for JSON):
        stored["weekly"][iso_date]              = value  (万トン)
        stored["monthly_prev"][iso_date]        = value  (previous-year month-end)
        stored["monthly_avg"]["YYYY-YYYY"][mon] = value  (5-yr avg month-end, e.g. "2021-2025")
    """
    stored.setdefault("weekly",       {}).update(parsed["weekly"])
    stored.setdefault("monthly_prev", {}).update(parsed["monthly_prev"])
    for key, months in parsed.get("monthly_avg", {}).items():
        existing = stored.setdefault("monthly_avg", {}).setdefault(key, {})
        for month, value in months.items():
            if month not in existing:   # don't overwrite manually-corrected avg values
                existing[month] = value
    return stored


# ── Chart helpers ─────────────────────────────────────────────────────────────

def _to_mn_t(v: float) -> float:
    """万トン → million tonnes (mn t)."""
    return v / 100.0


def _base_date(year: int, month: int, day: int) -> date:
    """Map a real date to the shared dummy base-year for seasonal comparison."""
    return date(_BASE_YEAR, month, day)


# ── Chart ────────────────────────────────────────────────────────────────────

def generate_chart(pub_year: int, avg_start: int, avg_end: int, stored: dict) -> None:
    prev_year = pub_year - 1

    # ── Current year weekly ──────────────────────────────────────────────────
    cy_raw = {
        d: v for d, v in stored.get("weekly", {}).items()
        if d.startswith(str(pub_year))
    }
    cy_pts = sorted(
        [(_base_date(*[int(x) for x in d.split("-")]), _to_mn_t(v))
         for d, v in cy_raw.items()]
    )

    # ── Previous year: month-end points ─────────────────────────────────────
    py_monthly_raw = {
        d: v for d, v in stored.get("monthly_prev", {}).items()
        if d.startswith(str(prev_year))
    }
    py_pts = sorted(
        [(_base_date(*[int(x) for x in d.split("-")]), _to_mn_t(v))
         for d, v in py_monthly_raw.items()]
    )
    py_label = f"{prev_year} (month-end)"

    # ── 5-year average: month-end points ────────────────────────────────────
    avg_key = f"{avg_start}-{avg_end}"
    avg_raw = stored.get("monthly_avg", {}).get(avg_key, {})
    # Fallback: old format stored values under pub_year key
    if not avg_raw:
        avg_raw = stored.get("monthly_avg", {}).get(str(pub_year), {})
    avg_pts = sorted([
        (_base_date(_BASE_YEAR, int(m), _last_day(_BASE_YEAR, int(m)).day),
         _to_mn_t(v))
        for m, v in avg_raw.items()
    ])
    avg_label = f"{avg_start}–{avg_end} avg (month-end)"

    # ── Build plot ────────────────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(13, 6.5))
    fig.subplots_adjust(bottom=0.22)

    if cy_pts:
        xs, ys = zip(*cy_pts)
        ax.plot(xs, ys, color="#1f77b4", linewidth=2.0, marker="o",
                markersize=3.5, label=f"{pub_year} (weekly)", zorder=3)
        ax.annotate(
            f"{ys[-1]:.2f}", xy=(xs[-1], ys[-1]),
            xytext=(8, 4), textcoords="offset points",
            fontsize=9, color="#1f77b4", fontweight="bold",
        )

    if py_pts:
        xs, ys = zip(*py_pts)
        ax.plot(xs, ys, color="#ff7f0e", linestyle="--", linewidth=1.8,
                marker="o", markersize=6, label=py_label, zorder=2)

    if avg_pts:
        xs, ys = zip(*avg_pts)
        ax.plot(xs, ys, color="#2ca02c", linestyle=":", linewidth=1.8,
                marker="s", markersize=5, label=avg_label, zorder=2)

    # ── Axes ──────────────────────────────────────────────────────────────────
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
    ax.xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=0))
    ax.set_xlim(date(_BASE_YEAR, 1, 1), date(_BASE_YEAR, 12, 31))

    ax.set_ylabel("mn t", fontsize=12)
    ax.set_title(
        f"Japan Power Sector LNG Inventory\n"
        f"Source: METI Agency for Natural Resources and Energy"
        f"  |  Latest data: {pub_year}",
        fontsize=13, fontweight="bold", pad=12,
    )
    # Legend below the x-axis (horizontal) so it never overlaps the data lines.
    ax.legend(fontsize=9.5, frameon=False, loc="upper center",
              bbox_to_anchor=(0.5, -0.09), ncol=3)
    ax.grid(True, which="major", alpha=0.35, linestyle="--")
    ax.grid(True, which="minor", alpha=0.12, linestyle=":")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.2f}"))

    note = (
        "Note: Prior-year and 5-yr average series are METI month-end reference values. "
        "Run with --bootstrap to backfill Wayback Machine archive data."
    )
    fig.text(0.5, 0.02, note, ha="center", va="bottom", fontsize=7.5, color="#555555")

    fig.savefig(str(CHART_FILE), dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Chart saved → {CHART_FILE}")


# ── Chart-image extraction ───────────────────────────────────────────────────

def extract_chart_image(pdf_bytes: bytes) -> Path | None:
    """
    Lift the bottom "参考" line chart out of the PDF as a PNG.

    That chart is a flat raster image inside the PDF (its data labels are pixels,
    not text), so we render its bounding box at high resolution and save it. The
    image is embedded in the Excel workbook as a visual source-of-truth.
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            pg = pdf.pages[0]
            if not pg.images:
                print("  No embedded chart image found in PDF.")
                return None
            # The chart is the largest image on the page.
            im = max(pg.images, key=lambda x: x["width"] * x["height"])
            bbox = (max(im["x0"], 0), max(im["top"], 0),
                    min(im["x1"], pg.width), min(im["bottom"], pg.height))
            pil = pg.crop(bbox).to_image(resolution=300).original
            pil.save(str(PDFCHART_FILE))
        print(f"  Chart image extracted → {PDFCHART_FILE}")
        return PDFCHART_FILE
    except Exception as e:
        print(f"  Chart image extraction failed: {e}")
        return None


# ── Excel workbook ───────────────────────────────────────────────────────────

# Series column colours matching the METI chart legend.
_XL_SERIES = [
    ("2024",  "2024",              "FFC000"),  # amber
    ("avg",   "5-yr avg (21-25)",  "808080"),  # grey
    ("2025",  "2025",              "1F77B4"),  # blue
    ("2026",  "2026",              "C00000"),  # red
]

_HDR_FILL   = PatternFill("solid", fgColor="1F4E78")
_HDR_FONT   = Font(bold=True, color="FFFFFF", size=11)
_TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
_NOTE_FONT  = Font(italic=True, size=9, color="808080")
_THIN       = Side(style="thin", color="D9D9D9")
_BORDER     = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER     = Alignment(horizontal="center", vertical="center")


def _refresh_monthly_from_refs(stored: dict, pub_year: int) -> None:
    """
    Overwrite the digitized 2025 and 5-yr-average columns with exact machine-read
    month-end values wherever the accumulated text tables provide them. The 2024
    and current-year columns are left as digitized (not in a 2026 edition's text).
    """
    mc = stored.setdefault("monthly_chart", {})
    prev_year = str(pub_year - 1)

    # Previous-year (=2025) month-end refs → the "2025" column.
    col = mc.setdefault(prev_year, {})
    for iso, val in stored.get("monthly_prev", {}).items():
        if iso.startswith(prev_year):
            col[str(int(iso[5:7]))] = int(round(val))

    # 5-yr average month-end refs → the "avg" column (latest range key wins).
    avg_keys = sorted(stored.get("monthly_avg", {}).keys())
    if avg_keys:
        latest = stored["monthly_avg"][avg_keys[-1]]
        col = mc.setdefault("avg", {})
        for mon, val in latest.items():
            col[str(int(mon))] = int(round(val))


def write_excel(stored: dict, pub_year: int, chart_img: Path | None) -> None:
    """Write the two-tab workbook: monthly month-end table + weekly readings."""
    wb = openpyxl.Workbook()

    # ── Tab 1: Monthly (month-end) ──────────────────────────────────────────
    ws = wb.active
    ws.title = "Monthly (month-end)"
    meta = stored.get("monthly_chart_meta", {})
    unit = meta.get("unit", "万トン")

    ws["A1"] = "Japan Power Sector LNG — Month-End Inventory"
    ws["A1"].font = _TITLE_FONT
    ws["A2"] = f"Source: METI 参考 chart (denryoku_LNG_stock.pdf)  |  Unit: {unit} (10,000 t)"
    ws["A2"].font = _NOTE_FONT

    header_row = 4
    headers = ["Month"] + [label for _, label, _ in _XL_SERIES]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=text)
        cell.fill = _HDR_FILL
        cell.font = _HDR_FONT
        cell.alignment = _CENTER
        cell.border = _BORDER

    mc = stored.get("monthly_chart", {})
    for i, mon_abbr in enumerate(_MONTH_ABBR):
        r = header_row + 1 + i
        mcell = ws.cell(row=r, column=1, value=f"{mon_abbr} (end)")
        mcell.font = Font(bold=True)
        mcell.border = _BORDER
        for c, (key, _, colour) in enumerate(_XL_SERIES, start=2):
            val = mc.get(key, {}).get(str(i + 1))
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = _CENTER
            cell.border = _BORDER
            cell.font = Font(color=colour, bold=(key == "2026"))

    # Column widths
    ws.column_dimensions["A"].width = 14
    for col_letter in ("B", "C", "D", "E"):
        ws.column_dimensions[col_letter].width = 16

    note_row = header_row + len(_MONTH_ABBR) + 2
    ws.cell(row=note_row, column=1,
            value=("Digitized from the PDF's bottom chart (a raster image). "
                   "2024 & current-year points are read from the chart labels; "
                   "2025 & 5-yr-avg self-refine from each edition's text table. "
                   "See the embedded source chart below to verify.")
            ).font = _NOTE_FONT

    # Embed the source chart image beneath the table.
    if chart_img and Path(chart_img).exists():
        try:
            img = XLImage(str(chart_img))
            # Scale to a comfortable width (~720px) preserving aspect ratio.
            scale = 720 / img.width
            img.width = int(img.width * scale)
            img.height = int(img.height * scale)
            ws.add_image(img, f"A{note_row + 2}")
        except Exception as e:
            print(f"  Could not embed chart image: {e}")

    # ── Tab 2: Weekly (current year) ────────────────────────────────────────
    ws2 = wb.create_sheet(f"Weekly ({pub_year})")
    ws2["A1"] = f"Japan Power Sector LNG — Weekly Inventory ({pub_year})"
    ws2["A1"].font = _TITLE_FONT
    ws2["A2"] = "Source: METI weekly readings (週末在庫量), parsed from PDF text"
    ws2["A2"].font = _NOTE_FONT

    wk_header = 4
    for c, text in enumerate(["Week ending", f"Inventory ({unit})", "Inventory (mn t)"], start=1):
        cell = ws2.cell(row=wk_header, column=c, value=text)
        cell.fill = _HDR_FILL
        cell.font = _HDR_FONT
        cell.alignment = _CENTER
        cell.border = _BORDER

    weekly = {k: v for k, v in stored.get("weekly", {}).items()
              if k.startswith(str(pub_year))}
    for i, iso in enumerate(sorted(weekly)):
        r = wk_header + 1 + i
        ws2.cell(row=r, column=1, value=iso).border = _BORDER
        c2 = ws2.cell(row=r, column=2, value=int(round(weekly[iso])))
        c2.alignment = _CENTER
        c2.border = _BORDER
        c3 = ws2.cell(row=r, column=3, value=round(weekly[iso] / 100.0, 2))
        c3.alignment = _CENTER
        c3.border = _BORDER

    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 18
    ws2.freeze_panes = f"A{wk_header + 1}"

    wb.save(str(XLSX_FILE))
    print(f"  Excel workbook saved → {XLSX_FILE}  "
          f"({len(weekly)} weekly rows, {len(_MONTH_ABBR)} monthly rows)")


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--bootstrap",
        action="store_true",
        help="Download all Wayback Machine archive snapshots to populate historical data.",
    )
    parser.add_argument(
        "--pdf",
        metavar="PATH",
        help="Path to a manually-downloaded METI PDF (skips live download).",
    )
    parser.add_argument(
        "--set-token",
        metavar="TOKEN",
        dest="set_token",
        help="Cache the aws-waf-token cookie value for automated live downloads.",
    )
    args = parser.parse_args()

    # --set-token: save WAF token and exit
    if args.set_token:
        TOKEN_FILE.write_text(args.set_token.strip(), encoding="utf-8")
        print(f"WAF token saved to {TOKEN_FILE}")
        print("Run  python3 japan_lng_chart.py  to download the live PDF.\n")
        return

    stored = load_data()

    if args.bootstrap:
        print("=== Bootstrapping from NDL WARP archive ===")
        stored = bootstrap_from_warp(stored)
        save_data(stored)
        # Also try the Internet Archive (older editions, when reachable).
        print("\n=== Supplementing from Internet Archive (Wayback) ===")
        stored = bootstrap_from_wayback(stored)
        save_data(stored)
        # Refresh reference columns from whatever editions we just parsed.
        _refresh_monthly_from_refs(stored, date.today().year)
        save_data(stored)
        print(f"\nBootstrap complete. Data saved → {DATA_FILE}")

    # Obtain PDF bytes
    if args.pdf:
        pdf_path = Path(args.pdf)
        if not pdf_path.exists():
            sys.exit(f"File not found: {pdf_path}")
        pdf_bytes = pdf_path.read_bytes()
        source = f"local:{pdf_path.name}"
        print(f"Using local PDF: {pdf_path}  ({len(pdf_bytes):,} bytes)")
    else:
        print("\nDownloading latest METI LNG inventory PDF ...")
        pdf_bytes, source = download_pdf_live()

    # Save the PDF into this folder (requirement: download the PDF first).
    PDF_FILE.write_bytes(pdf_bytes)
    print(f"  PDF saved → {PDF_FILE}")

    print("Parsing PDF ...")
    pub_year, avg_start, avg_end, parsed = parse_pdf(pdf_bytes)
    print(f"  Source           : {source}")
    print(f"  Publication year : {pub_year}")
    print(f"  5-yr avg range   : {avg_start}–{avg_end}")
    print(f"  Weekly dates     : {sorted(parsed['weekly'])}")
    print(f"  Prev-year dates  : {sorted(parsed['monthly_prev'])}")

    stored = merge(stored, parsed)
    # Refine the digitized 2025 & 5-yr-avg columns with exact text-parsed values.
    _refresh_monthly_from_refs(stored, pub_year)
    save_data(stored)
    print(f"  Data saved → {DATA_FILE}")

    print("Extracting bottom chart image from PDF ...")
    chart_img = extract_chart_image(pdf_bytes)

    print("Writing Excel workbook ...")
    write_excel(stored, pub_year, chart_img)

    print("Generating line chart (bonus) ...")
    generate_chart(pub_year, avg_start, avg_end, stored)


if __name__ == "__main__":
    main()
