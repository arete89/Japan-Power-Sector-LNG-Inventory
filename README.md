# Japan-Power-Sector-LNG-Inventory

A single-file Python tool that tracks **Japan's power-sector LNG inventory** published weekly by
METI (the Agency for Natural Resources and Energy). Each run downloads the latest official PDF,
parses it, accumulates the history, and produces an Excel workbook plus a chart.

**Data source:** METI «発電用ＬＮＧの在庫状況»
`https://www.enecho.meti.go.jp/category/electricity_and_gas/electricity_measures/pdf/denryoku_LNG_stock.pdf`
(updated weekly, usually Tuesdays). Values are in **万トン** (10,000 tonnes).

---

## What it does

On every run the script:

1. **Downloads** the latest METI PDF into the working folder (`denryoku_LNG_stock.pdf`).
2. **Parses** the PDF's text table — the weekly readings (`週末在庫量`) and the month-end
   reference values (previous year `前年同月末` and the N-year average `過去平均`).
3. **Extracts** the bottom "参考" line chart as an image (`denryoku_LNG_chart.png`).
4. **Accumulates** everything into `lng_data.json` so history builds up week by week.
5. **Writes** `japan_lng_inventory.xlsx` with two tabs:
   - **Monthly (month-end)** — the bottom chart digitized as a table
     (`2024` / `5-yr average` / `2025` / `2026`), with the source chart image embedded
     for visual verification.
   - **Weekly (<year>)** — the current year's weekly readings, in both 万トン and mn t.
6. **Renders** a bonus matplotlib line chart (`japan_lng_inventory.png`).

---

## Install

Requires Python 3.10+.

```bash
pip install pdfplumber matplotlib numpy requests openpyxl
```

## Usage

```bash
# Normal weekly run — download the latest PDF and rebuild everything:
python3 japan_lng_chart.py

# Parse a PDF you downloaded manually (skips the network download):
python3 japan_lng_chart.py --pdf /path/to/denryoku_LNG_stock.pdf

# One-time historical backfill from the web archives:
python3 japan_lng_chart.py --bootstrap

# Cache a browser aws-waf-token for automated downloads (see "WAF" below):
python3 japan_lng_chart.py --set-token "<paste value from browser DevTools>"
```

## Outputs

| File | Description |
|------|-------------|
| `japan_lng_inventory.xlsx` | Main deliverable — monthly table + weekly tab + embedded chart |
| `denryoku_LNG_stock.pdf`   | The latest downloaded source PDF |
| `denryoku_LNG_chart.png`   | The bottom chart lifted out of the PDF (raster) |
| `japan_lng_inventory.png`  | Bonus seasonal line chart (matplotlib) |
| `lng_data.json`            | Accumulated data store (weekly + month-end refs + archive cache) |

---

## How the data is sourced (important)

METI's PDF contains **two** things, handled differently:

- **The text table** (top of the page) — machine-readable. The weekly readings and the
  `前年同月末` / `過去平均` month-end references are parsed directly from text. These are exact.
  The table only shows a rolling ~6–8 week window, so full history is built up across editions.

- **The bottom "参考" chart** (`2024 / 5-yr avg / 2025 / 2026`) — a **flat raster image**.
  Its data labels are pixels, not text, so they cannot be parsed. The **2024 line and the
  current-year month-end points** in the Monthly tab are therefore **digitized** (read from the
  chart) rather than text-parsed. Where editions overlap, these digitized values have been
  cross-checked against the archive and match exactly. The source chart image is embedded in the
  workbook so every figure can be verified by eye.

## Historical backfill (`--bootstrap`)

Full year-round history is reconstructed from two web archives:

- **NDL WARP** (`warp.ndl.go.jp`) — primary. The National Diet Library's archive covers periods
  the Internet Archive misses (e.g. Sep 2025 → Apr 2026). It has no CDX API, so the script
  enumerates captures by probing dates and resolving each to the nearest real capture.
  **WARP rate-limits bursts (HTTP 403)**, so the bootstrap is *resumable*: the capture list and an
  "already ingested" set are cached in `lng_data.json`, requests are spaced out with back-off, and
  if it gets throttled part-way you simply run `--bootstrap` again later to finish — no lost work.
- **Internet Archive / Wayback Machine** — secondary supplement for older editions.

## About the "WAF" and `--set-token`

METI sits behind AWS WAF. In practice a plain HTTP `GET` for the PDF succeeds, so no setup is
usually needed. If METI ever tightens the rules and direct downloads start failing, the script
falls back to the latest WARP capture, and `--set-token` lets you paste your browser's
`aws-waf-token` cookie (DevTools → Application → Cookies) to restore direct downloads. The token
is stored locally in `.meti_waf_token` and is **never** committed — keep it out of version control.

---

## Notes & limitations

- The 5-year-average window shifts by edition (e.g. `2021–2025` vs the fiscal-year `2020–2024`);
  each window is stored under its own key so they never mix.
- `--bootstrap` should not be run repeatedly in quick succession, or WARP will throttle you for a
  few minutes.
- The Monthly tab's current-year month-end points are added as each month completes; the most
  recent point may reflect the latest weekly reading until METI publishes the true month-end.

## License

Provided as-is for personal/research use. METI data is subject to METI's own terms of use.
