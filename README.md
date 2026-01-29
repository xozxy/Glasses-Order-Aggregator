# Glasses Order Aggregator & Label Generator

A Cloudflare Workers + Pages project that aggregates glasses order CSVs into a single-row-per-order export and generates 5×4cm PDF labels.

## Features

- `/api/aggregate`: upload order-line CSV and export a one-row-per-order CSV
- `/api/labels`: upload CSV (raw or aggregated) and generate 5×4cm PDF labels
- Web UI: one-click upload/download for CSV/PDF, with automatic batch generation + merge

## Tech Stack

- Cloudflare Workers + Assets
- pdf-lib + fontkit (multilingual fonts)
- PapaParse (CSV parsing)

## Local Development

```bash
npm i
npx wrangler dev
```

Open `http://127.0.0.1:8787`.

## Endpoints

### 1) POST `/api/aggregate`

Upload the raw order-line CSV (one product per row) and get an aggregated CSV.

Request:

- `multipart/form-data`, field name: `file`
- Query: `format=csv` (CSV only for now)
- Optional:
  - `drop=default` (default): drop predefined columns
  - `drop=none`: keep all columns
  - `drop=custom&drop_cols=col1,col2`: add extra columns to drop

Response:

- `text/csv` download

### 2) POST `/api/labels`

Generate 5×4cm PDF labels (one page per record).

Request:

- `multipart/form-data`, field name: `file`
- Query:
  - `mode=bundle` (default): one label per bundle
  - `mode=order`: one label per order (aggregated CSV reads Bundle 1)
  - `start` / `limit`: pagination to avoid Worker limits (default 200, max 300)

Response:

- `application/pdf` download

## CSV Compatibility & Field Mapping

### A) Raw detail CSV (one product per row)

Required:

- `Order ID`
- `Bundle ID`
- `Line Item`
- `Quantity`

Prescription fields (priority order):

- OD: `OD SPH` / `OD CYL` / `OD Axis` / `OD ADD`
- OS: `OS SPH` / `OS CYL` / `OS Axis` / `OS ADD`
- PD: `PD OD` / `PD OS`
- Prescription type: `Prescription Type`
- Fallbacks: `OD_SPH` / `OS_SPH` underscore variants

### B) Aggregated CSV (`orders_one_row.csv`)

Prescription fields:

- `OD SPH (Bundle 1)` / `OD CYL (Bundle 1)` / `OD Axis (Bundle 1)` / `OD ADD (Bundle 1)`
- `OS SPH (Bundle 1)` / `OS CYL (Bundle 1)` / `OS Axis (Bundle 1)` / `OS ADD (Bundle 1)`
- `PD OD (Bundle 1)` / `PD OS (Bundle 1)`
- `Prescription Type (Bundle 1)`

Mode rules:

- `mode=order`: only reads Bundle 1
- `mode=bundle`: iterates all bundles (from `Bundle Count` or columns)

If bundle columns are missing, the API returns 400 with a message.

## PDF Generation Notes

- Helvetica is ASCII-only; any non-ASCII uses embedded Noto fonts
- Fonts are loaded on demand (JP/SC/KR) to reduce resource usage
- `subset:true` for smaller font payloads
- Over 200 pages will suggest batching (front-end auto-handles)

## Front-End Batch Merge

The “Generate & Download PDF” button:

1. Tries a single `/api/labels` request
2. If over the limit, batches `/api/labels?start=&limit=`
3. Merges PDFs in-browser via `pdf-lib` and downloads a single file

## Fonts

Place fonts in `public/fonts/`:

- `NotoSansJP.ttf`
- `NotoSansSC.ttf`
- `NotoSansTC.ttf` (fallback)
- `NotoSansKR.ttf`

Workers read fonts via `env.ASSETS.fetch("http://local/fonts/xxx")`.

## Quick Tests

1) Raw detail CSV  
```
curl -F "file=@/path/to/lensadvizor-orders.csv" \
  "http://127.0.0.1:8787/api/labels?mode=bundle" \
  -o labels_5x4cm.pdf
```

2) Aggregated CSV  
```
curl -F "file=@/path/to/orders_one_row.csv" \
  "http://127.0.0.1:8787/api/labels?mode=order" \
  -o labels_5x4cm.pdf
```

3) Generate aggregated CSV  
```
curl -F "file=@/path/to/lensadvizor-orders.csv" \
  "http://127.0.0.1:8787/api/aggregate" \
  -o orders_one_row.csv
```

## FAQ

**Q: Why did all values show as 0.00 / 0?**  
A: Field name mismatches produced empty strings, which previously converted to 0. Now real column names are prioritized and empty values stay empty.

**Q: Too many labels causes 503?**  
A: Use pagination with `start`/`limit`. The UI auto-batches and merges.
