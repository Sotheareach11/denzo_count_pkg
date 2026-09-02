# DENSO PKG Counter + PDF Combiner — Project Documentation

## 1. What this project is

A small **Flask web application** used by CKSN International Transport (a customs
broker / freight forwarder) to automate two repetitive document tasks around
DENSO shipments:

| Tool | URL | Input | Output |
|------|-----|-------|--------|
| **PKG Counter** | `/` | One or more **packing list PDFs** | An **Excel report** counting packages (PKG) and summarising net/gross weight per item, per invoice and per factory |
| **Combine PDFs** | `/combine` | A mix of **Invoice / Packing List / Freight PDFs** | A **single merged PDF**, grouped by invoice number and ordered INV → PL → Freight |

Everything runs in the browser. No desktop GUI, no database — files are uploaded,
processed in memory / on temp disk, and the result is offered as a download.

---

## 2. Tech stack

| Layer | Technology |
|-------|-----------|
| Web framework | Flask 3.0 |
| PDF text extraction | `pdfplumber` (built on `pdfminer.six`) |
| PDF page merging | `pypdf` |
| Excel generation | `openpyxl` |
| Production server | `gunicorn` (single worker, multiple threads) |
| Frontend | Plain HTML + vanilla JS in `templates/` (no framework) |
| Hosting targets | Render / CodeSandbox / Replit / any container |

### Files

```
app.py               Main Flask app: routes, PDF parsing, Excel building, background jobs
combinepdf.py        Standalone module: classify & merge INV/PL/Freight PDFs
templates/index.html   PKG Counter UI
templates/combine.html Combine PDFs UI
requirements.txt     Pinned dependencies
uploads/             Runtime scratch dir for uploaded files + generated reports
.devcontainer/, .codesandbox/  Cloud-IDE config (the app.py copy under
                     .devcontainer/ is an OLDER version — ignore it)
```

---

## 3. Tool 1 — PKG Counter

### 3.1 Purpose

DENSO packing lists list every carton ("CML No.") with its weight/volume, and
every item line inside that carton. The broker needs to know, per invoice and
per receiving factory:

- **PKG count** — how many cartons contain each item
- **Total quantity** shipped per item
- **Net weight / Gross weight** totals per carton, per invoice, per factory, and grand total

Doing this by hand across dozens of PDFs is slow and error-prone.

### 3.2 Request flow (why it uses background jobs)

Large batches take longer than the HTTP/gunicorn timeout allows, so the work is
split:

```
Browser                         Flask (app.py)                Background thread
  │                                   │                              │
  │  POST /count  (all PDFs) ───────► │                              │
  │                                   │ save each PDF to uploads/    │
  │                                   │ create JOBS[job_id]          │
  │                                   │ start daemon thread ───────► │ _process_count_job()
  │  ◄─── { job_id, total_files } ────│                              │  for each PDF:
  │                                   │                              │    parse_packing_list()
  │  GET /count/status/<job_id> ────► │ read JOBS[job_id]            │    summarize_by_name()
  │  ◄─── { status, progress… } ──────│  (polled every 2s)           │    append to preview
  │            …repeat…               │                              │  build_excel() → uploads/<id>.xlsx
  │  ◄─── { status:"done", preview,   │                              │  JOBS[job_id] = done + download_id
  │         download_id, grand… } ────│                              │
  │                                   │                              │
  │  GET /download/<report_id> ─────► │ send_file(uploads/<id>.xlsx) │
```

`JOBS` is an **in-memory dict** guarded by a lock. This only works with a
**single gunicorn worker** (hence `--workers 1 --threads 4` in the start
command). Scaling to multiple workers would require Redis/DB-backed job state.

### 3.3 PDF parsing — `parse_packing_list(pdf_path)`

Returns `(records, packages, ship_to)`:

- **`records`** — one dict per item/order line (feeds the per-invoice item sheet)
- **`packages`** — `OrderedDict` keyed by CML No. (carton), each with `vol`, `nw`, `gw`
- **`ship_to`** — the receiving factory name, used to group the summary sheet

#### Ship-To detection — `extract_ship_to()`

`pdfplumber` often merges the "Sold to" and "Ship to" columns onto one text
line. The function finds the `Ship to` marker, drops the merged-in
`Document Information` header, then grabs text up to the first `CO., LTD`
suffix — that is the factory name. Falls back to `"Unknown"` so an unusual
layout still groups instead of crashing.

#### Line-pattern matching

DENSO uses several packing-list templates. The parser walks the text lines and
tries a series of regexes **in priority order**:

| Regex | Matches | Example |
|-------|---------|---------|
| `CML_RE` | Strict carton line: `CML vol nw gw` | `STG002608060339 0.005 1.200 1.525` |
| `ORDER_RE` | "Normal" template — order/item/unit/qty/nw/cartons on one line | `...K123 ITEM01 pcs 60 1.200 3` |
| `MODEL_LINE_RE` | "Return Style Code" template — extra per-unit N/W column | `N73T pcs 10 16.700 167.000` |
| `ITEM_LINE_RE` | DENSO (Thailand) — item line with no order no. / no cartons; order + description on the **next** line | `TG022108-00509B pcs 60 1.200` |
| `CML_RE_LOOSE` | Carton line with free-text between CML and numbers (tried **last** so it can't swallow item lines) | `DTG0T11C607240160 PLASTIC PACKAGING ... 1.156 170.000 170.000` |
| `DESC_RE` | `custitem + description` on one line (first token must contain a digit) | `KN127314-3110 TUBE` |

`HEADER_PREFIXES` lines (column headers, page numbers, phone/fax, etc.) are
skipped entirely.

The parser keeps a `cur_cml` (current carton) and a `pending_order` (item line
waiting for its description). When a description line is found, a `record` is
emitted linking that item to the current carton.

### 3.4 Aggregation — `summarize_by_name(records)`

Groups records by CML, then by `(item no., description)` within each carton.

- **Package name**: within one carton, the item with the **largest net weight**
  "wins" and its description becomes that carton's label (`package_names[cml]`).
- **PKG count**: for each `(item, desc)` key, incremented **once per carton
  where that key is the winner** — i.e. "how many cartons is this item the main
  content of".
- **Total Qty**: summed across all cartons regardless of winner.

`display_name()` prefers the description text, falling back to the customer item
number, then `"(unknown)"`.

### 3.5 Excel output — `build_excel(all_results)`

The workbook contains:

1. **One sheet per PDF** (`sheet name = PDF filename`, sanitised to ≤31 chars):
   columns `Item No. | Customer Item No. | Description | Unit | PKG (Count) | Total Qty`.

2. **A `PKG Summary` sheet**, grouped hierarchically:
   ```
   Ship To: DENSO (THAILAND) CO., LTD.          ← dark section header
     <ship_to> | <invoice> | <CML No.> | <desc> | nw | gw     ← one row per carton
     …
     Subtotal (N pkg)                            ← light blue, per invoice
   Ship To: DENSO CORPORATION …
     …
   <factory> Total (N pkg)                       ← medium blue, per factory
   Grand Total (N pkg)                           ← dark blue, whole report
   ```
   Columns: `Ship To | Invoice | CML No. | Description | Net Weight (kg) | Gross Weight (kg)`.

Headers are frozen, columns auto-sized. The buffer is written to
`uploads/<report_id>.xlsx` and served by `/download/<report_id>` as
`PKG_Count_Report.xlsx`.

### 3.6 Robustness

- A single unreadable PDF is caught, recorded in the preview with an `error`
  field, and the batch continues.
- `gc.collect()` and `page.flush_cache()` are called aggressively to keep memory
  flat across large batches on low-RAM hosting.
- Temp PDFs are deleted after each file is processed.

---

## 4. Tool 2 — Combine PDFs

### 4.1 Purpose

For each shipment, the broker has up to three PDFs per invoice: the Invoice
(INV), the Packing List (PL), and the Freight document. They must be submitted
as **one PDF per invoice group**, in a fixed order, sorted by invoice number.

### 4.2 Request flow

```
Browser                              Flask
  │  POST /combine (all PDFs) ─────►  │ save PDFs to uploads/combine_<batch_id>/
  │                                   │ combine(folder, uploads/combined_<batch_id>.pdf)
  │                                   │ rmtree(batch folder)
  │  ◄── { invoice_count, details,    │
  │        unclassified, download_id }│
  │  GET /download_combine/<id> ────► │ send_file(combined_<id>.pdf)
```

This route is **synchronous** (no background job) — merging pages is fast.

### 4.3 Classification — `combinepdf.classify(filename)`

Role is detected **from the filename only**:

| Filename pattern | Role |
|------------------|------|
| Pure number, e.g. `651195.pdf` | `FREIGHT` |
| Contains `FREIGHT` or starts with `FRT` (+ a 5+ digit number) | `FREIGHT` |
| Starts with `INV` (+ a 5+ digit number) | `INV` |
| Starts with `PL` (+ a 5+ digit number) | `PL` |
| Anything else | unclassified (reported, not merged) |

The 5+ digit number found in the name is the **invoice number** used for
grouping.

### 4.4 Merging — `combinepdf.combine(folder, output_file)`

1. `collect_groups()` — build `{ invoice_no: {INV: path, PL: path, FREIGHT: path} }`.
2. Sort invoice numbers **numerically** (smallest → largest).
3. For each invoice, append pages in order `INV (0) → PL (1) → FREIGHT (2)`.
4. Missing document types are noted in `details` but do **not** cause an error
   (e.g. an invoice with no freight file is still merged as INV + PL).
5. Raises `ValueError` only if **nothing** could be classified.

Returns a summary dict: `output_file`, `invoice_count`, `details[]` (human-readable
per-file log), `unclassified[]`.

`combinepdf.py` also runs standalone: `python combinepdf.py <input_folder> [output_file]`.

---

## 5. Routes reference

| Method | Route | Purpose |
|--------|-------|---------|
| GET | `/` | PKG Counter page |
| POST | `/count` | Upload packing-list PDFs, start background job → `{ job_id }` |
| GET | `/count/status/<job_id>` | Poll job progress / result |
| GET | `/download/<report_id>` | Download the generated Excel report (file is deleted from `uploads/` once the download completes fully) |
| GET | `/combine` | Combine PDFs page |
| POST | `/combine` | Upload INV/PL/Freight PDFs, merge → `{ download_id }` |
| GET | `/download_combine/<download_id>` | Download the merged PDF (file is deleted from `uploads/` once the download completes fully) |

> **Download = delete.** Both download routes stream the file and then remove it
> from `uploads/` — but only after every byte has been sent. If the download is
> aborted or the connection drops, the file is left in place so it can be
> retried. A file that is generated but never downloaded still lingers (that is
> what the Phase 4 cleanup sweeper in `UPGRADE_PLAN.md` is for).

### Limits & errors

- `MAX_UPLOAD_MB` (env var, default **75 MB**) caps the total request body;
  exceeding it returns a clean `413` JSON error.
- Empty / non-PDF uploads return `400`.
- Unknown job IDs return `404`.

---

## 6. Running it

### Local development
```bash
pip install -r requirements.txt
python app.py           # serves http://0.0.0.0:8080  (debug=True)
```

### Production (e.g. Render start command)
```bash
gunicorn app:app --workers 1 --threads 4 --timeout 180 --bind 0.0.0.0:$PORT
```
> Keep `--workers 1`: the in-memory `JOBS` store and per-worker library memory
> footprint both assume a single worker. Use threads for concurrency.

### Environment variables
| Var | Default | Meaning |
|-----|---------|---------|
| `PORT` | `8080` | Port to bind (local dev) |
| `MAX_UPLOAD_MB` | `75` | Max total upload size per request |

---

## 7. Typical end-to-end usage

**PKG Counter**
1. Open `/`.
2. Click the drop zone, select all packing-list PDFs for the shipment.
3. Click **Count PKG**. Watch the `Processing X/Y` progress.
4. Review the on-page preview tables (items, PKG count, quantities, weights).
5. Click **Download Excel Report** → `PKG_Count_Report.xlsx` with per-invoice
   sheets + the grouped `PKG Summary`.

**Combine PDFs**
1. Open `/combine`.
2. Select the INV / PL / Freight PDFs. Filenames must follow the naming rules
   in §4.3 (`INV651195.pdf`, `PL651195.pdf`, `651195.pdf`).
3. Click **Combine PDFs**.
4. Read the details log (what was merged, what was skipped).
5. Click **Download Combined PDF** → `combined_output.pdf`.

---

## 8. Known constraints / gotchas

- **Filename-driven classification** in the combiner: a mislabelled file is
  silently put in the wrong invoice group or skipped. Always check the details log.
- **Single worker only** — do not scale horizontally without moving `JOBS` to
  shared storage.
- **`uploads/` is only partly self-cleaning** — a report/merged PDF is deleted
  right after its download finishes, but files that are generated and never
  downloaded (and the raw uploaded PDFs if a job crashes) still accumulate.
  Clear them periodically (ephemeral hosting wipes them on restart anyway); a
  proper retention sweeper is Phase 4 of `UPGRADE_PLAN.md`.
- **New DENSO template?** If a packing list uses a layout none of the regexes in
  §3.3 match, its items silently won't appear. Add a new regex following the
  existing priority-ordered pattern in `parse_packing_list()`.
- The `.devcontainer/app.py` file is a **stale earlier copy** (PKG Counter only,
  no combiner, no background jobs). The real app is the top-level `app.py`.
- `.codesandbox/tasks.json` references `python main.py`, but there is no
  `main.py` — the correct entrypoint is `app.py`.
