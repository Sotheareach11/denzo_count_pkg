# DENSO Tools — PKG Counter + PDF Combiner

A small Flask web app used by CKSN International Transport to automate two
document tasks for DENSO shipments:

| Tool | URL | Input | Output |
|------|-----|-------|--------|
| **PKG Counter** | `/` | Packing-list PDFs | Excel report: package counts + net/gross weight per item, invoice, and factory |
| **Combine PDFs** | `/combine` | Invoice / Packing List / Freight PDFs | One merged PDF per invoice, ordered INV → PL → Freight |

Full write-up: [`documents/PROJECT_DOCUMENTATION.md`](documents/PROJECT_DOCUMENTATION.md).
Roadmap: [`documents/UPGRADE_PLAN.md`](documents/UPGRADE_PLAN.md).

---

## Run locally

Requires Python 3.11+ (developed on 3.13).

```powershell
# from the project folder
py -m venv .venv
.\.venv\Scripts\Activate.ps1          # Windows PowerShell
# source .venv/bin/activate           # macOS / Linux

pip install -r requirements.txt
python app.py
```

Then open:

- PKG Counter — <http://localhost:8080/>
- Combine PDFs — <http://localhost:8080/combine>

Stop with `Ctrl+C`.

## Run the tests

```powershell
pip install -r requirements-dev.txt
pytest
```

> Note: the `tests/` suite is added in Phase 1 of the upgrade plan. Until then
> `pytest` reports "no tests ran".

## Production

```bash
gunicorn app:app -c gunicorn.conf.py
```

Keep it to **one worker** (already set in `gunicorn.conf.py`) — the PKG Counter
job store is in-process memory.

### Docker

```bash
docker build -t denso-tools .
docker run --rm -p 8080:8080 denso-tools
```

## Configuration

| Env var | Default | Meaning |
|---------|---------|---------|
| `PORT` | `8080` | Port to bind |
| `MAX_UPLOAD_MB` | `75` | Max total size of one upload batch |
| `GUNICORN_THREADS` | `4` | Worker threads (production) |
| `GUNICORN_TIMEOUT` | `180` | Request timeout seconds (production) |

## Layout

```
app.py             Flask app: routes, PDF parsing, Excel building, background jobs
combinepdf.py      INV/PL/Freight classification + merge (also runs standalone)
templates/         Server-rendered HTML for the two tools
uploads/           Runtime scratch — uploaded files + generated reports (git-ignored)
gunicorn.conf.py   Production server settings
documents/         Project documentation
```

## Filename rules for the Combine tool

| Filename | Detected as |
|----------|-------------|
| `INV651195.pdf` (starts with `INV` + 5+ digits) | Invoice |
| `PL651195.pdf` (starts with `PL` + 5+ digits) | Packing List |
| `651195.pdf` (digits only), or contains `FREIGHT` / starts with `FRT` | Freight |
| anything else | skipped (listed in the result as "unclassified") |
