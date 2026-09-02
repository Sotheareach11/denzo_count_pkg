# Upgrade Plan — from "working script" to "maintainable product"

Goal: keep the same two features (PKG Counter + Combine PDFs) but make the code
**faster to change**, **easier to use**, and **flexible** when DENSO changes a
form or a new factory/customer appears.

This document is a roadmap. Nothing here changes behaviour for the end user
except where marked **UX**.

---

## 1. Where the pain is today (review findings)

| # | Problem | Why it hurts maintenance |
|---|---------|--------------------------|
| 1 | **One 800-line `app.py`** mixing routing, PDF parsing, aggregation, Excel building, and the job store | Any change means reading the whole file; easy to break something unrelated |
| 2 | **Template parsing is a single regex chain** in `parse_packing_list()` with hidden priority ordering | Adding a new DENSO layout = editing the hottest, riskiest function with no safety net |
| 3 | **No tests / no sample PDFs** | You cannot tell if a change broke an existing template until a user complains |
| 4 | **Everything hardcoded** — factory-name rules, filename rules, output columns, size limits | Small business-rule tweaks require a code change + redeploy |
| 5 | **`JOBS` is an in-memory dict** | Locked to 1 worker; every restart loses in-flight jobs; no history |
| 6 | **`uploads/` grows forever** and **76 old test files are committed to git** | Repo is bloated; disk fills on a real server |
| 7 | **Two near-identical HTML files** with copy-pasted CSS + fetch/poll JS | Every UI fix must be done twice |
| 8 | **Silent failures** — an unmatched packing-list template just produces empty rows; combiner misfiles by filename with no content check | Wrong output looks like correct output |
| 9 | **Stale `.devcontainer/app.py`**, `.codesandbox` points at non-existent `main.py`, Python pinned to 3.8 in devcontainer vs 3.13 locally | New environment set-up is confusing and inconsistent |
| 10 | **`requirements.txt` mixes runtime + linters**, over-pinned (`cryptography`, `cffi`…) | Slow, fragile installs; unclear what is actually needed |

---

## 2. Target structure

```
denso_tools/
├── run.py                     # `python run.py` for local dev
├── gunicorn.conf.py           # production settings in one file, not a long CLI
├── pyproject.toml             # deps + tool config (replaces requirements.txt)
├── config.yaml                # business rules — see §5
├── Dockerfile / compose.yaml  # reproducible deploy
│
├── app/
│   ├── __init__.py            # Flask app factory, blueprint registration
│   ├── config.py              # load config.yaml + env overrides
│   ├── jobs.py                # JobStore interface (MemoryJobStore now, RedisJobStore later)
│   ├── cleanup.py             # background sweeper: delete files older than N hours
│   │
│   ├── pkgcounter/
│   │   ├── routes.py          # /  /count  /count/status  /download
│   │   ├── parsing.py         # orchestrates: pick parser -> parse -> aggregate
│   │   ├── aggregate.py       # summarize_by_name()
│   │   ├── excel.py           # build_excel()
│   │   └── parsers/
│   │       ├── base.py        # PackingListParser ABC + shared regex helpers
│   │       ├── registry.py    # ordered list of parser classes
│   │       ├── denso_normal.py
│   │       ├── denso_thailand.py
│   │       └── denso_return_style.py
│   │
│   ├── combiner/
│   │   ├── routes.py          # /combine  /download_combine
│   │   ├── classify.py        # filename rules + optional PDF-content fallback
│   │   └── merge.py
│   │
│   ├── templates/
│   │   ├── base.html          # ONE layout: header, styles, drop-zone, progress
│   │   ├── pkgcounter.html    # extends base
│   │   └── combiner.html      # extends base
│   └── static/
│       ├── app.css
│       └── upload.js          # ONE shared uploader/poller module
│
└── tests/
    ├── fixtures/
    │   ├── denso_normal/sample.pdf        + expected.json
    │   ├── denso_thailand/sample.pdf      + expected.json
    │   └── denso_return_style/sample.pdf  + expected.json
    ├── test_parsers.py        # golden-file: parse(sample.pdf) == expected.json
    ├── test_aggregate.py
    └── test_combiner.py
```

---

## 3. The key change: pluggable packing-list parsers

This is the single highest-value refactor — it makes "a new DENSO form" a
**10-minute, low-risk job** instead of a scary one.

### 3.1 The contract (`parsers/base.py`)

```python
from dataclasses import dataclass, field

@dataclass
class Package:
    cml: str
    vol: float
    nw: float
    gw: float

@dataclass
class ItemRecord:
    cml: str | None
    itemno: str
    custitem: str
    desc: str
    unit: str
    qty: int
    total_nw: float

@dataclass
class ParsedPackingList:
    ship_to: str
    packages: list[Package]
    records: list[ItemRecord]

class PackingListParser:
    name: str                       # "denso_thailand"

    def matches(self, first_page_text: str) -> bool:
        """Cheap check — does this template look like mine?"""
        raise NotImplementedError

    def parse(self, pages_text: list[str]) -> ParsedPackingList:
        raise NotImplementedError
```

### 3.2 The registry (`parsers/registry.py`)

```python
PARSERS = [
    DensoNormalParser(),
    DensoThailandParser(),
    DensoReturnStyleParser(),
]

def pick_parser(first_page_text: str) -> PackingListParser | None:
    for p in PARSERS:
        if p.matches(first_page_text):
            return p
    return None            # -> caller flags the file as "Unknown template"
```

### 3.3 Orchestration (`parsing.py`)

```python
def parse_packing_list(pdf_path):
    pages = extract_pages_text(pdf_path)          # pdfplumber, with flush_cache
    parser = pick_parser(pages[0])
    if parser is None:
        raise UnknownTemplateError(pdf_path, snippet=pages[0][:800])
    result = parser.parse(pages)
    result.parser_name = parser.name             # show it in the UI
    return result
```

### 3.4 Adding a new template becomes a checklist

1. Drop an anonymised `sample.pdf` in `tests/fixtures/denso_newthing/`.
2. Copy the closest existing parser module, adjust its regexes + `matches()`.
3. Add it to `PARSERS`.
4. Write `expected.json` (or generate it once and eyeball it).
5. `pytest` — green means every **other** template still works.

No touching of unrelated code, and the test suite is your seatbelt.

---

## 4. Tests — do this alongside §3, not later

Without fixtures you cannot refactor safely. Steps:

1. Take 1 real PDF per known template, **redact** customer-identifying data
   (keep the layout/structure — that's all the parser cares about).
2. Commit them under `tests/fixtures/`.
3. `test_parsers.py`:
   ```python
   @pytest.mark.parametrize("case", list_fixture_dirs())
   def test_parser_golden(case):
       result = parse_packing_list(case / "sample.pdf")
       assert asdict(result) == json.loads((case / "expected.json").read_text())
   ```
4. Add `test_combiner.py` for the filename classifier (pure function, trivial to test).

Target: **the whole suite runs in < 5 seconds** so you run it on every change.

---

## 5. Move business rules into `config.yaml`

```yaml
upload:
  max_mb: 75
  retention_hours: 24          # cleanup sweeper deletes older files

pkg_summary:
  columns: [ship_to, invoice, cml, description, net_weight, gross_weight]
  # normalise messy detected names to a canonical factory label
  ship_to_aliases:
    "DENSO (THAILAND) CO.,LTD": "DENSO (THAILAND) CO., LTD."

combiner:
  order: [INV, PL, FREIGHT]
  filename_rules:
    FREIGHT: '^\d+$'
    INV:     '^INV'
    PL:      '^PL'
  require_freight: false
```

`app/config.py` loads this, applies `MAX_UPLOAD_MB` / env overrides on top.
Now a rule change is a config edit + restart, not a code deploy.

---

## 6. Robustness / job store / cleanup

- **`JobStore` interface** in `jobs.py`. Ship `MemoryJobStore` (what exists now,
  just wrapped). The rest of the app never touches the dict directly, so
  swapping in `RedisJobStore` later is one class + one config line — no route
  changes.
- **Cleanup sweeper** (`cleanup.py`): a daemon thread (or APScheduler job) that
  every 30 min deletes files in `uploads/` older than `retention_hours`. Fixes
  "disk fills forever".
- **Structured logging** via `logging` — one line per job: files in, template
  matched per file, rows out, duration, errors. When a user reports "the numbers
  look wrong" you can actually see what happened.
- **`/health` endpoint** for uptime checks.
- **Loud "Unknown template"**: the preview must show a red banner naming the
  file and dumping the first ~10 detected lines, so you can build a parser for
  it fast. Currently it silently yields empty rows.

---

## 7. UX upgrades (optional, ordered by value)

| **UX** improvement | Effort | Payoff |
|--------------------|--------|--------|
| **One shared `base.html` + `upload.js`** — kill the duplication | S | Every future UI fix done once |
| **Show matched template + row count per file** in the preview | S | Instant "did it read this right?" confidence |
| **Editable preview grid** — fix a wrong description / qty inline, then "Regenerate Excel" from the corrected data | M | No more re-exporting from a hand-fixed spreadsheet |
| **Unified shipment flow** — drop *all* PDFs for a shipment once; app auto-splits packing lists (→ counter) from INV/PL/Freight (→ combiner) and returns both the Excel and the merged PDF | M | Half the clicks, one screen per shipment |
| **Recent reports list** — `/reports` showing the last N generated files with re-download links | S | Stop re-running a batch because you lost the download |
| **Drag-and-drop + folder upload** on the drop zone | S | Faster input for big batches |
| Replace hand-rolled fetch/poll with **Alpine.js or HTMX** (from the allowed CDN) | M | Much less JS to maintain |

---

## 8. Repo & environment hygiene (do this first — 1 hour, zero risk)

```bash
# 1. stop tracking generated files
git rm -r --cached uploads
printf '.venv/\n__pycache__/\n*.pyc\nuploads/*\n!uploads/.gitkeep\n*.xlsx\n*.pdf\n' > .gitignore
touch uploads/.gitkeep
git add .gitignore uploads/.gitkeep

# 2. delete the stale duplicate
git rm .devcontainer/app.py

# 3. split dependencies
#    pyproject.toml [project].dependencies  -> flask, pdfplumber, openpyxl, pypdf
#    [project.optional-dependencies].dev    -> pytest, ruff, black
```

- Fix `.codesandbox/tasks.json`: `python main.py` → `python run.py`.
- Bump `.devcontainer/devcontainer.json` image to `python:3.12` (match reality).
- Add a real `README.md`: what it does, `pip install -e .`, `python run.py`,
  `pytest`, deploy command.
- Add `.dockerignore` + a `Dockerfile` (python:3.12-slim, non-root, gunicorn).

---

## 9. Suggested execution order

| Phase | Work | Effort | Risk | Outcome |
|-------|------|--------|------|---------|
| **0** | §8 hygiene, `.gitignore`, split deps, fix devcontainer/codesandbox, README, Dockerfile | ~0.5 day | none | Clean base, reproducible setup |
| **1** | §4 collect + redact sample PDFs, stand up `pytest` with golden tests against the **current** code | ~1 day | none | Safety net before any refactor |
| **2** | §2 split `app.py` into `app/` package (blueprints, jobs, excel, aggregate) — no logic changes, tests stay green | ~1–2 days | low | Small files, clear ownership |
| **3** | §3 extract the regex chain into `parsers/` with the registry + `matches()` | ~1–2 days | medium (tests cover it) | New templates are a 10-min job |
| **4** | §5 `config.yaml` + §6 cleanup sweeper + logging + `/health` | ~1 day | low | Business rules & ops without redeploy |
| **5** | §7 UX: shared base template, per-file template badge, recent-reports list | ~1–2 days | low | Faster, clearer daily use |
| **6** | §7 bigger UX: editable preview grid, unified shipment flow | ~3–5 days | medium | The "master" experience |
| **7** | (only if multi-user/scale needed) RedisJobStore, auth, audit log | later | — | Team-ready |

Phases 0–3 are the core of "maintainable and flexible". Do those first; 4–6 are
where it starts to feel like a product.

---

## 10. What NOT to do

- Don't rewrite in a new framework/language — the logic is fine, the *packaging*
  is the problem.
- Don't build a fully data-driven (YAML-defined) parser engine yet — Python
  parser classes + tests give you 90% of the flexibility for 20% of the effort.
  Revisit only if you end up with 8+ templates.
- Don't add a database until more than one person needs shared job history.
