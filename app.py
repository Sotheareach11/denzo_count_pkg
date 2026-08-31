"""
PKG Counter + Combine PDFs - Web App (Flask)
=============================================
Runs in the browser - works great on CodeSandbox, Replit, or any
container with no display (no tkinter needed).

Install:
    pip install flask pdfplumber openpyxl pypdf

Run:
    python app.py

Then open the preview URL CodeSandbox gives you (or http://localhost:8080
if running locally).

Two tools in this app:
  1. PKG Counter ("/")        - upload packing list PDFs, get an Excel
                                 report of PKG counts. The "PKG Summary"
                                 sheet is grouped first by Ship To
                                 (factory) detected in each PDF, then by
                                 invoice within each factory.
  2. Combine PDFs ("/combine") - upload INV / PL / Freight PDFs, get them
                                 merged into one PDF, grouped by invoice
                                 number and ordered INV -> PL -> Freight.
"""

import os
import re
import io
import uuid
import shutil
from collections import OrderedDict
from combinepdf import combine

from flask import Flask, request, render_template, send_file, jsonify
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# PDF PARSING (PKG Counter) 111
# ---------------------------------------------------------------------------

# CML No.   Volume(m3)   Net Weight(kg)   Gross Weight(kg)
CML_RE = re.compile(
    r'^(?P<cml>\S+)\s+(?P<vol>[\d.]+)\s+(?P<nw>[\d.,]+)\s+(?P<gw>[\d.,]+)$')

ORDER_RE = re.compile(
    r'^(?P<order>\S+K\d+)\s+(?P<itemno>\S+)\s+(?P<unit>\S+)\s+'
    r'(?P<qty>[\d,]+)\s+(?P<totalnw>[\d.,]+)\s+(?P<cartons>\d+)$'
)

# "KN127314-3110 TUBE" -> custitem + desc, both on the same line
DESC_RE = re.compile(r'^(?P<custitem>\S+)\s+(?P<desc>.+)$')

HEADER_PREFIXES = (
    "CML No.", "Customer Order No.", "Customer Item No.",
    "Page", "Total Package",
)

# Marker used to find the "Ship to" company name on page 1.
SHIP_TO_MARKER_RE = re.compile(r'Ship\s*to', re.IGNORECASE)
COMPANY_HINT_RE = re.compile(r'CO\.,?\s*LTD', re.IGNORECASE)


def extract_ship_to(first_page_text):
    """
    Detects the "Ship to" company / factory name from the raw text of a
    packing list's first page.

    pdfplumber often merges the "Sold to" and "Ship to" table columns onto
    the same line (since they sit at the same y-position in the PDF), e.g.:

        Ship to Document Information DENSO (THAILAND) CO., LTD.
        DENSO (THAILAND) CO., LTD. Document Type NORMAL ...

    So instead of taking a whole line, we find the "Ship to" marker and
    then grab text up to (and including) the *first* "CO., LTD." style
    company suffix that follows it - that's the factory name, and it
    stops before any of the following merged-in text.

    Returns the company name (e.g. "DENSO (THAILAND) CO., LTD.") or
    "Unknown" if it can't be found, so a PDF with an unrecognised layout
    still gets grouped (under "Unknown") instead of crashing the report.
    """
    text = " ".join(first_page_text.split()
                    )  # collapse all whitespace/newlines

    m = SHIP_TO_MARKER_RE.search(text)
    if not m:
        return "Unknown"

    tail = text[m.end():]
    # "Document Information" is the merged-in column header next to "Ship to" - drop it.
    tail = re.sub(r'^\s*Document Information\s*',
                  '', tail, flags=re.IGNORECASE)

    company_match = re.match(r'\s*(.*?CO\.,?\s*LTD\.?)', tail, re.IGNORECASE)
    if company_match:
        return company_match.group(1).strip()

    return "Unknown"


def parse_packing_list(pdf_path):
    """
    Returns (records, packages, ship_to)

    records  - one dict per order line (used for the item summary sheet).
               Always has both "custitem" and "desc" filled in as best as
               possible (desc may be "" if truly not found anywhere).
    packages - OrderedDict keyed by CML No., each value:
               {"cml": ..., "vol": float, "nw": float, "gw": float}
               (used for the combined PKG weight sheet)
    ship_to  - detected Ship To company / factory name (string), used to
               group the PKG Summary sheet by factory.
    """
    lines = []
    ship_to = "Unknown"
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if page_idx == 0:
                ship_to = extract_ship_to(text)
            for raw_line in text.split("\n"):
                line = raw_line.strip()
                if not line:
                    continue
                if line.startswith(HEADER_PREFIXES):
                    continue
                lines.append(line)

    records = []
    packages = OrderedDict()
    cur_cml = None
    pending_order = None

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]

        m = CML_RE.match(line)
        if m:
            cur_cml = m.group("cml")
            packages[cur_cml] = {
                "cml": cur_cml,
                "vol": float(m.group("vol")),
                "nw": float(m.group("nw").replace(",", "")),
                "gw": float(m.group("gw").replace(",", "")),
            }
            i += 1
            continue

        m = ORDER_RE.match(line)
        if m:
            pending_order = m.groupdict()
            i += 1
            continue

        if pending_order is not None:
            custitem = ""
            desc = ""

            m = DESC_RE.match(line)
            if m:
                custitem = m.group("custitem")
                desc = m.group("desc").strip()
                i += 1
            elif " " not in line:
                custitem = line
                i += 1
                if i < n:
                    nxt = lines[i]
                    if not CML_RE.match(nxt) and not ORDER_RE.match(nxt) \
                            and not nxt.startswith(HEADER_PREFIXES):
                        desc = nxt.strip()
                        i += 1
            else:
                desc = line.strip()
                i += 1

            rec = dict(pending_order)
            rec["cml"] = cur_cml
            rec["custitem"] = custitem
            rec["desc"] = desc
            records.append(rec)
            pending_order = None
            continue

        i += 1

    return records, packages, ship_to


def display_name(custitem, desc):
    desc = (desc or "").strip()
    if desc:
        return desc
    return (custitem or "").strip() or "(unknown)"


def summarize_by_name(records):
    cml_groups = OrderedDict()
    for r in records:
        cml_groups.setdefault(r["cml"], []).append(r)

    summary = OrderedDict()
    package_names = OrderedDict()

    for cml, lines in cml_groups.items():
        within = OrderedDict()
        for r in lines:
            key = (r["itemno"], r["desc"])
            if key not in within:
                within[key] = {
                    "itemno": r["itemno"], "desc": r["desc"], "custitem": r["custitem"],
                    "unit": r["unit"], "qty": 0, "netweight": 0.0,
                }
            within[key]["qty"] += int(r["qty"].replace(",", ""))
            within[key]["netweight"] += float(r["totalnw"].replace(",", ""))

        winner_key = max(within.keys(), key=lambda k: within[k]["netweight"])
        winner = within[winner_key]
        package_names[cml] = display_name(winner["custitem"], winner["desc"])

        for key, data in within.items():
            if key not in summary:
                summary[key] = {
                    "itemno": data["itemno"], "custitem": data["custitem"],
                    "desc": data["desc"], "unit": data["unit"],
                    "count": 0, "total_qty": 0,
                }
            summary[key]["total_qty"] += data["qty"]
            if key == winner_key:
                summary[key]["count"] += 1

    return list(summary.values()), package_names


# ---------------------------------------------------------------------------
# EXCEL EXPORT (PKG Counter)
# ---------------------------------------------------------------------------

def safe_sheet_name(name, used_names):
    base = re.sub(r'[\\/*?:\[\]]', "_", name)[:31]
    candidate = base
    i = 1
    while candidate in used_names:
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        i += 1
    used_names.add(candidate)
    return candidate


def style_header(ws, headers, header_fill, header_font):
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")


def autosize(ws, headers, rows_as_strs):
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows_as_strs:
            max_len = max(max_len, len(row[col_idx - 1]))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4


def build_excel(all_results):
    """
    all_results: list of tuples
        (pdf_name, item_rows, packages, package_names, ship_to)
    """
    wb = Workbook()
    wb.remove(wb.active)

    used_names = set()
    header_fill = PatternFill(start_color="305496",
                              end_color="305496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    sub_fill = PatternFill(start_color="D9E1F2",
                           end_color="D9E1F2", fill_type="solid")
    sub_font = Font(bold=True)
    factory_fill = PatternFill(start_color="8EA9DB",
                               end_color="8EA9DB", fill_type="solid")
    factory_font = Font(bold=True)
    section_fill = PatternFill(start_color="203864",
                               end_color="203864", fill_type="solid")
    section_font = Font(bold=True, color="FFFFFF")
    grand_fill = PatternFill(start_color="305496",
                             end_color="305496", fill_type="solid")
    grand_font = Font(bold=True, color="FFFFFF")

    item_headers = ["Item No.", "Customer Item No.",
                    "Description", "Unit", "PKG (Count)", "Total Qty"]

    # ---- Per-invoice item summary sheets ----
    for pdf_name, rows, packages, package_names, ship_to in all_results:
        base_title = os.path.splitext(pdf_name)[0]
        sheet_title = safe_sheet_name(base_title, used_names)
        ws = wb.create_sheet(title=sheet_title)
        style_header(ws, item_headers, header_fill, header_font)
        for row in rows:
            desc_display = display_name(row["custitem"], row["desc"])
            ws.append([
                row["itemno"], row["custitem"], desc_display,
                row["unit"], row["count"], row["total_qty"],
            ])
        row_strs = [
            [row["itemno"], row["custitem"], display_name(row["custitem"], row["desc"]),
             row["unit"], str(row["count"]), str(row["total_qty"])]
            for row in rows
        ]
        autosize(ws, item_headers, row_strs)
        ws.freeze_panes = "A2"

    # ---- PKG Summary sheet, grouped by Ship To (factory), then invoice ----
    pkg_headers = ["Ship To", "Invoice", "CML No.", "Description",
                   "Net Weight (kg)", "Gross Weight (kg)"]
    ws_pkg = wb.create_sheet(title="PKG Summary")
    style_header(ws_pkg, pkg_headers, header_fill, header_font)

    # Group results by detected Ship To, preserving first-seen order.
    grouped_by_ship_to = OrderedDict()
    for pdf_name, rows, packages, package_names, ship_to in all_results:
        grouped_by_ship_to.setdefault(ship_to, []).append(
            (pdf_name, packages, package_names))

    pkg_row_strs = []
    grand_nw = 0.0
    grand_gw = 0.0
    grand_pkg_count = 0

    for ship_to, items in grouped_by_ship_to.items():
        # --- Factory section header row ---
        section_row_idx = ws_pkg.max_row + 1
        ws_pkg.append([f"Ship To: {ship_to}", "", "", "", "", ""])
        ws_pkg.merge_cells(start_row=section_row_idx, start_column=1,
                           end_row=section_row_idx, end_column=len(pkg_headers))
        for col_idx in range(1, len(pkg_headers) + 1):
            cell = ws_pkg.cell(row=section_row_idx, column=col_idx)
            cell.fill = section_fill
            cell.font = section_font
        pkg_row_strs.append([f"Ship To: {ship_to}", "", "", "", "", ""])

        ship_nw = 0.0
        ship_gw = 0.0
        ship_pkg_count = 0

        for pdf_name, packages, package_names in items:
            invoice_label = os.path.splitext(pdf_name)[0]
            inv_nw = 0.0
            inv_gw = 0.0

            for pkg in packages.values():
                name = package_names.get(pkg["cml"], "")
                ws_pkg.append([ship_to, invoice_label, pkg["cml"], name,
                               pkg["nw"], pkg["gw"]])
                pkg_row_strs.append(
                    [ship_to, invoice_label, pkg["cml"], name,
                     f'{pkg["nw"]:.3f}', f'{pkg["gw"]:.3f}'])
                inv_nw += pkg["nw"]
                inv_gw += pkg["gw"]

            sub_row_idx = ws_pkg.max_row + 1
            ws_pkg.append([ship_to, invoice_label,
                           f"Subtotal ({len(packages)} pkg)", "",
                           round(inv_nw, 3), round(inv_gw, 3)])
            for col_idx in range(1, len(pkg_headers) + 1):
                cell = ws_pkg.cell(row=sub_row_idx, column=col_idx)
                cell.fill = sub_fill
                cell.font = sub_font
            pkg_row_strs.append(
                [ship_to, invoice_label, f"Subtotal ({len(packages)} pkg)", "",
                 f"{inv_nw:.3f}", f"{inv_gw:.3f}"])

            ship_nw += inv_nw
            ship_gw += inv_gw
            ship_pkg_count += len(packages)

        # --- Factory (Ship To) subtotal row ---
        factory_row_idx = ws_pkg.max_row + 1
        ws_pkg.append(["", f"{ship_to} Total ({ship_pkg_count} pkg)", "", "",
                       round(ship_nw, 3), round(ship_gw, 3)])
        for col_idx in range(1, len(pkg_headers) + 1):
            cell = ws_pkg.cell(row=factory_row_idx, column=col_idx)
            cell.fill = factory_fill
            cell.font = factory_font
        pkg_row_strs.append(
            ["", f"{ship_to} Total ({ship_pkg_count} pkg)", "", "",
             f"{ship_nw:.3f}", f"{ship_gw:.3f}"])

        grand_nw += ship_nw
        grand_gw += ship_gw
        grand_pkg_count += ship_pkg_count

    # --- Grand total row across all factories ---
    grand_row_idx = ws_pkg.max_row + 1
    ws_pkg.append(["", f"Grand Total ({grand_pkg_count} pkg)", "", "",
                   round(grand_nw, 3), round(grand_gw, 3)])
    for col_idx in range(1, len(pkg_headers) + 1):
        cell = ws_pkg.cell(row=grand_row_idx, column=col_idx)
        cell.fill = grand_fill
        cell.font = grand_font
    pkg_row_strs.append(
        ["", f"Grand Total ({grand_pkg_count} pkg)", "", "",
         f"{grand_nw:.3f}", f"{grand_gw:.3f}"])

    autosize(ws_pkg, pkg_headers, pkg_row_strs)
    ws_pkg.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# ROUTES - PKG Counter
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/count", methods=["POST"])
def count_pkg():
    files = request.files.getlist("pdfs")
    if not files or files[0].filename == "":
        return jsonify({"error": "No files uploaded."}), 400

    all_results = []
    preview = []

    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            continue
        temp_path = os.path.join(
            UPLOAD_DIR, f"{uuid.uuid4().hex}_{f.filename}")
        f.save(temp_path)
        try:
            records, packages, ship_to = parse_packing_list(temp_path)
            summary, package_names = summarize_by_name(records)
            all_results.append(
                (f.filename, summary, packages, package_names, ship_to))

            total_nw = sum(p["nw"] for p in packages.values())
            total_gw = sum(p["gw"] for p in packages.values())

            preview.append({
                "filename": f.filename,
                "ship_to": ship_to,
                "items": [
                    {"itemno": r["itemno"],
                        "desc": display_name(r["custitem"], r["desc"]),
                        "pkg": r["count"], "qty": r["total_qty"]}
                    for r in summary
                ],
                "packages": [
                    {"cml": p["cml"], "name": package_names.get(p["cml"], ""),
                     "nw": p["nw"], "gw": p["gw"]}
                    for p in packages.values()
                ],
                "total_pkg": len(packages),
                "total_nw": round(total_nw, 3),
                "total_gw": round(total_gw, 3),
            })
        finally:
            os.remove(temp_path)

    if not all_results:
        return jsonify({"error": "No valid PDF data extracted."}), 400

    excel_buffer = build_excel(all_results)

    report_id = uuid.uuid4().hex
    report_path = os.path.join(UPLOAD_DIR, f"{report_id}.xlsx")
    with open(report_path, "wb") as out:
        out.write(excel_buffer.getvalue())

    return jsonify({"preview": preview, "download_id": report_id})


@app.route("/download/<report_id>")
def download(report_id):
    report_path = os.path.join(UPLOAD_DIR, f"{report_id}.xlsx")
    if not os.path.isfile(report_path):
        return "Report not found or expired.", 404
    return send_file(
        report_path,
        as_attachment=True,
        download_name="PKG_Count_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# ROUTES - Combine PDFs
# ---------------------------------------------------------------------------

@app.route("/combine")
def combine_page():
    return render_template("combine.html")


@app.route("/combine", methods=["POST"])
def combine_pdfs_route():
    files = request.files.getlist("pdfs")
    if not files or files[0].filename == "":
        return jsonify({"error": "No files uploaded."}), 400

    batch_id = uuid.uuid4().hex
    batch_folder = os.path.join(UPLOAD_DIR, f"combine_{batch_id}")
    os.makedirs(batch_folder, exist_ok=True)

    for f in files:
        if f.filename.lower().endswith(".pdf"):
            f.save(os.path.join(batch_folder, f.filename))

    output_path = os.path.join(UPLOAD_DIR, f"combined_{batch_id}.pdf")

    try:
        result = combine(batch_folder, output_path)
    except ValueError as e:
        shutil.rmtree(batch_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 400

    shutil.rmtree(batch_folder, ignore_errors=True)

    return jsonify({
        "invoice_count": result["invoice_count"],
        "details": result["details"],
        "unclassified": result["unclassified"],
        "download_id": f"combined_{batch_id}",
    })


@app.route("/download_combine/<download_id>")
def download_combine(download_id):
    file_path = os.path.join(UPLOAD_DIR, f"{download_id}.pdf")
    if not os.path.isfile(file_path):
        return "File not found or expired.", 404
    return send_file(
        file_path,
        as_attachment=True,
        download_name="combined_output.pdf",
        mimetype="application/pdf",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)
