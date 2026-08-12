"""
PKG Counter - Web App (Flask)
=============================
Runs in the browser - works great on CodeSandbox, Replit, or any
container with no display (no tkinter needed).

Install:
    pip install flask pdfplumber openpyxl

Run:
    python app.py

Then open the preview URL CodeSandbox gives you (or http://localhost:5000
if running locally). Select one or more PDF packing lists, click
"Count PKG", and an Excel file downloads automatically.

Output workbook layout:
  - One "item summary" sheet PER PDF (same as before: Item No.,
    Description, PKG count, Qty) - named after the PDF file.
  - ONE single combined "PKG Summary" sheet (not per PDF) listing every
    package (CML No.) from every uploaded PDF, with a Description column
    (falls back to Customer Item No. if the description text couldn't be
    detected), Net Weight, Gross Weight, a subtotal per invoice, and a
    Grand Total at the end.

Parsing note on Description / Customer Item No.:
  Normally a line looks like "KN127314-3110 TUBE" -> Customer Item No.
  "KN127314-3110", Description "TUBE", both on the same line. Sometimes
  the description text is missing from that line (only the Customer Item
  No. token is present) and the actual name appears by itself on the
  next line instead. The parser now looks ahead for that case, and if it
  still can't find a description anywhere, it falls back to displaying
  the Customer Item No. as the name so the row is never blank.



================ for run python ====================
python app.py
===================================================
"""

import os
import re
import io
import uuid
from collections import OrderedDict

from flask import Flask, request, render_template, send_file, jsonify
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# PDF PARSING
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


def parse_packing_list(pdf_path):
    """
    Returns (records, packages)

    records  - one dict per order line (used for the item summary sheet).
               Always has both "custitem" and "desc" filled in as best as
               possible (desc may be "" if truly not found anywhere).
    packages - OrderedDict keyed by CML No., each value:
               {"cml": ..., "vol": float, "nw": float, "gw": float}
               (used for the combined PKG weight sheet)
    """
    # Flatten all pages into one list of stripped, non-header lines so we
    # can look ahead across line boundaries when a description is missing.
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
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
                # Normal case: "KN127314-3110 TUBE" on one line
                custitem = m.group("custitem")
                desc = m.group("desc").strip()
                i += 1
            elif " " not in line:
                # Only the Customer Item No. token is on this line - the
                # description (if any) may be sitting by itself on the
                # next line.
                custitem = line
                i += 1
                if i < n:
                    nxt = lines[i]
                    if not CML_RE.match(nxt) and not ORDER_RE.match(nxt) \
                            and not nxt.startswith(HEADER_PREFIXES):
                        desc = nxt.strip()
                        i += 1
            else:
                # Unexpected shape - treat the whole line as the description
                desc = line.strip()
                i += 1

            rec = dict(pending_order)
            rec["cml"] = cur_cml
            rec["custitem"] = custitem
            rec["desc"] = desc
            records.append(rec)
            pending_order = None
            continue

        # Unrecognized line (stray text) - skip it
        i += 1

    return records, packages


def display_name(custitem, desc):
    """Description if we have one, otherwise fall back to Customer Item No."""
    desc = (desc or "").strip()
    if desc:
        return desc
    return (custitem or "").strip() or "(unknown)"


def summarize_by_name(records):
    """
    Groups line items by their CML (physical package) first, then by item name.

    - If a package contains the SAME item split across multiple order lines
      (e.g. two order numbers for the same part in one CML), they are merged
      and counted as ONE pkg for that package, not two.
    - If a package contains DIFFERENT items mixed together in one CML, only
      the item with the largest net weight in that package gets +1 pkg;
      the other items in that same package get +0 (that package is already
      accounted for by the winner).

    The final "count" per item is the sum of its pkg contributions across
    every package (CML) it appeared in.

    Also returns package_names: OrderedDict of CML No. -> display name
    (Description, falling back to Customer Item No.) for the item that
    "won" that package - used by the PKG Summary sheet.
    """
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
# EXCEL EXPORT
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
    """rows_as_strs: list of lists of strings, one list per data row."""
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows_as_strs:
            max_len = max(max_len, len(row[col_idx - 1]))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4


def build_excel(all_results):
    """
    all_results: list of (pdf_name, item_summary_rows, package_rows, package_names)
    Returns an in-memory Excel file (BytesIO):

      - one item-summary sheet PER PDF (same as the old code)
      - ONE single combined "PKG Summary" sheet with Description (falls
        back to Customer Item No.), Net Weight, Gross Weight per package,
        subtotal per invoice, and a grand total
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
    grand_fill = PatternFill(start_color="305496",
                             end_color="305496", fill_type="solid")
    grand_font = Font(bold=True, color="FFFFFF")

    item_headers = ["Item No.", "Customer Item No.",
                     "Description", "Unit", "PKG (Count)", "Total Qty"]

    # --- One item-summary sheet per PDF (unchanged, like the old code) ---
    for pdf_name, rows, packages, package_names in all_results:
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

    # --- One single combined PKG Summary sheet ---
    pkg_headers = ["Invoice", "CML No.", "Description",
                   "Net Weight (kg)", "Gross Weight (kg)"]
    ws_pkg = wb.create_sheet(title="PKG Summary")
    style_header(ws_pkg, pkg_headers, header_fill, header_font)

    pkg_row_strs = []
    grand_nw = 0.0
    grand_gw = 0.0
    grand_pkg_count = 0

    for pdf_name, rows, packages, package_names in all_results:
        invoice_label = os.path.splitext(pdf_name)[0]
        inv_nw = 0.0
        inv_gw = 0.0

        for pkg in packages.values():
            name = package_names.get(pkg["cml"], "")
            ws_pkg.append([invoice_label, pkg["cml"], name, pkg["nw"], pkg["gw"]])
            pkg_row_strs.append(
                [invoice_label, pkg["cml"], name,
                 f'{pkg["nw"]:.3f}', f'{pkg["gw"]:.3f}'])
            inv_nw += pkg["nw"]
            inv_gw += pkg["gw"]

        # Subtotal row for this invoice
        sub_row_idx = ws_pkg.max_row + 1
        ws_pkg.append([invoice_label, f"Subtotal ({len(packages)} pkg)", "",
                        round(inv_nw, 3), round(inv_gw, 3)])
        for col_idx in range(1, len(pkg_headers) + 1):
            cell = ws_pkg.cell(row=sub_row_idx, column=col_idx)
            cell.fill = sub_fill
            cell.font = sub_font
        pkg_row_strs.append(
            [invoice_label, f"Subtotal ({len(packages)} pkg)", "",
             f"{inv_nw:.3f}", f"{inv_gw:.3f}"])

        grand_nw += inv_nw
        grand_gw += inv_gw
        grand_pkg_count += len(packages)

    # Grand total row across all invoices
    grand_row_idx = ws_pkg.max_row + 1
    ws_pkg.append(["", f"Grand Total ({grand_pkg_count} pkg)", "",
                    round(grand_nw, 3), round(grand_gw, 3)])
    for col_idx in range(1, len(pkg_headers) + 1):
        cell = ws_pkg.cell(row=grand_row_idx, column=col_idx)
        cell.fill = grand_fill
        cell.font = grand_font
    pkg_row_strs.append(
        ["", f"Grand Total ({grand_pkg_count} pkg)", "",
         f"{grand_nw:.3f}", f"{grand_gw:.3f}"])

    autosize(ws_pkg, pkg_headers, pkg_row_strs)
    ws_pkg.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# ROUTES
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
    preview = []  # for showing counts back on the page

    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            continue
        temp_path = os.path.join(
            UPLOAD_DIR, f"{uuid.uuid4().hex}_{f.filename}")
        f.save(temp_path)
        try:
            records, packages = parse_packing_list(temp_path)
            summary, package_names = summarize_by_name(records)
            all_results.append((f.filename, summary, packages, package_names))

            total_nw = sum(p["nw"] for p in packages.values())
            total_gw = sum(p["gw"] for p in packages.values())

            preview.append({
                "filename": f.filename,
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

    # Store the excel temporarily so the frontend can fetch it via a download link
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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)