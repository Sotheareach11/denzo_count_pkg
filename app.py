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
"Count PKG", and an Excel file (one sheet per PDF) downloads automatically.



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

CML_RE = re.compile(
    r'^(?P<cml>\S+)\s+(?P<vol>[\d.]+)\s+(?P<nw>[\d.,]+)\s+(?P<gw>[\d.,]+)$')

ORDER_RE = re.compile(
    r'^(?P<order>\S+K\d+)\s+(?P<itemno>\S+)\s+(?P<unit>\S+)\s+'
    r'(?P<qty>[\d,]+)\s+(?P<totalnw>[\d.,]+)\s+(?P<cartons>\d+)$'
)

DESC_RE = re.compile(r'^(?P<custitem>\S+)\s+(?P<desc>.+)$')


def parse_packing_list(pdf_path):
    records = []
    cur_cml = None
    pending_order = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.split("\n"):
                line = raw_line.strip()
                if not line:
                    continue
                if line.startswith("CML No.") or line.startswith("Customer Order No.") \
                        or line.startswith("Customer Item No.") or line.startswith("Page"):
                    continue

                m = CML_RE.match(line)
                if m:
                    cur_cml = m.group("cml")
                    continue

                m = ORDER_RE.match(line)
                if m:
                    pending_order = m.groupdict()
                    continue

                m = DESC_RE.match(line)
                if m and pending_order:
                    rec = dict(pending_order)
                    rec["cml"] = cur_cml
                    rec["custitem"] = m.group("custitem")
                    rec["desc"] = m.group("desc")
                    records.append(rec)
                    pending_order = None

    return records


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
    """
    cml_groups = OrderedDict()
    for r in records:
        cml_groups.setdefault(r["cml"], []).append(r)

    summary = OrderedDict()

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

    return list(summary.values())


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


def build_excel(all_results):
    """Returns an in-memory Excel file (BytesIO), one sheet per PDF."""
    wb = Workbook()
    wb.remove(wb.active)

    used_names = set()
    header_fill = PatternFill(start_color="305496",
                              end_color="305496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    headers = ["Item No.", "Customer Item No.",
               "Description", "Unit", "PKG (Count)", "Total Qty"]

    for pdf_name, rows in all_results:
        sheet_title = safe_sheet_name(
            os.path.splitext(pdf_name)[0], used_names)
        ws = wb.create_sheet(title=sheet_title)

        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for row in rows:
            ws.append([
                row["itemno"], row["custitem"], row["desc"],
                row["unit"], row["count"], row["total_qty"],
            ])

        for col_idx, header in enumerate(headers, start=1):
            max_len = len(header)
            for row in rows:
                val = [row["itemno"], row["custitem"], row["desc"],
                       row["unit"], str(row["count"]), str(row["total_qty"])][col_idx - 1]
                max_len = max(max_len, len(str(val)))
            ws.column_dimensions[get_column_letter(
                col_idx)].width = max_len + 4

        ws.freeze_panes = "A2"

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
            records = parse_packing_list(temp_path)
            summary = summarize_by_name(records)
            all_results.append((f.filename, summary))
            preview.append({
                "filename": f.filename,
                "items": [
                    {"itemno": r["itemno"], "desc": r["desc"],
                        "pkg": r["count"], "qty": r["total_qty"]}
                    for r in summary
                ],
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
