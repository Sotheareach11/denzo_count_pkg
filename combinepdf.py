"""
combinepdf.py
----------------
Combines DENSO-style shipping PDFs (Invoice / Packing List / Freight)
into one PDF per invoice number, ordered INV -> PL -> Freight,
sorted from the smallest invoice number to the largest.

Detects each file's role from its filename:
  - starts with "INV"                          -> Invoice
  - starts with "PL"                            -> Packing List
  - filename is JUST a number (e.g. "651195.pdf")
    OR contains "freight" / starts with "FRT"   -> Freight

If an invoice number has no freight file, it's skipped (no error) -
the combined output just has INV + PL for that invoice.

Can be run standalone:
    python combinepdf.py <input_folder> [output_file]

Or imported:
    from combinepdf import combine
    summary = combine("uploads", "outputs/combined.pdf")
"""

import os
import re
import sys
from pypdf import PdfWriter, PdfReader

TYPE_ORDER = {"INV": 0, "PL": 1, "FREIGHT": 2}
TYPE_LABEL = {"INV": "Invoice", "PL": "Packing List", "FREIGHT": "Freight"}


def classify(filename: str):
    """Return (invoice_number:str, doc_type:str) or (None, None) if unclassifiable."""
    name = os.path.basename(filename)
    stem = os.path.splitext(name)[0]

    # Pure-number filename -> Freight doc, e.g. "651195.pdf"
    if re.fullmatch(r"\d+", stem):
        return stem, "FREIGHT"

    upper = stem.upper()

    if "FREIGHT" in upper or upper.startswith("FRT"):
        m = re.search(r"(\d{5,})", stem)
        if m:
            return m.group(1), "FREIGHT"

    if upper.startswith("INV"):
        m = re.search(r"(\d{5,})", stem)
        if m:
            return m.group(1), "INV"

    if upper.startswith("PL"):
        m = re.search(r"(\d{5,})", stem)
        if m:
            return m.group(1), "PL"

    return None, None


def collect_groups(folder: str):
    groups = {}  # invoice_number -> {"INV": path, "PL": path, "FREIGHT": path}
    unclassified = []
    for entry in os.listdir(folder):
        if not entry.lower().endswith(".pdf"):
            continue
        full_path = os.path.join(folder, entry)
        inv_no, doc_type = classify(entry)
        if inv_no is None:
            unclassified.append(entry)
            continue
        groups.setdefault(inv_no, {})[doc_type] = full_path
    return groups, unclassified


def combine(folder: str, output_file: str):
    """
    Combines all classifiable PDFs in `folder` into `output_file`.
    Returns a dict summary: {
        "output_file": str,
        "invoice_count": int,
        "details": [ "Invoice 651195: [INV] filename.pdf (2 pages)", ... ],
        "unclassified": [ "some_random_file.pdf", ... ]
    }
    Raises ValueError if nothing could be classified.
    """
    groups, unclassified = collect_groups(folder)
    if not groups:
        raise ValueError("No classifiable PDFs found (need INV_*, PL_*, or a plain-number freight file).")

    sorted_invoice_numbers = sorted(groups.keys(), key=lambda x: int(x))

    writer = PdfWriter()
    details = []

    for inv_no in sorted_invoice_numbers:
        docs = groups[inv_no]
        ordered_types = sorted(docs.keys(), key=lambda t: TYPE_ORDER.get(t, 99))
        for doc_type in ordered_types:
            path = docs[doc_type]
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
            details.append(
                f"Invoice {inv_no}: [{TYPE_LABEL[doc_type]}] {os.path.basename(path)} ({len(reader.pages)} pages)"
            )
        missing = [TYPE_LABEL[t] for t in ("INV", "PL", "FREIGHT") if t not in docs]
        if missing:
            details.append(f"Invoice {inv_no}: no {', '.join(missing)} file found - skipped")

    os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)
    with open(output_file, "wb") as f:
        writer.write(f)

    return {
        "output_file": output_file,
        "invoice_count": len(sorted_invoice_numbers),
        "details": details,
        "unclassified": unclassified,
    }


if __name__ == "__main__":
    in_folder = sys.argv[1] if len(sys.argv) > 1 else "."
    out_file = sys.argv[2] if len(sys.argv) > 2 else os.path.join(in_folder, "combined_output.pdf")
    result = combine(in_folder, out_file)
    print(f"Combined {result['invoice_count']} invoice group(s) -> {result['output_file']}\n")
    print("Order used:")
    for line in result["details"]:
        print(f"  {line}")
    if result["unclassified"]:
        print("\nSkipped (could not classify):")
        for f in result["unclassified"]:
            print(f"  {f}")