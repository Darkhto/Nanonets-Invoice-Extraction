"""
Tests for invoice_filter.filter_invoice_pages.

Builds synthetic 3-page PDFs (Invoice + PR + GRN) in several page orders,
runs them through the filter, and asserts that ONLY the invoice page
survives and that its number/items are the ones extracted.
"""

import fitz  # PyMuPDF

from invoice_filter import (
    classify_page_text,
    filter_invoice_pages,
    score_page_text,
)

# ------------------------------------------------------------------
# Page text fixtures (mirrors a real Malaysian supplier pack)
# ------------------------------------------------------------------
INVOICE_TEXT = """\
ACME FOOD SUPPLIES SDN BHD
TAX INVOICE
Invoice No: INV01
Date: 12/05/2025
Bill To: Mokky's Pizza

Description            Qty    Unit Price
Mozzarella Cheese 5kg    4        45.00
Pizza Flour 25kg         2        80.00
Tomato Paste 3kg         6        15.00

Total Due: MYR 470.00
GST 6%
"""

PR_TEXT = """\
MOKKY'S PIZZA SDN BHD
PURCHASE REQUISITION
PR No: PR-2025-114
Requested By: Store Manager
Reference Invoice No: (to be filled)

Description            Qty
Mozzarella Cheese 5kg    4
Pizza Flour 25kg         2
Pepperoni 2kg            3
"""

GRN_TEXT = """\
MOKKY'S PIZZA SDN BHD
GOODS RECEIVED NOTE
GRN No: GRN-889
Delivery Order No: DO-553
Received By: Warehouse

Description            Qty Received
Mozzarella Cheese 5kg    4
Pizza Flour 25kg         2
Napkins box              1
"""


def make_pdf(pages_text):
    """Create an in-memory PDF (bytes) from a list of page text strings."""
    doc = fitz.open()
    for text in pages_text:
        page = doc.new_page()  # A4
        page.insert_text((72, 72), text, fontsize=11)
    data = doc.tobytes()
    doc.close()
    return data


def page_count(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = doc.page_count
    doc.close()
    return n


def page_texts(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texts = [doc.load_page(i).get_text("text") for i in range(doc.page_count)]
    doc.close()
    return texts


# ------------------------------------------------------------------
# Unit tests: text classification
# ------------------------------------------------------------------
def test_classify_each_document_type():
    assert classify_page_text(INVOICE_TEXT)[0] == "invoice"
    assert classify_page_text(PR_TEXT)[0] == "pr"
    assert classify_page_text(GRN_TEXT)[0] == "grn"


def test_invoice_reference_on_pr_does_not_win():
    # The PR mentions "Invoice No:" as a reference but must still be a PR.
    label, scores = classify_page_text(PR_TEXT)
    assert label == "pr"
    assert scores["pr"] > scores["invoice"]


def test_empty_page_is_unknown():
    assert classify_page_text("")[0] == "unknown"
    assert classify_page_text("   \n  ")[0] == "unknown"


# ------------------------------------------------------------------
# Integration tests: full filtering across every page ordering
# ------------------------------------------------------------------
ORDERINGS = [
    ("GRN, PR, Invoice", [GRN_TEXT, PR_TEXT, INVOICE_TEXT], 3),
    ("Invoice, PR, GRN", [INVOICE_TEXT, PR_TEXT, GRN_TEXT], 1),
    ("PR, GRN, Invoice", [PR_TEXT, GRN_TEXT, INVOICE_TEXT], 3),
    ("PR, Invoice, GRN", [PR_TEXT, INVOICE_TEXT, GRN_TEXT], 2),
    ("GRN, Invoice, PR", [GRN_TEXT, INVOICE_TEXT, PR_TEXT], 2),
]


def test_only_invoice_page_survives_in_every_order():
    for name, pages, expected_page in ORDERINGS:
        pdf = make_pdf(pages)
        out_bytes, info = filter_invoice_pages(pdf, filename="merged.pdf")

        assert info["filtered"] is True, f"[{name}] expected filtering"
        assert info["total_pages"] == 3, f"[{name}] page count"
        assert info["kept_pages"] == [expected_page], (
            f"[{name}] kept {info['kept_pages']}, expected [{expected_page}]"
        )
        assert page_count(out_bytes) == 1, f"[{name}] output should be 1 page"

        # The surviving page must be the invoice: it has INV01 and must NOT
        # carry the PR/GRN numbers or the items unique to those documents.
        text = page_texts(out_bytes)[0]
        assert "INV01" in text, f"[{name}] invoice number missing"
        assert "PR-2025-114" not in text, f"[{name}] PR number leaked"
        assert "GRN-889" not in text, f"[{name}] GRN number leaked"
        assert "Pepperoni" not in text, f"[{name}] PR-only item leaked"
        assert "Napkins" not in text, f"[{name}] GRN-only item leaked"
        print(f"  [OK] {name}: kept page {info['kept_pages']}")


def test_single_page_invoice_passes_through():
    pdf = make_pdf([INVOICE_TEXT])
    out_bytes, info = filter_invoice_pages(pdf, filename="single.pdf")
    assert info["filtered"] is False
    assert info["total_pages"] == 1
    assert page_count(out_bytes) == 1


def test_non_pdf_passes_through_untouched():
    raw = b"\x89PNG fake image bytes"
    out_bytes, info = filter_invoice_pages(raw, filename="photo.png")
    assert out_bytes == raw
    assert info["filtered"] is False
    assert "non-pdf" in info["reason"]


def test_no_invoice_page_keeps_full_document():
    # Two pages, neither is an invoice -> never silently drop data.
    pdf = make_pdf([PR_TEXT, GRN_TEXT])
    out_bytes, info = filter_invoice_pages(pdf, filename="no_invoice.pdf")
    assert info["filtered"] is False
    assert page_count(out_bytes) == 2


if __name__ == "__main__":
    print("Unit: classification")
    test_classify_each_document_type()
    test_invoice_reference_on_pr_does_not_win()
    test_empty_page_is_unknown()
    print("  [OK] classification unit tests")

    print("Integration: page ordering")
    test_only_invoice_page_survives_in_every_order()

    print("Edge cases")
    test_single_page_invoice_passes_through()
    test_non_pdf_passes_through_untouched()
    test_no_invoice_page_keeps_full_document()
    print("  [OK] edge cases")

    print("\nALL TESTS PASSED")
