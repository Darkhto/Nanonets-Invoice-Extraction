"""
invoice_filter.py
-----------------
Pre-processing step that runs BEFORE the Nanonets extraction call.

A user uploads a single merged PDF that contains three different documents:
    - the supplier INVOICE   (the only one we actually want)
    - a Purchase Request / Purchase Order (PR / PO)
    - a Goods Received Note (GRN / Delivery Order)

Because all three contain a "number" and a list of items, Nanonets often
grabs the PR number instead of the invoice number, or merges line items
from all three documents. The fix is to never send the PR/GRN pages to
Nanonets in the first place.

`filter_invoice_pages()` splits the PDF, classifies every page, and rebuilds
a new PDF that contains only the invoice page(s).

Classification is keyword based on the page's text layer (deterministic and
unit-testable). Pages with no usable text layer (scanned images) optionally
fall back to a Gemini vision call.
"""

import io
import logging

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)

# ------------------------------------------------------------------
# Keyword tables
# ------------------------------------------------------------------
# Each tuple is (keyword, weight). Title-style phrases get a high weight
# because they are the strongest signal for what a page actually *is*.
# A cross-reference such as "Invoice No:" printed on a GRN gets a low
# weight, so the dominant document type still wins the argmax.

INVOICE_KEYWORDS = [
    ("tax invoice", 6),
    ("invois cukai", 6),        # Malay
    ("invoice", 4),
    ("invois", 4),
    ("inbois", 4),
    ("bill to", 2),
    ("amount due", 2),
    ("total due", 2),
    ("balance due", 2),
    ("gst", 1),
    ("sst", 1),
]

PR_KEYWORDS = [
    ("purchase requisition", 6),
    ("purchase request", 6),
    ("purchase order", 6),
    ("requisition", 4),
    ("pesanan belian", 6),      # Malay: purchase order
    ("permohonan belian", 6),   # Malay: purchase request
    ("p.o. no", 4),
    ("po no", 4),
    ("po number", 4),
    ("pr no", 4),
    ("pr number", 4),
    ("requested by", 2),
]

GRN_KEYWORDS = [
    ("goods received note", 6),
    ("goods receipt note", 6),
    ("goods receipt", 5),
    ("goods received", 5),
    ("delivery order", 6),
    ("delivery note", 6),
    ("nota terima barang", 6),  # Malay: GRN
    ("nota penghantaran", 6),   # Malay: delivery note
    ("grn", 5),
    ("d.o. no", 4),
    ("do no", 4),
    ("received by", 2),
    ("delivered by", 2),
]

CATEGORY_KEYWORDS = {
    "invoice": INVOICE_KEYWORDS,
    "pr": PR_KEYWORDS,
    "grn": GRN_KEYWORDS,
}


def score_page_text(text):
    """Return a {category: score} dict for a page's text.

    Matching is case-insensitive on a lowercased copy of the text. Each
    keyword contributes ``weight * occurrences`` to its category.
    """
    lowered = (text or "").lower()
    scores = {category: 0 for category in CATEGORY_KEYWORDS}
    for category, keywords in CATEGORY_KEYWORDS.items():
        for keyword, weight in keywords:
            count = lowered.count(keyword)
            if count:
                scores[category] += weight * count
    return scores


def classify_page_text(text):
    """Classify a page from its text.

    Returns (label, scores) where label is one of
    'invoice' | 'pr' | 'grn' | 'unknown'. 'unknown' means there was no
    usable signal (e.g. an empty / scanned page) and the caller should
    decide whether to use a vision fallback.
    """
    scores = score_page_text(text)
    best_category = max(scores, key=scores.get)
    if scores[best_category] == 0:
        return "unknown", scores
    return best_category, scores


# ------------------------------------------------------------------
# Optional Gemini vision fallback for scanned pages (no text layer)
# ------------------------------------------------------------------
def _classify_page_image_gemini(page, genai_client, model="gemini-2.5-flash"):
    """Classify a single rendered page with Gemini vision.

    Returns 'invoice' | 'pr' | 'grn' | 'unknown'. Never raises – any
    failure returns 'unknown' so the caller can fall back safely.
    """
    if genai_client is None:
        return "unknown"
    try:
        from google.genai import types

        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        prompt = (
            "This is one page from a scanned business document. Classify it as "
            "exactly one of: INVOICE (a supplier's sales/tax invoice), "
            "PR (a purchase request / purchase order / requisition), or "
            "GRN (a goods received note / delivery order). "
            "Answer with only one word: INVOICE, PR, or GRN."
        )
        response = genai_client.models.generate_content(
            model=model,
            contents=[
                types.Part.from_bytes(data=img_bytes, mime_type="image/png"),
                prompt,
            ],
        )
        answer = (response.text or "").strip().lower()
        if "invoice" in answer:
            return "invoice"
        if "grn" in answer:
            return "grn"
        if "pr" in answer or "purchase" in answer:
            return "pr"
        return "unknown"
    except Exception as e:  # pragma: no cover - network/SDK dependent
        logger.warning(f"Gemini page classification failed: {e}")
        return "unknown"


# ------------------------------------------------------------------
# Public entry point
# ------------------------------------------------------------------
def filter_invoice_pages(file_bytes, filename="document.pdf", genai_client=None):
    """Return PDF bytes containing only the invoice page(s).

    Parameters
    ----------
    file_bytes : bytes
        Raw bytes of the uploaded file.
    filename : str
        Original filename (used only to detect non-PDF uploads).
    genai_client : optional
        A configured ``google.genai`` client. When provided it is used to
        classify scanned pages that have no extractable text layer.

    Returns
    -------
    (out_bytes, info) : (bytes, dict)
        ``out_bytes`` is a new PDF containing only invoice pages (or the
        original bytes when filtering is not applicable / not safe).
        ``info`` describes what happened, useful for logging / debugging:
            {
              "filtered": bool,        # True if we actually dropped pages
              "total_pages": int,
              "kept_pages": [int, ...],# 1-based page numbers kept
              "labels": [str, ...],    # per-page label
              "reason": str,           # human readable explanation
            }
    """
    info = {
        "filtered": False,
        "total_pages": 0,
        "kept_pages": [],
        "labels": [],
        "reason": "",
    }

    # Non-PDF uploads (single images) cannot be page-split – pass through.
    if not filename.lower().endswith(".pdf"):
        info["reason"] = "non-pdf upload, passed through unchanged"
        return file_bytes, info

    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        logger.warning(f"Could not open PDF '{filename}': {e}; passing through")
        info["reason"] = f"could not open pdf: {e}"
        return file_bytes, info

    total_pages = doc.page_count
    info["total_pages"] = total_pages

    # A single page document has nothing to filter.
    if total_pages <= 1:
        doc.close()
        info["reason"] = "single page, nothing to filter"
        info["labels"] = ["invoice"]
        info["kept_pages"] = [1] if total_pages == 1 else []
        return file_bytes, info

    labels = []
    for page_index in range(total_pages):
        page = doc.load_page(page_index)
        text = page.get_text("text")
        label, scores = classify_page_text(text)

        # Scanned page with no text layer -> try the vision fallback.
        if label == "unknown" and genai_client is not None:
            label = _classify_page_image_gemini(page, genai_client)

        labels.append(label)

    info["labels"] = labels
    invoice_pages = [i for i, lbl in enumerate(labels) if lbl == "invoice"]

    # Fallback: no page is clearly an invoice (e.g. all scanned with no text
    # and no vision client, or a pack that genuinely has no invoice). Keep the
    # whole document so we never silently drop the invoice. We deliberately do
    # NOT guess a "best invoice-scoring page" here, because a stray
    # "Reference Invoice No:" field on a PR/GRN would fool that guess.
    if not invoice_pages:
        doc.close()
        info["reason"] = (
            "could not identify an invoice page; sent full document unchanged"
        )
        return file_bytes, info

    info["kept_pages"] = [i + 1 for i in invoice_pages]

    # If every page is an invoice there is nothing to drop.
    if len(invoice_pages) == total_pages:
        doc.close()
        info["reason"] = "all pages look like invoice; sent unchanged"
        return file_bytes, info

    # Build a new PDF with only the invoice pages.
    out_doc = fitz.open()
    for i in invoice_pages:
        out_doc.insert_pdf(doc, from_page=i, to_page=i)
    out_bytes = out_doc.tobytes()
    out_doc.close()
    doc.close()

    info["filtered"] = True
    if not info["reason"]:
        info["reason"] = (
            f"kept invoice page(s) {info['kept_pages']} of {total_pages}"
        )
    return out_bytes, info
