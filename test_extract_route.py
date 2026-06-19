"""
End-to-end test of the /extract route.

Uploads a merged 3-page PDF (GRN, PR, Invoice) and intercepts the outbound
Nanonets HTTP call to prove that ONLY the invoice page's bytes are sent.
No real network calls are made.
"""

import io
import os
import sys
from unittest import mock

os.environ.setdefault("NANONETS_API_KEY", "dummy")

import fitz  # PyMuPDF

from test_invoice_filter import GRN_TEXT, PR_TEXT, INVOICE_TEXT, make_pdf
import app


def page_texts(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texts = [doc.load_page(i).get_text("text") for i in range(doc.page_count)]
    doc.close()
    return texts


def run():
    # Build the worst-case ordering: invoice is the LAST of three pages.
    merged_pdf = make_pdf([GRN_TEXT, PR_TEXT, INVOICE_TEXT])

    captured = {}

    class FakeResp:
        status_code = 200

        def __init__(self, payload):
            self._payload = payload

        def raise_for_status(self):
            pass

        def json(self):
            return self._payload

    def fake_post(url, headers=None, files=None, data=None, timeout=None):
        # Capture exactly what would be sent to Nanonets.
        # files is a list of ('files', (name, content, ctype)) tuples.
        sent = files[0][1]
        captured["filename"] = sent[0]
        captured["bytes"] = sent[1]
        return FakeResp({"records": [{"record_id": "rec_123"}]})

    # The background polling thread calls requests.get; stub it as "completed"
    # so the thread exits cleanly without a real network call.
    def fake_get(url, headers=None, timeout=None):
        return FakeResp({
            "status": "completed",
            "result": {"json": {"content": {"invoice_id": "INV01",
                                            "line_items": []},
                                "metadata": {"confidence_score": {}}}},
        })

    app.app.config["TESTING"] = True
    client = app.app.test_client()

    with mock.patch.object(app.requests, "post", side_effect=fake_post), \
         mock.patch.object(app.requests, "get", side_effect=fake_get):
        resp = client.post(
            "/extract",
            data={"files": (io.BytesIO(merged_pdf), "merged_invoice.pdf")},
            content_type="multipart/form-data",
        )

    assert resp.status_code == 200, f"unexpected status {resp.status_code}: {resp.data}"
    assert "bytes" in captured, "Nanonets was never called"

    sent_pages = page_texts(captured["bytes"])
    sent_text = "\n".join(sent_pages)

    print(f"Uploaded merged PDF: 3 pages (GRN, PR, Invoice)")
    print(f"Bytes sent to Nanonets: {len(sent_pages)} page(s)")

    problems = []
    if len(sent_pages) != 1:
        problems.append(f"expected 1 page sent, got {len(sent_pages)}")
    if "INV01" not in sent_text:
        problems.append("invoice number INV01 missing from sent page")
    if "PR-2025-114" in sent_text:
        problems.append("PR number leaked into the page sent to Nanonets")
    if "GRN-889" in sent_text:
        problems.append("GRN number leaked into the page sent to Nanonets")
    if "Pepperoni" in sent_text:
        problems.append("PR-only item (Pepperoni) leaked")
    if "Napkins" in sent_text:
        problems.append("GRN-only item (Napkins) leaked")

    if problems:
        print("\nFAILED:")
        for p in problems:
            print("  -", p)
        sys.exit(1)

    print("\n[OK] Only the invoice page reached Nanonets:")
    print("       - contains INV01")
    print("       - no PR/GRN numbers")
    print("       - no PR/GRN-only line items")
    print("\nEND-TO-END TEST PASSED")


if __name__ == "__main__":
    run()
