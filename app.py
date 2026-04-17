import os
import json
import time
import requests
import io
import traceback
import re
import logging
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file, session
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import PatternFill

# --- Setup Logging ---
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'weibfewb21712897')  # Use env var in production
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25 MB limit

# --- API Configuration ---
NANONETS_API_KEY = os.environ.get('NANONETS_API_KEY')
if not NANONETS_API_KEY:
    raise ValueError("NANONETS_API_KEY environment variable not set")

BATCH_ENDPOINT = "https://extraction-api.nanonets.com/api/v1/extract/batch"
RESULTS_ENDPOINT = "https://extraction-api.nanonets.com/api/v1/extract/results/{}"

EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "supplier_name": {"type": "string"},
        "invoice_id": {"type": "string"},
        "invoice_date": {"type": "string"},
        "payment_term": {"type": "string"},
        "line_items": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "description": {"type": "string"},
                    "quantity": {"type": "number"},
                    "unit_price": {"type": "number"},
                    "discount": {"type": "string"},
                    "tax": {"type": "string"}
                }
            }
        }
    }
}

# ------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------
def standardize_date(date_str):
    """Convert various date formats to DD/MM/YYYY."""
    if not date_str or not isinstance(date_str, str):
        return ''
    date_str = date_str.strip()
    formats = [
        "%d/%b/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y", "%d %b %Y", "%d %B %Y",
        "%b %d, %Y", "%B %d, %Y"
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            if dt.year < 100:
                dt = dt.replace(year=2000 + dt.year if dt.year < 50 else 1900 + dt.year)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    pattern = r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})'
    match = re.search(pattern, date_str)
    if match:
        day, month, year = match.groups()
        try:
            day, month, year = int(day), int(month), int(year)
            if year < 100:
                year = 2000 + year if year < 50 else 1900 + year
            dt = datetime(year, month, day)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            pass
    return date_str

def standardize_payment_term(term_str):
    """Convert detected payment term to NETxx or COD with highlight flag."""
    if not term_str or not isinstance(term_str, str):
        return "COD", True
    term_str = term_str.upper()
    numbers = re.findall(r'\d+', term_str)
    if numbers:
        num = int(numbers[0])
        mapping = {7: "NET7", 14: "NET14", 20: "NET20", 25: "NET25", 30: "NET30", 60: "NET60"}
        if num in mapping:
            return mapping[num], False
    return "COD", True

# ------------------------------------------------------------
# Routes
# ------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract_invoices():
    try:
        if 'files' not in request.files:
            return jsonify({"error": "No files uploaded"}), 400

        files = request.files.getlist('files')
        if not files:
            return jsonify({"error": "Empty file list"}), 400

        multipart_data = {
            'output_format': 'json',
            'json_options': json.dumps(EXTRACTION_SCHEMA),
            'include_metadata': 'confidence_score'
        }

        file_tuples = []
        for f in files:
            if f.filename == '':
                continue
            filename = secure_filename(f.filename)
            file_tuples.append(('files', (filename, f.read(), f.content_type)))

        if not file_tuples:
            return jsonify({"error": "No valid files"}), 400

        headers = {"Authorization": f"Bearer {NANONETS_API_KEY}"}

        batch_resp = requests.post(
            BATCH_ENDPOINT,
            headers=headers,
            files=file_tuples,
            data=multipart_data
        )
        batch_resp.raise_for_status()
        batch_data = batch_resp.json()
        logger.info(f"Batch submitted: {batch_data.get('batch_id')}")

        records = batch_data.get("records", [])
        if not records:
            return jsonify({"error": "No records in batch response"}), 500

        formatted_results = []
        max_poll_attempts = 30
        poll_interval = 2

        for idx, record in enumerate(records):
            record_id = record.get("record_id")
            original_filename = files[idx].filename if idx < len(files) else record.get("filename", f"file_{idx}.pdf")

            if not record_id:
                formatted_results.append({
                    "file": original_filename,
                    "error": "No record_id returned",
                    "extracted_data": None,
                    "confidence": None
                })
                continue

            result_data = None
            for attempt in range(max_poll_attempts):
                try:
                    poll_resp = requests.get(
                        RESULTS_ENDPOINT.format(record_id),
                        headers=headers
                    )
                    poll_resp.raise_for_status()
                    poll_data = poll_resp.json()
                    status = poll_data.get("status")
                    if status == "completed":
                        result_data = poll_data
                        break
                    elif status == "failed":
                        formatted_results.append({
                            "file": original_filename,
                            "error": poll_data.get("message", "Processing failed"),
                            "extracted_data": None,
                            "confidence": None
                        })
                        break
                    else:
                        time.sleep(poll_interval)
                except Exception as e:
                    formatted_results.append({
                        "file": original_filename,
                        "error": f"Polling error: {str(e)}",
                        "extracted_data": None,
                        "confidence": None
                    })
                    break
            else:
                formatted_results.append({
                    "file": original_filename,
                    "error": "Polling timed out",
                    "extracted_data": None,
                    "confidence": None
                })
                continue

            if result_data is None:
                continue

            json_result = result_data.get("result", {}).get("json", {})
            extracted = json_result.get("content", {})
            metadata = json_result.get("metadata", {})
            confidence_scores = metadata.get("confidence_score", {})

            extracted['invoice_date'] = standardize_date(extracted.get('invoice_date'))

            field_confidences = {}
            for field in ["supplier_name", "invoice_id", "invoice_date", "payment_term"]:
                field_confidences[field] = confidence_scores.get(field)

            line_items_conf = []
            line_items = extracted.get("line_items", [])
            line_items_conf_raw = confidence_scores.get("line_items", {})
            ref_keys = list(line_items_conf_raw.keys())
            for i, item in enumerate(line_items):
                item_conf = {}
                if i < len(ref_keys):
                    ref_key = ref_keys[i]
                    item_conf_data = line_items_conf_raw.get(ref_key, {})
                    for subfield in ["description", "quantity", "unit_price", "discount", "tax"]:
                        item_conf[subfield] = item_conf_data.get(subfield)
                else:
                    for subfield in ["description", "quantity", "unit_price", "discount", "tax"]:
                        item_conf[subfield] = None
                line_items_conf.append(item_conf)

            all_confidences = list(field_confidences.values())
            for ic in line_items_conf:
                all_confidences.extend(ic.values())
            valid_confs = [c for c in all_confidences if c is not None]
            avg_confidence = sum(valid_confs) / len(valid_confs) if valid_confs else None

            formatted_results.append({
                "file": original_filename,
                "extracted_data": extracted,
                "field_confidences": field_confidences,
                "line_items_confidences": line_items_conf,
                "average_confidence": avg_confidence
            })

        session['extraction_results'] = json.dumps(formatted_results)
        return jsonify({"results": formatted_results})

    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP error from Nanonets: {e.response.text if e.response else str(e)}")
        return jsonify({"error": f"API error: {str(e)}"}), 500
    except Exception as e:
        logger.error(f"Extraction failed: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route('/export', methods=['POST'])
def export_excel():
    try:
        data = request.get_json(silent=True)  # FIXED: 'silent' not 'silent'
        if data and 'results' in data:
            results = data['results']
        elif 'extraction_results' in session:
            results = json.loads(session['extraction_results'])
        else:
            return jsonify({"error": "No extraction results available"}), 400

        # Use absolute path based on this file's location
        template_path = os.path.join(os.path.dirname(__file__), 'purchase_bills.xlsx')
        if not os.path.exists(template_path):
            logger.error(f"Template file not found at {template_path}")
            return jsonify({"error": "Template file purchase_bills.xlsx not found on server"}), 500

        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        start_row = 31
        current_row = start_row

        yellow_fill_conf = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        red_fill_conf = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        orange_fill_cod = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')

        def get_fill_for_confidence(conf):
            if conf is None:
                return None
            if conf < 90:
                return red_fill_conf
            elif 90 <= conf < 95:
                return yellow_fill_conf
            else:
                return None

        def safe_set_fill(cell, fill):
            if fill is not None:
                try:
                    cell.fill = fill
                except Exception as e:
                    logger.warning(f"Could not set fill on {cell.coordinate}: {e}")

        for invoice in results:
            extracted = invoice.get('extracted_data', {})
            if not extracted:
                continue

            supplier = extracted.get('supplier_name', '')
            invoice_date = extracted.get('invoice_date', '')
            invoice_id = extracted.get('invoice_id', '')
            line_items = extracted.get('line_items', [])

            field_conf = invoice.get('field_confidences', {})
            line_items_conf = invoice.get('line_items_confidences', [])

            raw_payment_term = extracted.get('payment_term', '')
            payment_term, cod_highlight = standardize_payment_term(raw_payment_term)

            first_item = True
            for i, item in enumerate(line_items):
                if first_item:
                    cell_supplier = ws.cell(row=current_row, column=2)
                    cell_supplier.value = supplier
                    safe_set_fill(cell_supplier, get_fill_for_confidence(field_conf.get('supplier_name')))

                    cell_invoice_id = ws.cell(row=current_row, column=3)
                    cell_invoice_id.value = invoice_id
                    safe_set_fill(cell_invoice_id, get_fill_for_confidence(field_conf.get('invoice_id')))

                    cell_date = ws.cell(row=current_row, column=5)
                    cell_date.value = invoice_date
                    safe_set_fill(cell_date, get_fill_for_confidence(field_conf.get('invoice_date')))

                    cell_payment = ws.cell(row=current_row, column=6)
                    cell_payment.value = payment_term
                    if cod_highlight:
                        safe_set_fill(cell_payment, orange_fill_cod)
                    else:
                        safe_set_fill(cell_payment, get_fill_for_confidence(field_conf.get('payment_term')))

                    ws.cell(row=current_row, column=7).value = ''
                    ws.cell(row=current_row, column=8).value = 'MYR'
                    ws.cell(row=current_row, column=10).value = 'Non-Interco, Warehouse Picorp'

                    avg_conf = invoice.get('average_confidence')
                    cell_avg = ws.cell(row=current_row, column=23)  # Column W (23rd)
                    cell_avg.value = avg_conf
                    safe_set_fill(cell_avg, get_fill_for_confidence(avg_conf))

                    first_item = False

                item_conf = line_items_conf[i] if i < len(line_items_conf) else {}

                cell_desc = ws.cell(row=current_row, column=14)
                cell_desc.value = item.get('description', '')
                safe_set_fill(cell_desc, get_fill_for_confidence(item_conf.get('description')))

                qty = item.get('quantity')
                if qty is not None:
                    cell_qty = ws.cell(row=current_row, column=15)
                    cell_qty.value = qty
                    safe_set_fill(cell_qty, get_fill_for_confidence(item_conf.get('quantity')))

                unit_price = item.get('unit_price')
                if unit_price is not None:
                    cell_price = ws.cell(row=current_row, column=18)
                    cell_price.value = unit_price
                    safe_set_fill(cell_price, get_fill_for_confidence(item_conf.get('unit_price')))

                discount = item.get('discount')
                if discount is not None:
                    cell_disc = ws.cell(row=current_row, column=19)
                    cell_disc.value = discount
                    safe_set_fill(cell_disc, get_fill_for_confidence(item_conf.get('discount')))

                tax = item.get('tax')
                if tax is not None:
                    cell_tax = ws.cell(row=current_row, column=20)
                    cell_tax.value = tax
                    safe_set_fill(cell_tax, get_fill_for_confidence(item_conf.get('tax')))

                current_row += 1

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='extracted_invoices.xlsx'
        )

    except Exception as e:
        logger.error(f"Export failed: {e}", exc_info=True)
        return jsonify({"error": f"Export failed: {str(e)}"}), 500

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Unhandled exception: {e}", exc_info=True)
    return jsonify({"error": str(e), "type": type(e).__name__}), 500

# ------------------------------------------------------------
# Main Entry Point
# ------------------------------------------------------------
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)