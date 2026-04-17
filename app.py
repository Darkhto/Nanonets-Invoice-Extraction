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

        # Submit batch job with a timeout
        batch_resp = requests.post(
            BATCH_ENDPOINT,
            headers=headers,
            files=file_tuples,
            data=multipart_data,
            timeout=30  # 30 seconds for submission
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
                        headers=headers,
                        timeout=10  # each poll request times out after 10 seconds
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
                except requests.exceptions.Timeout:
                    logger.warning(f"Polling timeout for record {record_id}, attempt {attempt+1}")
                    continue
                except Exception as e:
                    formatted_results.append({
                        "file": original_filename,
                        "error": f"Polling error: {str(e)}",
                        "extracted_data": None,
                        "confidence": None
                    })
                    break
            else:
                # Loop completed without break -> timeout
                formatted_results.append({
                    "file": original_filename,
                    "error": "Polling timed out after 60 seconds",
                    "extracted_data": None,
                    "confidence": None
                })
                continue

            if result_data is None:
                continue

            # Extract data and confidence
            json_result = result_data.get("result", {}).get("json", {})
            extracted = json_result.get("content", {})
            metadata = json_result.get("metadata", {})
            confidence_scores = metadata.get("confidence_score", {})

            # Standardize date format
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

            # Calculate average confidence
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

    except requests.exceptions.Timeout:
        logger.error("Batch submission or polling timed out")
        return jsonify({"error": "Request timed out. Please try with fewer files or check your connection."}), 504
    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP error from Nanonets: {e.response.text if e.response else str(e)}")
        return jsonify({"error": f"API error: {str(e)}"}), 500
    except Exception as e:
        logger.error(f"Extraction failed: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Unhandled exception: {e}", exc_info=True)
    response = jsonify({"error": "Internal server error", "details": str(e)})
    response.status_code = 500
    return response

# ------------------------------------------------------------
# Main Entry Point
# ------------------------------------------------------------
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)