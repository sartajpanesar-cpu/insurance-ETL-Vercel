import json
import math
import base64
import io
import os
import re
from datetime import datetime
from http.server import BaseHTTPRequestHandler

import openpyxl


# ── Load mapping ──────────────────────────────────────────────────────────────
MAPPING_PATH = os.path.join(os.path.dirname(__file__), "..", "mapping", "field_mapping.json")
with open(MAPPING_PATH) as f:
    MAPPING = json.load(f)

FIELDS = MAPPING["fields"]
RECORD_LENGTH = MAPPING["record_length"]

# ── Hardcoded transform functions ─────────────────────────────────────────────

def transform_zero_pad(value, field):
    v = str(value).strip()
    if not v.isdigit():
        return None, f"{field['source_col']}: '{v}' is not numeric"
    padded = v.zfill(field["encoded_length"])
    if len(padded) > field["encoded_length"]:
        return None, f"{field['source_col']}: '{v}' exceeds {field['encoded_length']} digits"
    return padded, None


def transform_state_recode(value, field):
    key = str(value).strip().lower()
    code = field["recode_map"].get(key)
    if not code:
        return None, f"state_code: '{value}' is not a recognized state name or abbreviation"
    return code, None


DATE_FORMATS = [
    "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%-m/%-d/%Y",
    "%B %d, %Y", "%B %-d, %Y", "%b %d %Y", "%b %d, %Y",
    "%d-%b-%Y", "%m-%d-%Y",
]

def parse_date(value):
    v = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(v, fmt)
        except ValueError:
            continue
    # Try dateutil as fallback
    try:
        from dateutil import parser as du
        return du.parse(v, dayfirst=False)
    except Exception:
        pass
    return None


def transform_date_normalize(value, field):
    dt = parse_date(value)
    if not dt:
        return None, f"{field['source_col']}: '{value}' could not be parsed as a date"
    return dt.strftime("%m%d%y"), None


def transform_recode(value, field):
    key = str(value).strip().lower()
    code = field["recode_map"].get(key)
    if not code:
        return None, f"{field['source_col']}: '{value}' is not a recognized value"
    return code, None


def transform_amount_encode(value, field):
    try:
        amount = float(str(value).replace(",", "").strip())
    except ValueError:
        return None, f"amount: '{value}' is not numeric"

    neg_map = field["negative_encoding"]
    is_negative = amount < 0
    abs_val = abs(amount)
    ceiled = math.ceil(abs_val)
    padded = str(ceiled).zfill(field["encoded_length"])

    if len(padded) > field["encoded_length"]:
        return None, f"amount: {value} exceeds 8-digit capacity after ceiling"

    if is_negative:
        last = padded[-1]
        replacement = neg_map.get(last, last)
        padded = padded[:-1] + replacement

    return padded, None


def transform_uppercase_passthrough(value, field):
    v = str(value).strip().upper()
    if len(v) != field["encoded_length"]:
        return None, f"{field['source_col']}: '{v}' must be exactly {field['encoded_length']} characters (got {len(v)})"
    if not re.match(r'^[A-Z0-9]+$', v):
        return None, f"{field['source_col']}: '{v}' contains invalid characters (alphanumeric only)"
    return v, None


TRANSFORM_MAP = {
    "zero_pad":              transform_zero_pad,
    "state_recode":          transform_state_recode,
    "date_normalize":        transform_date_normalize,
    "recode":                transform_recode,
    "amount_encode":         transform_amount_encode,
    "uppercase_passthrough": transform_uppercase_passthrough,
}


# ── Agent functions ───────────────────────────────────────────────────────────

def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None, "Uploaded file is empty"
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        data.append(dict(zip(headers, row)))
    return headers, data


def agent_quality(headers, data):
    required_cols = [f["source_col"] for f in FIELDS]
    errors = []

    # Check all required columns exist
    missing_cols = [c for c in required_cols if c not in headers]
    if missing_cols:
        errors.append({"type": "missing_column", "severity": "blocking",
                       "message": f"Required columns not found: {', '.join(missing_cols)}"})
        return errors

    for i, row in enumerate(data, 2):  # row 2 = first data row
        for col in required_cols:
            val = row.get(col)
            if val is None or str(val).strip() == "":
                errors.append({"row": i, "field": col, "severity": "blocking",
                               "message": f"Row {i}: '{col}' is required but missing or null"})

    return errors


def agent_transform_and_encode(data):
    encoded_lines = []
    row_errors = []
    successful = 0

    # Pre-parse dates for cross-field validation
    incept_field = next(f for f in FIELDS if f["source_col"] == "inception_date")
    expir_field  = next(f for f in FIELDS if f["source_col"] == "expiration_date")

    for i, row in enumerate(data, 2):
        record_errors = []
        record = {}

        for field in FIELDS:
            val = row.get(field["source_col"])
            if val is None or str(val).strip() == "":
                record_errors.append(f"Row {i}: '{field['source_col']}' is null")
                continue

            transform_fn = TRANSFORM_MAP.get(field["transform"])
            if not transform_fn:
                record_errors.append(f"Row {i}: unknown transform '{field['transform']}'")
                continue

            encoded, err = transform_fn(val, field)
            if err:
                record_errors.append(f"Row {i}: {err}")
            else:
                record[field["source_col"]] = encoded

        # Cross-field: expiry > inception
        if "inception_date" in record and "expiration_date" in record:
            try:
                inc_dt = datetime.strptime(record["inception_date"], "%m%d%y")
                exp_dt = datetime.strptime(record["expiration_date"], "%m%d%y")
                if exp_dt <= inc_dt:
                    record_errors.append(f"Row {i}: expiration_date must be after inception_date")
            except Exception:
                pass

        if record_errors:
            row_errors.extend(record_errors)
            continue

        # Assemble 38-char fixed-width record
        line = ""
        for field in FIELDS:
            line += record.get(field["source_col"], "")

        if len(line) == RECORD_LENGTH:
            encoded_lines.append(line)
            successful += 1
        else:
            row_errors.append(f"Row {i}: assembled record is {len(line)} chars, expected {RECORD_LENGTH}")

    return encoded_lines, row_errors, successful


def agent_validate(encoded_lines):
    validation_errors = []
    for i, line in enumerate(encoded_lines, 1):
        if len(line) != RECORD_LENGTH:
            validation_errors.append(f"Output record {i}: length {len(line)}, expected {RECORD_LENGTH}")
    checksum = f"TRAILER RECCOUNT={len(encoded_lines):06d}"
    return validation_errors, checksum


# ── Main pipeline ─────────────────────────────────────────────────────────────

def run_pipeline(file_bytes, filename):
    result = {
        "filename": filename,
        "status": "error",
        "dq_errors": [],
        "transform_errors": [],
        "validation_errors": [],
        "total_input_rows": 0,
        "successful_records": 0,
        "encoded_output": "",
        "checksum": "",
        "summary": "",
    }

    # Agent 1 — Ingestion
    headers, data = parse_excel(file_bytes)
    if headers is None:
        result["dq_errors"].append({"message": data})
        result["summary"] = "File could not be parsed."
        return result

    result["total_input_rows"] = len(data)

    # Agent 2 — Quality
    dq_errors = agent_quality(headers, data)
    result["dq_errors"] = dq_errors
    blocking = [e for e in dq_errors if e.get("severity") == "blocking" and "row" not in e]
    if blocking:
        result["summary"] = "Pipeline halted: required columns missing."
        return result

    # Agents 3 + 4 — Transform & Encode
    encoded_lines, transform_errors, successful = agent_transform_and_encode(data)
    result["transform_errors"] = transform_errors
    result["successful_records"] = successful

    if not encoded_lines:
        result["summary"] = f"No records encoded. {len(transform_errors)} error(s) found."
        return result

    # Agent 5 — Validation
    validation_errors, checksum = agent_validate(encoded_lines)
    result["validation_errors"] = validation_errors
    result["checksum"] = checksum

    output_lines = encoded_lines + [checksum]
    result["encoded_output"] = "\n".join(output_lines)
    result["status"] = "success" if not validation_errors else "partial"
    result["summary"] = (
        f"{successful} of {len(data)} records encoded successfully. "
        f"{len(transform_errors)} row error(s). "
        f"{len(validation_errors)} validation error(s)."
    )
    return result


# ── Vercel handler ────────────────────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):

    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)

        try:
            payload = json.loads(body)
            file_b64 = payload.get("file")
            filename  = payload.get("filename", "upload.xlsx")

            if not file_b64:
                self._respond(400, {"error": "No file provided"})
                return

            file_bytes = base64.b64decode(file_b64)
            result = run_pipeline(file_bytes, filename)
            self._respond(200, result)

        except Exception as e:
            self._respond(500, {"error": str(e)})

    def do_GET(self):
        self._respond(200, {"status": "Insurance ETL API is running"})

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _respond(self, status, data):
        body = json.dumps(data).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self._cors_headers()
        self.end_headers()
        self.wfile.write(body)
