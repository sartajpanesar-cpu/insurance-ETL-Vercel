"""
Insurance ETL pipeline — Vercel serverless function.

Mirrors the five-agent design from the local CLI project, but in a single
self-contained module suitable for Vercel's Python runtime. Deterministic
transforms run regardless; Claude is called at four narrow inflection points
(header resolution, DQ summarization, fuzzy recode matching, validation root
cause) and is gracefully skipped if ANTHROPIC_API_KEY is not set.
"""
import base64
import io
import json
import logging
import math
import os
import re
from collections import defaultdict
from datetime import datetime
from http.server import BaseHTTPRequestHandler

import openpyxl

try:
    import anthropic
    _ANTHROPIC_IMPORT_OK = True
except ImportError:
    _ANTHROPIC_IMPORT_OK = False

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# ── Config ────────────────────────────────────────────────────────────────────
MODEL = "claude-sonnet-4-6"
RECODE_CONFIDENCE_THRESHOLD = 0.85
LLM_MAX_TOKENS = 1024

MAPPING_PATH = os.path.join(os.path.dirname(__file__), "..", "mapping", "field_mapping.json")
with open(MAPPING_PATH) as f:
    MAPPING = json.load(f)

FIELDS = MAPPING["fields"]
RECORD_LENGTH = MAPPING["record_length"]


# ── LLM helpers ───────────────────────────────────────────────────────────────

_client = None


def llm_available() -> bool:
    return _ANTHROPIC_IMPORT_OK and bool(os.environ.get("ANTHROPIC_API_KEY"))


def _get_client():
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


def _call_llm_json(system_prompt: str, user_content: str, max_tokens: int = LLM_MAX_TOKENS):
    """Call Claude and parse a JSON response. Raises on any failure."""
    msg = _get_client().messages.create(
        model=MODEL,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": user_content}],
    )
    raw = msg.content[0].text.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"```(?:json)?\s*([\s\S]+?)```", raw)
        if m:
            return json.loads(m.group(1))
        raise ValueError(f"LLM response was not valid JSON: {raw[:200]}")


# ── Deterministic transforms ──────────────────────────────────────────────────

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
    if isinstance(value, datetime):
        return value
    v = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(v, fmt)
        except ValueError:
            continue
    try:
        from dateutil import parser as du
        return du.parse(v, dayfirst=False)
    except Exception:
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
        padded = padded[:-1] + neg_map.get(last, last)

    return padded, None


def transform_uppercase_passthrough(value, field):
    v = str(value).strip().upper()
    if len(v) != field["encoded_length"]:
        return None, f"{field['source_col']}: '{v}' must be exactly {field['encoded_length']} characters (got {len(v)})"
    if not re.match(r"^[A-Z0-9]+$", v):
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


# ── Agent 1 — Ingestion ───────────────────────────────────────────────────────

_KNOWN_CANONICAL = {f["source_col"] for f in FIELDS}


def _row_to_headers(row):
    return [str(h).strip() if h is not None else "" for h in row]


def parse_excel(file_bytes):
    """Parse first sheet; skip a leading title row if header row 1 has no canonical matches."""
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None, "Uploaded file is empty"

    header_idx = 0
    candidate = _row_to_headers(rows[0])
    if not (set(c.lower() for c in candidate) & _KNOWN_CANONICAL) and len(rows) > 1:
        # Row 1 looks like a title, not headers — try row 2
        header_idx = 1

    headers = _row_to_headers(rows[header_idx])
    data = []
    for row in rows[header_idx + 1:]:
        if all(v is None for v in row):
            continue
        data.append(dict(zip(headers, row)))
    return headers, data


INGESTION_PROMPT = """You are an insurance ETL ingestion agent.
Your job is to map raw input column headers to their canonical source_col names from the field mapping.
Return ONLY a valid JSON object mapping raw header names to canonical source_col names.
Only include columns that need renaming. If a column name already matches exactly, omit it.
Example: {"Company Code": "company_code", "Eff Date": "inception_date"}"""


def llm_resolve_headers(headers, canonical_cols, sample_row):
    """Ask Claude to map non-matching headers to canonical names. Returns a rename dict."""
    needs_resolution = [h for h in headers if h and h not in canonical_cols]
    if not needs_resolution:
        return {}, None
    if not llm_available():
        return {}, "LLM unavailable (ANTHROPIC_API_KEY not set) — using literal header names"

    schema_sample = {h: (sample_row.get(h) if sample_row else None) for h in needs_resolution}
    user_content = json.dumps({
        "raw_headers": needs_resolution,
        "canonical_source_cols": canonical_cols,
        "schema_sample": {k: str(v) for k, v in schema_sample.items()},
    }, indent=2, default=str)
    try:
        resolved = _call_llm_json(INGESTION_PROMPT, user_content)
        if isinstance(resolved, dict):
            return resolved, None
        return {}, "LLM returned non-dict for header resolution"
    except Exception as e:
        return {}, f"LLM header resolution failed: {e}"


def rename_headers(data, col_map):
    if not col_map:
        return data
    return [{col_map.get(k, k): v for k, v in row.items()} for row in data]


# ── Agent 2 — Quality ─────────────────────────────────────────────────────────

QUALITY_PROMPT = """You are an insurance data quality agent.
Given a structured list of DQ issues, return a JSON array summarizing them by severity.
Each item must have: {"severity": "blocking"|"warning", "field": str, "issue": str, "count": int}
Group similar issues. Be concise. Return ONLY valid JSON."""


def agent_quality(headers, data):
    required_cols = [f["source_col"] for f in FIELDS]
    errors = []

    missing_cols = [c for c in required_cols if c not in headers]
    if missing_cols:
        errors.append({"type": "missing_column", "severity": "blocking",
                       "message": f"Required columns not found: {', '.join(missing_cols)}"})
        return errors

    for i, row in enumerate(data, 2):
        for col in required_cols:
            val = row.get(col)
            if val is None or str(val).strip() == "":
                errors.append({"row": i, "field": col, "severity": "blocking",
                               "message": f"Row {i}: '{col}' is required but missing or null"})
    return errors


def llm_summarize_dq(dq_errors):
    """Ask Claude to group DQ errors by severity. Returns (summary_list, error_or_None)."""
    if not dq_errors or not llm_available():
        return [], None
    try:
        summary = _call_llm_json(QUALITY_PROMPT, json.dumps(dq_errors[:200], indent=2, default=str))
        if isinstance(summary, list):
            return summary, None
        return [], "LLM returned non-list for DQ summary"
    except Exception as e:
        return [], f"LLM DQ summary failed: {e}"


# ── Agent 3 — Transformation & Encoding ───────────────────────────────────────

TRANSFORM_PROMPT = """You are an insurance ETL transformation agent.
You will receive field values that could not be matched to any known recode value.
For each unmatched value, suggest the closest valid recode key and provide a confidence score 0.0–1.0.
Return ONLY a JSON array: [{"field": str, "raw_value": str, "suggested_key": str, "confidence": float}]
Only suggest if you are confident. If no good match exists, set confidence to 0.0."""


def _transform_row(row, row_num):
    """Returns (record_dict, errors, unrecognized_recodes).
    unrecognized_recodes is a dict {source_col: raw_value} for recode/state_recode failures only."""
    record = {}
    errors = []
    unrecognized = {}

    for field in FIELDS:
        val = row.get(field["source_col"])
        if val is None or str(val).strip() == "":
            errors.append(f"Row {row_num}: '{field['source_col']}' is null")
            continue

        transform_fn = TRANSFORM_MAP.get(field["transform"])
        if not transform_fn:
            errors.append(f"Row {row_num}: unknown transform '{field['transform']}'")
            continue

        encoded, err = transform_fn(val, field)
        if err:
            # Stash unrecognized recode values for an LLM second-pass
            if field["transform"] in ("recode", "state_recode"):
                unrecognized[field["source_col"]] = str(val).strip()
            errors.append(f"Row {row_num}: {err}")
        else:
            record[field["source_col"]] = encoded

    # Cross-field rule: expiry > inception
    if "inception_date" in record and "expiration_date" in record:
        try:
            inc_dt = datetime.strptime(record["inception_date"], "%m%d%y")
            exp_dt = datetime.strptime(record["expiration_date"], "%m%d%y")
            if exp_dt <= inc_dt:
                errors.append(f"Row {row_num}: expiration_date must be after inception_date")
        except Exception:
            pass

    return record, errors, unrecognized


def _assemble_line(record):
    return "".join(record.get(f["source_col"], "") for f in FIELDS)


def llm_resolve_recodes(unrecognized_per_field):
    """Ask Claude to fuzzy-match unrecognized recode values to known keys.
    Returns {(source_col, raw_value_lower): encoded_value} and an info note string."""
    if not unrecognized_per_field or not llm_available():
        return {}, None

    entries = []
    field_index = {f["source_col"]: f for f in FIELDS}
    for source_col, raw_vals in unrecognized_per_field.items():
        f = field_index.get(source_col)
        if not f or not f.get("recode_map"):
            continue
        for raw in raw_vals:
            entries.append({
                "field": source_col,
                "raw_value": raw,
                "recode_map_keys": list(f["recode_map"].keys()),
            })
    if not entries:
        return {}, None

    try:
        suggestions = _call_llm_json(TRANSFORM_PROMPT, json.dumps(entries, indent=2))
    except Exception as e:
        return {}, f"LLM recode resolution failed: {e}"

    if not isinstance(suggestions, list):
        return {}, "LLM returned non-list for recode suggestions"

    corrections = {}
    applied = 0
    for s in suggestions:
        try:
            if s.get("confidence", 0) < RECODE_CONFIDENCE_THRESHOLD:
                continue
            f = field_index.get(s["field"])
            if not f or not f.get("recode_map"):
                continue
            key = s["suggested_key"].strip().lower()
            recode_map = {k.lower(): v for k, v in f["recode_map"].items()}
            if key not in recode_map:
                continue
            corrections[(s["field"], s["raw_value"].strip().lower())] = recode_map[key]
            applied += 1
        except Exception:
            continue

    note = f"LLM auto-corrected {applied} unrecognized recode value(s)" if applied else None
    return corrections, note


def agent_transform_and_encode(data):
    encoded_lines = []
    row_errors = []
    successful = 0

    # First pass — collect unrecognized recode values per field
    unrecognized_per_field = defaultdict(set)
    pending = []  # rows that may benefit from a retry after LLM correction
    cleanly_failed = []

    for i, row in enumerate(data, 2):
        record, errors, unrecognized = _transform_row(row, i)
        if not errors:
            line = _assemble_line(record)
            if len(line) == RECORD_LENGTH:
                encoded_lines.append(line)
                successful += 1
            else:
                row_errors.append(f"Row {i}: assembled record is {len(line)} chars, expected {RECORD_LENGTH}")
        elif unrecognized:
            for col, raw in unrecognized.items():
                unrecognized_per_field[col].add(raw)
            pending.append((i, row, errors))
        else:
            cleanly_failed.append((i, errors))

    # LLM second pass — fuzzy-match unrecognized recode values, retry pending rows
    corrections, llm_note = llm_resolve_recodes(unrecognized_per_field)

    if corrections and pending:
        for i, row, original_errors in pending:
            patched_row = dict(row)
            for col in unrecognized_per_field.keys():
                raw = patched_row.get(col)
                if raw is None:
                    continue
                key = (col, str(raw).strip().lower())
                if key in corrections:
                    # Replace the source value with the *encoded* output directly via a sentinel field map
                    patched_row[col] = _ReverseLookupSentinel(corrections[key])

            record, retry_errors, _ = _transform_row_with_sentinels(patched_row, i)
            if not retry_errors:
                line = _assemble_line(record)
                if len(line) == RECORD_LENGTH:
                    encoded_lines.append(line)
                    successful += 1
                    continue
            row_errors.extend(original_errors)
    else:
        for i, _row, original_errors in pending:
            row_errors.extend(original_errors)

    for _i, errs in cleanly_failed:
        row_errors.extend(errs)

    return encoded_lines, row_errors, successful, llm_note


class _ReverseLookupSentinel:
    """Marker for a value that's already been encoded — bypasses transform_recode."""
    def __init__(self, encoded):
        self.encoded = encoded


def _transform_row_with_sentinels(row, row_num):
    """Same as _transform_row but honors _ReverseLookupSentinel for already-resolved fields."""
    record = {}
    errors = []

    for field in FIELDS:
        val = row.get(field["source_col"])
        if isinstance(val, _ReverseLookupSentinel):
            record[field["source_col"]] = val.encoded
            continue
        if val is None or str(val).strip() == "":
            errors.append(f"Row {row_num}: '{field['source_col']}' is null")
            continue

        transform_fn = TRANSFORM_MAP.get(field["transform"])
        if not transform_fn:
            errors.append(f"Row {row_num}: unknown transform '{field['transform']}'")
            continue

        encoded, err = transform_fn(val, field)
        if err:
            errors.append(f"Row {row_num}: {err}")
        else:
            record[field["source_col"]] = encoded

    if "inception_date" in record and "expiration_date" in record:
        try:
            inc_dt = datetime.strptime(record["inception_date"], "%m%d%y")
            exp_dt = datetime.strptime(record["expiration_date"], "%m%d%y")
            if exp_dt <= inc_dt:
                errors.append(f"Row {row_num}: expiration_date must be after inception_date")
        except Exception:
            pass

    return record, errors, {}


# ── Agent 5 — Validation ──────────────────────────────────────────────────────

VALIDATION_PROMPT = """You are an insurance ETL validation agent.
Given a list of validation failures from a fixed-width flat file, provide:
1. Root-cause summary for each failure type
2. Recommendation: "reprocess" (can be fixed by re-running) or "manual_review" (needs human intervention)
Return ONLY a JSON array: [{"issue": str, "root_cause": str, "recommendation": "reprocess"|"manual_review"}]"""


def agent_validate(encoded_lines):
    validation_errors = []
    for i, line in enumerate(encoded_lines, 1):
        if len(line) != RECORD_LENGTH:
            validation_errors.append(f"Output record {i}: length {len(line)}, expected {RECORD_LENGTH}")
    # Trailer = zero-padded record count (per project spec)
    checksum = f"{len(encoded_lines):08d}"
    return validation_errors, checksum


def llm_validation_root_cause(validation_errors):
    if not validation_errors or not llm_available():
        return [], None
    try:
        payload = [{"issue": e, "severity": "blocking"} for e in validation_errors[:50]]
        analysis = _call_llm_json(VALIDATION_PROMPT, json.dumps(payload, indent=2))
        if isinstance(analysis, list):
            return analysis, None
        return [], "LLM returned non-list for validation analysis"
    except Exception as e:
        return [], f"LLM validation analysis failed: {e}"


# ── Main pipeline ─────────────────────────────────────────────────────────────

def run_pipeline(file_bytes, filename):
    result = {
        "filename": filename,
        "status": "error",
        "llm_enabled": llm_available(),
        "llm_notes": [],
        "header_mapping": {},
        "dq_errors": [],
        "dq_llm_summary": [],
        "transform_errors": [],
        "validation_errors": [],
        "validation_llm_analysis": [],
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

    canonical_cols = [f["source_col"] for f in FIELDS]
    sample_row = data[0] if data else {}
    col_map, note = llm_resolve_headers(headers, canonical_cols, sample_row)
    if note:
        result["llm_notes"].append(note)
    if col_map:
        result["header_mapping"] = col_map
        data = rename_headers(data, col_map)
        headers = [col_map.get(h, h) for h in headers]
        result["llm_notes"].append(f"Renamed {len(col_map)} column(s) via LLM")

    # Agent 2 — Quality
    dq_errors = agent_quality(headers, data)
    result["dq_errors"] = dq_errors
    blocking_struct = [e for e in dq_errors if e.get("severity") == "blocking" and "row" not in e]
    if blocking_struct:
        # Try LLM summary even for blocking errors so the user gets a clean overview
        summary, note = llm_summarize_dq(dq_errors)
        if note:
            result["llm_notes"].append(note)
        result["dq_llm_summary"] = summary
        result["summary"] = "Pipeline halted: required columns missing."
        return result

    if dq_errors:
        summary, note = llm_summarize_dq(dq_errors)
        if note:
            result["llm_notes"].append(note)
        result["dq_llm_summary"] = summary

    # Agents 3 + 4 — Transform & Encode
    encoded_lines, transform_errors, successful, llm_recode_note = agent_transform_and_encode(data)
    result["transform_errors"] = transform_errors
    result["successful_records"] = successful
    if llm_recode_note:
        result["llm_notes"].append(llm_recode_note)

    if not encoded_lines:
        result["summary"] = f"No records encoded. {len(transform_errors)} error(s) found."
        return result

    # Agent 5 — Validation
    validation_errors, checksum = agent_validate(encoded_lines)
    result["validation_errors"] = validation_errors
    result["checksum"] = checksum

    if validation_errors:
        analysis, note = llm_validation_root_cause(validation_errors)
        if note:
            result["llm_notes"].append(note)
        result["validation_llm_analysis"] = analysis

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
            filename = payload.get("filename", "upload.xlsx")

            if not file_b64:
                self._respond(400, {"error": "No file provided"})
                return

            file_bytes = base64.b64decode(file_b64)
            result = run_pipeline(file_bytes, filename)
            self._respond(200, result)

        except Exception as e:
            logger.exception("Pipeline error")
            self._respond(500, {"error": str(e)})

    def do_GET(self):
        self._respond(200, {
            "status": "Insurance ETL API is running",
            "llm_enabled": llm_available(),
            "model": MODEL if llm_available() else None,
        })

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _respond(self, status, data):
        body = json.dumps(data, default=str).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self._cors_headers()
        self.end_headers()
        self.wfile.write(body)
