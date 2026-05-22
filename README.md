# Insurance ETL Pipeline

A multi-agent LLM-assisted ETL pipeline for insurance transaction data. Upload a carrier Excel file, run five specialized agents in sequence, and download a fixed-width encoded output file.

**Live demo:** [https://insuranceetlvercel.vercel.app/]

---

## What it does

| Agent | Role |
|---|---|
| **Agent 1 — Ingestion** | Parses Excel input, detects schema, resolves column name ambiguity |
| **Agent 2 — Quality** | Validates all 8 required fields; halts pipeline on blocking errors |
| **Agent 3 — Transform** | Recodes states, normalizes dates, maps LOB/transaction types |
| **Agent 4 — Encoding** | Assembles 38-character fixed-width records per field spec |
| **Agent 5 — Validation** | Verifies output record lengths and appends checksum trailer |

## Field spec (38-char fixed-width output)

| Pos | Field | Transform |
|---|---|---|
| 1–4 | Company Code | Zero-padded 4-digit numeric |
| 5–6 | State Code | Name or abbreviation → 01–52 |
| 7–12 | Inception Date | Any date format → MMDDYY |
| 13–18 | Expiration Date | Any date format → MMDDYY |
| 19–20 | Line of Business | Auto→01, Homeowners→02, Dwelling→03 |
| 21–22 | Transaction Type | Premium→01, Paid Loss→05, Outstanding→06 |
| 23–30 | Amount | Ceiling + zero-pad; negatives use letter suffix |
| 31–38 | Policy Code | 8-char alphanumeric, uppercased |

## Stack

- **Frontend:** Vanilla HTML/CSS/JS — drag-and-drop file upload, live agent progress, error display, output download
- **Backend:** Python serverless function (Vercel) — deterministic transform functions, all encoding logic hardcoded per spec
- **Mapping:** `field_mapping.json` drives all field rules at runtime — no encoding logic is hardcoded in the agents themselves

## Local development

```bash
npm i -g vercel
vercel dev
```

Then open `http://localhost:3000`.

## Deploy to Vercel

```bash
vercel --prod
```

Set `ANTHROPIC_API_KEY` in your Vercel project environment variables if extending agents with LLM calls.

## Files

```
├── index.html              # Frontend UI
├── api/pipeline.py         # Python ETL pipeline (Vercel serverless)
├── mapping/
│   ├── field_mapping.json  # Runtime field rules (source of truth)
│   └── field_mapping.xlsx  # Editable Excel version of mapping doc
├── vercel.json
└── requirements.txt
```
