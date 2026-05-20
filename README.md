# Insurance ETL — Vercel deployment 

LLM-assisted multi-agent ETL pipeline for insurance transaction data. Upload a carrier Excel file; the pipeline runs five agents (ingestion → quality → transform → encode → validate) and returns a 38-character fixed-width output file plus a record-count trailer.

This is the **Vercel-deployable** version of the project. The full local CLI pipeline (with pandas, pytest suite, and modular `agents/`/`functions/` layout) lives in the sibling `insurance_etl/` directory and is unchanged.

---

## What the LLM does

Claude (`claude-sonnet-4-6`) is called at four narrow inflection points. The deterministic encoding logic runs regardless — if `ANTHROPIC_API_KEY` is missing, the pipeline still works, you just lose the LLM-assisted features below.

| Agent | LLM role | Falls back to |
|---|---|---|
| **1. Ingestion** | Maps fuzzy column headers (`Eff Date` → `inception_date`) | Literal header names; missing columns become DQ errors |
| **2. Quality** | Groups raw DQ issues into a concise severity summary | Raw DQ error list |
| **3. Transform** | Suggests closest valid recode key for unknown values (e.g., `Autos` → `Auto`); applied when confidence ≥ 0.85 | Row marked as transform error |
| **5. Validation** | Provides root-cause notes and reprocess vs. manual-review recommendation for any validation failures | Raw failure list |

## Field spec (38-char fixed-width output)

| Pos | Field | Transform |
|---|---|---|
| 1–4 | Company Code | Zero-padded 4-digit numeric |
| 5–6 | State Code | Name or abbreviation → 01–52 |
| 7–12 | Inception Date | Any date format → MMDDYY |
| 13–18 | Expiration Date | Any date format → MMDDYY (must be after inception) |
| 19–20 | Line of Business | Auto→01, Homeowners→02, Dwelling→03 |
| 21–22 | Transaction Type | Premium→01, Paid Loss→05, Outstanding→06 |
| 23–30 | Amount | Ceiling + zero-pad; negatives use letter suffix |
| 31–38 | Policy Code | 8-char alphanumeric, uppercased |

Trailer record: zero-padded 8-digit record count (e.g., `00000025`).

## Stack

- **Frontend:** Vanilla HTML/CSS/JS — drag-and-drop upload, live agent progress, error display, output download, LLM-insight panels
- **Backend:** Python serverless function on Vercel — deterministic transforms in `api/pipeline.py`, anthropic SDK for the four LLM hooks
- **Mapping:** `mapping/field_mapping.json` is the single source of truth for field rules

## Deploy to Vercel

### 1. Push this folder to GitHub

```bash
cd insurance_etl_vercel
git push origin main     # if the remote is already set
```

If the existing remote points to a repo you don't want to overwrite, push to a fresh branch or a new GitHub repo instead:

```bash
git remote set-url origin https://github.com/<you>/<new-repo>.git
git push -u origin main
```

### 2. Import the repo in Vercel

- New Project → Import GitHub repo → leave framework as **Other**
- Root directory: leave at default
- Build & output settings: leave at defaults (Vercel reads `vercel.json`)

### 3. Set the API key as an env var

In the Vercel project: **Settings → Environment Variables**, add:

| Key | Value |
|---|---|
| `ANTHROPIC_API_KEY` | your Anthropic API key (sk-ant-…) |

Apply to Production, Preview, and Development. Redeploy.

## Local dev

```bash
npm i -g vercel
ANTHROPIC_API_KEY=sk-ant-... vercel dev
```

Then open `http://localhost:3000`.

## Health check

```
GET /api/pipeline
```

Returns `{"status": "Insurance ETL API is running", "llm_enabled": true, "model": "claude-sonnet-4-6"}` when configured.

## Files

```
├── index.html                  # Frontend UI
├── api/pipeline.py             # Python ETL pipeline (Vercel serverless)
├── mapping/
│   ├── field_mapping.json      # Runtime field rules (source of truth)
│   └── field_mapping.xlsx      # Editable Excel version
├── vercel.json
└── requirements.txt
```
