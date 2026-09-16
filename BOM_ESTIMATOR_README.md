# Drawing → BOM & Cost Estimator

A small web app: upload a drawing (PDF, image, or Excel/CSV), Claude reads it
and extracts a bill of materials and cost estimate, you review/edit the
numbers in the browser, then download it as a formatted Excel workbook or
print it directly.

## How it works

1. **Upload** — drag a drawing (door schedule, shop drawing, spec sheet…)
   onto the page.
2. **Analyze** — the backend sends the file to Claude (vision, for PDFs and
   images; extracted text, for spreadsheets) with a prompt that asks for a
   structured JSON bill of materials: project info, a top-level item/BOQ
   list, a detailed parts breakdown, a hardware/accessories schedule, and a
   list of every assumption it had to make (inferred prices, unclear
   dimensions, etc.).
3. **Review & edit** — every field renders as an editable table in the
   browser. Quantities and unit costs recalculate totals live. Rows can be
   added or removed.
4. **Export** — "Download Excel" builds a 4-sheet workbook (Summary, BOQ &
   Quotation, Detailed BOM Breakdown, Hardware & Accessories) with live
   formulas for line/column totals, or "Print" opens a clean, landscape print
   view of the same data.

The commercial grand total is computed only from the top-level BOQ items
(quantity × unit price). The BOM and hardware sheets are supplementary detail
for production planning — their costs are already included inside each
item's unit price, not added on top, matching how the reference workbook
this app was modeled on is structured.

## Project layout

```
backend/            FastAPI app
  app/
    main.py          API routes + static file serving
    claude_client.py Claude API call: builds the file content block, prompt, parses JSON
    file_parsing.py  Extracts text from uploaded spreadsheets
    excel_export.py  Builds the downloadable .xlsx workbook
    schema.py        Shared column/field definitions
  requirements.txt
  .env.example
frontend/            Static HTML/CSS/JS (no build step)
  index.html
  app.js
  styles.css
```

## Running it locally

```bash
cd backend
pip install -r requirements.txt
cp .env.example .env        # then fill in ANTHROPIC_API_KEY
export $(grep -v '^#' .env | xargs)   # or use your preferred env loader
python3 -m uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Open `http://localhost:8000`.

### Configuration

| Variable | Required | Purpose |
|---|---|---|
| `ANTHROPIC_API_KEY` | Yes | Used to call Claude for drawing analysis. Get one at https://console.anthropic.com/ |
| `ANTHROPIC_MODEL` | No | Overrides the model used for extraction (default `claude-sonnet-5`) |

Without `ANTHROPIC_API_KEY` set, the app still serves the UI and the Excel
export/print flow works on manually entered data — only the "Analyze
Drawing" step needs the key.

## Notes on the Excel export

Exported cells use live formulas (`=qty*unitCost`, `=SUM(...)`) rather than
baked-in numbers, so quantities or prices edited after download recalculate
automatically the moment the file is opened in Excel, Google Sheets, or
LibreOffice — all of which recompute formulas on open.
