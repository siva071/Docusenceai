# DocuSense AI — Self-Auditing PDF-to-Excel Agent

**Extract. Verify. Export.**

DocuSense AI is a Streamlit application for PDF intake and content extraction, with the goal of converting digital and scanned PDFs into reviewable Excel workbooks. It extracts selectable text, locally OCRs likely scanned pages, detects tables, stitches likely continuations across adjacent pages, runs conservative data checks, and exports results and available source evidence to an in-memory Excel workbook. Evidence and validation are review aids; they do not prove a value is correct.

## Implemented features

- PDF-only upload and safe checks for empty, unreadable, encrypted, zero-page, and over-50-page documents.
- Page-by-page selectable text extraction using PyMuPDF, with source page numbers and text character counts.
- Low-text pages with substantial embedded image coverage are identified as likely scans. OCR is only run on those candidate pages, not on pages with sufficient selectable text.
- Local Tesseract OCR provides page-aware text, positioned words, OCR engine/language/DPI metadata, progress and per-page errors. Processing continues after an individual OCR page fails.
- OCR word layout can supply a conservative fallback table when pdfplumber cannot detect a table, including wider tables with 10+ aligned columns.
- Table detection/extraction uses pdfplumber, with supplementary PyMuPDF layout word counts.
- Detected tables retain source page, row and column counts, bounding box, and per-row page references before stitching decisions are applied.
- Likely table continuations across consecutive pages are stitched only when headers, column geometry, structure, and page-break placement provide supporting evidence.
- Repeated headers are omitted from merged data; genuine rows retain original order and source-page references. Uncertain pairs remain separate with a warning.
- Fully empty table rows are removed; partially populated rows and original header cell values are preserved.
- A non-mutating validation pass reports suspicious blank cells, duplicate and near-duplicate records, numeric text/format issues, date anomalies, uneven row lengths, possible misalignment, potential total discrepancies, and low-confidence OCR cells.
- Arithmetic checks run only when exactly one total column and at least two recognized component headers (such as `Subtotal` and `Tax`) are present, and all component values are parseable. Blank or ambiguous components remain unknown; they are never assumed to be zero.
- Positive-magnitude and already-negative discounts are both treated as deductions. Arithmetic differences are review warnings rather than confirmed errors because formulas are inferred from headers; the check allows one cent for rounding and records component, actual, and expected values.
- The Validation Report UI summarizes severity and category counts and supports severity filtering. The workbook includes a separate `Validation Report` worksheet, including an explicit no-findings message where applicable.
- Proof Mode lets users select an extracted table, row, and cell to inspect available page, extraction method, original positions, mapped selectable/OCR text, and bounding boxes.
- A `Source Evidence` worksheet records each extracted cell, its value, source page, extraction method, evidence availability, and any source text/word boxes or limitations. Missing provenance is marked unavailable instead of being inferred.
- The UI displays processing progress, page-level warnings, table counts, and table previews.
- Stage 8 shows page-level success, failure, or manual-review status. Failed OCR pages can be selectively retried after the app confirms Tesseract, the Python OCR adapter, and the requested language data are available; successful page text is not re-OCRed.
- Failed OCR/page processing does not silently count as successful. Other pages continue, and unexpected diagnostics are logged without showing raw stack traces to users.
- Invalid PDF, missing OCR runtime/language, empty extraction, and workbook failures provide corrective guidance.
- Download an in-memory `.xlsx` workbook with a separate worksheet per extracted table.
- Excel worksheet names are sanitized and unique; source-page metadata is retained for each row.
- The download filename is `docusense_extracted_data.xlsx`. No workbook is offered when no table was detected or export fails.

## Windows setup

1. Install Python 3.10 or later and make sure the Python launcher (`py`) is available.
2. Open PowerShell in this project folder.
3. Create and activate a virtual environment:

   ```powershell
   py -3 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

   If PowerShell blocks activation, allow scripts for the current terminal session or activate the environment from Command Prompt with `.venv\Scripts\activate.bat`.
4. Install the dependencies:

   ```powershell
   python -m pip install --upgrade pip
   python -m pip install -r requirements.txt
   ```
5. Install the **Tesseract OCR application** (the Python package alone does not include the OCR engine). On Windows, use a maintained Windows Tesseract installer such as the UB Mannheim build. Install English language data (`eng`) and note the executable location, commonly:

   ```text
   C:\Program Files\Tesseract-OCR\tesseract.exe
   ```

   Add `C:\Program Files\Tesseract-OCR` to the Windows `PATH`, then open a new terminal. Alternatively, set the `TESSERACT_CMD` environment variable to the full `tesseract.exe` path before running the app. For additional languages, install the matching trained-data file and set `TESSDATA_PREFIX` to its containing tessdata directory if Tesseract cannot locate it.
6. Start the app:

   ```powershell
   streamlit run app.py
   ```

Streamlit prints a local URL (typically `http://localhost:8501`) to open in your browser. To leave the virtual environment, run `deactivate`.

OCR is optional for digital-only PDFs. If the Tesseract application is not installed or discoverable, digital text and tables still process; scanned-page OCR reports an actionable page error and the rest of the PDF continues.

## Run the tests

Synthetic PDFs are generated in-memory by the tests, including scanned and mixed PDFs and Streamlit upload workflows. OCR responses are mocked in automated tests, so running the test suite does not require the Tesseract executable. From the project folder with the virtual environment active, run:

```powershell
python -m pytest -v
```

Last verified in this workspace: **46 passed, 0 failed, 0 skipped** (Python 3.10.11, pytest 9.1.1). This includes mocked OCR tests; it does not constitute real Tesseract OCR verification.

## Rebuild the challenge sample cases

From the project root and with the project environment active:

```powershell
python samples\build_demo_samples.py
```

The script creates three PDFs and runs each through the actual extraction, OCR-attempt, stitching, validation, evidence, and Excel pipeline. It reopens generated workbooks and checks known extracted values and source-page columns. It writes `samples/outputs/sample_run_manifest.json` with results. When local OCR is unavailable, the scanned PDF remains a failed/manual-review case and the script deliberately does not generate a workbook for invented OCR results.

In this workspace, real processing succeeded for the 10-column digital table and the two-page repeated-header continuation; their `.xlsx` workbooks were generated and checked. The scanned sample was detected as image-based, but OCR failed because Tesseract and its `eng` trained data are not installed here; no scanned workbook was generated.

## Validation example

For a table with headers `Subtotal`, `Tax`, and `Total`, the validator compares rows only when all recognized components are parseable. A row where `100.00 + 5.00` does not equal `110.00` is reported as a potential arithmetic mismatch warning with the actual and expected values, component evidence, table row, and source page. It is never changed automatically. A blank component prevents reconciliation; a blank cell in a column that has values in other rows remains a review warning, while a column that is entirely blank is distinguished as an informational blank field.

## Proof Mode example

For a selectable-text PDF table, Proof Mode may show a detected cell region and the digital words whose bounding boxes fall inside it. For an OCR table, it may show the Tesseract-recognized text, word boxes, and OCR confidence. If the extractor did not provide a source cell region or words, the UI and workbook say evidence is unavailable; table/page association alone is not presented as exact cell-level provenance.

## Project layout

```text
app.py
requirements.txt
README.md
modules/
  __init__.py
  pdf_extractor.py
  ocr_engine.py
  table_stitcher.py
  validator.py
  excel_exporter.py
samples/
   build_demo_samples.py
   DEMO_SCRIPT.md
   digital_10_column.pdf
   scanned_table.pdf
   multipage_continuation.pdf
   outputs/
      digital_10_column.xlsx
      multipage_continuation.xlsx
      sample_run_manifest.json
tests/
   test_app_workflow.py
   test_excel_exporter.py
   test_ocr_engine.py
   test_pdf_extractor.py
   test_table_stitcher.py
   test_validator.py
```

## Current limitations

- OCR needs the separate Tesseract executable and the requested language data installed locally; it is not installed by pip. OCR text and inferred values require manual review, and no accuracy score is claimed.
- Scan detection uses selectable text length and embedded-image coverage. Vector-only scans, tiny image scans, low-quality scans, unusual languages, and pages with sparse digital text over image backgrounds can be misclassified.
- OCR table layout inference is heuristic and relies on repeated aligned text rows and whitespace. Complex, borderless, merged-cell, rotated, or tightly packed tables may be missed or mis-segmented. OCRed values should be reviewed against the source PDF.
- Table detection uses pdfplumber's default layout-based detection, so borderless, irregular, nested, or visually complex tables may be missed or mis-segmented.
- Stitching is conservative and applies only to tables on consecutive pages. Uncertain candidates remain separate with a warning; complex layouts may still require manual review.
- Header identification is heuristic; the original extracted rows, including a likely header row, remain in the returned row data.
- Validation uses conservative header and format heuristics; it can miss issues or report false positives. Date interpretation is ambiguous for some locale-specific forms, and numeric separators may be ambiguous.
- Arithmetic checks are limited to explicit total-like headers and recognized component labels. They do not reconcile ledgers, infer arbitrary formulas, or establish financial correctness.
- Duplicate-like detection is a similarity hint, not proof that records are duplicates. OCR confidence thresholds are review cues, not a measured accuracy estimate.
- Source evidence depends on geometry supplied by pdfplumber, PyMuPDF, or Tesseract. Spatial word-to-cell mapping can be imperfect, and some rows/cells have only page/table association or no exact evidence. Evidence is not independent verification.
- Stitched rows retain their original source-page, source-row, and cell-evidence records. Repeated headers excluded from stitched data do not appear as data rows in the final evidence sheet.
- PDF input checks and validation do not certify extracted values or financial correctness.
- The 50-page check is enforced after upload/open and does not impose a separate upload byte-size limit.
- Advanced financial reconciliation and verification of extracted values are not implemented.
- No external or paid API is used.

## OCR troubleshooting

- **“Tesseract was not found”**: install the Windows Tesseract application, add its folder to `PATH`, restart VS Code/PowerShell, or set `TESSERACT_CMD` to the full executable path.
- **Language data unavailable**: install the requested `.traineddata` file (English uses `eng`) and set `TESSDATA_PREFIX` to the directory containing `tessdata` if needed.
- **Page OCR timed out or returns no text**: check scan resolution/orientation and contrast. The page remains identified for review and other pages continue processing.

## Competition demo

Follow [samples/DEMO_SCRIPT.md](samples/DEMO_SCRIPT.md) for a 3–4 minute walkthrough of digital extraction, stitching, validation, Proof Mode, workbook download, and the honest OCR recovery state. No demo recording or competition submission has been made by this project tooling.

## Recovery behavior

- **Invalid or encrypted PDF**: repair/re-export an unlocked PDF, or split documents above the 50-page limit.
- **Missing OCR runtime/dependencies/language**: install the Tesseract Windows application, verify `pytesseract`/Pillow from `requirements.txt`, install the selected trained-data pack, restart the app, and retry only the failed pages.
- **No tables detected**: review page status and the PDF. Table extraction may not recognize borderless or visually complex layouts; no workbook is offered without extracted tables.
- **Excel export failure**: extracted values and validation/evidence remain available in the session for review; fix the workbook/table input and reprocess to generate a new download.
- **Unexpected internal failure**: the user sees a concise recovery message; detailed exception context is sent to the Python application logger for developer diagnosis.
