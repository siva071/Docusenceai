"""Streamlit interface for DocuSense AI's Stage 8 reliable recovery workflow."""

from __future__ import annotations

import hashlib
import logging
import copy
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from modules.pdf_extractor import (
    PdfInputError,
    PdfMetadata,
    extract_pdf_tables,
    extract_pdf_text,
    extract_ocr_layout_tables,
    inspect_pdf_document,
    retry_failed_ocr_pages,
)
from modules.excel_exporter import ExcelExportError, export_tables_to_excel
from modules.ocr_engine import check_ocr_runtime
from modules.table_stitcher import stitch_tables
from modules.validator import validate_tables

MAX_PDF_PAGES = 50
LOGGER = logging.getLogger(__name__)


def _format_file_size(size_bytes: int) -> str:
    """Format a byte count using a readable binary unit."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024**2:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024**2):.2f} MB"


def _set_page_recovery_status(
    pages: list[dict[str, object]],
    table_page_statuses: list[dict[str, object]],
) -> None:
    """Assign explicit success/failure/manual-review states to every page."""
    table_status_by_page = {
        int(page["page_number"]): page
        for page in table_page_statuses
        if page.get("page_number") is not None
    }
    for page in pages:
        page_number = int(page["page_number"])
        table_status = table_status_by_page.get(page_number, {})
        if page.get("status") == "extraction_failed" or table_status.get("status") == "extraction_failed":
            page["recovery_status"] = "failed"
            page["recovery_reason"] = page.get("error") or table_status.get("error") or "Page extraction failed."
        elif page.get("ocr_status") in {"ocr_failed", "ocr_no_text"}:
            page["recovery_status"] = "failed"
            page["recovery_reason"] = page.get("error") or "OCR did not produce usable text."
        elif page.get("ocr_status") == "ocr_text":
            page["recovery_status"] = "manual_review"
            page["recovery_reason"] = "OCR output should be compared with the scanned source page."
        elif page.get("likely_requires_ocr") or table_status.get("status") == "no_tables_detected":
            page["recovery_status"] = "manual_review"
            page["recovery_reason"] = "No extractable table or selectable text was confirmed for this page."
        else:
            page["recovery_status"] = "succeeded"
            page["recovery_reason"] = "Digital text/table extraction completed without a page-level error."


def _recovery_guidance(message: str) -> str:
    """Return concise corrective instructions based on a user-facing error."""
    lowered = message.casefold()
    if "tesseract was not found" in lowered:
        return "Install Tesseract, add its folder to PATH or set TESSERACT_CMD, then retry the failed OCR pages."
    if "language" in lowered or "trained-data" in lowered or "tessdata" in lowered:
        return "Install the required Tesseract language pack and configure TESSDATA_PREFIX, then retry only the failed pages."
    if "empty" in lowered:
        return "Choose a non-empty PDF file and upload it again."
    if "encrypted" in lowered or "password" in lowered:
        return "Remove PDF password protection and upload the unlocked document."
    if "corrupt" in lowered or "readable pdf" in lowered or "invalid" in lowered:
        return "Re-export or repair the PDF with a PDF reader, then upload the repaired file."
    if "page limit" in lowered or "pages;" in lowered:
        return "Split the document into PDFs of 50 pages or fewer and process each part separately."
    if "excel" in lowered or "workbook" in lowered:
        return "Review the table structure and cell values, then run processing again. The extracted values remain available for review."
    return "Check the PDF and page warnings, correct the file or local OCR setup, and retry processing."


def _ocr_retry_blocked(
    pages: list[dict[str, object]],
    runtime_status: dict[str, object],
) -> tuple[bool, str]:
    """Return whether the page-level OCR retry path is blocked by missing runtime support."""
    for page in pages:
        if page.get("ocr_status") in {"ocr_failed", "ocr_no_text"}:
            error_message = str(page.get("error") or "")
            if error_message:
                return True, _recovery_guidance(error_message)

    if not runtime_status.get("available", True):
        return True, str(runtime_status.get("message") or "OCR retry is unavailable.")

    for page in pages:
        error_message = str(page.get("error") or "")
        if not error_message:
            continue
        lowered = error_message.casefold()
        if (
            "tesseract was not found" in lowered
            or "install it locally" in lowered
            or "ocr dependencies are missing" in lowered
            or "could not be started" in lowered
            or "language" in lowered and "tesseract" in lowered
            or "trained-data" in lowered
            or "tessdata" in lowered
        ):
            return True, _recovery_guidance(error_message)

    return False, ""

st.set_page_config(
    page_title="DocuSense AI | Extract. Verify. Export.",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
            :root { --ink: #142b4a; --muted: #64748b; --blue: #2563eb; --line: #e2e8f0; }
            .stApp { background: #f7f9fc; color: var(--ink); }
            .block-container { max-width: 1180px; padding-top: 1.5rem; padding-bottom: 3rem; }
            html, body, [class*="css"] { font-family: Inter, ui-sans-serif, system-ui, -apple-system, 'Segoe UI', sans-serif; }
            h1, h2, h3 { color: var(--ink); letter-spacing: -.02em; }
            .hero { padding: 1.45rem 1.7rem; margin-bottom: 1.1rem; border: 1px solid var(--line);
                border-radius: 16px; background: #fff; box-shadow: 0 2px 10px rgba(20,43,74,.035); }
            .eyebrow { color: var(--blue); font-size: .74rem; letter-spacing: .12em; font-weight: 700; text-transform: uppercase; }
            .hero h1 { margin: .25rem 0 .15rem; font-size: clamp(1.8rem, 3vw, 2.45rem); }
            .hero p { margin: 0; color: var(--muted); font-size: 1rem; }
            .section-nav { display: flex; flex-wrap: wrap; gap: .5rem; margin: .2rem 0 1.4rem; }
            .section-nav a { color: #24466f; text-decoration: none; background: #fff; border: 1px solid var(--line);
                border-radius: 999px; padding: .42rem .75rem; font-size: .83rem; font-weight: 600; }
            .section-nav a:hover { color: var(--blue); border-color: #93b4f8; }
            .section-kicker { color: var(--blue); text-transform: uppercase; letter-spacing: .1em; font-size: .72rem; font-weight: 700; }
            div[data-testid="stFileUploader"] { background: #fff; border: 1px dashed #9bb5d7; border-radius: 14px; padding: .6rem; }
            div[data-testid="stMetric"] { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: .75rem .9rem; }
      div[data-testid="stMetricLabel"] { color: var(--muted); }
      div[data-testid="stMetricValue"] { color: var(--ink); }
      .muted { color: var(--muted); font-size: .91rem; }
      .status-label { color: var(--muted); font-size: .78rem; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; }
            .stButton > button[kind="primary"] { background: var(--blue); border-color: var(--blue); border-radius: 9px; min-height: 2.7rem; font-weight: 700; }
            .stButton > button[kind="primary"]:hover { background: #1d4ed8; border-color: #1d4ed8; }
            @media (max-width: 640px) { .block-container { padding: .8rem .8rem 2rem; } .hero { padding: 1.1rem; } .section-nav { gap: .35rem; } }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
                <div class="eyebrow">Document intelligence</div>
      <h1>DocuSense AI</h1>
      <p>Extract. Verify. Export.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
        '<nav class="section-nav" aria-label="Workflow sections">'
        '<a href="#pdf-upload">PDF Upload</a><a href="#extraction-results">Extraction Results</a>'
        '<a href="#validation-report">Validation</a><a href="#source-evidence">Source Evidence</a>'
        '<a href="#excel-export">Excel Export</a></nav>',
        unsafe_allow_html=True,
)

st.markdown('<div id="pdf-upload"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-kicker">01 · Upload document</div>', unsafe_allow_html=True)
st.subheader("Bring in a PDF")
st.markdown(
    '<p class="muted">Upload a digital or scanned PDF to validate the file and inspect its page count. '
    'Processing is limited to documents of 50 pages or fewer.</p>',
    unsafe_allow_html=True,
)
uploaded_file = st.file_uploader(
    "Choose a PDF file",
    type=["pdf"],
    accept_multiple_files=False,
    help="PDF files up to 50 pages are accepted. Low-text image pages are processed with local Tesseract OCR.",
)

metadata: PdfMetadata | None = None
validation_error: str | None = None
fingerprint: str | None = None

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()
    fingerprint = hashlib.sha256(pdf_bytes).hexdigest()
    try:
        metadata = inspect_pdf_document(
            pdf_bytes,
            filename=uploaded_file.name,
            max_pages=MAX_PDF_PAGES,
        )
    except PdfInputError as exc:
        validation_error = str(exc)
    except Exception:
        # Keep unexpected parser failures from taking down the Streamlit session.
        validation_error = "The PDF could not be checked due to an unexpected error. Try another file."

    if validation_error:
        st.error(validation_error, icon="⚠️")
        st.info(f"Recovery: {_recovery_guidance(validation_error)}")
    elif metadata is not None:
        st.success("PDF checks passed. The document is ready for extraction.", icon="✅")
        col_name, col_size, col_pages = st.columns([2.2, 1, 1])
        with col_name:
            st.metric("Uploaded filename", metadata.filename)
        with col_size:
            st.metric("File size", _format_file_size(metadata.size_bytes))
        with col_pages:
            st.metric("Page count", f"{metadata.page_count} / {MAX_PDF_PAGES}")

        st.markdown('<div class="section-kicker">02 · Processing status</div>', unsafe_allow_html=True)
        st.subheader("Extract document content")
        st.caption("Selectable text is preferred; local Tesseract OCR is limited to likely scanned pages.")
        start_clicked = st.button("Start Processing", type="primary", use_container_width=True)
        if start_clicked:
            progress_bar = st.progress(0, text="Preparing PDF extraction…")
            page_status = st.empty()
            progress_state = [0]

            def update_progress(page_number: int, total_pages: int, stage: str) -> None:
                stage_offsets = {"text": (0, 15), "ocr": (15, 35), "tables": (50, 30)}
                stage_offset, stage_span = stage_offsets.get(stage, (50, 30))
                percentage = min(95, stage_offset + int(page_number / total_pages * stage_span))
                percentage = max(progress_state[0], percentage)
                progress_state[0] = percentage
                progress_bar.progress(
                    percentage,
                    text=f"{stage.title()} processing · page {page_number} of {total_pages}",
                )
                page_status.caption(
                    f"{stage.title()} processing: page {page_number} of {total_pages}"
                )

            try:
                text_result = extract_pdf_text(
                    pdf_bytes, progress_callback=update_progress, filename=metadata.filename
                )
                table_result = extract_pdf_tables(
                    pdf_bytes,
                    progress_callback=update_progress,
                    filename=metadata.filename,
                    ocr_pages=text_result["pages"],
                )
                stitch_result = stitch_tables(table_result["tables"])
                table_result["source_tables"] = copy.deepcopy(table_result["tables"])
                table_result["retry_ocr_tables"] = []
                table_result["tables"] = stitch_result.tables
                table_result["stitch_warnings"] = stitch_result.warnings
                table_result["stitched_count"] = stitch_result.stitched_count
                _set_page_recovery_status(text_result["pages"], table_result["page_statuses"])
                validation_result = validate_tables(table_result["tables"])
                excel_bytes: bytes | None = None
                excel_error: str | None = None
                progress_bar.progress(90, text="Preparing Excel export…")
                if table_result["tables"]:
                    try:
                        excel_bytes = export_tables_to_excel(
                            table_result["tables"], validation_report=validation_result
                        )
                    except ExcelExportError as exc:
                        LOGGER.warning("Workbook export was rejected: %s", exc)
                        excel_error = str(exc)
                    except Exception:
                        LOGGER.exception("Unexpected workbook export failure")
                        excel_error = "Workbook export failed unexpectedly. Developer diagnostics were logged."
                progress_bar.progress(100, text="Extraction, OCR, stitching, and export preparation finished.")
                st.session_state["extraction_result"] = {
                    "fingerprint": fingerprint,
                    "text": text_result,
                    "tables": table_result,
                    "validation": validation_result,
                    "excel_bytes": excel_bytes,
                    "excel_error": excel_error,
                    "error": None,
                }
            except PdfInputError as exc:
                LOGGER.warning("PDF processing rejected input %s: %s", metadata.filename, exc)
                st.session_state["extraction_result"] = {
                    "fingerprint": fingerprint,
                    "text": None,
                    "tables": None,
                    "validation": None,
                    "excel_bytes": None,
                    "excel_error": None,
                    "error": str(exc),
                }
            except Exception as exc:
                LOGGER.exception("Unhandled PDF workflow failure for %s", metadata.filename)
                st.session_state["extraction_result"] = {
                    "fingerprint": fingerprint,
                    "text": None,
                    "tables": None,
                    "validation": None,
                    "excel_bytes": None,
                    "excel_error": None,
                    "error": "Processing could not be completed because of an unexpected internal error. Developer diagnostics were logged.",
                }

        extraction_result = st.session_state.get("extraction_result")
        if extraction_result and extraction_result.get("fingerprint") == fingerprint:
            if extraction_result["error"]:
                st.error(extraction_result["error"], icon="⚠️")
                st.markdown('<div class="status-label">Status</div><div>Extraction failed; no complete results are available.</div>', unsafe_allow_html=True)
            else:
                text_result = extraction_result["text"]
                table_result = extraction_result["tables"]
                validation_result = extraction_result["validation"]
                failed_ocr_pages = [
                    int(page["page_number"])
                    for page in text_result["pages"]
                    if page.get("ocr_status") in {"ocr_failed", "ocr_no_text"}
                ]
                if failed_ocr_pages:
                    runtime_status = check_ocr_runtime()
                    retry_blocked, retry_message = _ocr_retry_blocked(text_result["pages"], runtime_status)
                    if not retry_blocked:
                        selected_retry_pages = st.multiselect(
                            "Failed OCR pages to retry",
                            failed_ocr_pages,
                            default=failed_ocr_pages,
                            key=f"retry_ocr_selection_{fingerprint}",
                        )
                        if st.button(
                            "Retry selected OCR pages",
                            disabled=not selected_retry_pages,
                            key=f"retry_ocr_button_{fingerprint}",
                        ):
                            with st.spinner("Retrying only the selected OCR pages…"):
                                retry_result = retry_failed_ocr_pages(
                                    pdf_bytes,
                                    text_result["pages"],
                                    page_numbers=selected_retry_pages,
                                )
                                text_result = retry_result["pages"]
                                existing_statuses = {
                                    int(page["page_number"]): page
                                    for page in table_result["page_statuses"]
                                }
                                fallback_candidates = [
                                    page_number
                                    for page_number in retry_result["retried_pages"]
                                    if existing_statuses.get(page_number, {}).get("status")
                                    in {"no_tables_detected", "extraction_failed"}
                                ]
                                fallback_tables = extract_ocr_layout_tables(
                                    text_result,
                                    page_numbers=fallback_candidates,
                                )
                                previous_retry_tables = [
                                    table
                                    for table in table_result.get("retry_ocr_tables", [])
                                    if int(table.get("source_page", -1)) not in selected_retry_pages
                                ]
                                table_result["retry_ocr_tables"] = previous_retry_tables + fallback_tables
                                base_tables = table_result.get("source_tables", table_result["tables"])
                                tables_to_stitch = list(base_tables) + list(table_result["retry_ocr_tables"])
                                if fallback_tables:
                                    for recovered_table in fallback_tables:
                                        page_status = existing_statuses.get(recovered_table["source_page"])
                                        if page_status is not None:
                                            page_status["status"] = "tables_detected_ocr_layout"
                                            page_status["table_count"] = max(1, int(page_status.get("table_count", 0)))
                                            page_status["error"] = None
                                stitch_result = stitch_tables(tables_to_stitch)
                                table_result["tables"] = stitch_result.tables
                                table_result["stitch_warnings"] = stitch_result.warnings
                                table_result["stitched_count"] = stitch_result.stitched_count
                                _set_page_recovery_status(text_result, table_result["page_statuses"])
                                validation_result = validate_tables(table_result["tables"])
                                excel_bytes: bytes | None = None
                                excel_error: str | None = None
                                if table_result["tables"]:
                                    try:
                                        excel_bytes = export_tables_to_excel(
                                            table_result["tables"],
                                            validation_report=validation_result,
                                        )
                                    except ExcelExportError as exc:
                                        LOGGER.warning("Workbook export after OCR retry was rejected: %s", exc)
                                        excel_error = str(exc)
                                    except Exception:
                                        LOGGER.exception("Unexpected workbook export failure after OCR retry")
                                        excel_error = "Workbook export failed unexpectedly. Developer diagnostics were logged."
                                extraction_result.update(
                                    {
                                        "text": text_result,
                                        "tables": table_result,
                                        "validation": validation_result,
                                        "excel_bytes": excel_bytes,
                                        "excel_error": excel_error,
                                    }
                                )
                                st.session_state["extraction_result"] = extraction_result
                                recovered_pages = [
                                    page_number
                                    for page_number in retry_result["retried_pages"]
                                    if next(
                                        page for page in text_result
                                        if int(page["page_number"]) == page_number
                                    ).get("ocr_status") == "ocr_text"
                                ]
                                if recovered_pages:
                                    st.success(
                                        "OCR succeeded on page(s) "
                                        + ", ".join(map(str, recovered_pages))
                                        + ". Existing successful pages were retained. Review the recovered OCR text."
                                    )
                                else:
                                    st.warning(
                                        "The selected pages still need review. Existing successful page results were retained."
                                    )
                                if retry_result["warnings"]:
                                    for warning in retry_result["warnings"]:
                                        st.warning(warning)
                                # The page-state and result widgets below should use the refreshed state.
                                extraction_result = st.session_state["extraction_result"]
                                text_result = extraction_result["text"]
                                table_result = extraction_result["tables"]
                                validation_result = extraction_result["validation"]
                    else:
                        st.button(
                            "Retry selected OCR pages",
                            disabled=True,
                            key=f"retry_ocr_button_{fingerprint}",
                        )
                        st.info(
                            f"OCR retry is unavailable. {retry_message}",
                            icon="ℹ️",
                        )
                processed_pages = len(text_result["pages"])
                detected_tables = len(table_result["tables"])
                severity_counts = validation_result["severity_counts"]
                total_rows = sum(
                    1
                    for table in table_result["tables"]
                    for row in table.get("rows", [])
                    if not isinstance(row, dict) or not row.get("is_inferred_header")
                )
                total_columns = sum(
                    int(table.get("column_count", 0) or 0)
                    for table in table_result["tables"]
                )
                ocr_success_count = sum(page["ocr_status"] == "ocr_text" for page in text_result["pages"])
                ocr_failure_count = sum(
                    page["ocr_status"] in {"ocr_failed", "ocr_no_text"}
                    for page in text_result["pages"]
                )
                summary_top = st.columns(4)
                summary_top[0].metric("Pages processed", f"{processed_pages} / {metadata.page_count}")
                summary_top[1].metric("Tables after stitching", detected_tables)
                summary_top[2].metric("Rows extracted", total_rows)
                summary_top[3].metric("Columns extracted", total_columns)
                summary_bottom = st.columns(4)
                summary_bottom[0].metric(
                    "OCR pages · succeeded / needs review",
                    f"{ocr_success_count} / {ocr_failure_count}",
                )
                summary_bottom[1].metric("Validation errors", severity_counts["error"])
                summary_bottom[2].metric("Validation warnings", severity_counts["warning"])
                summary_bottom[3].metric("Continuations stitched", table_result["stitched_count"])
                page_failures = any(
                    page["status"] == "extraction_failed" for page in text_result["pages"]
                ) or any(
                    page["status"] == "extraction_failed"
                    for page in table_result["page_statuses"]
                )
                if page_failures or table_result["stitch_warnings"] or ocr_failure_count:
                    st.warning(
                        "Processing finished with page-level issues or uncertain table continuations; "
                        "review the warnings below."
                    )
                else:
                    st.markdown(
                        '<div class="status-label">Status</div>'
                        '<div>Extraction run finished; review page warnings and table previews below.</div>',
                        unsafe_allow_html=True,
                    )
                with st.expander("Page recovery status", expanded=bool(page_failures or ocr_failure_count)):
                    st.dataframe(
                        pd.DataFrame(
                            [
                                {
                                    "Page": page["page_number"],
                                    "Status": page.get("recovery_status", "manual_review"),
                                    "OCR": page.get("ocr_status", "not_needed"),
                                    "Reason / next step": page.get("recovery_reason") or page.get("error") or "",
                                }
                                for page in text_result["pages"]
                            ]
                        ),
                        hide_index=True,
                        use_container_width=True,
                    )
        else:
            st.caption("Start processing to extract selectable page text and detect tables.")

st.divider()
st.markdown('<div id="extraction-results"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-kicker">03 · Results workspace</div>', unsafe_allow_html=True)
st.subheader("Extraction Results, Validation & Export")

left, middle, right = st.columns(3, gap="medium")
with left:
    with st.container(border=True):
        st.markdown("### Extracted Tables")
        st.caption("Table detection and source-page traceability")
        extraction_result = st.session_state.get("extraction_result")
        has_current_result = extraction_result and extraction_result.get("fingerprint") == fingerprint
        if has_current_result and not extraction_result["error"]:
            detected_tables = extraction_result["tables"]["tables"]
            if not detected_tables:
                st.warning("No extractable tables were detected in this PDF.", icon="🔎")
                st.info(
                    "Recovery: check page status, confirm the PDF contains a table, and retry failed OCR pages after fixing the local OCR setup.",
                    icon="ℹ️",
                )
            for table in detected_tables:
                source_pages = table.get("source_pages", [table["source_page"]])
                table_rows = table.get("rows", [])
                data_rows = [
                    (row_index, row)
                    for row_index, row in enumerate(table_rows)
                    if not isinstance(row, dict) or not row.get("is_inferred_header")
                ]
                extracted_headers = (
                    list(table.get("headers") or [])
                    if table.get("headers_inferred")
                    else []
                )
                visible_column_count = max(
                    int(table.get("column_count", 0) or 0),
                    len(extracted_headers),
                    max(
                        (
                            len(row.get("cells", []))
                            if isinstance(row, dict)
                            else len(row)
                            for _, row in data_rows
                        ),
                        default=0,
                    ),
                )
                display_headers = extracted_headers + [
                    f"Column {index + 1}"
                    for index in range(len(extracted_headers), visible_column_count)
                ]
                st.markdown(
                    f"**{table['table_id']}** · page(s) {', '.join(map(str, source_pages))} · "
                    f"{len(data_rows)} rows × {visible_column_count} columns"
                )
                if table["headers_inferred"]:
                    st.caption("Likely column headers: " + " · ".join(table["headers"]))
                preview_rows = []
                for row_index, row in data_rows:
                    cells = row.get("cells", []) if isinstance(row, dict) else row
                    source_page = (
                        row.get("source_page", table.get("source_page"))
                        if isinstance(row, dict)
                        else table.get("source_page")
                    )
                    source_row = (
                        row.get("source_row_index", row.get("row_index", row_index) + 1)
                        if isinstance(row, dict)
                        else row_index + 1
                    )
                    preview_rows.append(
                        [source_page, source_row]
                        + list(cells)
                        + [None] * max(0, visible_column_count - len(cells))
                    )
                st.dataframe(
                    pd.DataFrame(
                        preview_rows,
                        columns=["Source page", "Source row", *display_headers],
                    ),
                    hide_index=True,
                    use_container_width=True,
                    height=320,
                )
            if detected_tables:
                st.markdown('<div id="source-evidence"></div>', unsafe_allow_html=True)
                with st.expander("Proof Mode · Source Evidence", expanded=False):
                    st.caption(
                        "Inspect available source context for an extracted cell. Evidence indicates where extraction came from; it does not certify that the value is correct."
                    )
                    table_choice = st.selectbox(
                        "Table",
                        list(range(len(detected_tables))),
                        format_func=lambda index: str(detected_tables[index].get("table_id", f"Table {index + 1}")),
                        key=f"proof_table_{fingerprint}",
                    )
                    selected_table = detected_tables[table_choice]
                    selected_rows = selected_table.get("rows", [])
                    if selected_rows:
                        row_choice = st.selectbox(
                            "Extracted row",
                            list(range(len(selected_rows))),
                            format_func=lambda index: (
                                f"Table row {index + 1} · source page "
                                f"{selected_rows[index].get('source_page', selected_table.get('source_page', 'unknown'))}"
                            ),
                            key=f"proof_row_{fingerprint}_{table_choice}",
                        )
                        selected_row = selected_rows[row_choice]
                        selected_cells = selected_row.get("cells", [])
                        if selected_cells:
                            headers = selected_table.get("headers") or []
                            cell_choice = st.selectbox(
                                "Cell",
                                list(range(len(selected_cells))),
                                format_func=lambda index: (
                                    f"{headers[index] if index < len(headers) else f'Column {index + 1}'} · "
                                    f"{selected_cells[index]!s}"
                                ),
                                key=f"proof_cell_{fingerprint}_{table_choice}_{row_choice}",
                            )
                            cell_records = selected_row.get("cell_evidence", [])
                            evidence = (
                                cell_records[cell_choice]
                                if isinstance(cell_records, list)
                                and cell_choice < len(cell_records)
                                and isinstance(cell_records[cell_choice], dict)
                                else {}
                            )
                            st.markdown(f"**Extracted value:** {selected_cells[cell_choice]!s}")
                            st.markdown(
                                f"**Source page:** {evidence.get('source_page', selected_row.get('source_page', selected_table.get('source_page', 'Unavailable')))}"
                            )
                            st.markdown(
                                f"**Extraction method:** {evidence.get('extraction_method', selected_row.get('extraction_method', selected_table.get('extraction_method', 'unknown')))}"
                            )
                            st.markdown(
                                f"**Source text method:** {evidence.get('source_text_method', 'unavailable')}"
                            )
                            if evidence.get("ocr_confidence") is not None:
                                st.markdown(f"**OCR confidence (engine score):** {evidence['ocr_confidence']}")
                            evidence_status = evidence.get("evidence_status", "unavailable")
                            if evidence_status == "unavailable":
                                st.warning(
                                    "Source cell evidence is unavailable for this value. Its page/table association is not an exact cell-level citation.",
                                    icon="⚠️",
                                )
                            else:
                                st.info(
                                    f"Evidence available: {evidence_status}. This is source context, not independent verification.",
                                    icon="🔎",
                                )
                            st.markdown(
                                f"**Table row / column:** {selected_row.get('source_row_index', selected_row.get('row_index', row_choice) + 1)} / {evidence.get('table_column', cell_choice + 1)}"
                            )
                            if evidence.get("source_row") is not None or evidence.get("source_column") is not None:
                                st.markdown(
                                    f"**Original source row / word column:** {evidence.get('source_row', 'unavailable')} / {evidence.get('source_column', 'unavailable')}"
                                )
                            if evidence.get("bbox"):
                                st.markdown(f"**Cell bounding box:** {evidence['bbox']} ({evidence.get('coordinate_system', 'coordinate system unavailable')})")
                            if evidence.get("source_text"):
                                st.markdown("**Mapped source text:**")
                                st.code(str(evidence["source_text"]), language=None)
                            if evidence.get("word_boxes"):
                                st.json(evidence["word_boxes"])
                            st.caption(
                                str(evidence.get("confidence_or_limitations") or "No additional cell-level evidence was recorded.")
                            )
                        else:
                            st.info("This row has no extracted cells to inspect.")
        elif not has_current_result and uploaded_file is None:
            st.info("Upload a PDF and start processing to see extracted table previews.", icon="📋")
        elif has_current_result and extraction_result["error"]:
            st.warning("No table preview is available because extraction failed.", icon="⚠️")
        else:
            st.info("Extraction has not run yet.", icon="📋")
st.markdown('<div id="validation-report"></div>', unsafe_allow_html=True)
with middle:
    with st.container(border=True):
        st.markdown("### Validation Report")
        st.caption("Data checks and source-linked review findings")
        if metadata is not None:
            st.success(f"PDF input checks passed · {metadata.page_count} pages", icon="✅")
        elif uploaded_file is None:
            st.info("Upload a PDF to see its file validation status.", icon="🔎")
        if metadata is not None:
            extraction_result = st.session_state.get("extraction_result")
            has_current_result = extraction_result and extraction_result.get("fingerprint") == fingerprint
            if has_current_result and not extraction_result["error"]:
                text_result = extraction_result["text"]
                table_result = extraction_result["tables"]
                validation_result = extraction_result["validation"]
                for page in text_result["pages"]:
                    if page["error"]:
                        st.warning(f"Page {page['page_number']}: {page['error']}")
                    elif page["ocr_status"] == "ocr_text":
                        st.caption(f"Page {page['page_number']}: local OCR completed ({page['ocr_metadata']['word_count']} words; review OCR text).")
                    elif page["likely_requires_ocr"] and not page["ocr_attempted"]:
                        st.warning(f"Page {page['page_number']} has no selectable text or image content suitable for OCR; review this page.")
                for page in table_result["page_statuses"]:
                    if page["status"] == "extraction_failed":
                        st.warning(f"Page {page['page_number']}: table extraction failed.")
                    elif page["status"] == "no_tables_detected":
                        st.caption(f"Page {page['page_number']}: no table detected.")
                for warning in table_result["stitch_warnings"]:
                    st.warning(warning, icon="⚠️")
                severity_counts = validation_result["severity_counts"]
                counts = st.columns(3)
                counts[0].metric("Errors", severity_counts["error"])
                counts[1].metric("Warnings", severity_counts["warning"])
                counts[2].metric("Info", severity_counts["info"])
                if not table_result["tables"]:
                    st.info("No tables were detected, so there is no table data to validate.", icon="🔎")
                elif not validation_result["findings"]:
                    st.success(validation_result["summary"], icon="✅")
                else:
                    st.markdown("**Findings by category**")
                    st.caption(
                        " · ".join(
                            f"{category.replace('_', ' ').title()}: {count}"
                            for category, count in validation_result["category_counts"].items()
                        )
                    )
                    severity_filter = st.selectbox(
                        "Filter findings by severity",
                        ["All", "Error", "Warning", "Info"],
                        key="validation_severity_filter",
                    )
                    selected_findings = [
                        finding
                        for finding in validation_result["findings"]
                        if severity_filter == "All"
                        or finding["severity"] == severity_filter.casefold()
                    ]
                    for severity in ("error", "warning", "info"):
                        severity_findings = [
                            finding
                            for finding in selected_findings
                            if finding["severity"] == severity
                        ]
                        if severity_findings:
                            icon = {"error": "🛑", "warning": "⚠️", "info": "ℹ️"}[severity]
                            st.markdown(f"**{severity.title()} ({len(severity_findings)})** {icon}")
                            st.dataframe(
                                pd.DataFrame(
                                    [
                                        {
                                            "Table": finding["table_name"],
                                            "Page": finding["source_page"],
                                            "Row": finding["row"],
                                            "Related row": finding["related_row"],
                                            "Column": finding["column"],
                                            "Category": finding["category"].replace("_", " ").title(),
                                            "Actual value": finding["actual_value"],
                                            "Expected value": finding["expected_value"],
                                            "Explanation": finding["issue"],
                                            "Evidence": finding["evidence"],
                                            "Suggested review": finding["suggested_action"],
                                        }
                                        for finding in severity_findings
                                    ]
                                ),
                                hide_index=True,
                                use_container_width=True,
                            )
                with st.expander("Selectable text by page", expanded=False):
                    for page in text_result["pages"]:
                        st.markdown(
                            f"**Page {page['page_number']}** · {page['character_count']} characters · "
                            f"{page['status']} · method: {page['extraction_method']}"
                        )
                        if page["text"].strip():
                            st.text(page["text"])
                        else:
                            st.caption("No selectable text found on this page.")
            elif has_current_result and extraction_result["error"]:
                st.error("Extraction failed. No table data is presented as verified.")
st.markdown('<div id="excel-export"></div>', unsafe_allow_html=True)
with right:
    with st.container(border=True):
        st.markdown("### Excel Export")
        st.caption("Download the extracted tables as an Excel workbook")
        extraction_result = st.session_state.get("extraction_result")
        has_current_result = extraction_result and extraction_result.get("fingerprint") == fingerprint
        if has_current_result and not extraction_result["error"]:
            if not extraction_result["tables"]["tables"]:
                st.info("No tables were detected, so there is no Excel workbook to download.", icon="📊")
                st.caption("Recovery: review the source PDF and page-level extraction status before retrying.")
            elif extraction_result["excel_error"]:
                st.error(f"Workbook export failed: {extraction_result['excel_error']}", icon="⚠️")
                st.info(f"Recovery: {_recovery_guidance(extraction_result['excel_error'])}")
            elif extraction_result["excel_bytes"]:
                st.download_button(
                    "Download Excel",
                    data=extraction_result["excel_bytes"],
                    file_name="docusense_extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption("One worksheet per detected table; source page is included for each row.")
                with st.expander("Workbook contents", expanded=False):
                    workbook = load_workbook(
                        BytesIO(extraction_result["excel_bytes"]),
                        read_only=True,
                        data_only=True,
                    )
                    for worksheet in workbook.worksheets:
                        if worksheet.title == "Validation Report":
                            description = "Validation findings, severity, source location, and review guidance."
                        elif worksheet.title == "Source Evidence":
                            description = "Extracted cell values and available page, text, OCR, and geometry evidence."
                        else:
                            description = "Extracted table values with source-page metadata."
                        st.markdown(
                            f"**{worksheet.title}** · {max(worksheet.max_row - 1, 0)} rows · "
                            f"{worksheet.max_column} columns — {description}"
                        )
                    workbook.close()
        elif has_current_result and extraction_result["error"]:
            st.warning("Excel export is unavailable because PDF extraction failed.", icon="⚠️")
        else:
            st.info("Process a PDF with at least one detected table to enable Excel export.", icon="📊")

st.markdown(
    '<p class="muted" style="text-align:center; margin-top:1.5rem">'
    'Stage 8 makes page failures recoverable; extraction, validation, and evidence remain review aids, not proof of correctness.</p>',
    unsafe_allow_html=True,
)
