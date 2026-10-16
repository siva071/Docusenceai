"""In-memory PDF validation, text extraction, and table extraction."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import copy
from io import BytesIO
import logging
from pathlib import Path
from typing import Callable

import pdfplumber

try:  # PyMuPDF exposed the recommended module name after earlier releases.
    import pymupdf as fitz
except ImportError:  # Keep compatibility with the declared PyMuPDF 1.23 minimum.
    import fitz

from modules.ocr_engine import _image_coverage, extract_text_with_ocr, is_likely_scanned_page

MAX_PDF_PAGES = 50
ProgressCallback = Callable[[int, int, str], None]
LOGGER = logging.getLogger(__name__)


class PdfInputError(ValueError):
    """Raised when an uploaded PDF cannot safely enter the processing flow."""


@dataclass(frozen=True)
class PdfMetadata:
    """Basic metadata established during PDF intake."""

    filename: str
    size_bytes: int
    page_count: int


def _open_pdf(pdf_bytes: bytes, max_pages: int = MAX_PDF_PAGES) -> fitz.Document:
    """Open and validate PDF bytes, closing invalid documents before raising."""
    if not isinstance(pdf_bytes, bytes) or not pdf_bytes:
        raise PdfInputError("The uploaded file is empty. Choose a non-empty PDF and try again.")
    if max_pages < 1:
        raise ValueError("max_pages must be at least 1")
    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
        if document.needs_pass:
            document.close()
            raise PdfInputError(
                "This PDF is encrypted or password-protected. Remove its protection and upload it again."
            )
        if document.page_count < 1:
            document.close()
            raise PdfInputError("This PDF contains no pages and cannot be processed.")
        if document.page_count > max_pages:
            page_count = document.page_count
            document.close()
            raise PdfInputError(
                f"This PDF has {page_count} pages. The limit is {max_pages} pages; upload a shorter PDF."
            )
        return document
    except PdfInputError:
        raise
    except Exception as exc:
        LOGGER.exception("Failed to open or validate uploaded PDF bytes")
        raise PdfInputError(
            "The uploaded file is not a readable PDF or is corrupted. Check the file and try again."
        ) from exc


def inspect_pdf_document(
    pdf_bytes: bytes,
    filename: str,
    max_pages: int = MAX_PDF_PAGES,
) -> PdfMetadata:
    """Validate a PDF and return its filename, byte size, and page count."""
    if not filename or Path(filename).suffix.lower() != ".pdf":
        raise PdfInputError("Only files with a .pdf extension are accepted.")
    document = _open_pdf(pdf_bytes, max_pages=max_pages)
    try:
        page_count = document.page_count
    finally:
        document.close()
    return PdfMetadata(Path(filename).name, len(pdf_bytes), page_count)


def get_pdf_metadata(
    pdf_bytes: bytes,
    filename: str = "upload.pdf",
    max_pages: int = MAX_PDF_PAGES,
) -> PdfMetadata:
    """Return validated PDF metadata without writing the upload to disk."""
    return inspect_pdf_document(pdf_bytes, filename, max_pages)


def extract_pdf_text(
    pdf_bytes: bytes,
    progress_callback: ProgressCallback | None = None,
    max_pages: int = MAX_PDF_PAGES,
    filename: str = "upload.pdf",
) -> dict[str, object]:
    """Extract page text and identify pages that likely require OCR.

    Pages are numbered from 1. Only low-text pages with substantial embedded
    raster content are sent to local Tesseract OCR. Per-page OCR errors are
    reported without aborting extraction on other pages.
    """
    document = _open_pdf(pdf_bytes, max_pages=max_pages)
    pages: list[dict[str, object]] = []
    warnings: list[str] = []
    ocr_candidates: list[int] = []
    try:
        for index in range(document.page_count):
            page_number = index + 1
            try:
                pdf_page = document.load_page(index)
                page_text = pdf_page.get_text("text") or ""
                digital_words = _digital_word_records(pdf_page, page_number)
                has_digital_text = bool(page_text.strip())
                image_coverage = _image_coverage(pdf_page)
                is_scanned = is_likely_scanned_page(page_text, image_coverage)
                page_result: dict[str, object] = {
                    "page_number": page_number,
                    "text": page_text,
                    "selectable_text": page_text,
                    "digital_words": digital_words,
                    "digital_evidence": {
                        "source_page": page_number,
                        "extraction_method": "digital_text",
                        "word_count": len(digital_words),
                        "coordinate_system": "PDF points, top-left origin",
                        "availability": "available" if digital_words else "text_only_or_unavailable",
                    },
                    "character_count": len(page_text.strip()),
                    "has_digital_text": has_digital_text,
                    "likely_requires_ocr": is_scanned or not has_digital_text,
                    "is_likely_scanned": is_scanned,
                    "image_coverage": round(image_coverage, 3),
                    "extraction_method": "digital" if has_digital_text else "none",
                    "ocr_attempted": False,
                    "ocr_status": "not_needed",
                    "ocr_words": [],
                    "ocr_metadata": None,
                    "ocr_page_width": None,
                    "ocr_page_height": None,
                    "status": "digital_text" if has_digital_text else "no_selectable_text",
                    "error": None,
                }
                if is_scanned:
                    page_result["status"] = "ocr_pending"
                    page_result["ocr_status"] = "pending"
                    ocr_candidates.append(page_number)
                elif not has_digital_text:
                    warnings.append(
                        f"Page {page_number} has no selectable text or substantial embedded images; review it."
                    )
            except Exception as exc:
                LOGGER.exception("Digital text extraction failed on PDF page %s", page_number)
                page_result = {
                    "page_number": page_number,
                    "text": "",
                    "selectable_text": "",
                    "digital_words": [],
                    "digital_evidence": {
                        "source_page": page_number,
                        "extraction_method": "digital_text",
                        "word_count": 0,
                        "coordinate_system": "PDF points, top-left origin",
                        "availability": "unavailable",
                    },
                    "character_count": 0,
                    "has_digital_text": False,
                    "likely_requires_ocr": True,
                    "is_likely_scanned": False,
                    "image_coverage": 0.0,
                    "extraction_method": "none",
                    "ocr_attempted": False,
                    "ocr_status": "not_attempted",
                    "ocr_words": [],
                    "ocr_metadata": None,
                    "ocr_page_width": None,
                    "ocr_page_height": None,
                    "status": "extraction_failed",
                    "error": "Digital text extraction failed on this page. Review the PDF or retry OCR if this page is a scan.",
                }
                warnings.append(f"Page {page_number}: text extraction failed; manual review is required.")
            pages.append(page_result)
            if progress_callback is not None:
                progress_callback(page_number, document.page_count, "text")
        result_metadata = PdfMetadata(Path(filename).name, len(pdf_bytes), document.page_count)
    finally:
        document.close()

    ocr_result = extract_text_with_ocr(
        pdf_bytes,
        page_numbers=ocr_candidates,
        progress_callback=progress_callback,
    )
    if ocr_result.get("fatal_error"):
        LOGGER.error("OCR engine could not initialize for candidate pages: %s", ocr_result["fatal_error"])
        ocr_result["pages"] = [
            {
                "page_number": page_number,
                "text": "",
                "words": [],
                "word_count": 0,
                "status": "ocr_failed",
                "error": "OCR could not start for this PDF. Check the file and local OCR setup, then retry.",
                "engine": "Tesseract",
                "language": "eng",
                "dpi": None,
                "page_width": None,
                "page_height": None,
            }
            for page_number in ocr_candidates
        ]
    _merge_ocr_page_results(pages, ocr_result["pages"], warnings)
    warnings.extend(ocr_result["warnings"])
    return {
        "metadata": asdict(result_metadata),
        "pages": pages,
        "warnings": list(dict.fromkeys(warnings)),
        "ocr_page_count": len(ocr_candidates),
    }


def _merge_ocr_page_results(
    page_results: list[dict[str, object]],
    ocr_pages: list[dict[str, object]],
    warnings: list[str],
) -> None:
    """Merge OCR outcomes into existing page metadata without replacing good text."""
    ocr_by_page = {int(page["page_number"]): page for page in ocr_pages}
    for page in page_results:
        ocr_page = ocr_by_page.get(int(page["page_number"]))
        if ocr_page is None:
            continue
        page["ocr_attempted"] = True
        page["ocr_status"] = ocr_page["status"]
        page["ocr_words"] = ocr_page.get("words", [])
        page["ocr_metadata"] = {
            key: ocr_page.get(key)
            for key in ("engine", "language", "dpi", "word_count", "page_width", "page_height")
        }
        page["ocr_page_width"] = ocr_page.get("page_width")
        page["ocr_page_height"] = ocr_page.get("page_height")
        if ocr_page["status"] == "ocr_text":
            digital_lines = str(page.get("selectable_text", "")).splitlines()
            existing_keys = {" ".join(line.casefold().split()) for line in digital_lines if line.strip()}
            additions = [
                line
                for line in str(ocr_page.get("text", "")).splitlines()
                if line.strip() and " ".join(line.casefold().split()) not in existing_keys
            ]
            page["text"] = "\n".join(digital_lines + additions).strip()
            page["character_count"] = len(str(page["text"]))
            page["has_ocr_text"] = True
            page["extraction_method"] = "digital+ocr" if page.get("has_digital_text") else "ocr"
            page["status"] = "ocr_text"
            page["error"] = None
        else:
            page["has_ocr_text"] = False
            page["status"] = ocr_page.get("status", "ocr_failed")
            page["error"] = ocr_page.get("error") or "OCR did not produce usable text."
            warnings.append(f"Page {page['page_number']}: {page['error']}")


def retry_failed_ocr_pages(
    pdf_bytes: bytes,
    existing_pages: list[dict[str, object]],
    page_numbers: list[int] | None = None,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, object]:
    """Retry only previously failed OCR pages while retaining other page results.

    Successful pages are deep-copied unchanged and are never sent back to
    Tesseract. The return includes the actual attempted page numbers so callers
    can report which recoveries were tried.
    """
    eligible = {
        int(page["page_number"])
        for page in existing_pages
        if page.get("ocr_status") in {"ocr_failed", "ocr_no_text"}
    }
    requested = sorted(eligible if page_numbers is None else eligible.intersection(page_numbers))
    updated_pages = copy.deepcopy(existing_pages)
    if not requested:
        return {
            "pages": updated_pages,
            "retried_pages": [],
            "warnings": [],
            "message": "No failed OCR pages were selected for retry.",
        }

    ocr_result = extract_text_with_ocr(
        pdf_bytes,
        page_numbers=requested,
        progress_callback=progress_callback,
    )
    retry_warnings: list[str] = []
    _merge_ocr_page_results(updated_pages, ocr_result.get("pages", []), retry_warnings)
    retry_warnings.extend(ocr_result.get("warnings", []))
    return {
        "pages": updated_pages,
        "retried_pages": requested,
        "warnings": list(dict.fromkeys(retry_warnings)),
        "message": "OCR retry completed for selected failed pages.",
    }


def _clean_cell(cell: object) -> str:
    """Normalize a pdfplumber cell's presentation whitespace."""
    if cell is None:
        return ""
    return str(cell).replace("\r", "").strip()


def _digital_word_records(page: fitz.Page, page_number: int) -> list[dict[str, object]]:
    """Return selectable words with source page and PDF-point bounding boxes."""
    records: list[dict[str, object]] = []
    try:
        for word in page.get_text("words"):
            if len(word) < 5 or not str(word[4]).strip():
                continue
            records.append(
                {
                    "text": str(word[4]),
                    "page_number": page_number,
                    "bbox": [float(value) for value in word[:4]],
                    "block_number": int(word[5]) if len(word) > 5 else None,
                    "line_number": int(word[6]) if len(word) > 6 else None,
                    "word_number": int(word[7]) if len(word) > 7 else None,
                }
            )
    except Exception:
        return []
    return records


def _cell_evidence(
    cell_bbox: object,
    words: list[dict[str, object]],
    page_number: int,
    source_row: int,
    source_column: int,
    extraction_method: str,
) -> dict[str, object]:
    """Describe cell geometry and mapped text without asserting verification."""
    evidence: dict[str, object] = {
        "source_page": page_number,
        "source_row": source_row,
        "source_column": source_column,
        "extraction_method": extraction_method,
        "source_text_method": "digital_text" if words else "unavailable",
        "table_column": source_column,
        "coordinate_system": "PDF points, top-left origin",
        "bbox": None,
        "source_text": None,
        "word_boxes": [],
        "evidence_status": "unavailable",
        "confidence_or_limitations": "Source cell location was not provided by the extractor; this value is not independently verified.",
    }
    if not isinstance(cell_bbox, (list, tuple)) or len(cell_bbox) != 4 or any(value is None for value in cell_bbox):
        return evidence
    try:
        x0, top, x1, bottom = (float(value) for value in cell_bbox)
    except (TypeError, ValueError):
        return evidence
    evidence["bbox"] = [x0, top, x1, bottom]
    contained_words: list[dict[str, object]] = []
    for word in words:
        word_bbox = word.get("bbox")
        if not isinstance(word_bbox, (list, tuple)) or len(word_bbox) != 4:
            continue
        wx0, wy0, wx1, wy1 = (float(value) for value in word_bbox)
        center_x, center_y = (wx0 + wx1) / 2, (wy0 + wy1) / 2
        if x0 <= center_x <= x1 and top <= center_y <= bottom:
            contained_words.append({"text": word.get("text", ""), "bbox": [wx0, wy0, wx1, wy1]})
    evidence["source_text"] = " ".join(str(word["text"]) for word in contained_words) or None
    evidence["word_boxes"] = contained_words
    evidence["source_text_method"] = "digital_text" if contained_words else "unavailable"
    evidence["evidence_status"] = "source_text_mapped" if contained_words else "cell_region_only"
    evidence["confidence_or_limitations"] = (
        "Selectable words were spatially mapped into the detected table-cell region; review against the PDF. This is not independent verification."
        if contained_words
        else "A detected table-cell region is available, but selectable source words could not be mapped to it; exact cell provenance is unavailable."
    )
    return evidence


def _ocr_source_row(row: list[dict[str, object]], fallback: int) -> int:
    """Return the dominant original OCR line number for an inferred row."""
    source_rows = [int(cell["source_row"]) for cell in row if cell.get("source_row") is not None]
    return max(set(source_rows), key=source_rows.count) if source_rows else fallback


def _is_likely_header(rows: list[list[str]]) -> bool:
    """Conservatively identify a first row that looks like column labels."""
    if len(rows) < 2 or not rows[0]:
        return False
    first_row = rows[0]
    non_empty = [value for value in first_row if value]
    if len(non_empty) != len(first_row) or len(set(non_empty)) != len(first_row):
        return False
    if not all(any(character.isalpha() for character in value) for value in first_row):
        return False
    next_row = rows[1]
    return any(value and value != first_row[index] for index, value in enumerate(next_row[:len(first_row)]))


def _ocr_line_cells(words: list[dict[str, object]]) -> list[dict[str, object]]:
    """Split one OCR text line into cells using large horizontal whitespace gaps."""
    ordered_words = sorted(words, key=lambda word: float(word["bbox"][0]))
    heights = [float(word["bbox"][3]) - float(word["bbox"][1]) for word in ordered_words]
    typical_height = sorted(heights)[len(heights) // 2] if heights else 0.0
    gap_threshold = max(8.0, typical_height * 1.5)
    cells: list[list[dict[str, object]]] = []
    for word in ordered_words:
        if not cells:
            cells.append([word])
            continue
        previous_right = float(cells[-1][-1]["bbox"][2])
        word_left = float(word["bbox"][0])
        if word_left - previous_right > gap_threshold:
            cells.append([word])
        else:
            cells[-1].append(word)
    return [
        {
            "text": " ".join(str(word["text"]) for word in cell),
            "left": min(float(word["bbox"][0]) for word in cell),
            "right": max(float(word["bbox"][2]) for word in cell),
            "top": min(float(word["bbox"][1]) for word in cell),
            "bottom": max(float(word["bbox"][3]) for word in cell),
            "confidence": sum(float(word["confidence"]) for word in cell) / len(cell),
            "word_boxes": [
                {
                    "text": str(word["text"]),
                    "bbox": [float(value) for value in word["bbox"]],
                    "confidence": float(word["confidence"]),
                }
                for word in cell
            ],
            "source_row": cell[0].get("source_row"),
            "source_columns": [word.get("source_column") for word in cell],
        }
        for cell in cells
    ]


def _extract_ocr_layout_table(ocr_page: dict[str, object]) -> dict[str, object] | None:
    """Infer a conservative OCR table from repeated aligned word rows.

    This layout fallback is used only when pdfplumber did not detect a table.
    At least two text lines must have the same multi-cell structure; column
    counts are not capped, so wide OCR-recognized tables remain representable.
    """
    words = ocr_page.get("words", ocr_page.get("ocr_words"))
    if not isinstance(words, list) or not words:
        return None
    line_words: dict[tuple[int, int, int], list[dict[str, object]]] = {}
    for word in words:
        if not isinstance(word, dict) or not isinstance(word.get("bbox"), (list, tuple)):
            continue
        key = (
            int(word.get("block_number", 0)),
            int(word.get("paragraph_number", 0)),
            int(word.get("line_number", 0)),
        )
        line_words.setdefault(key, []).append(word)

    row_cells = [
        cells
        for cells in (_ocr_line_cells(line) for line in line_words.values())
        if cells
    ]
    row_cells.sort(key=lambda cells: min(float(cell["top"]) for cell in cells))
    if len(row_cells) < 2:
        return None

    row_counts: dict[int, int] = {}
    for row in row_cells:
        if len(row) >= 2:
            row_counts[len(row)] = row_counts.get(len(row), 0) + 1
    if not row_counts:
        return None
    column_count, consistent_row_count = max(row_counts.items(), key=lambda item: item[1])
    if column_count < 2 or consistent_row_count < 2:
        return None
    full_rows = [row for row in row_cells if len(row) == column_count]
    max_page_width = float(
        ocr_page.get("ocr_page_width")
        or ocr_page.get("page_width")
        or (ocr_page.get("ocr_metadata") or {}).get("page_width")
        or 0
    )
    if max_page_width <= 0:
        return None
    alignment_tolerance = max(4.0, max_page_width * 0.025)
    first_starts = [float(cell["left"]) for cell in full_rows[0]]
    if any(
        any(
            abs(float(cell["left"]) - first_starts[index]) > alignment_tolerance
            for index, cell in enumerate(row)
        )
        for row in full_rows[1:]
    ):
        return None

    column_starts = [
        sum(float(row[column]["left"]) for row in full_rows) / len(full_rows)
        for column in range(column_count)
    ]
    aligned_rows: list[list[dict[str, object]]] = []
    full_table_top = min(float(cell["top"]) for row in full_rows for cell in row)
    full_table_bottom = max(float(cell["bottom"]) for row in full_rows for cell in row)
    full_row_centers = sorted(
        sum(float(cell["top"]) + float(cell["bottom"]) for cell in row) / (2 * len(row))
        for row in full_rows
    )
    row_spacing = (
        sorted(
            right - left
            for left, right in zip(full_row_centers, full_row_centers[1:])
        )[len(full_row_centers[1:]) // 2]
        if len(full_row_centers) > 1
        else 0.0
    )
    vertical_margin = max(alignment_tolerance, row_spacing * 1.25)
    for row in row_cells:
        row_center = sum(float(cell["top"]) + float(cell["bottom"]) for cell in row) / (2 * len(row))
        if row_center < full_table_top - vertical_margin or row_center > full_table_bottom + vertical_margin:
            continue
        aligned: list[dict[str, object] | None] = [None] * column_count
        valid = True
        for cell in row:
            nearest_column = min(
                range(column_count),
                key=lambda column: abs(float(cell["left"]) - column_starts[column]),
            )
            if (
                abs(float(cell["left"]) - column_starts[nearest_column]) > alignment_tolerance
                or aligned[nearest_column] is not None
            ):
                valid = False
                break
            aligned[nearest_column] = cell
        if valid and aligned and any(cell is not None for cell in aligned):
            aligned_rows.append(
                [
                    cell
                    if cell is not None
                    else {
                        "text": "",
                        "left": column_starts[index],
                        "right": column_starts[index],
                        "top": 0.0,
                        "bottom": 0.0,
                        "confidence": 0.0,
                        "word_boxes": [],
                        "source_row": None,
                        "source_columns": [],
                    }
                    for index, cell in enumerate(aligned)
                ]
            )
    if len(aligned_rows) < 2:
        return None

    rows = [[str(cell["text"]) for cell in row] for row in aligned_rows]
    page_number = int(ocr_page["page_number"])
    non_empty_cells = [cell for row in aligned_rows for cell in row if cell["text"]]
    left_edges = [float(cell["left"]) for cell in non_empty_cells]
    right_edges = [float(cell["right"]) for cell in non_empty_cells]
    top_edges = [float(cell["top"]) for cell in non_empty_cells]
    bottom_edges = [float(cell["bottom"]) for cell in non_empty_cells]
    column_ends = [
        sum(float(row[column]["right"]) for row in full_rows) / len(full_rows)
        for column in range(column_count)
    ]
    boundaries = [column_starts[0]] + [
        (column_ends[index] + column_starts[index + 1]) / 2
        for index in range(column_count - 1)
    ] + [column_ends[-1]]
    headers = rows[0] if _is_likely_header(rows) else None
    return {
        "table_id": f"page-{page_number}-ocr-table-1",
        "source_page": page_number,
        "table_index_on_page": 1,
        "row_count": len(rows),
        "column_count": column_count,
        "headers": headers,
        "headers_inferred": headers is not None,
        "rows": [
            {
                "row_index": index,
                "cells": row,
                "source_page": page_number,
                "is_inferred_header": index == 0 and headers is not None,
                "extraction_method": "ocr",
                "source_row_index": _ocr_source_row(aligned_rows[index], index + 1),
                "cell_evidence": [
                    {
                        "source_page": page_number,
                        "source_row": aligned_rows[index][column_index].get("source_row")
                        or _ocr_source_row(aligned_rows[index], index + 1),
                        "source_column": aligned_rows[index][column_index].get("source_columns", [None])[0]
                        if aligned_rows[index][column_index].get("source_columns")
                        else None,
                        "extraction_method": "ocr",
                        "source_text_method": "ocr" if aligned_rows[index][column_index]["word_boxes"] else "unavailable",
                        "table_column": column_index + 1,
                        "coordinate_system": "PDF points, top-left origin",
                        "bbox": [
                            float(aligned_rows[index][column_index]["left"]),
                            float(aligned_rows[index][column_index]["top"]),
                            float(aligned_rows[index][column_index]["right"]),
                            float(aligned_rows[index][column_index]["bottom"]),
                        ] if aligned_rows[index][column_index]["word_boxes"] else None,
                        "source_text": aligned_rows[index][column_index]["text"] or None,
                        "word_boxes": aligned_rows[index][column_index]["word_boxes"],
                        "ocr_confidence": float(aligned_rows[index][column_index]["confidence"])
                        if aligned_rows[index][column_index]["word_boxes"]
                        else None,
                        "evidence_status": "source_text_mapped" if aligned_rows[index][column_index]["word_boxes"] else "unavailable",
                        "confidence_or_limitations": (
                            "OCR-recognized word text and bounding boxes are available; compare against the source image. This is not independent verification."
                            if aligned_rows[index][column_index]["word_boxes"]
                            else "No OCR word was mapped to this blank cell; no source value was inferred."
                        ),
                    }
                    for column_index in range(column_count)
                ],
                "cell_confidences": [float(cell["confidence"]) for cell in aligned_rows[index]],
                "ocr_confidence_mean": (
                    sum(float(cell["confidence"]) for cell in aligned_rows[index] if cell["text"])
                    / max(1, sum(1 for cell in aligned_rows[index] if cell["text"]))
                ),
            }
            for index, row in enumerate(rows)
        ],
        "layout_word_count": len(words),
        "bbox": [min(left_edges), min(top_edges), max(right_edges), max(bottom_edges)],
        "column_edges": boundaries,
        "page_width": max_page_width,
        "page_height": float(
            ocr_page.get("ocr_page_height")
            or ocr_page.get("page_height")
            or (ocr_page.get("ocr_metadata") or {}).get("page_height")
            or 0
        ),
        "extraction_method": "ocr_layout",
        "ocr_confidence_mean": sum(float(cell["confidence"]) for cell in non_empty_cells)
        / max(1, len(non_empty_cells)),
    }


def extract_ocr_layout_tables(
    ocr_pages: list[dict[str, object]],
    page_numbers: list[int] | None = None,
) -> list[dict[str, object]]:
    """Infer layout tables only from supplied successful OCR page results."""
    requested = set(page_numbers) if page_numbers is not None else None
    tables: list[dict[str, object]] = []
    for page in ocr_pages:
        try:
            page_number = int(page["page_number"])
        except (KeyError, TypeError, ValueError):
            continue
        if requested is not None and page_number not in requested:
            continue
        if page.get("status") != "ocr_text":
            continue
        try:
            table = _extract_ocr_layout_table(page)
        except Exception:
            LOGGER.exception("OCR table layout inference failed on page %s", page_number)
            continue
        if table is not None:
            tables.append(table)
    return tables


def extract_pdf_tables(
    pdf_bytes: bytes,
    progress_callback: ProgressCallback | None = None,
    max_pages: int = MAX_PDF_PAGES,
    filename: str = "upload.pdf",
    ocr_pages: list[dict[str, object]] | None = None,
) -> dict[str, object]:
    """Extract per-page tables with pdfplumber and PyMuPDF layout metadata.

    Each distinct detected table remains separate. Completely empty rows are
    removed; partially populated rows and original header values are retained.
    A page-level warning is included when no table is found or parsing fails.
    """
    fitz_document = _open_pdf(pdf_bytes, max_pages=max_pages)
    tables: list[dict[str, object]] = []
    page_statuses: list[dict[str, object]] = []
    warnings: list[str] = []
    ocr_by_page = {
        int(page["page_number"]): page
        for page in (ocr_pages or [])
        if isinstance(page, dict) and page.get("page_number") is not None
    }
    try:
        try:
            plumber_document = pdfplumber.open(BytesIO(pdf_bytes))
        except Exception as exc:
            LOGGER.exception("pdfplumber could not open uploaded PDF")
            raise PdfInputError(
                "The PDF opened in PyMuPDF but could not be read by the table extraction engine."
            ) from exc

        with plumber_document:
            page_count = fitz_document.page_count
            if len(plumber_document.pages) != page_count:
                warnings.append("PDF parsers reported different page counts; common pages were processed.")
                page_count = min(page_count, len(plumber_document.pages))
            for index in range(page_count):
                page_number = index + 1
                try:
                    # Supplementary PyMuPDF layout context; pdfplumber remains the table detector.
                    fitz_page = fitz_document.load_page(index)
                    layout_word_count = len(fitz_page.get_text("words"))
                    plumber_page = plumber_document.pages[index]
                    detected = plumber_page.find_tables()
                    detected_tables = detected.tables if hasattr(detected, "tables") else detected
                    page_table_count = 0
                    for table_index, detected_table in enumerate(detected_tables, start=1):
                        extracted = detected_table.extract() or []
                        rows = [
                            [_clean_cell(cell) for cell in row]
                            for row in extracted
                            if row is not None
                        ]
                        rows = [row for row in rows if any(cell for cell in row)]
                        if not rows:
                            continue
                        column_count = max(len(row) for row in rows)
                        first_row_cells = (
                            detected_table.rows[0].cells if detected_table.rows else []
                        )
                        column_edges = sorted(
                            {
                                float(edge)
                                for cell_bbox in first_row_cells
                                if cell_bbox is not None
                                for edge in (cell_bbox[0], cell_bbox[2])
                            }
                        )
                        digital_words = _digital_word_records(fitz_page, page_number)
                        row_records: list[dict[str, object]] = []
                        retained_row_index = 0
                        for source_row_index, extracted_row in enumerate(extracted):
                            if extracted_row is None:
                                continue
                            cleaned_cells = [_clean_cell(cell) for cell in extracted_row]
                            if not any(cleaned_cells):
                                continue
                            source_cell_boxes = (
                                detected_table.rows[source_row_index].cells
                                if source_row_index < len(detected_table.rows)
                                else []
                            )
                            row_records.append(
                                {
                                    "row_index": retained_row_index,
                                    "source_row_index": source_row_index + 1,
                                    "cells": cleaned_cells,
                                    "source_page": page_number,
                                    "extraction_method": "pdfplumber",
                                    "cell_evidence": [
                                        _cell_evidence(
                                            source_cell_boxes[column_index]
                                            if column_index < len(source_cell_boxes)
                                            else None,
                                            digital_words,
                                            page_number,
                                            source_row_index + 1,
                                            column_index + 1,
                                            "pdfplumber",
                                        )
                                        for column_index in range(len(cleaned_cells))
                                    ],
                                    "is_inferred_header": retained_row_index == 0 and _is_likely_header(rows),
                                }
                            )
                            retained_row_index += 1
                        headers = rows[0] if _is_likely_header(rows) else None
                        tables.append(
                            {
                                "table_id": f"page-{page_number}-table-{table_index}",
                                "source_page": page_number,
                                "table_index_on_page": table_index,
                                "row_count": len(rows),
                                "column_count": column_count,
                                "headers": headers,
                                "headers_inferred": headers is not None,
                                "rows": row_records,
                                "layout_word_count": layout_word_count,
                                "bbox": list(detected_table.bbox),
                                "column_edges": column_edges,
                                "page_width": float(plumber_page.width),
                                "page_height": float(plumber_page.height),
                                "extraction_method": "pdfplumber",
                                "source_row_positions_preserved": True,
                            }
                        )
                        page_table_count += 1
                    if page_table_count == 0:
                        ocr_table = _extract_ocr_layout_table(ocr_by_page.get(page_number, {}))
                        if ocr_table is not None:
                            ocr_table["layout_word_count"] = layout_word_count
                            tables.append(ocr_table)
                            page_table_count += 1
                    if page_table_count:
                        status = "tables_detected"
                    else:
                        status = "no_tables_detected"
                        warnings.append(f"Page {page_number}: no extractable table was detected.")
                    page_statuses.append(
                        {
                            "page_number": page_number,
                            "status": status,
                            "table_count": page_table_count,
                            "layout_word_count": layout_word_count,
                            "error": None,
                        }
                    )
                except Exception as exc:
                    LOGGER.exception("Table extraction failed on PDF page %s", page_number)
                    page_statuses.append(
                        {
                            "page_number": page_number,
                            "status": "extraction_failed",
                            "table_count": 0,
                            "layout_word_count": None,
                            "error": "Table extraction failed on this page. Review the source PDF and developer logs.",
                        }
                    )
                    warnings.append(f"Page {page_number}: table extraction failed.")
                if progress_callback is not None:
                    progress_callback(page_number, page_count, "tables")
        metadata = PdfMetadata(Path(filename).name, len(pdf_bytes), fitz_document.page_count)
    finally:
        fitz_document.close()
    return {
        "metadata": asdict(metadata),
        "tables": tables,
        "page_statuses": page_statuses,
        "warnings": warnings,
    }
