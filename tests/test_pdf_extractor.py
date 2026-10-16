"""Synthetic-PDF tests for DocuSense AI's Stage 2 extraction engine."""

from __future__ import annotations

import unittest

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from modules.pdf_extractor import (
    PdfInputError,
    extract_pdf_tables,
    extract_pdf_text,
    get_pdf_metadata,
)


def make_pdf(page_count: int = 1) -> tuple[fitz.Document, list[fitz.Page]]:
    """Create a small in-memory PDF document for a test case."""
    document = fitz.open()
    for _ in range(page_count):
        document.new_page(width=792, height=612)
    pages = [document.load_page(index) for index in range(page_count)]
    return document, pages


def make_wide_table_pdf() -> tuple[bytes, list[list[str]]]:
    """Build a ruled 11-column PDF table with one partially populated row."""
    document, pages = make_pdf()
    page = pages[0]
    headers = [f"Field {index + 1}" for index in range(11)]
    values = [f"Value {index + 1}" for index in range(10)] + [""]
    partial_values = ["Partial row"] + [""] * 10
    rows = [headers, values, partial_values]
    x0, y0, cell_width, cell_height = 28, 80, 66, 38
    for column in range(12):
        x = x0 + column * cell_width
        page.draw_line((x, y0), (x, y0 + cell_height * len(rows)), color=(0, 0, 0), width=0.7)
    for row_index in range(len(rows) + 1):
        y = y0 + row_index * cell_height
        page.draw_line((x0, y), (x0 + 11 * cell_width, y), color=(0, 0, 0), width=0.7)
    for row_index, row in enumerate(rows):
        for column_index, value in enumerate(row):
            if value:
                page.insert_text(
                    (x0 + column_index * cell_width + 3, y0 + row_index * cell_height + 24),
                    value,
                    fontsize=7,
                )
    result = document.tobytes()
    document.close()
    return result, rows


class PdfTextExtractionTests(unittest.TestCase):
    def test_extracts_digital_text_and_metadata(self) -> None:
        document, pages = make_pdf()
        pages[0].insert_text((72, 72), "Known digital PDF text")
        pdf_bytes = document.tobytes()
        document.close()

        result = extract_pdf_text(pdf_bytes)
        metadata = get_pdf_metadata(pdf_bytes, "known.pdf")

        self.assertEqual(result["pages"][0]["text"].strip(), "Known digital PDF text")
        self.assertTrue(result["pages"][0]["has_digital_text"])
        self.assertFalse(result["pages"][0]["likely_requires_ocr"])
        self.assertGreater(len(result["pages"][0]["digital_words"]), 0)
        digital_word = result["pages"][0]["digital_words"][0]
        self.assertEqual(digital_word["page_number"], 1)
        self.assertEqual(len(digital_word["bbox"]), 4)
        self.assertEqual(result["pages"][0]["digital_evidence"]["availability"], "available")
        self.assertEqual(metadata.filename, "known.pdf")
        self.assertEqual(metadata.page_count, 1)

    def test_multi_page_text_keeps_one_based_page_numbers(self) -> None:
        document, pages = make_pdf(2)
        pages[0].insert_text((72, 72), "Text on first page")
        pages[1].insert_text((72, 72), "Text on second page")
        pdf_bytes = document.tobytes()
        document.close()

        result = extract_pdf_text(pdf_bytes)

        self.assertEqual([page["page_number"] for page in result["pages"]], [1, 2])
        self.assertIn("Text on first page", result["pages"][0]["text"])
        self.assertIn("Text on second page", result["pages"][1]["text"])

    def test_empty_page_is_marked_as_likely_needing_ocr(self) -> None:
        document, _ = make_pdf()
        pdf_bytes = document.tobytes()
        document.close()

        page = extract_pdf_text(pdf_bytes)["pages"][0]

        self.assertFalse(page["has_digital_text"])
        self.assertTrue(page["likely_requires_ocr"])
        self.assertEqual(page["status"], "no_selectable_text")

    def test_invalid_pdf_is_rejected_without_crashing(self) -> None:
        with self.assertRaises(PdfInputError):
            extract_pdf_text(b"this is not a PDF")

        with self.assertRaises(PdfInputError):
            extract_pdf_tables(b"")


class PdfTableExtractionTests(unittest.TestCase):
    def test_extracts_eleven_columns_and_preserves_partial_row(self) -> None:
        pdf_bytes, known_rows = make_wide_table_pdf()

        result = extract_pdf_tables(pdf_bytes)

        self.assertEqual(len(result["tables"]), 1)
        table = result["tables"][0]
        self.assertEqual(table["column_count"], 11)
        self.assertEqual(table["source_page"], 1)
        self.assertEqual(table["headers"], known_rows[0])
        self.assertEqual(table["rows"][0]["cells"], known_rows[0])
        self.assertEqual(table["rows"][1]["cells"], known_rows[1])
        self.assertEqual(table["rows"][2]["cells"][0], "Partial row")
        self.assertEqual(len(table["rows"][2]["cells"]), 11)
        self.assertEqual(table["rows"][2]["source_page"], 1)
        cell_evidence = table["rows"][1]["cell_evidence"][0]
        self.assertEqual(cell_evidence["source_page"], 1)
        self.assertEqual(cell_evidence["source_row"], 2)
        self.assertEqual(cell_evidence["source_column"], 1)
        self.assertEqual(cell_evidence["extraction_method"], "pdfplumber")
        self.assertIsNotNone(cell_evidence["bbox"])
        self.assertEqual(cell_evidence["source_text"], "Value 1")
        self.assertEqual(cell_evidence["evidence_status"], "source_text_mapped")

    def test_page_without_table_returns_clear_status(self) -> None:
        document, pages = make_pdf()
        pages[0].insert_text((72, 72), "No table on this page")
        pdf_bytes = document.tobytes()
        document.close()

        result = extract_pdf_tables(pdf_bytes)

        self.assertEqual(result["tables"], [])
        self.assertEqual(result["page_statuses"][0]["status"], "no_tables_detected")
        self.assertTrue(any("no extractable table" in warning for warning in result["warnings"]))


if __name__ == "__main__":
    unittest.main()
