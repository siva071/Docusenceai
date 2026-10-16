"""Unit tests for safe, non-mutating table validation and reconciliation."""

from __future__ import annotations

import unittest

from modules.validator import validate_tables


def make_table(
    headers: list[str],
    data_rows: list[list[object]],
    *,
    source_pages: list[int] | None = None,
    table_id: str = "test-table",
    extraction_method: str = "digital",
) -> dict[str, object]:
    """Build a table in the extraction engine's current data shape."""
    source_pages = source_pages or [1] * len(data_rows)
    rows: list[dict[str, object]] = [
        {
            "row_index": 0,
            "cells": headers,
            "source_page": source_pages[0] if source_pages else 1,
            "is_inferred_header": True,
        }
    ]
    for index, values in enumerate(data_rows):
        rows.append(
            {
                "row_index": index + 1,
                "cells": list(values),
                "source_page": source_pages[index],
                "is_inferred_header": False,
                "extraction_method": extraction_method,
            }
        )
    return {
        "table_id": table_id,
        "source_page": source_pages[0] if source_pages else 1,
        "column_count": len(headers),
        "headers": headers,
        "headers_inferred": True,
        "rows": rows,
    }


class ValidatorTests(unittest.TestCase):
    def test_missing_cell_is_reported_with_source_page_and_location(self) -> None:
        table = make_table(["Account", "Balance"], [["A-1", "10"], ["A-2", ""]], source_pages=[2, 4])

        report = validate_tables([table])
        finding = next(item for item in report["findings"] if item["category"] == "missing_cell")

        self.assertEqual(finding["severity"], "warning")
        self.assertEqual(finding["source_page"], 4)
        self.assertEqual(finding["row"], 3)
        self.assertIn("Balance", finding["column"])
        self.assertIn("source page", finding["suggested_action"])
        self.assertEqual(table["rows"][2]["cells"], ["A-2", ""])

    def test_exact_and_near_duplicate_rows_are_reported_not_removed(self) -> None:
        table = make_table(
            ["Record", "Name", "Department", "Amount"],
            [
                ["R-1", "Ada", "Finance", "100"],
                ["R-1", "Ada", "Finance", "100"],
                ["R-1", "Ada", "Finance", "101"],
            ],
        )

        report = validate_tables([table])
        categories = [finding["category"] for finding in report["findings"]]

        self.assertIn("duplicate_row", categories)
        self.assertIn("duplicate_like", categories)
        self.assertEqual(len(table["rows"]), 4)
        self.assertEqual(table["rows"][3]["cells"][-1], "101")

    def test_numeric_text_and_mixed_numeric_formats_are_flagged(self) -> None:
        table = make_table(
            ["Amount"],
            [["1,234.50"], ["1.234,50"], ["not reported"], ["12..50"]],
            source_pages=[1, 1, 3, 3],
        )

        report = validate_tables([table])
        categories = [finding["category"] for finding in report["findings"]]

        self.assertIn("numeric_text", categories)
        self.assertIn("invalid_numeric", categories)
        self.assertIn("number_format", categories)
        numeric_finding = next(item for item in report["findings"] if item["category"] == "numeric_text")
        self.assertEqual(numeric_finding["source_page"], 3)
        self.assertEqual(numeric_finding["row"], 4)

    def test_date_column_reports_invalid_and_inconsistent_formats(self) -> None:
        table = make_table(
            ["Due Date"],
            [["2026-10-16"], ["16/10/2026"], ["2026-99-99"]],
            source_pages=[1, 2, 2],
        )

        report = validate_tables([table])
        findings = [item for item in report["findings"] if item["category"] == "suspicious_date"]

        self.assertGreaterEqual(len(findings), 2)
        self.assertTrue(any(item["source_page"] == 2 and item["row"] == 4 for item in findings))

    def test_correct_and_mismatched_totals_only_when_components_identified(self) -> None:
        table = make_table(
            ["Subtotal", "Tax", "Total"],
            [["100.00", "5.00", "105.00"], ["100.00", "5.00", "110.00"]],
            source_pages=[1, 3],
        )

        report = validate_tables([table])
        mismatches = [item for item in report["findings"] if item["category"] == "arithmetic_mismatch"]

        self.assertEqual(len(mismatches), 1)
        self.assertEqual(mismatches[0]["severity"], "warning")
        self.assertEqual(mismatches[0]["source_page"], 3)
        self.assertEqual(mismatches[0]["row"], 3)
        self.assertEqual(mismatches[0]["actual_value"], "110.00")
        self.assertEqual(mismatches[0]["expected_value"], "105.00")

    def test_discount_reduces_total_for_positive_and_negative_conventions(self) -> None:
        table = make_table(
            ["Subtotal", "Tax", "Discount", "Total"],
            [["100.00", "5.00", "10.00", "95.00"], ["100.00", "5.00", "-10.00", "95.00"]],
            source_pages=[2, 3],
        )

        report = validate_tables([table])

        self.assertFalse(any(item["category"] == "arithmetic_mismatch" for item in report["findings"]))

    def test_blank_component_is_unknown_and_does_not_trigger_mismatch(self) -> None:
        table = make_table(
            ["Subtotal", "Tax", "Discount", "Total"],
            [["100.00", "5.00", "", "105.00"], ["100.00", "5.00", "10.00", "95.00"]],
            source_pages=[4, 5],
        )

        report = validate_tables([table])

        self.assertFalse(any(item["category"] == "arithmetic_mismatch" for item in report["findings"]))
        finding = next(item for item in report["findings"] if item["category"] == "arithmetic_incomplete")
        self.assertEqual(finding["severity"], "info")
        self.assertEqual(finding["source_page"], 4)
        self.assertIn("Discount", finding["evidence"])
        self.assertTrue(any(item["category"] == "missing_cell" for item in report["findings"]))
        self.assertEqual(table["rows"][1]["cells"], ["100.00", "5.00", "", "105.00"])

    def test_large_total_does_not_get_a_proportionally_large_tolerance(self) -> None:
        table = make_table(
            ["Subtotal", "Tax", "Total"],
            [["10000.00", "5.00", "10000.00"]],
            source_pages=[5],
        )

        report = validate_tables([table])
        mismatch = next(item for item in report["findings"] if item["category"] == "arithmetic_mismatch")

        self.assertEqual(mismatch["expected_value"], "10005.00")
        self.assertEqual(mismatch["actual_value"], "10000.00")

    def test_nonfinancial_numeric_table_does_not_infer_total_relationship(self) -> None:
        table = make_table(["X coordinate", "Y coordinate"], [["10", "20"], ["12", "25"]])

        report = validate_tables([table])

        self.assertFalse(any(item["category"] == "arithmetic_mismatch" for item in report["findings"]))

    def test_ocr_confidence_issue_keeps_source_page_metadata(self) -> None:
        table = make_table(["Code", "Amount"], [["A-1", "10.50"]], source_pages=[7], extraction_method="ocr")
        table["rows"][1]["cell_confidences"] = [48.0, 92.0]

        report = validate_tables([table])
        finding = next(item for item in report["findings"] if item["category"] == "ocr_review")

        self.assertEqual(finding["source_page"], 7)
        self.assertEqual(finding["row"], 2)
        self.assertEqual(finding["column"], "Code (column 1)")

    def test_intentionally_blank_ocr_cell_is_not_reported_as_low_confidence(self) -> None:
        table = make_table(["Code", "Description"], [["A-1", ""]], source_pages=[2])
        table["rows"][1]["cell_confidences"] = [91.0, 0.0]

        report = validate_tables([table])

        self.assertFalse(any(item["category"] == "ocr_review" for item in report["findings"]))

    def test_row_shape_empty_tables_and_no_issue_report(self) -> None:
        malformed = make_table(["A", "B"], [["only one cell"]])
        overlong = make_table(["A", "B"], [["one", "two", "unexpected third"]])
        empty = {"table_id": "empty-table", "source_page": 5, "column_count": 2, "rows": []}
        clean = make_table(["Name", "Status"], [["Ada", "Active"], ["Lin", "Inactive"]])

        malformed_report = validate_tables([malformed])
        overlong_report = validate_tables([overlong])
        empty_report = validate_tables([empty])
        no_tables_report = validate_tables([])
        clean_report = validate_tables([clean])

        self.assertTrue(any(item["category"] == "row_length" for item in malformed_report["findings"]))
        self.assertTrue(
            any(item["category"] == "possible_misalignment" for item in malformed_report["findings"])
        )
        self.assertTrue(any(item["category"] == "row_length" for item in overlong_report["findings"]))
        self.assertTrue(
            any(item["category"] == "possible_misalignment" for item in overlong_report["findings"])
        )
        self.assertTrue(any(item["category"] == "empty_table" for item in empty_report["findings"]))
        self.assertEqual(no_tables_report["findings"], [])
        self.assertIn("No tables were available", no_tables_report["summary"])
        self.assertEqual(clean_report["findings"], [])
        self.assertIn("No validation issues were detected", clean_report["summary"])
        self.assertEqual(clean["rows"][1]["cells"], ["Ada", "Active"])


if __name__ == "__main__":
    unittest.main()
