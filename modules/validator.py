"""Conservative, non-mutating validation and reconciliation for extracted tables."""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
import math
import re
from typing import Any


_NUMERIC_HEADER_TERMS = {
    "amount", "price", "cost", "balance", "quantity", "qty", "count", "rate",
    "tax", "fee", "subtotal", "total", "debit", "credit", "payment", "value",
    "units", "percent", "percentage", "discount", "charge", "shipping",
}
_DATE_HEADER_TERMS = {"date", "dated", "timestamp", "birthdate", "duedate"}
_COMPONENT_HEADER_TERMS = {"subtotal", "tax", "fee", "fees", "shipping", "charge", "charges", "discount", "rebate"}
_DATE_FORMATS = (
    ("iso", "%Y-%m-%d"),
    ("iso_slash", "%Y/%m/%d"),
    ("us_slash", "%m/%d/%Y"),
    ("eu_slash", "%d/%m/%Y"),
    ("us_short", "%m/%d/%y"),
    ("eu_short", "%d/%m/%y"),
    ("month_name", "%b %d, %Y"),
    ("month_name_long", "%B %d, %Y"),
    ("day_month_name", "%d %b %Y"),
)


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _normalized(value: object) -> str:
    return re.sub(r"\s+", " ", _text(value)).casefold()


def _is_blank(value: object) -> bool:
    return value is None or not _text(value)


def _parse_number(value: object) -> tuple[Decimal | None, str | None]:
    """Parse common numeric conventions without interpreting arbitrary text."""
    if isinstance(value, bool) or value is None:
        return None, None
    if isinstance(value, (int, float, Decimal)):
        if isinstance(value, float) and not math.isfinite(value):
            return None, None
        try:
            number = Decimal(str(value))
            return (number, "native") if number.is_finite() else (None, None)
        except InvalidOperation:
            return None, None

    raw = _text(value).replace("\u00a0", " ")
    if not raw:
        return None, None
    negative_parentheses = raw.startswith("(") and raw.endswith(")")
    if negative_parentheses:
        raw = raw[1:-1].strip()
    raw = raw.replace("−", "-").replace("–", "-").replace("—", "-")
    raw = re.sub(r"[\s$€£¥₹]", "", raw)
    percent = raw.endswith("%")
    if percent:
        raw = raw[:-1]
    if not re.fullmatch(r"[+-]?\d[\d.,]*", raw):
        return None, None

    style: str
    if "," in raw and "." in raw:
        decimal_separator = "." if raw.rfind(".") > raw.rfind(",") else ","
        grouping_separator = "," if decimal_separator == "." else "."
        integer_part, decimal_part = raw.rsplit(decimal_separator, 1)
        if len(decimal_part) not in {1, 2} or not re.fullmatch(
            rf"[+-]?\d{{1,3}}(?:{re.escape(grouping_separator)}\d{{3}})+", integer_part
        ):
            return None, None
        canonical = integer_part.replace(grouping_separator, "") + "." + decimal_part
        style = "comma_decimal" if decimal_separator == "," else "dot_decimal"
    elif "," in raw:
        if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", raw):
            canonical = raw.replace(",", "")
            style = "comma_grouping"
        else:
            integer_part, decimal_part = raw.rsplit(",", 1)
            if len(decimal_part) not in {1, 2}:
                return None, None
            canonical = integer_part + "." + decimal_part
            style = "comma_decimal"
    elif "." in raw:
        canonical = raw
        style = "dot_decimal"
    else:
        canonical = raw
        style = "integer"
    if negative_parentheses:
        canonical = canonical.lstrip("+-")
        canonical = "-" + canonical
    try:
        number = Decimal(canonical)
    except InvalidOperation:
        return None, None
    if percent:
        number /= Decimal(100)
    return number, style


def _header_terms(header: object) -> set[str]:
    return set(re.findall(r"[a-z]+", _text(header).casefold()))


def _looks_numeric_column(header: object, values: list[object]) -> bool:
    non_blank = [value for value in values if not _is_blank(value)]
    if not non_blank:
        return False
    numeric_count = sum(_parse_number(value)[0] is not None for value in non_blank)
    explicit_numeric = bool(_header_terms(header) & _NUMERIC_HEADER_TERMS)
    if explicit_numeric:
        return numeric_count >= 1 and (len(non_blank) < 3 or numeric_count / len(non_blank) >= 0.5)
    return len(non_blank) >= 2 and numeric_count >= 2 and numeric_count / len(non_blank) >= 0.75


def _parse_date(value: object) -> tuple[datetime | None, str | None]:
    text = _text(value)
    if not text:
        return None, None
    for label, date_format in _DATE_FORMATS:
        try:
            return datetime.strptime(text, date_format), label
        except ValueError:
            continue
    return None, None


def _table_rows(table: dict[str, Any]) -> list[dict[str, Any]]:
    rows = table.get("rows")
    if not isinstance(rows, list):
        return []
    return [row for row in rows if isinstance(row, dict) and isinstance(row.get("cells"), (list, tuple))]


def _table_headers(table: dict[str, Any], column_count: int) -> list[str]:
    headers = table.get("headers")
    if table.get("headers_inferred") and isinstance(headers, (list, tuple)):
        result = [_text(header) for header in headers]
    else:
        result = []
    return result + [f"Column {index + 1}" for index in range(len(result), column_count)]


def _finding(
    *,
    table_name: str,
    category: str,
    severity: str,
    issue: str,
    action: str,
    source_page: object = None,
    row: int | None = None,
    column: str | None = None,
    related_row: int | None = None,
    evidence: str | None = None,
) -> dict[str, Any]:
    return {
        "table_name": table_name,
        "source_page": source_page,
        "row": row,
        "column": column,
        "related_row": related_row,
        "severity": severity,
        "category": category,
        "issue": issue,
        "suggested_action": action,
        "evidence": evidence,
        "actual_value": None,
        "expected_value": None,
    }


def _row_is_header(row: dict[str, Any], index: int) -> bool:
    return index == 0 and bool(row.get("is_inferred_header"))


def _row_signature(cells: list[object]) -> tuple[str, ...]:
    return tuple(_normalized(cell) for cell in cells)


def _validate_table(table: dict[str, Any], table_index: int) -> list[dict[str, Any]]:
    table_name = str(table.get("table_id") or f"Table {table_index}")
    raw_rows = table.get("rows")
    findings: list[dict[str, Any]] = []
    if not isinstance(raw_rows, list) or not raw_rows:
        return [
            _finding(
                table_name=table_name,
                category="empty_table",
                severity="info",
                issue="No extracted rows were available for validation.",
                action="Review the source PDF to confirm whether this table contains data.",
                source_page=table.get("source_page"),
            )
        ]

    rows = _table_rows(table)
    if not rows:
        return [
            _finding(
                table_name=table_name,
                category="invalid_structure",
                severity="warning",
                issue="The extracted table rows do not use a supported row structure.",
                action="Review the table extraction and source PDF.",
                source_page=table.get("source_page"),
            )
        ]

    inferred_header_count = (
        len(table.get("headers", []))
        if table.get("headers_inferred") and isinstance(table.get("headers"), (list, tuple))
        else 0
    )
    try:
        declared_column_count = int(table.get("column_count", 0) or 0)
    except (TypeError, ValueError):
        declared_column_count = 0
    # Trust headers and explicit extraction metadata as structural evidence;
    # do not let an unexpectedly wide row redefine the expected column count.
    column_count = inferred_header_count or declared_column_count or max(
        (len(row["cells"]) for row in rows), default=0
    )
    if column_count == 0:
        return [
            _finding(
                table_name=table_name,
                category="empty_table",
                severity="info",
                issue="No table columns were available for validation.",
                action="Review the source PDF and table extraction.",
                source_page=table.get("source_page"),
            )
        ]
    headers = _table_headers(table, column_count)
    data_rows = [(index, row) for index, row in enumerate(rows) if not _row_is_header(row, index)]
    columns = [
        [row["cells"][column] if column < len(row["cells"]) else None for _, row in data_rows]
        for column in range(column_count)
    ]

    for row_index, row in enumerate(rows):
        cells = list(row["cells"])
        source_page = row.get("source_page", table.get("source_page"))
        if len(cells) != column_count and not _row_is_header(row, row_index):
            findings.append(
                _finding(
                    table_name=table_name,
                    category="row_length",
                    severity="warning",
                    issue=f"Row has {len(cells)} cells; table structure suggests {column_count} columns.",
                    action="Compare this row with adjacent columns in the source PDF; extraction may have shifted or omitted cells.",
                    source_page=source_page,
                    row=row_index + 1,
                )
            )
            findings.append(
                _finding(
                    table_name=table_name,
                    category="possible_misalignment",
                    severity="warning",
                    issue="This row's cell count differs from the expected table structure, which may indicate an omitted, split, or shifted cell.",
                    action="Compare the row boundaries with the source page; no values were changed.",
                    source_page=source_page,
                    row=row_index + 1,
                )
            )
        if _row_is_header(row, row_index):
            continue
        row_has_value = any(not _is_blank(cell) for cell in cells)
        for column_index in range(column_count):
            value = cells[column_index] if column_index < len(cells) else None
            if not _is_blank(value):
                continue
            column_has_values = any(not _is_blank(item) for item in columns[column_index])
            if not row_has_value:
                continue
            findings.append(
                _finding(
                    table_name=table_name,
                    category="missing_cell" if column_has_values else "blank_field",
                    severity="warning" if column_has_values else "info",
                    issue=(
                        f"Cell is blank in a column that contains values elsewhere ({headers[column_index]})."
                        if column_has_values
                        else f"Column {headers[column_index]} is blank throughout the extracted data."
                    ),
                    action=(
                        "Check the source page to determine whether the value is missing or intentionally blank."
                        if column_has_values
                        else "Confirm this field is intentionally blank in the source document."
                    ),
                    source_page=source_page,
                    row=row_index + 1,
                    column=f"{headers[column_index]} (column {column_index + 1})",
                )
            )

    for column_index, header in enumerate(headers):
        values = columns[column_index]
        if not _looks_numeric_column(header, values):
            continue
        non_blank = [value for value in values if not _is_blank(value)]
        parsed = [_parse_number(value) for value in non_blank]
        numeric_count = sum(number is not None for number, _ in parsed)
        explicit_numeric = bool(_header_terms(header) & _NUMERIC_HEADER_TERMS)
        if numeric_count >= 2 or (explicit_numeric and numeric_count >= 1 and len(non_blank) >= 2):
            for row_index, row in data_rows:
                value = row["cells"][column_index] if column_index < len(row["cells"]) else None
                if _is_blank(value) or _parse_number(value)[0] is not None:
                    continue
                value_text = _text(value)
                looks_malformed_numeric = bool(re.search(r"\d", value_text))
                findings.append(
                    _finding(
                        table_name=table_name,
                        category="invalid_numeric" if looks_malformed_numeric else "numeric_text",
                        severity="warning",
                        issue=(
                            f"Value {_text(value)!r} resembles a number but could not be parsed in numeric-like column {header}."
                            if looks_malformed_numeric
                            else f"Unexpected text {_text(value)!r} appears in numeric-like column {header}."
                        ),
                        action="Check OCR or column alignment against the source; retain the extracted value unless manually corrected.",
                        source_page=row.get("source_page", table.get("source_page")),
                        row=row_index + 1,
                        column=f"{header} (column {column_index + 1})",
                    )
                )

        styles = {
            style
            for number, style in parsed
            if number is not None and style in {"comma_decimal", "dot_decimal"}
        }
        if len(styles) > 1:
            findings.append(
                _finding(
                    table_name=table_name,
                    category="number_format",
                    severity="warning",
                    issue=f"Numeric formatting is inconsistent in column {header} (comma and dot decimal conventions both appear).",
                    action="Review the source values and locale; the validator does not normalize or rewrite them.",
                    source_page=next((row.get("source_page", table.get("source_page")) for _, row in data_rows), table.get("source_page")),
                    column=f"{header} (column {column_index + 1})",
                )
            )

    for column_index, header in enumerate(headers):
        if not (_header_terms(header) & _DATE_HEADER_TERMS):
            continue
        formats: set[str] = set()
        for row_index, row in data_rows:
            value = row["cells"][column_index] if column_index < len(row["cells"]) else None
            if _is_blank(value):
                continue
            parsed_date, format_name = _parse_date(value)
            if parsed_date is None:
                findings.append(
                    _finding(
                        table_name=table_name,
                        category="suspicious_date",
                        severity="warning",
                        issue=f"Date-like column {header} contains a value that did not match supported date formats: {_text(value)!r}.",
                        action="Check the original date and expected locale/format.",
                        source_page=row.get("source_page", table.get("source_page")),
                        row=row_index + 1,
                        column=f"{header} (column {column_index + 1})",
                    )
                )
            elif format_name:
                formats.add(format_name)
        if len(formats) > 1:
            findings.append(
                _finding(
                    table_name=table_name,
                    category="suspicious_date",
                    severity="warning",
                    issue=f"Date formats vary within column {header}: {', '.join(sorted(formats))}.",
                    action="Review date conventions and locale; no date values were converted.",
                    source_page=table.get("source_page"),
                    column=f"{header} (column {column_index + 1})",
                )
            )

    seen_rows: dict[tuple[str, ...], tuple[int, dict[str, Any]]] = {}
    for row_index, row in data_rows:
        cells = list(row["cells"])
        signature = _row_signature(cells)
        if signature in seen_rows:
            first_index, first_row = seen_rows[signature]
            findings.append(
                _finding(
                    table_name=table_name,
                    category="duplicate_row",
                    severity="warning",
                    issue=f"This row exactly matches row {first_index + 1}; duplicates are reported but not removed.",
                    action="Compare the source pages and decide whether this is a legitimate repeated record.",
                    source_page=row.get("source_page", table.get("source_page")),
                    row=row_index + 1,
                    related_row=first_index + 1,
                    evidence=f"Repeated source pages: {first_row.get('source_page')} and {row.get('source_page')}.",
                )
            )
        else:
            seen_rows[signature] = (row_index, row)

    for position, (row_index, row) in enumerate(data_rows):
        cells = list(row["cells"])
        if len(cells) < 3:
            continue
        for prior_index, prior_row in data_rows[:position]:
            prior_cells = list(prior_row["cells"])
            if len(prior_cells) != len(cells):
                continue
            comparable = [
                (first, second)
                for first, second in zip(cells, prior_cells)
                if not _is_blank(first) and not _is_blank(second)
            ]
            if len(comparable) < 3:
                continue
            equal_count = sum(_normalized(first) == _normalized(second) for first, second in comparable)
            changed_count = len(comparable) - equal_count
            if changed_count == 1 and equal_count / len(comparable) >= 0.75:
                differing_columns = [
                    headers[column_index]
                    for column_index, (first, second) in enumerate(zip(cells, prior_cells))
                    if _normalized(first) != _normalized(second)
                ]
                findings.append(
                    _finding(
                        table_name=table_name,
                        category="duplicate_like",
                        severity="info",
                        issue=f"This row is similar to row {prior_index + 1}; values differ in {', '.join(differing_columns) or 'one or more cells'}.",
                        action="Review both records for a possible near-duplicate; neither row was changed or removed.",
                        source_page=row.get("source_page", table.get("source_page")),
                        row=row_index + 1,
                        related_row=prior_index + 1,
                    )
                )
                break

    component_columns = [
        index
        for index, header in enumerate(headers)
        if _header_terms(header) & _COMPONENT_HEADER_TERMS
        and not (_header_terms(header) & {"rate", "percent", "percentage"})
        and "%" not in _text(header)
    ]
    total_columns = [
        index
        for index, header in enumerate(headers)
        if "total" in _header_terms(header) and "subtotal" not in _header_terms(header)
    ]
    if len(component_columns) >= 2 and len(total_columns) == 1:
        for row_index, row in data_rows:
            cells = row["cells"]
            total_value = (
                _parse_number(cells[total_columns[0]])[0]
                if total_columns[0] < len(cells)
                else None
            )
            if total_value is None:
                continue
            component_values: list[tuple[int, Decimal]] = []
            unavailable_components: list[str] = []
            for component_index in component_columns:
                parsed_value = (
                    _parse_number(cells[component_index])[0]
                    if component_index < len(cells)
                    else None
                )
                if parsed_value is None:
                    unavailable_components.append(headers[component_index])
                else:
                    component_values.append((component_index, parsed_value))
            if unavailable_components:
                findings.append(
                    _finding(
                        table_name=table_name,
                        category="arithmetic_incomplete",
                        severity="info",
                        issue=(
                            "Total reconciliation was not performed because one or more "
                            "recognized component values are blank or ambiguous."
                        ),
                        action=(
                            "Review the source values and decide whether blank components "
                            "are intentionally zero; no value was assumed or changed."
                        ),
                        source_page=row.get("source_page", table.get("source_page")),
                        row=row_index + 1,
                        column=f"{headers[total_columns[0]]} (column {total_columns[0] + 1})",
                        evidence="Unavailable components: " + ", ".join(unavailable_components),
                    )
                )
                continue
            expected_total = Decimal(0)
            for component_index, value in component_values:
                if "discount" in _header_terms(headers[component_index]) or "rebate" in _header_terms(headers[component_index]):
                    # A discount reduces the total whether shown as a positive
                    # magnitude or as an already-negative amount.
                    expected_total -= abs(value)
                else:
                    expected_total += value
            # Percentage-based tolerance grows with large totals and can hide
            # material differences. Allow only ordinary cent rounding.
            tolerance = Decimal("0.01")
            difference = abs(total_value - expected_total)
            if difference > tolerance:
                component_summary = "; ".join(
                    f"{headers[component_index]}={value}"
                    for component_index, value in component_values
                )
                mismatch = _finding(
                    table_name=table_name,
                    category="arithmetic_mismatch",
                    severity="warning",
                    issue=(
                        f"Potential total mismatch: extracted {total_value} in "
                        f"{headers[total_columns[0]]}, while identified components "
                        f"sum to {expected_total}."
                    ),
                    action=(
                        "Verify the inferred formula and each extracted value against "
                        "the source page. No values were reconciled automatically."
                    ),
                    source_page=row.get("source_page", table.get("source_page")),
                    row=row_index + 1,
                    column=f"{headers[total_columns[0]]} (column {total_columns[0] + 1})",
                    evidence=(
                        f"Components: {component_summary}; absolute difference: {difference}; "
                        f"tolerance: {tolerance}."
                    ),
                )
                mismatch["actual_value"] = str(total_value)
                mismatch["expected_value"] = str(expected_total)
                findings.append(mismatch)

    for row_index, row in enumerate(rows):
        confidences = row.get("cell_confidences")
        if not isinstance(confidences, (list, tuple)):
            continue
        for column_index, confidence in enumerate(confidences[:column_count]):
            cell_value = row["cells"][column_index] if column_index < len(row["cells"]) else None
            if _is_blank(cell_value):
                continue
            try:
                confidence_value = float(confidence)
            except (TypeError, ValueError):
                continue
            if confidence_value < 70:
                findings.append(
                    _finding(
                        table_name=table_name,
                        category="ocr_review",
                        severity="warning",
                        issue=f"OCR confidence for this cell is {confidence_value:.1f}, below the review threshold of 70.",
                        action="Compare this OCR-derived value with the source image; the value was not corrected.",
                        source_page=row.get("source_page", table.get("source_page")),
                        row=row_index + 1,
                        column=f"{headers[column_index]} (column {column_index + 1})",
                    )
                )

    return findings


def validate_tables(tables: list[dict[str, Any]] | None) -> dict[str, Any]:
    """Validate extracted tables without modifying or deleting any source data.

    The checks are heuristic. Arithmetic totals are compared only when one
    total column and at least two recognized component headers (for example,
    Subtotal and Tax) are present and every component value is parseable.
    """
    if not tables:
        findings: list[dict[str, Any]] = []
        tables_checked = 0
    else:
        findings = []
        tables_checked = 0
        for table_index, table in enumerate(tables, start=1):
            if not isinstance(table, dict):
                findings.append(
                    _finding(
                        table_name=f"Table {table_index}",
                        category="invalid_structure",
                        severity="warning",
                        issue="Table entry has an unsupported format and could not be checked.",
                        action="Review the extraction output for this table.",
                    )
                )
                continue
            tables_checked += 1
            findings.extend(_validate_table(table, table_index))

    severities = Counter(finding["severity"] for finding in findings)
    categories = Counter(finding["category"] for finding in findings)
    return {
        "tables_checked": tables_checked,
        "findings": findings,
        "severity_counts": {severity: severities.get(severity, 0) for severity in ("error", "warning", "info")},
        "category_counts": dict(sorted(categories.items())),
        "has_findings": bool(findings),
        "summary": (
            f"{len(findings)} validation finding(s) across {tables_checked} table(s)."
            if findings
            else ("No tables were available for validation." if not tables else f"No validation issues were detected by the implemented checks across {tables_checked} table(s).")
        ),
    }
