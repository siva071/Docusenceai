"""Conservative stitching for tables that continue across adjacent PDF pages."""

from __future__ import annotations

from dataclasses import dataclass, field
from difflib import SequenceMatcher
import re
import unicodedata
from typing import Any


@dataclass(frozen=True)
class StitchResult:
    """Stitched table groups and warnings for plausible but uncertain pairs."""

    tables: list[dict[str, Any]]
    warnings: list[str]
    stitched_count: int


@dataclass
class _TableGroup:
    """Internal state for a table and any pages already joined to it."""

    table: dict[str, Any]
    tail: dict[str, Any]
    anchor_headers: list[str] | None
    identifiers: list[str] = field(default_factory=list)
    original_bboxes: list[object] = field(default_factory=list)


def _page_number(table: dict[str, Any]) -> int | None:
    try:
        return int(table.get("source_page"))
    except (TypeError, ValueError):
        return None


def _normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).casefold()
    text = "".join(character for character in text if not unicodedata.combining(character))
    return re.sub(r"[^a-z0-9]+", "", text)


def _headers(table: dict[str, Any]) -> list[str] | None:
    raw_headers = table.get("headers")
    if not table.get("headers_inferred") or not isinstance(raw_headers, (list, tuple)):
        return None
    return [str(header or "") for header in raw_headers]


def _headers_match(left: list[str] | None, right: list[str] | None) -> bool:
    if not left or not right or len(left) != len(right):
        return False
    for first, second in zip(left, right):
        first_normalized = _normalize_header(first)
        second_normalized = _normalize_header(second)
        if not first_normalized or not second_normalized:
            return False
        if first_normalized != second_normalized:
            # Only tolerate tiny extraction/formatting discrepancies; short labels
            # require exact normalized equality to avoid conflating distinct fields.
            if min(len(first_normalized), len(second_normalized)) < 6:
                return False
            if SequenceMatcher(None, first_normalized, second_normalized).ratio() < 0.92:
                return False
    return True


def _first_row_matches_headers(table: dict[str, Any], headers: list[str] | None) -> bool:
    rows = table.get("rows")
    if not headers or not isinstance(rows, list) or not rows or not isinstance(rows[0], dict):
        return False
    first_cells = rows[0].get("cells")
    if not isinstance(first_cells, (list, tuple)) or not _headers_match(
        headers, [str(cell or "") for cell in first_cells]
    ):
        return False
    # The extractor only marks a header when followed by a distinct data row.
    # A one-row continuation header has no following row on that page and thus
    # is not marked, but remains safe to classify as repeated from the anchor.
    return bool(rows[0].get("is_inferred_header")) or len(rows) == 1 or (
        isinstance(rows[1], dict)
        and not _headers_match(headers, [str(cell or "") for cell in rows[1].get("cells", [])])
    )


def _normalized_edges(table: dict[str, Any]) -> list[float] | None:
    edges = table.get("column_edges")
    page_width = table.get("page_width")
    if not isinstance(edges, (list, tuple)) or len(edges) < 2 or not page_width:
        return None
    try:
        width = float(page_width)
        if width <= 0:
            return None
        return [float(edge) / width for edge in edges]
    except (TypeError, ValueError):
        return None


def _columns_align(left: dict[str, Any], right: dict[str, Any]) -> bool:
    left_edges = _normalized_edges(left)
    right_edges = _normalized_edges(right)
    if left_edges is None or right_edges is None or len(left_edges) != len(right_edges):
        return False
    return max(abs(first - second) for first, second in zip(left_edges, right_edges)) <= 0.015


def _horizontal_overlap(left: dict[str, Any], right: dict[str, Any]) -> float:
    left_box = left.get("bbox")
    right_box = right.get("bbox")
    left_width = left.get("page_width")
    right_width = right.get("page_width")
    try:
        left_x0, left_x1 = float(left_box[0]) / float(left_width), float(left_box[2]) / float(left_width)
        right_x0, right_x1 = float(right_box[0]) / float(right_width), float(right_box[2]) / float(right_width)
    except (TypeError, ValueError, IndexError, ZeroDivisionError):
        return 0.0
    intersection = max(0.0, min(left_x1, right_x1) - max(left_x0, right_x0))
    smallest_width = min(left_x1 - left_x0, right_x1 - right_x0)
    return intersection / smallest_width if smallest_width > 0 else 0.0


def _at_page_break(left: dict[str, Any], right: dict[str, Any]) -> bool:
    """Check that the earlier table ends low and the next begins high on its page."""
    try:
        left_bottom = float(left["bbox"][3]) / float(left["page_height"])
        right_top = float(right["bbox"][1]) / float(right["page_height"])
    except (KeyError, TypeError, ValueError, IndexError, ZeroDivisionError):
        return False
    return left_bottom >= 0.62 and right_top <= 0.38


def _row_structure_is_valid(table: dict[str, Any], column_count: int) -> bool:
    rows = table.get("rows")
    if not isinstance(rows, list) or not rows:
        return False
    return all(
        isinstance(row, dict)
        and isinstance(row.get("cells"), (list, tuple))
        and len(row["cells"]) <= column_count
        for row in rows
    )


def _assessment(group: _TableGroup, current: dict[str, Any]) -> tuple[bool, str | None]:
    """Return whether a candidate continuation is safe to merge and a warning."""
    previous = group.tail
    previous_page = _page_number(previous)
    current_page = _page_number(current)
    if previous_page is None or current_page != previous_page + 1:
        return False, None

    previous_columns = previous.get("column_count")
    current_columns = current.get("column_count")
    try:
        same_columns = int(previous_columns) == int(current_columns)
    except (TypeError, ValueError):
        same_columns = False
    at_break = _at_page_break(previous, current)
    overlap = _horizontal_overlap(previous, current)
    if not at_break or overlap < 0.75:
        # Similar headers/alignment on adjacent pages are still useful to flag,
        # but do not alone establish that a table crosses a page boundary.
        if same_columns and _headers_match(group.anchor_headers, _headers(current)):
            return False, (
                f"Adjacent tables {previous.get('table_id')} and {current.get('table_id')} "
                "look similar but are not positioned at a clear page break; kept separate."
            )
        return False, None

    if not same_columns:
        return False, (
            f"Tables {previous.get('table_id')} and {current.get('table_id')} meet near a page break "
            "but have different column counts; kept separate."
        )

    if not _columns_align(previous, current):
        return False, (
            f"Adjacent tables {previous.get('table_id')} and {current.get('table_id')} have matching "
            "column counts but their column alignments differ; kept separate."
        )

    if not _row_structure_is_valid(previous, int(previous_columns)) or not _row_structure_is_valid(
        current, int(current_columns)
    ):
        return False, (
            f"Adjacent tables {previous.get('table_id')} and {current.get('table_id')} have incomplete "
            "row structure; kept separate because continuation is uncertain."
        )

    current_headers = _headers(current)
    current_first_row_matches = _first_row_matches_headers(current, group.anchor_headers)
    if current_headers is not None and not _headers_match(group.anchor_headers, current_headers):
        if current_first_row_matches:
            return False, (
                f"Tables {previous.get('table_id')} and {current.get('table_id')} contain conflicting "
                "header evidence; kept separate."
            )
        return False, (
            f"Tables {previous.get('table_id')} and {current.get('table_id')} have different headers; kept separate."
        )
    repeated_header = _headers_match(group.anchor_headers, current_headers) or current_first_row_matches
    if repeated_header:
        return True, None

    if group.anchor_headers and current_headers is None:
        # Repeatedly aligned columns plus placement at opposite page boundaries
        # provides sufficient evidence for a continuation without a repeated header.
        return True, None

    return False, (
        f"Tables {previous.get('table_id')} and {current.get('table_id')} are adjacent and aligned, "
        "but their headers do not reliably match; kept separate."
    )


def _make_group(table: dict[str, Any]) -> _TableGroup:
    copied = dict(table)
    copied["rows"] = [dict(row) if isinstance(row, dict) else row for row in table.get("rows", [])]
    table_id = str(table.get("table_id", "table"))
    return _TableGroup(
        table=copied,
        tail=table,
        anchor_headers=_headers(table),
        identifiers=[table_id],
        original_bboxes=[table.get("bbox")],
    )


def _append_continuation(group: _TableGroup, continuation: dict[str, Any]) -> None:
    current_rows = continuation.get("rows", [])
    rows_to_add = [dict(row) if isinstance(row, dict) else row for row in current_rows]
    if rows_to_add and isinstance(rows_to_add[0], dict):
        first_row = rows_to_add[0]
        repeated_header = bool(first_row.get("is_inferred_header")) or _first_row_matches_headers(
            continuation, group.anchor_headers
        )
        if repeated_header and _headers_match(group.anchor_headers, [
            str(cell or "") for cell in first_row.get("cells", [])
        ]):
            rows_to_add.pop(0)

    combined_rows = list(group.table.get("rows", [])) + rows_to_add
    for index, row in enumerate(combined_rows):
        if isinstance(row, dict):
            row["row_index"] = index
    group.table["rows"] = combined_rows
    group.table["row_count"] = len(combined_rows)
    group.table["stitched"] = True
    group.identifiers.append(str(continuation.get("table_id", "table")))
    group.table["stitched_table_ids"] = list(group.identifiers)
    group.original_bboxes.append(continuation.get("bbox"))
    group.table["source_bboxes"] = list(group.original_bboxes)
    group.table["source_pages"] = sorted(
        {
            page
            for page in (_page_number(group.tail), _page_number(continuation))
            if page is not None
        }
        | set(group.table.get("source_pages", []))
    )
    group.tail = continuation


def stitch_tables(tables: list[dict[str, Any]]) -> StitchResult:
    """Join only strongly supported continuations across consecutive pages.

    Candidate tables are compared by page adjacency, inferred header names,
    column count, normalized column-boundary alignment, horizontal overlap,
    row structure, and top/bottom page-break placement. Repeated headers are
    removed only when they match the initial table headers. Uncertain candidates
    remain separate and receive an explanatory warning.
    """
    if not tables:
        return StitchResult([], [], 0)

    ordered = sorted(
        (dict(table) for table in tables),
        key=lambda table: (
            _page_number(table) if _page_number(table) is not None else 10**9,
            int(table.get("table_index_on_page", 0) or 0),
        ),
    )
    groups: list[_TableGroup] = []
    warnings: list[str] = []
    stitched_count = 0

    for current in ordered:
        candidates: list[_TableGroup] = []
        uncertain_messages: list[str] = []
        current_page = _page_number(current)
        for group in groups:
            if current_page is None or _page_number(group.tail) != current_page - 1:
                continue
            can_merge, warning = _assessment(group, current)
            if can_merge:
                candidates.append(group)
            elif warning:
                uncertain_messages.append(warning)

        if len(candidates) == 1:
            _append_continuation(candidates[0], current)
            stitched_count += 1
        else:
            if len(candidates) > 1:
                uncertain_messages.append(
                    f"Table {current.get('table_id')} has multiple plausible prior-page matches; kept separate."
                )
            warnings.extend(uncertain_messages)
            groups.append(_make_group(current))

    result_tables: list[dict[str, Any]] = []
    for group in groups:
        if group.table.get("stitched"):
            first_page = _page_number(group.table)
            last_page = _page_number(group.tail)
            group.table["table_id"] = f"stitched-pages-{first_page}-{last_page}-table"
            group.table["source_page"] = first_page
            actual_source_pages = sorted(
                {
                    int(row["source_page"])
                    for row in group.table.get("rows", [])
                    if isinstance(row, dict) and row.get("source_page") is not None
                }
            )
            group.table["source_pages"] = actual_source_pages or sorted(
                {
                    page
                    for page in (_page_number(group.table), _page_number(group.tail))
                    if page is not None
                }
            )
        result_tables.append(group.table)

    return StitchResult(result_tables, list(dict.fromkeys(warnings)), stitched_count)
