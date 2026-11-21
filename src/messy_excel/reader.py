from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.merge import MergedCellRange
from openpyxl.worksheet.worksheet import Worksheet

from .detection import HeaderCandidate, TailCandidate, TailDecision, detect_header, detect_tail, has_label

_MAX_GRID_CELLS = 1_000_000


@dataclass(frozen=True)
class FlattenReport:
    """Parsing decisions using Excel row numbers and cell ranges."""

    sheet_name: str
    header_rows: tuple[int, ...]
    filled_merge_ranges: tuple[str, ...]
    skipped_merge_ranges: tuple[str, ...]
    dropped_tail_rows: tuple[int, ...]
    header_source: str
    candidates: tuple[HeaderCandidate, ...] = ()
    skipped_prefix: tuple[int, int] | None = None
    trimmed_blank_tail: tuple[int, int] | None = None
    tail_decisions: tuple[TailDecision, ...] = ()
    warnings: tuple[str, ...] = ()


class HeaderDetectionError(ValueError):
    """Automatic header detection could not choose a supported header."""

    def __init__(self, report: FlattenReport):
        super().__init__("Could not detect a header reliably. Specify header_row and optionally header_depth.")
        self.report = report


def flatten(
    path: str | Path,
    *,
    header_row: int | None = None,
    header_depth: int | None = None,
    sheet_name: str | int = 0,
    drop_tail: int | None = None,
) -> tuple[pd.DataFrame, FlattenReport]:
    """Read one XLSX sheet, detecting boundaries unless explicitly overridden."""
    rows, report = _read(path, header_row, header_depth, sheet_name, drop_tail)
    return _to_frame(rows, report), report


def _to_frame(rows: list[list[object]], report: FlattenReport) -> pd.DataFrame:
    depth = len(report.header_rows)
    names = _table_columns(rows, depth)
    return pd.DataFrame([row[:len(names)] for row in rows[depth:]], columns=names)


def _table_columns(rows: list[list[object]], depth: int) -> list[str]:
    width = max(
        column
        for row in rows
        for column, value in enumerate(row, start=1)
        if value is not None
    )
    return _column_names(rows[:depth], width)


def _read(
    path: str | Path,
    header_row: int | None,
    header_depth: int | None,
    sheet_name: str | int,
    drop_tail: int | None,
) -> tuple[list[list[object]], FlattenReport]:
    _validate_options(header_row, header_depth, drop_tail)
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx files are supported")

    workbook = load_workbook(path, data_only=True, keep_links=False)
    try:
        sheet = _select_sheet(workbook.worksheets, sheet_name)
        rows, report, _ = _analyze_sheet(sheet, header_row, header_depth, drop_tail)
        return rows, report
    finally:
        workbook.close()


def _validate_options(header_row: int | None, header_depth: int | None, drop_tail: int | None) -> None:
    for name, value, minimum in (
        ("header_row", header_row, 1),
        ("header_depth", header_depth, 1),
        ("drop_tail", drop_tail, 0),
    ):
        if value is not None and (type(value) is not int or value < minimum):
            raise ValueError(f"{name} must be an integer >= {minimum}")
    if header_row is None and header_depth is not None:
        raise ValueError("header_depth requires an explicit header_row")


def _select_sheet(sheets: list[Worksheet], name: str | int) -> Worksheet:
    if type(name) is int and 0 <= name < len(sheets):
        return sheets[name]
    for sheet in sheets:
        if sheet.title == name:
            return sheet
    raise ValueError(f"Unknown worksheet: {name!r}")


def _analyze_sheet(
    sheet: Worksheet, header_row: int | None, header_depth: int | None, drop_tail: int | None,
    *, apply_detected_tail: bool = True,
) -> tuple[list[list[object]], FlattenReport, tuple[TailCandidate, ...]]:
    last_row, last_column = sheet.max_row, sheet.max_column
    first_row = header_row or 1
    depth = header_depth or 1
    if first_row + depth - 1 > last_row:
        raise ValueError("Header extends beyond the worksheet")
    if (last_row - first_row + 1) * last_column > _MAX_GRID_CELLS:
        raise ValueError("Worksheet range exceeds the 1,000,000-cell limit")

    rows = [
        list(row)
        for row in sheet.iter_rows(
            min_row=first_row, max_row=last_row, max_col=last_column, values_only=True
        )
    ]
    merges = sorted(sheet.merged_cells.ranges, key=lambda item: (item.min_row, item.min_col))
    header_source = "manual"
    candidates = ()
    if header_row is None:
        header_source = "auto"
        best, candidates, warning = detect_header(rows, merges)
        if best is None:
            raise HeaderDetectionError(FlattenReport(
                sheet.title, (), (), (), (), header_source,
                candidates=candidates, warnings=(warning,) if warning else (),
            ))
        first_row, depth = best.row, best.depth
        del rows[:first_row - 1]

    header_end = first_row + depth - 1
    filled, skipped = _restore_merges(merges, rows, first_row, header_end)
    if not any(has_label(value) for row in rows[:depth] for value in row):
        raise ValueError("Selected header has no labels")

    # Keep internal empty rows and Excel row positions intact.
    while len(rows) > depth and all(value is None for value in rows[-1]):
        rows.pop()
    last_data_row = first_row + len(rows) - 1
    tail_candidates = ()
    if drop_tail is None:
        tail_candidates = detect_tail(rows[depth:], header_end + 1, merges)
        action = "dropped" if apply_detected_tail else "suggested"
        decisions = tuple(
            TailDecision(candidate.row, action if candidate.eligible else "kept", candidate.reason)
            for candidate in tail_candidates
        )
        drop_tail = sum(candidate.eligible for candidate in tail_candidates) if apply_detected_tail else 0
    elif drop_tail > len(rows) - depth:
        raise ValueError("drop_tail exceeds the number of data rows")
    else:
        decisions = tuple(
            TailDecision(number, "dropped", "manual drop_tail override")
            for number in range(last_data_row - drop_tail + 1, last_data_row + 1)
        )
    dropped = tuple(range(last_data_row - drop_tail + 1, last_data_row + 1))
    if drop_tail:
        del rows[-drop_tail:]

    report = FlattenReport(
        sheet_name=sheet.title,
        header_rows=tuple(range(first_row, header_end + 1)),
        filled_merge_ranges=tuple(filled),
        skipped_merge_ranges=tuple(skipped),
        dropped_tail_rows=dropped,
        header_source=header_source,
        candidates=candidates,
        skipped_prefix=(1, first_row - 1) if first_row > 1 else None,
        trimmed_blank_tail=(last_data_row + 1, last_row) if last_data_row < last_row else None,
        tail_decisions=decisions,
    )
    return rows, report, tail_candidates


def _restore_merges(
    merges: list[MergedCellRange], rows: list[list[object]], header_row: int, header_end: int
) -> tuple[list[str], list[str]]:
    filled = []
    skipped = []
    for area in merges:
        if area.max_row < header_row:
            continue
        if area.min_row < header_row or area.min_row <= header_end < area.max_row:
            raise ValueError(f"Merged range {area} crosses a header boundary")

        anchor = rows[area.min_row - header_row][area.min_col - 1]
        if anchor is None or (area.min_row > header_end and area.min_col != area.max_col):
            skipped.append(str(area))
            continue
        for row_number in range(area.min_row, area.max_row + 1):
            row = rows[row_number - header_row]
            row[area.min_col - 1 : area.max_col] = [anchor] * (area.max_col - area.min_col + 1)
        filled.append(str(area))
    return filled, skipped


def _column_names(headers: list[list[object]], width: int) -> list[str]:
    base_names = []
    for column in range(width):
        parts = []
        for row in headers:
            value = row[column]
            label = "" if value is None else str(value).strip()
            if label and (not parts or parts[-1] != label):
                parts.append(label)
        base_names.append(" / ".join(parts) or f"column_{get_column_letter(column + 1)}")

    reserved = set(base_names)
    next_suffix: dict[str, int] = {}
    names = []
    for base in base_names:
        if base not in next_suffix:
            names.append(base)
            next_suffix[base] = 2
            continue
        suffix = next_suffix[base]
        while f"{base}__{suffix}" in reserved:
            suffix += 1
        names.append(f"{base}__{suffix}")
        next_suffix[base] = suffix + 1
    return names
