from __future__ import annotations

from bisect import bisect_right
from collections import Counter
from dataclasses import dataclass, replace
from datetime import date, time, timedelta
from itertools import islice, takewhile

from openpyxl.worksheet.merge import MergedCellRange

_HEADER_SCAN_ROWS = 30
_DATA_SAMPLE_ROWS = 10
_MIN_SCORE = 0.65
_MIN_MARGIN = 0.12
_TOTAL_LABELS = {"合计", "小计", "总计", "total", "subtotal"}


@dataclass(frozen=True)
class HeaderCandidate:
    row: int
    depth: int
    score: float
    type_stability: float
    label_coverage: float
    type_boundary: float
    layer_evidence: bool
    sample_rows: tuple[int, ...]


@dataclass(frozen=True)
class TailDecision:
    row: int
    action: str
    reason: str


@dataclass(frozen=True)
class TailCandidate:
    row: int
    eligible: bool
    reason: str
    evidence: tuple[tuple[int, object], ...]


def _kind(value: object) -> str:
    if value is None or isinstance(value, str) and not value.strip():
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (float, int)):
        return "number"
    if isinstance(value, (date, time, timedelta)):
        return "date"
    return "text"


def has_label(value: object) -> bool:
    return value is not None and bool(str(value).strip())


def _total_label(row: list[object]) -> int | None:
    for column, value in enumerate(row):
        if has_label(value):
            if isinstance(value, str) and value.strip().casefold() in _TOTAL_LABELS:
                return column
            break
    return None


def detect_header(
    rows: list[list[object]], merges: list[MergedCellRange]
) -> tuple[HeaderCandidate | None, tuple[HeaderCandidate, ...], str | None]:
    limit = min(_HEADER_SCAN_ROWS, len(rows))
    header_grid = [row.copy() for row in rows[:limit]]
    header_merges = list(takewhile(lambda area: area.min_row <= limit, merges))
    blocked_cuts: set[int] = set()
    for area in header_merges:
        blocked_cuts.update(range(area.min_row, min(area.max_row, limit + 1)))
        anchor = rows[area.min_row - 1][area.min_col - 1]
        for number in range(area.min_row, min(area.max_row, limit) + 1):
            header_grid[number - 1][area.min_col - 1 : area.max_col] = [anchor] * area.size["columns"]

    # Precompute the bounded nonempty lookahead once for all candidate depths.
    samples = []
    samples_past_window = 0
    for number, row in enumerate(rows, start=1):
        if not any(map(has_label, row)) or _total_label(row) is not None:
            continue
        samples.append((number, row))
        if number > limit:
            samples_past_window += 1
            if samples_past_window == _DATA_SAMPLE_ROWS:
                break

    candidates = []
    for end in range(1, limit + 1):
        if end in blocked_cuts:
            continue
        sample = list(islice((entry for entry in samples if entry[0] > end), _DATA_SAMPLE_ROWS))
        if not sample:
            continue
        sample_rows, data = zip(*sample)
        active, stability = _profile(data)
        for depth in range(1, min(3, end) + 1):
            start = end - depth + 1
            if start - 1 in blocked_cuts:
                continue
            headers = header_grid[start - 1 : end]
            if any(not any(map(has_label, row)) for row in headers):
                continue
            if _contains_title(rows, header_merges, active, start, end):
                continue
            candidate = _score(headers, active, stability, sample_rows, data[0], header_merges, start, end)
            if depth > 1 and len(active) > 1 and not candidate.layer_evidence and any(
                sum(map(has_label, row)) == 1 for row in headers[:-1]
            ):
                continue
            candidates.append(candidate)

    candidates.sort(key=lambda candidate: -candidate.score)
    top = tuple(candidates[:3])
    if not top:
        return None, top, "No usable header candidates in the first 30 rows."
    best = top[0]
    if len(best.sample_rows) < 3:
        return None, top, "Fewer than three data rows support the top candidate."
    if not best.type_boundary and not best.layer_evidence:
        return None, top, "No type transition or merged header hierarchy supports the top candidate."
    if best.score < _MIN_SCORE:
        return None, top, "The top header candidate scored below the minimum."

    end = best.row + best.depth - 1
    for other in candidates[1:]:
        # A verified merged parent beats the same leaf row without that parent.
        if best.layer_evidence and other.row > best.row and other.row + other.depth - 1 == end:
            continue
        if best.score - other.score < _MIN_MARGIN:
            return None, top, "The leading header candidates have similar scores."
        break
    return best, top, None


def _profile(data: tuple[list[object], ...]) -> tuple[dict[int, str], float]:
    active = {}
    ratios = []
    for column, values in enumerate(zip(*data)):
        kinds = Counter(map(_kind, values))
        kinds.pop("blank", None)
        count = sum(kinds.values())
        if not count:
            continue
        dominant, frequency = kinds.most_common(1)[0]
        active[column] = dominant if count >= 2 and frequency / count >= 0.75 else "mixed"
        if count >= 2:
            ratios.append(frequency / count)
    return active, sum(ratios) / len(ratios) if ratios else 0.0


def _contains_title(
    rows: list[list[object]], merges: list[MergedCellRange], active: dict[int, str],
    start: int, end: int,
) -> bool:
    if not active:
        return False
    first, last = min(active) + 1, max(active) + 1
    return any(
        start <= area.min_row <= end
        and area.min_col <= first < last <= area.max_col
        and sum(map(has_label, rows[area.min_row - 1])) == 1
        for area in merges
    )


def _score(
    headers: list[list[object]], active: dict[int, str], stability: float,
    sample_rows: tuple[int, ...], first_data: list[object], merges: list[MergedCellRange],
    start: int, end: int,
) -> HeaderCandidate:
    coverage = sum(any(has_label(row[col]) for row in headers) for col in active) / len(active)
    comparable = {col: kind for col, kind in active.items() if kind in {"number", "boolean", "date"}}
    # Compare the first few data rows too, so a stray second header cannot be the first detail.
    leaf = headers[-1]
    # A parent merge earns the hierarchy signal only if distinct child labels lie below it.
    hierarchy = False
    for area in merges:
        if area.min_row > end:
            break
        if start <= area.min_row and area.max_row < end and area.min_col < area.max_col:
            children = {
                str(leaf[col]).strip() for col in active
                if area.min_col - 1 <= col < area.max_col and has_label(leaf[col])
            }
            if len(children) >= 2:
                hierarchy = True
                break
    boundary = (
        sum(_kind(leaf[col]) == "text" and _kind(first_data[col]) == kind for col, kind in comparable.items())
        / len(comparable) if comparable else 0.0
    )
    score = 0.40 * stability + 0.30 * coverage + 0.20 * boundary + 0.10 * hierarchy
    return HeaderCandidate(start, end - start + 1, score, stability, coverage, boundary, hierarchy, sample_rows)


def detect_tail(
    rows: list[list[object]], data_start: int, merges: list[MergedCellRange]
) -> tuple[TailCandidate, ...]:
    """Find a contiguous, corroborated suffix using Excel row coordinates."""
    suffix = []
    for index in range(len(rows) - 1, -1, -1):
        label_column = _total_label(rows[index])
        if label_column is None:
            break
        suffix.append((index, label_column))
    if not suffix:
        return ()

    # Derive the detail pattern from the nearest ordinary rows above the candidate suffix.
    detail = list(islice((
        row for row in reversed(rows[:suffix[-1][0]])
        if _total_label(row) is None and any(map(has_label, row))
    ), _DATA_SAMPLE_ROWS))
    detail_columns = (
        tuple(
            column for column, values in enumerate(zip(*detail))
            if all(_kind(value) in {"text", "date", "boolean"} for value in values)
        ) if len(detail) >= 3 else ()
    )
    candidates = []
    count = 0
    for index, label_column in suffix:
        row = rows[index]
        number = data_start + index
        label = ((label_column + 1, row[label_column]),)
        numeric = next((column for column, value in enumerate(row) if _kind(value) == "number"), None)
        missing = next(
            (column for column in detail_columns if column != label_column and not has_label(row[column])), None,
        )
        if numeric is None or missing is None:
            candidates.append(TailCandidate(number, False, "insufficient summary evidence", label))
            break
        evidence = (*label, (numeric + 1, row[numeric]), (missing + 1, row[missing]))
        candidates.append(TailCandidate(number, True, "summary label, numeric value, and missing detail field", evidence))
        count += 1

    cut = data_start + len(rows) - count
    blocking_reason = None
    if count and any(area.min_row < cut <= area.max_row for area in merges):
        blocking_reason = "merged range crosses the proposed tail boundary"
    labels_by_column: dict[int, list[int]] = {}
    for index, column in reversed(suffix[:count]):
        labels_by_column.setdefault(column + 1, []).append(data_start + index)
    if blocking_reason is None and labels_by_column:
        for area in merges:
            labels = labels_by_column.get(area.min_col)
            if labels and area.min_col == area.max_col:
                position = bisect_right(labels, area.min_row)
                if position < len(labels) and labels[position] <= area.max_row:
                    blocking_reason = "summary label was filled from a merged range"
                    break
    if blocking_reason is not None:
        candidates = [
            replace(item, eligible=False, reason=blocking_reason) if item.eligible else item
            for item in candidates
        ]
    return tuple(reversed(candidates))
