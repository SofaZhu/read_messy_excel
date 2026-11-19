from __future__ import annotations

import argparse
import sys
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .reader import FlattenReport, HeaderDetectionError, _read


def _show(report: FlattenReport) -> None:
    print(f"Worksheet: {report.sheet_name}")
    header = ", ".join(map(str, report.header_rows)) or "undetermined"
    print(f"Header rows ({report.header_source}): {header}")
    for candidate in report.candidates:
        end = candidate.row + candidate.depth - 1
        print(
            f"  Rows {candidate.row}-{end}: score {candidate.score:.2f}, "
            f"stability {candidate.type_stability:.2f}, coverage {candidate.label_coverage:.2f}, "
            f"type change {candidate.type_boundary:.2f}, merged hierarchy {candidate.layer_evidence}, "
            f"sample rows {', '.join(map(str, candidate.sample_rows))}"
        )
    if report.candidates and not report.header_rows:
        best = report.candidates[0]
        print(f"Override syntax: --header-row {best.row} --header-depth {best.depth}")
    for name, interval in (("Skipped prefix", report.skipped_prefix), ("Blank tail", report.trimmed_blank_tail)):
        if interval:
            print(f"{name}: {interval[0]}-{interval[1]}")
    print(f"Filled merges: {', '.join(report.filled_merge_ranges) or 'none'}")
    print(f"Skipped merges: {', '.join(report.skipped_merge_ranges) or 'none'}")
    print(f"Dropped tail rows: {', '.join(map(str, report.dropped_tail_rows)) or 'none'}")
    for item in report.tail_decisions:
        print(f"  Row {item.row}: {item.action} ({item.reason})")
    for warning in report.warnings:
        print(f"Warning: {warning}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="messy-excel", description="Inspect business Excel layouts.")
    subcommands = parser.add_subparsers(dest="command", required=True)
    inspect = subcommands.add_parser("inspect", help="Show parsing decisions for one .xlsx file.")
    inspect.add_argument("path", help="Path to a local .xlsx workbook.")
    sheets = inspect.add_mutually_exclusive_group()
    sheets.add_argument("--sheet-name", help="Worksheet name.")
    sheets.add_argument("--sheet-index", type=int, default=0, help="Zero-based worksheet index (default: 0).")
    inspect.add_argument("--header-row", type=int, help="One-based first header row (default: detect).")
    inspect.add_argument("--header-depth", type=int, help="Number of header rows with --header-row (default: 1).")
    inspect.add_argument("--drop-tail", type=int, help="Number of tail rows to drop (0 keeps them, default: detect).")
    gui = subcommands.add_parser("gui", help="Open the local workbook preview.")
    gui.add_argument("--port", type=int, default=0, help="Local port (default: choose an available port).")
    gui.add_argument("--no-browser", action="store_true", help="Print the local URL without opening a browser.")
    args = parser.parse_args(argv)
    if args.command == "gui":
        from .gui import run

        try:
            run(port=args.port, open_browser=not args.no_browser)
        except (OSError, OverflowError) as error:
            print(f"Error: {error}", file=sys.stderr)
            return 1
        return 0

    sheet = args.sheet_name if args.sheet_name is not None else args.sheet_index
    try:
        _, report = _read(args.path, args.header_row, args.header_depth, sheet, args.drop_tail)
    except HeaderDetectionError as error:
        _show(error.report)
        print(f"Error: {error}", file=sys.stderr)
        return 1
    except (OSError, ValueError, KeyError, BadZipFile, InvalidFileException, ParseError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    _show(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
