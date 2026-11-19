from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from messy_excel import flatten


def main() -> None:
    with TemporaryDirectory() as folder:
        path = Path(folder) / "sales.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sales"
        for row in (
            ["Sales report"],
            ["Item", "Sales", None, "Owner"],
            [None, "Units", "Amount", None],
            ["Pears", 2, 20, "Chen"],
            ["Oranges", 3, 30, "Wang"],
            ["Apples", 4, 40, "Lin"],
            ["Total", 9, 90, None],
        ):
            sheet.append(row)
        for area in ("A1:D1", "A2:A3", "B2:C2", "D2:D3"):
            sheet.merge_cells(area)
        workbook.save(path)
        workbook.close()

        frame, report = flatten(path)
        print(frame.to_string(index=False))
        print("Header rows:", report.header_rows)
        print("Filled merges:", report.filled_merge_ranges)
        print("Dropped tail rows:", report.dropped_tail_rows)


if __name__ == "__main__":
    main()
