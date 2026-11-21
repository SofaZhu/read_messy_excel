from collections.abc import Callable
from io import BytesIO
from itertools import islice
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.worksheet._writer import WorksheetWriter
from openpyxl.writer.excel import ExcelWriter


class _ScopedWorksheetWriter(WorksheetWriter):
    def cleanup(self) -> None:
        # An explicit output path belongs to this request, outside openpyxl's temp-file registry.
        Path(self.out).unlink(missing_ok=True)


def export_xlsx(
    columns: list[str], rows: list[list[object]], depth: int, sheet_name: str,
    *, check_client: Callable[[], None],
) -> bytes:
    if any(len(name) > 32_767 for name in columns):
        raise ValueError("A column name is too long to export to Excel")

    check_client()
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet(sheet_name)
    sheet.freeze_panes = "A2"

    def export_cell(value: object):
        if isinstance(value, str):
            # Write strings such as '=1+1' as text.
            cell = WriteOnlyCell(sheet, value=value)
            cell.data_type = "s"
            return cell
        return value

    with TemporaryDirectory(prefix="messy-excel-") as directory:
        # The write-only sheet otherwise allocates XML in the process-wide temporary directory.
        writer = _ScopedWorksheetWriter(sheet, out=str(Path(directory) / "sheet.xml"))
        sheet._writer = writer
        try:
            writer.write_top()
            sheet.append(export_cell(name) for name in columns)
            for index, row in enumerate(islice(rows, depth, None), start=1):
                if index % 128 == 0:
                    check_client()
                values = row[:len(columns)]
                if all(value is None for value in values):
                    # Persist even a blank row that becomes the final row after cutting the tail.
                    blank = WriteOnlyCell(sheet)
                    blank.number_format = "@"
                    sheet.append([blank])
                else:
                    sheet.append(export_cell(value) for value in values)
            check_client()
            sheet.close()
        finally:
            try:
                if not sheet.closed and sheet._rows is not None:
                    sheet._rows.close()
            finally:
                writer.close()

        check_client()
        with BytesIO() as output:
            with ZipFile(output, "w", ZIP_DEFLATED) as archive:
                ExcelWriter(workbook, archive).write_data()
            check_client()
            return output.getvalue()
