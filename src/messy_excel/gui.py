from __future__ import annotations

import json
import webbrowser
from collections.abc import Callable
from dataclasses import asdict, replace
from datetime import date, datetime, time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib.resources import files
from io import BytesIO
from select import select
from socket import MSG_PEEK
from threading import BoundedSemaphore
from urllib.parse import parse_qs, urlsplit
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

from .detection import TailCandidate
from .reader import FlattenReport, HeaderDetectionError, _analyze_sheet, _select_sheet, _table_columns, _validate_options
from .xlsx_export import export_xlsx

_ASSETS = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/app.css": ("app.css", "text/css; charset=utf-8"),
    "/app.js": ("app.js", "text/javascript; charset=utf-8"),
}
_MAX_BYTES = 32 * 1024 * 1024
_MAX_WORKBOOK_JOBS = 2
_PREVIEW_ROWS, _PREVIEW_COLUMNS = 50, 30
_XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class _ClientDisconnected(Exception):
    pass


def _demo_workbook() -> Workbook:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "销售明细"
    sheet.append(["区域销售周报"])
    sheet.append(["地区", "订单编号", "商品", "销售数据", None, None, "交付信息"])
    sheet.append([None, None, None, "数量", "单价", "金额", "负责人", "日期"])
    orders = (
        ("华东", "无线键盘", 24, 179, "林悦"),
        (None, "便携显示器", 8, 799, "林悦"),
        (None, "桌面支架", 16, 129, "周驰"),
        ("华南", "机械键盘", 20, 329, "陈宁"),
        (None, "无线鼠标", 32, 109, "陈宁"),
        (None, "扩展坞", 12, 249, "许瑶"),
        ("华北", "便携显示器", 11, 799, "沈桐"),
        (None, "桌面支架", 15, 129, "沈桐"),
        (None, "无线鼠标", 21, 109, "宋文"),
        ("西南", "扩展坞", 9, 249, "李珊"),
        (None, "无线键盘", 18, 179, "李珊"),
        (None, "机械键盘", 13, 329, "顾青"),
    )
    for index, (region, item, quantity, unit_price, owner) in enumerate(orders):
        sheet.append([region, f"SO-{2401 + index}", item, quantity, unit_price, quantity * unit_price,
                      owner, date(2026, 11, 9 + index // 3)])
    sheet.append(["合计", None, None, sum(row[2] for row in orders), None, sum(row[2] * row[3] for row in orders)])
    for area in ("A1:H1", "A2:A3", "B2:B3", "C2:C3", "D2:F2", "G2:H2", "A4:A6", "A7:A9", "A10:A12", "A13:A15"):
        sheet.merge_cells(area)
    notes = workbook.create_sheet("纯文本示例")
    for row in (("部门", "对接人"), ("销售", "林悦"), ("物流", "陈宁"), ("财务", "周驰")):
        notes.append(row)
    return workbook


def _cell_text(value: object) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, datetime) and value.time() == time.min:
        return value.date().isoformat()
    text = str(value)
    return text if len(text) <= 160 else text[:159] + "…"


def _report_data(
    report: FlattenReport, sheet: Worksheet, candidates: tuple[TailCandidate, ...] = ()
) -> dict:
    fields = ("filled_merge_ranges", "skipped_merge_ranges", "dropped_tail_rows", "tail_decisions")
    values = {name: getattr(report, name) for name in fields}
    details = asdict(replace(report, **{name: items[:30] for name, items in values.items()}))
    details.update({f"{name}_count": len(items) for name, items in values.items()})
    details["suggested_tail_count"] = sum(item.action == "suggested" for item in report.tail_decisions)
    # Candidates and decisions are in the same row order for automatic tail detection.
    for decision, candidate in zip(details["tail_decisions"], candidates):
        decision["evidence"] = [
            {
                "coordinate": f"{get_column_letter(column)}{candidate.row}",
                "value": _cell_text(value),
                "filled": value is not None and sheet.cell(candidate.row, column).value is None,
            }
            for column, value in candidate.evidence
        ]
    return details


def _source_data(sheet: Worksheet, report: FlattenReport | None = None) -> dict:
    last_row, width = sheet.max_row, sheet.max_column
    header = range(0)
    prefix = range(0)
    end = last_row
    cut = None
    if report:
        if report.header_rows:
            header = range(report.header_rows[0], report.header_rows[-1] + 1)
        if report.skipped_prefix:
            prefix = range(report.skipped_prefix[0], report.skipped_prefix[1] + 1)
        if report.trimmed_blank_tail:
            end = report.trimmed_blank_tail[0] - 1
        cut = next((item.row for item in report.tail_decisions if item.action in {"dropped", "suggested"}), None)

    visible = set(range(1, min(last_row, 24) + 1))
    visible.update(range(max(1, end - 11), end + 1))
    for boundary in (header.start or None, cut):
        if boundary is not None:
            visible.update(range(max(1, boundary - 2), min(last_row, boundary + 4) + 1))
    actions = {item.row: item.action for item in report.tail_decisions if item.row in visible} if report else {}
    columns = range(1, min(width, _PREVIEW_COLUMNS) + 1)
    return {
        "shape": [last_row, width],
        "columns": [get_column_letter(column) for column in columns],
        "rows": [
            {
                "number": number,
                "kind": actions.get(number) or ("header" if number in header else "skipped" if number in prefix else ""),
                "values": [_cell_text(sheet.cell(number, column).value) for column in columns],
            }
            for number in sorted(visible)
        ],
    }


def _workbook_response(
    workbook: Workbook, params: dict[str, list[str]], *, export: bool, check_client: Callable[[], None]
) -> tuple[HTTPStatus, dict | bytes]:
    sheets = workbook.worksheets
    result = {"sheets": [sheet.title for sheet in sheets]}
    sheet = None

    def number(key: str) -> int | None:
        value = params.get(key, [""])[0]
        return int(value) if value else None

    try:
        header_row, header_depth, drop_tail = (number(key) for key in ("header_row", "header_depth", "drop_tail"))
        _validate_options(header_row, header_depth, drop_tail)
        sheet_index = number("sheet") or 0
        sheet = _select_sheet(sheets, sheet_index)
        if export and drop_tail is None:
            drop_tail = 0
        check_client()
        rows, report, tail_candidates = _analyze_sheet(sheet, header_row, header_depth, drop_tail, apply_detected_tail=False)
        check_client()
        depth = len(report.header_rows)
        columns = _table_columns(rows, depth)
        if export:
            return HTTPStatus.OK, export_xlsx(columns, rows, depth, report.sheet_name, check_client=check_client)
    except HeaderDetectionError as error:
        return HTTPStatus.UNPROCESSABLE_ENTITY, {
            **result, "source": _source_data(sheet), "report": _report_data(error.report, sheet),
            "error": "无法可靠识别表头。请指定表头起始行。",
        }
    except ValueError as error:
        if sheet is not None and not export:
            result["source"] = _source_data(sheet)
        return HTTPStatus.UNPROCESSABLE_ENTITY, {**result, "error": str(error)}

    return HTTPStatus.OK, {
        **result, "sheet": sheet_index, "report": _report_data(report, sheet, tail_candidates),
        "source": _source_data(sheet, report), "shape": [len(rows) - depth, len(columns)],
        "columns": [_cell_text(name) for name in columns[:_PREVIEW_COLUMNS]],
        "rows": [
            [_cell_text(value) for value in row[:min(len(columns), _PREVIEW_COLUMNS)]]
            for row in rows[depth:depth + _PREVIEW_ROWS]
        ],
    }


class _Server(ThreadingHTTPServer):
    def __init__(self, *args, **kwargs):
        self.workbook_slots = BoundedSemaphore(_MAX_WORKBOOK_JOBS)
        super().__init__(*args, **kwargs)

    @property
    def origin(self) -> str:
        return f"http://127.0.0.1:{self.server_port}"


class _Handler(BaseHTTPRequestHandler):
    server: _Server

    def log_message(self, format: str, *args: object) -> None:
        pass

    def _send(self, status: HTTPStatus, kind: str, data: bytes, *, download: str | None = None) -> None:
        try:
            self.send_response(status)
            self.send_header("Content-Type", kind)
            self.send_header("Content-Length", str(len(data)))
            if download is not None:
                self.send_header("Content-Disposition", f'attachment; filename="{download}"')
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Content-Security-Policy", "default-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'")
            self.end_headers()
            self.wfile.write(data)
        except (ConnectionError, TimeoutError):
            pass

    def _json(self, status: HTTPStatus, payload: dict) -> None:
        self._send(status, "application/json; charset=utf-8", json.dumps(payload).encode("utf-8"))

    def do_GET(self) -> None:
        asset = _ASSETS.get(urlsplit(self.path).path)
        if not asset or self.headers.get("Host") != f"127.0.0.1:{self.server.server_port}":
            self._json(HTTPStatus.NOT_FOUND, {"error": "页面不存在。"})
            return
        filename, kind = asset
        self._send(HTTPStatus.OK, kind, files("messy_excel").joinpath("gui_assets", filename).read_bytes())

    def do_POST(self) -> None:
        url = urlsplit(self.path)
        if url.path not in {"/preview", "/export"}:
            self._json(HTTPStatus.NOT_FOUND, {"error": "接口不存在。"})
            return
        if self.headers.get("Origin") != self.server.origin or self.headers.get("Host") != f"127.0.0.1:{self.server.server_port}":
            self._json(HTTPStatus.FORBIDDEN, {"error": "请从本地工作台发起请求。"})
            return
        if not self.server.workbook_slots.acquire(blocking=False):
            self._json(HTTPStatus.SERVICE_UNAVAILABLE, {"error": "有其他工作簿正在处理，请稍后重试。"})
            return
        try:
            self.connection.settimeout(30)
            self._process_workbook(parse_qs(url.query), export=url.path == "/export")
        except (_ClientDisconnected, ConnectionError):
            pass
        finally:
            self.server.workbook_slots.release()

    def _check_client(self) -> None:
        try:
            readable, _, _ = select([self.connection], [], [], 0)
            if readable and not self.connection.recv(1, MSG_PEEK):
                raise _ClientDisconnected
        except ConnectionError:
            raise _ClientDisconnected from None

    def _process_workbook(self, params: dict[str, list[str]], *, export: bool) -> None:
        try:
            workbook = _demo_workbook() if params.get("demo") == ["1"] else self._read_workbook()
        except ConnectionError:
            raise
        except (BadZipFile, InvalidFileException, KeyError, OSError, ParseError, ValueError):
            self._json(HTTPStatus.UNPROCESSABLE_ENTITY, {"error": "无法打开工作簿。请选择可正常打开的 .xlsx 文件。"})
            return
        if workbook is None:
            return
        try:
            status, result = _workbook_response(workbook, params, export=export, check_client=self._check_client)
            if isinstance(result, bytes):
                self._send(status, _XLSX_TYPE, result, download="messy-excel-cleaned.xlsx")
            else:
                self._json(status, result)
        except OSError:
            message = "无法生成导出文件。请检查临时目录是否可写、磁盘空间是否充足。" if export else "无法完成工作簿预览。"
            self._json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": message})
        finally:
            workbook.close()

    def _read_workbook(self) -> Workbook | None:
        try:
            length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            length = 0
        if self.headers.get_content_type() != "application/octet-stream" or length <= 0:
            self._json(HTTPStatus.BAD_REQUEST, {"error": "请选择 .xlsx 文件。"})
            return None
        if length > _MAX_BYTES:
            self._json(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, {"error": "文件超过 32 MB，无法在界面中预览。"})
            return None
        upload = self.rfile.read(length)
        if len(upload) != length:
            raise _ClientDisconnected
        self._check_client()
        return load_workbook(BytesIO(upload), data_only=True, keep_links=False)


def run(*, port: int = 0, open_browser: bool = True) -> None:
    """Serve the preview on the loopback interface until interrupted."""
    with _Server(("127.0.0.1", port), _Handler) as server:
        print(f"Messy Excel: {server.origin}", flush=True)
        if open_browser:
            webbrowser.open(server.origin)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
