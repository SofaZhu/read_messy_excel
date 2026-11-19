from __future__ import annotations

import json
import webbrowser
from dataclasses import asdict, replace
from datetime import date
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib.resources import files
from io import BytesIO
from urllib.parse import parse_qs, urlsplit
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .reader import FlattenReport, HeaderDetectionError, _analyze_sheet, _select_sheet, _to_frame, _validate_options

_ASSETS = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/app.css": ("app.css", "text/css; charset=utf-8"),
    "/app.js": ("app.js", "text/javascript; charset=utf-8"),
}
_MAX_BYTES = 32 * 1024 * 1024
_PREVIEW_ROWS, _PREVIEW_COLUMNS = 50, 30


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
    if isinstance(value, pd.Timestamp) and value == value.normalize():
        return value.date().isoformat()
    text = str(value)
    return text if len(text) <= 160 else text[:159] + "…"


def _report_data(report: FlattenReport) -> dict:
    fields = ("filled_merge_ranges", "skipped_merge_ranges", "dropped_tail_rows", "tail_decisions")
    values = {name: getattr(report, name) for name in fields}
    details = asdict(replace(report, **{name: items[:30] for name, items in values.items()}))
    details.update({f"{name}_count": len(items) for name, items in values.items()})
    return details


def _send_preview(workbook: Workbook, params: dict[str, list[str]]) -> tuple[HTTPStatus, dict]:
    sheets = workbook.worksheets
    result = {"sheets": [sheet.title for sheet in sheets]}

    def number(key: str) -> int | None:
        value = params.get(key, [""])[0]
        return int(value) if value else None

    try:
        header_row, header_depth, drop_tail = (number(key) for key in ("header_row", "header_depth", "drop_tail"))
        _validate_options(header_row, header_depth, drop_tail)
        sheet_index = number("sheet") or 0
        sheet = _select_sheet(sheets, sheet_index)
        rows, report = _analyze_sheet(sheet, header_row, header_depth, drop_tail)
        frame = _to_frame(rows, report)
    except HeaderDetectionError as error:
        return HTTPStatus.UNPROCESSABLE_ENTITY, {
            **result, "report": _report_data(error.report), "error": "无法可靠识别表头。请指定表头起始行。",
        }
    except ValueError as error:
        return HTTPStatus.UNPROCESSABLE_ENTITY, {**result, "error": str(error)}

    preview = frame.iloc[:_PREVIEW_ROWS, :_PREVIEW_COLUMNS]
    return HTTPStatus.OK, {
        **result, "sheet": sheet_index, "report": _report_data(report),
        "shape": list(frame.shape), "columns": [_cell_text(value) for value in preview.columns],
        "rows": [[_cell_text(value) for value in row] for row in preview.itertuples(index=False, name=None)],
    }


class _Server(ThreadingHTTPServer):
    @property
    def origin(self) -> str:
        return f"http://127.0.0.1:{self.server_port}"


class _Handler(BaseHTTPRequestHandler):
    server: _Server

    def log_message(self, format: str, *args: object) -> None:
        pass

    def _send(self, status: HTTPStatus, kind: str, data: bytes) -> None:
        self.send_response(status)
        self.send_header("Content-Type", kind)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", "default-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'")
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
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
        if url.path != "/preview":
            self._json(HTTPStatus.NOT_FOUND, {"error": "接口不存在。"})
            return
        if self.headers.get("Origin") != self.server.origin or self.headers.get("Host") != f"127.0.0.1:{self.server.server_port}":
            self._json(HTTPStatus.FORBIDDEN, {"error": "请从本地工作台发起请求。"})
            return
        params = parse_qs(url.query)
        try:
            if params.get("demo") == ["1"]:
                workbook = _demo_workbook()
            else:
                workbook = self._read_workbook()
                if workbook is None:
                    return
        except (BadZipFile, InvalidFileException, KeyError, OSError, ParseError, ValueError):
            self._json(HTTPStatus.UNPROCESSABLE_ENTITY, {"error": "无法打开工作簿。请选择可正常打开的 .xlsx 文件。"})
            return
        try:
            status, result = _send_preview(workbook, params)
            self._json(status, result)
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
        self.connection.settimeout(30)
        return load_workbook(BytesIO(self.rfile.read(length)), data_only=True, keep_links=False)


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
