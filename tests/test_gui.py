import json
from datetime import datetime
from functools import partial
from http.client import HTTPConnection
from io import BytesIO
from pathlib import Path
from socket import SHUT_RDWR
from tempfile import TemporaryDirectory
from threading import Event, Thread
from unittest import TestCase
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

from messy_excel import gui
from messy_excel.gui import _ClientDisconnected, _Handler, _MAX_WORKBOOK_JOBS, _Server, _XLSX_TYPE
from messy_excel.xlsx_export import export_xlsx


class WorkbenchTests(TestCase):
    def setUp(self):
        self.server = _Server(("127.0.0.1", 0), _Handler)
        thread = Thread(target=self.server.serve_forever, kwargs={"poll_interval": 0.01}, daemon=True)
        thread.start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(thread.join)
        self.addCleanup(self.server.shutdown)

    def post(self, path, workbook=None, *, origin=None):
        connection = HTTPConnection("127.0.0.1", self.server.server_port, timeout=5)
        try:
            connection.request("POST", path, body=workbook, headers={
                "Origin": origin or self.server.origin,
                "Content-Type": "application/octet-stream",
            })
            response = connection.getresponse()
            body = response.read()
            if response.getheader("Content-Type") == _XLSX_TYPE:
                self.assertIn("attachment", response.getheader("Content-Disposition"))
                return response.status, load_workbook(BytesIO(body), data_only=False)
            return response.status, json.loads(body)
        finally:
            connection.close()

    def source_workbook(self, rows):
        book = Workbook()
        book.active.title = "订单"
        for row in rows:
            book.active.append(row)
        return book

    def upload_bytes(self, book):
        with BytesIO() as output:
            book.save(output)
            return output.getvalue()

    def test_demo_tail_is_suggested_until_confirmed_and_can_be_restored(self):
        status, preview = self.post("/preview?demo=1")
        self.assertEqual(status, 200)
        self.assertEqual(preview["shape"], [13, 8])
        self.assertEqual(preview["report"]["suggested_tail_count"], 1)
        self.assertEqual(preview["report"]["dropped_tail_rows"], [])
        source = {row["number"]: row for row in preview["source"]["rows"]}
        self.assertEqual(source[1]["kind"], "skipped")
        self.assertEqual(source[2]["kind"], "header")
        self.assertEqual(source[16]["kind"], "suggested")
        self.assertIsNone(source[5]["values"][0])
        self.assertEqual(preview["rows"][1][0], "华东")

        for tail_query, count, final in (("", 13, "合计"), ("&drop_tail=1", 12, "西南"), ("", 13, "合计")):
            with self.subTest(tail_query=tail_query):
                status, preview = self.post(f"/preview?demo=1{tail_query}")
                self.assertEqual(status, 200)
                self.assertEqual(preview["shape"][0], count)
                if tail_query:
                    self.assertEqual(preview["source"]["rows"][-1]["kind"], "dropped")
                    self.assertEqual(preview["report"]["dropped_tail_rows"], [16])
                status, export = self.post(f"/export?demo=1{tail_query}")
                self.assertEqual(status, 200)
                self.addCleanup(export.close)
                self.assertEqual(export.sheetnames, ["销售明细"])
                self.assertEqual(export.active.max_row, count + 1)
                self.assertEqual(export.active.cell(count + 1, 1).value, final)
                self.assertEqual(export.active["D1"].value, "销售数据 / 数量")

    def test_default_review_does_not_drop_detail_named_total(self):
        gap = (None,) * 28
        workbook = self.source_workbook((
            ("Item", "Owner", *gap, "Amount"), ("Pears", "Lin", *gap, 10), ("Apples", "Chen", *gap, 20),
            ("Oranges", "Wang", *gap, 30), ("Total", None, *gap, 7),
        ))
        upload = self.upload_bytes(workbook)
        status, preview = self.post("/preview", upload)
        self.assertEqual(status, 200)
        self.assertEqual(preview["rows"][-1][:2], ["Total", None])
        self.assertEqual(preview["source"]["columns"][-1], "AD")
        decision, = preview["report"]["tail_decisions"]
        self.assertEqual(decision["action"], "suggested")
        self.assertEqual(decision["evidence"], [
            {"coordinate": "A5", "value": "Total", "filled": False},
            {"coordinate": "AE5", "value": "7", "filled": False},
            {"coordinate": "B5", "value": None, "filled": False},
        ])
        status, export = self.post("/export", upload)
        self.assertEqual(status, 200)
        self.addCleanup(export.close)
        self.assertEqual(export.active["A5"].value, "Total")
        self.assertEqual(export.active["AE5"].value, 7)

    def test_export_includes_unpreviewed_cells_dates_and_literal_text(self):
        columns = ["=订单", "已发货", "日期", "备注", *[f"指标{i}" for i in range(5, 34)]]
        workbook = self.source_workbook([columns])
        sheet = workbook.active
        for number in range(1, 66):
            sheet.append([f"{number:05}", True, datetime(2026, 11, 20), None, *range(5, 34)])
        sheet["A1"].data_type = "s"
        sheet["D2"] = "=1+1"
        sheet["D2"].data_type = "s"
        sheet["D3"] = "#N/A"
        sheet["D3"].data_type = "s"
        sheet["D66"] = "末行" * 90
        sheet["AG66"] = 987
        upload = self.upload_bytes(workbook)

        status, preview = self.post("/preview?header_row=1", upload)
        self.assertEqual(status, 200)
        self.assertEqual(preview["shape"], [65, 33])
        self.assertEqual((len(preview["rows"]), len(preview["columns"])), (50, 30))
        source_rows = preview["source"]["rows"]
        self.assertEqual(source_rows[-1]["number"], 66)
        self.assertTrue(source_rows[-1]["values"][3].endswith("…"))
        self.assertLessEqual(len(source_rows), 50)

        status, export = self.post("/export?header_row=1", upload)
        self.assertEqual(status, 200)
        self.addCleanup(export.close)
        sheet = export.active
        self.assertEqual((sheet.max_row, sheet.max_column), (66, 33))
        self.assertEqual((sheet["A1"].value, sheet["A1"].data_type), ("=订单", "s"))
        self.assertEqual(sheet["A2"].value, "00001")
        self.assertIs(sheet["B2"].value, True)
        self.assertEqual(sheet["C2"].value, datetime(2026, 11, 20))
        self.assertEqual((sheet["D2"].value, sheet["D2"].data_type), ("=1+1", "s"))
        self.assertEqual((sheet["D3"].value, sheet["D3"].data_type), ("#N/A", "s"))
        self.assertEqual(sheet["D66"].value, "末行" * 90)
        self.assertEqual(sheet["AG66"].value, 987)

    def test_manual_cut_preserves_preceding_blank_row(self):
        upload = self.upload_bytes(self.source_workbook((("列名",), ("明细",), (None,), ("尾行",))))
        status, export = self.post("/export?header_row=1&drop_tail=1", upload)
        self.assertEqual(status, 200)
        self.addCleanup(export.close)
        self.assertEqual(list(export.active.values), [("列名",), ("明细",), (None,)])

    def test_source_help_is_available_when_header_is_uncertain(self):
        status, preview = self.post("/preview?demo=1&sheet=1")
        self.assertEqual(status, 422)
        self.assertEqual(preview["source"]["rows"][0]["values"], ["部门", "对接人"])
        self.assertTrue(preview["report"]["candidates"])
        status, failed = self.post("/export?demo=1&sheet=1")
        self.assertEqual(status, 422)
        self.assertIn("表头", failed["error"])
        status, export = self.post("/export?demo=1&sheet=1&header_row=1")
        self.assertEqual(status, 200)
        self.addCleanup(export.close)
        self.assertEqual(export.sheetnames, ["纯文本示例"])
        self.assertEqual(export.active.max_row, 4)

    def test_export_keeps_the_loopback_origin_check(self):
        status, _ = self.post("/export?demo=1", origin="https://example.org")
        self.assertEqual(status, 403)

    def test_export_write_failure_cleans_scratch_and_returns_readable_error(self):
        with TemporaryDirectory() as scratch:
            with patch("messy_excel.xlsx_export.TemporaryDirectory", partial(TemporaryDirectory, dir=scratch)):
                with patch("messy_excel.xlsx_export.ZipFile.write", side_effect=OSError("simulated archive failure")):
                    status, failure = self.post("/export?demo=1")
                self.assertEqual(status, 500)
                self.assertIn("临时目录", failure["error"])
                self.assertEqual(list(Path(scratch).iterdir()), [])

                status, export = self.post("/export?demo=1")
                self.assertEqual(status, 200)
                self.addCleanup(export.close)
                self.assertEqual(list(Path(scratch).iterdir()), [])

    def test_client_disconnect_during_parse_skips_xlsx_generation(self):
        analyzing, resume, finished = Event(), Event(), Event()
        analyze_sheet = gui._analyze_sheet
        workbook_response = gui._workbook_response

        def pause_analysis(*args, **kwargs):
            analyzing.set()
            if not resume.wait(5):
                raise AssertionError("test did not release workbook analysis")
            return analyze_sheet(*args, **kwargs)

        def observe_response(*args, **kwargs):
            try:
                return workbook_response(*args, **kwargs)
            finally:
                finished.set()

        with (
            patch.object(gui, "_analyze_sheet", side_effect=pause_analysis),
            patch.object(gui, "_workbook_response", side_effect=observe_response),
            patch.object(gui, "export_xlsx") as generate,
        ):
            connection = HTTPConnection("127.0.0.1", self.server.server_port, timeout=5)
            try:
                connection.request("POST", "/export?demo=1", headers={"Origin": self.server.origin})
                self.assertTrue(analyzing.wait(5))
                connection.sock.shutdown(SHUT_RDWR)
            finally:
                connection.close()
                resume.set()
            self.assertTrue(finished.wait(5))
            generate.assert_not_called()

    def test_cancelled_export_cleans_partially_written_sheet(self):
        with TemporaryDirectory() as scratch:
            def check_client():
                if any(Path(scratch).iterdir()):
                    raise _ClientDisconnected

            rows = [["Item"], *[[f"Item {number}"] for number in range(300)]]
            with patch("messy_excel.xlsx_export.TemporaryDirectory", partial(TemporaryDirectory, dir=scratch)):
                with self.assertRaises(_ClientDisconnected):
                    export_xlsx(["Item"], rows, 1, "Orders", check_client=check_client)
            self.assertEqual(list(Path(scratch).iterdir()), [])

    def test_windows_abort_while_writing_response_is_handled(self):
        handler = _Handler.__new__(_Handler)
        handler.send_response = Mock()
        handler.send_header = Mock()
        handler.end_headers = Mock()
        handler.wfile = Mock()
        handler.wfile.write.side_effect = ConnectionAbortedError("simulated Windows socket close")

        handler._send(200, _XLSX_TYPE, b"content")

        handler.wfile.write.assert_called_once_with(b"content")

    def test_workbook_jobs_are_bounded(self):
        slots = self.server.workbook_slots
        for _ in range(_MAX_WORKBOOK_JOBS):
            self.assertTrue(slots.acquire(blocking=False))
        try:
            with patch.object(gui, "_demo_workbook") as open_workbook:
                status, response = self.post("/preview?demo=1")
                self.assertEqual(status, 503)
                self.assertIn("重试", response["error"])
                open_workbook.assert_not_called()
        finally:
            for _ in range(_MAX_WORKBOOK_JOBS):
                slots.release()

        status, _ = self.post("/preview?demo=1")
        self.assertEqual(status, 200)
