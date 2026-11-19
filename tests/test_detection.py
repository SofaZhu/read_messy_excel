from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

from openpyxl import Workbook

from messy_excel import HeaderDetectionError, flatten
from messy_excel.cli import main


class DetectionTests(TestCase):
    def setUp(self):
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name) / "example.xlsx"

    def make_workbook(self, rows, *, merges=(), blank_tail=None):
        workbook = Workbook()
        self.addCleanup(workbook.close)
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        for area in merges:
            sheet.merge_cells(area)
        if blank_tail:
            sheet[blank_tail].number_format = "0.00"
        workbook.save(self.path)

    def test_auto_selects_merged_parent_and_ignores_merged_title(self):
        self.make_workbook((
            ["销售晨报"],
            ["区域", "销售"],
            ["地区", "数量", "金额"],
            ["华东", 1, 20],
            ["华南", 2, 40],
            ["华北", 3, 60],
            ["华西", 1, 10],
        ), merges=("A1:C1", "B2:C2"), blank_tail="H9")

        frame, report = flatten(self.path)

        self.assertEqual(frame.columns.tolist(), ["区域 / 地区", "销售 / 数量", "销售 / 金额"])
        self.assertEqual(frame["销售 / 数量"].tolist(), [1, 2, 3, 1])
        self.assertEqual(report.header_source, "auto")
        self.assertEqual(report.header_rows, (2, 3))
        self.assertEqual(report.skipped_prefix, (1, 1))
        self.assertEqual(report.trimmed_blank_tail, (8, 9))
        self.assertTrue(report.candidates[0].layer_evidence)
        self.assertEqual(report.candidates[0].sample_rows, (4, 5, 6, 7))

    def test_auto_finds_single_header_and_only_drops_contiguous_total_suffix(self):
        self.make_workbook((
            ["销售周报"],
            ["品类", "经办人", "数量"],
            ["西瓜", "陈宁", 2],
            ["小计", None, 2],
            ["橘子", "王明", 3],
            ["苹果", "林晓", 4],
            ["小计", None, 7],
            [" GRAND TOTAL ", "李芳", 1],
            [" 小计 ", None, 8],
            [" ToTAL ", None, 10],
        ))

        frame, report = flatten(self.path)

        self.assertEqual(report.header_rows, (2,))
        self.assertEqual(report.dropped_tail_rows, (9, 10))
        self.assertEqual([item.row for item in report.tail_decisions], [9, 10])
        self.assertEqual(frame["品类"].tolist(), ["西瓜", "小计", "橘子", "苹果", "小计", " GRAND TOTAL "])
        all_rows, report = flatten(self.path, header_row=2, drop_tail=0)
        self.assertEqual(len(all_rows), 8)
        self.assertEqual(report.dropped_tail_rows, ())

    def test_tail_preserves_complete_detail_named_total_and_text_lookalikes(self):
        for value in ("合计", "小计说明", "Total Care"):
            with self.subTest(value=value):
                self.make_workbook((
                    ["地区", "负责人", "数量"],
                    ["华南", "刘欣", 2],
                    ["华东", "杨明", 3],
                    ["华北", "高远", 4],
                    [value, "刘欣", 5],
                ))
                frame, report = flatten(self.path, header_row=1)
                self.assertEqual(frame["地区"].iloc[-1], value)
                self.assertEqual(report.dropped_tail_rows, ())
                if value == "合计":
                    self.assertEqual(report.tail_decisions[0].action, "kept")

    def test_tail_refuses_to_cut_a_merge_or_use_a_copied_summary_label(self):
        detail = (
            ["地区", "负责人", "数量", "备注"],
            ["华南", "刘欣", 2, "审核"],
            ["华东", "杨明", 3, None],
            ["华北", "高远", 4, None],
        )
        for suffix, area, reason in (
            ((["合计", None, 9, None],), "D4:D5", "crosses"),
            ((["合计", None, 9, None], [None, None, 9, None]), "A5:A6", "filled"),
        ):
            with self.subTest(area=area):
                self.make_workbook(detail + suffix, merges=(area,))
                frame, report = flatten(self.path, header_row=1)
                self.assertEqual(len(frame), 3 + len(suffix))
                self.assertEqual(report.dropped_tail_rows, ())
                self.assertIn(reason, report.tail_decisions[-1].reason)

    def test_ambiguous_short_text_numeric_and_nested_layouts_can_be_overridden(self):
        for rows, start, depth, cause in (
            ((["公司", "地区"], ["甲", "南"], ["乙", "北"], ["丙", "东"]), 1, 1, "type transition"),
            ((["名称", "数量"], ["甲", 1], ["乙", 2]), 1, 1, "three data rows"),
            (([1, 10], [2, 20], [3, 30], [4, 40]), 1, 1, "type transition"),
            ((["公司", "销售"], ["名称", "数量"], ["甲", 1], ["乙", 2], ["丙", 3]), 1, 2, "similar scores"),
        ):
            with self.subTest(cause=cause):
                self.make_workbook(rows)
                with self.assertRaises(HeaderDetectionError) as raised:
                    flatten(self.path)
                report = raised.exception.report
                self.assertEqual(report.header_rows, ())
                self.assertTrue(report.candidates)
                self.assertIn(cause, report.warnings[0])
                frame, _ = flatten(self.path, header_row=start, header_depth=depth)
                self.assertEqual(len(frame), len(rows) - depth)

    def test_cli_reports_identical_coordinates_and_omits_cell_values(self):
        self.make_workbook((
            ["项目", "金额"], ["private-detail-1", 1], ["p2", 2], ["p3", 3],
        ))
        original = self.path.read_bytes()
        _, report = flatten(self.path)
        out, err = StringIO(), StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            result = main(["inspect", str(self.path)])
        self.assertEqual(result, 0)
        self.assertIn(f"Header rows (auto): {report.header_rows[0]}", out.getvalue())
        self.assertIn(f"score {report.candidates[0].score:.2f}", out.getvalue())
        self.assertNotIn("private-detail-1", out.getvalue())
        self.assertEqual(err.getvalue(), "")
        self.assertEqual(self.path.read_bytes(), original)

    def test_cli_returns_nonzero_with_diagnostics_for_ambiguous_and_invalid_files(self):
        self.make_workbook((["名称", "数量"], ["甲", 1]))
        for target, expected in ((self.path, "Could not detect a header"), (self.path.with_name("missing.xlsx"), "missing.xlsx")):
            with self.subTest(target=target):
                out, err = StringIO(), StringIO()
                with redirect_stdout(out), redirect_stderr(err):
                    self.assertEqual(main(["inspect", str(target)]), 1)
                self.assertIn(expected, err.getvalue())
                if target == self.path:
                    self.assertIn("--header-row 1 --header-depth 1", out.getvalue())
