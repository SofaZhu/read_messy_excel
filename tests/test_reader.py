from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

import pandas as pd
from openpyxl import Workbook

from messy_excel import flatten


class FlattenTests(TestCase):
    def setUp(self):
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name) / "sample.xlsx"
        self.workbook = Workbook()
        self.addCleanup(self.workbook.close)
        self.sheet = self.workbook.active

    def test_report_with_title_two_header_rows_and_vertical_merge(self):
        self.sheet.title = "销售"
        for row in (
            ["销售周报"],
            ["地区", "销售"],
            [None, "数量", "金额"],
            ["华南", 2, 200],
            [None, 3, 300],
            ["华东", None, 100],
            ["合计", 5, 600],
        ):
            self.sheet.append(row)
        for area in ("A1:C1", "A2:A3", "B2:C2", "A4:A5"):
            self.sheet.merge_cells(area)
        self.sheet["J8"].number_format = "0.00"
        self.workbook.save(self.path)
        original = self.path.read_bytes()

        frame, report = flatten(self.path, sheet_name="销售", header_row=2, header_depth=2, drop_tail=1)

        self.assertEqual(frame.columns.tolist(), ["地区", "销售 / 数量", "销售 / 金额"])
        self.assertEqual(frame["地区"].tolist(), ["华南", "华南", "华东"])
        self.assertTrue(pd.isna(frame.iloc[2, 1]))
        self.assertEqual(report.sheet_name, "销售")
        self.assertEqual(report.header_rows, (2, 3))
        self.assertEqual(report.filled_merge_ranges, ("A2:A3", "B2:C2", "A4:A5"))
        self.assertEqual(report.dropped_tail_rows, (7,))
        self.assertEqual(self.path.read_bytes(), original)

    def test_names_are_unique_and_keep_real_suffixes(self):
        self.sheet.append(["名称", "名称", "名称__2", "名称", None])
        self.sheet.append(["001", "NA", "0003", "小计说明", "note"])
        self.workbook.save(self.path)

        frame, _ = flatten(self.path, header_row=1)

        self.assertEqual(frame.columns.tolist(), ["名称", "名称__3", "名称__2", "名称__4", "column_E"])
        self.assertEqual(frame.iloc[0].tolist(), ["001", "NA", "0003", "小计说明", "note"])

    def test_horizontal_body_merge_and_interior_blank_row_are_preserved(self):
        self.sheet.append(["品类", "数目"])
        self.sheet.append(["汇总说明"])
        self.sheet.append([])
        self.sheet.append(["产品", 3])
        self.sheet.merge_cells("A2:B2")
        self.workbook.save(self.path)

        frame, report = flatten(self.path, header_row=1)

        self.assertEqual(frame.shape, (3, 2))
        self.assertTrue(pd.isna(frame.iloc[0, 1]))
        self.assertTrue(frame.iloc[1].isna().all())
        self.assertEqual(report.skipped_merge_ranges, ("A2:B2",))

    def test_merge_crossing_header_boundary_is_rejected(self):
        self.sheet.append(["品类", "数目"])
        self.sheet.append([None, 1])
        self.sheet.merge_cells("A1:A2")
        self.workbook.save(self.path)

        with self.assertRaisesRegex(ValueError, "A1:A2 crosses a header boundary"):
            flatten(self.path, header_row=1)
        with self.assertRaisesRegex(ValueError, "A1:A2 crosses a header boundary"):
            flatten(self.path, header_row=2)

    def test_sheet_selection_tail_override_and_validation(self):
        other = self.workbook.create_sheet("details")
        other.append(["条目"])
        other["J2"] = "尾部说明"
        self.workbook.save(self.path)

        frame, report = flatten(self.path, sheet_name=1, header_row=1, drop_tail=1)
        self.assertEqual(frame.shape, (0, 1))
        self.assertEqual(report.dropped_tail_rows, (2,))

        for override, message in (
            ({"header_row": 0}, "header_row"),
            ({"header_depth": True}, "header_depth"),
            ({"header_depth": 3}, "extends beyond"),
            ({"sheet_name": -1}, "Unknown worksheet"),
            ({"sheet_name": "missing"}, "Unknown worksheet"),
            ({"drop_tail": 2}, "drop_tail exceeds"),
            ({"sheet_name": 0}, "no labels"),
        ):
            with self.subTest(override=override), self.assertRaisesRegex(ValueError, message):
                flatten(self.path, **({"header_row": 1, "sheet_name": 1} | override))

    def test_oversized_dimension_is_rejected_before_grid_expansion(self):
        self.sheet.append(["品类"])
        self.sheet["XFD1000"].number_format = "0.00"
        self.workbook.save(self.path)

        with self.assertRaisesRegex(ValueError, "1,000,000-cell limit"):
            flatten(self.path, header_row=1)
