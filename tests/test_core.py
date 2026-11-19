# -*- coding: utf-8 -*-
"""回归测试：把 2026-11-19 修掉的 5 个缺陷钉死，改代码时先跑这个。

    python -m unittest discover -s tests -v
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402

import main as tool  # noqa: E402


def make_xlsx(path: Path, header, rows, title: str = "Sheet1") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def read_rows(path: Path):
    ws = load_workbook(path).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


class MergeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="excel-test-"))

    def test_aligns_columns_by_header_not_position(self):
        """两个文件列顺序不同，数据不能串列。"""
        make_xlsx(self.tmp / "a.xlsx", ["姓名", "金额"], [["张三", 1200]])
        make_xlsx(self.tmp / "b.xlsx", ["金额", "姓名"], [[980, "李四"]])
        res = tool.merge_files(self.tmp, self.tmp / "out.xlsx")
        rows = read_rows(self.tmp / "out.xlsx")
        self.assertEqual(rows[0], ["姓名", "金额"])
        self.assertIn(["张三", 1200], rows)
        self.assertIn(["李四", 980], rows)
        self.assertEqual(res.rows, 2)

    def test_output_file_is_excluded_from_input(self):
        """连跑两次不能把上次的结果再合并一遍。"""
        make_xlsx(self.tmp / "src.xlsx", ["姓名", "金额"], [["张三", 1], ["李四", 2]])
        tool.merge_files(self.tmp, self.tmp / "合并结果.xlsx")
        second = tool.merge_files(self.tmp, self.tmp / "合并结果.xlsx")
        self.assertEqual(second.rows, 2)

    def test_extra_column_extends_output(self):
        make_xlsx(self.tmp / "a.xlsx", ["姓名", "金额"], [["张三", 1]])
        make_xlsx(self.tmp / "b.xlsx", ["姓名", "金额", "门店"], [["李四", 2, "A店"]])
        tool.merge_files(self.tmp, self.tmp / "out.xlsx")
        rows = read_rows(self.tmp / "out.xlsx")
        self.assertEqual(rows[0], ["姓名", "金额", "门店"])
        self.assertIn(["李四", 2, "A店"], rows)

    def test_skips_non_tabular_sheet(self):
        """说明页/封面页不能被当作数据灌进总表。"""
        wb = Workbook()
        ws = wb.active
        ws.title = "数据"
        ws.append(["姓名", "金额"])
        ws.append(["张三", 1])
        note = wb.create_sheet("说明")
        note.append(["本表由系统导出"])
        note.append(["如有疑问联系财务"])
        wb.save(self.tmp / "a.xlsx")
        make_xlsx(self.tmp / "b.xlsx", ["姓名", "金额"], [["李四", 2]])

        res = tool.merge_files(self.tmp, self.tmp / "out.xlsx")
        self.assertEqual(res.rows, 2)
        self.assertIn("a.xlsx[说明]", res.skipped_sheets)


class SplitTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="excel-test-"))
        self.src = self.tmp / "src.xlsx"
        make_xlsx(self.src, ["部门", "姓名"],
                  [["A/B", "张三"], ["A_B", "李四"], [1.0, "王五"], [1, "赵六"]])

    def test_colliding_keys_do_not_overwrite(self):
        out = self.tmp / "out"
        result = tool.split_by_column(self.src, "部门", out)
        files = list(out.glob("*.xlsx"))
        self.assertEqual(len(files), len(result))
        self.assertEqual(len(files), 3)          # A/B、A_B、1

    def test_numeric_key_normalized(self):
        """1 与 1.0 属于同一组，不能拆成两个文件。"""
        result = tool.split_by_column(self.src, "部门", self.tmp / "out")
        self.assertIn("1", result)
        self.assertEqual(result["1"], 2)


class CsvTests(unittest.TestCase):
    def test_exports_every_sheet(self):
        tmp = Path(tempfile.mkdtemp(prefix="excel-test-"))
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "第一页"
        ws1.append(["a"])
        ws1.append([1])
        ws2 = wb.create_sheet("第二页")
        ws2.append(["b"])
        ws2.append([2])
        wb.save(tmp / "book.xlsx")

        n = tool.to_csv_files(tmp, tmp / "csv")
        self.assertEqual(n, 2)
        self.assertTrue((tmp / "csv" / "book.csv").exists())
        self.assertTrue((tmp / "csv" / "book__第二页.csv").exists())


if __name__ == "__main__":
    unittest.main()
