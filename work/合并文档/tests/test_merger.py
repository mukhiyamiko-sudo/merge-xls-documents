from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import xlrd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from merger import discover_files, merge_files, read_source


HEADER = [
    "CODE",
    "DESCRIPTION",
    "ITEM #",
    "ITEM DESCRIPTION",
    "UPC",
    "MODEL&nbsp;#",
    "UNIT COST",
    "CLAIM QTY",
    "UOM",
    "CLAIM AMT",
]


def write_html_xls(
    path: Path,
    rows: list[list[str]],
    *,
    bad_header: bool = False,
    missing_footer: bool = False,
) -> None:
    header = HEADER.copy()
    if bad_header:
        header[-1] = "OTHER"
    body = ["<table><tbody>"]
    body.append("<tr>" + "".join(f"<td>{value}</td>" for value in header) + "</tr>")
    for row in rows:
        body.append("<tr>" + "".join(f"<td>{value}</td>" for value in row) + "</tr>")
    body.append("<tr><td colspan='10'>&nbsp;</td></tr>")
    footers = [
        ("TOTAL QUANTITY", "1"),
        ("TOTAL FOR CLAIM # 123", "-1.00"),
        ("TOTAL FREIGHT FEE", "0.00"),
        ("TOTAL HANDLING COST", "0.00"),
        ("TOTAL DISPOSAL FEE", "0.00"),
    ]
    if missing_footer:
        footers.pop()
    for label, value in footers:
        body.append(
            f"<tr><td colspan='5'>&nbsp;</td><td colspan='3'>{label}</td>"
            f"<td colspan='2'>{value}</td></tr>"
        )
    body.append("</tbody></table>")
    path.write_text("\n".join(body), encoding="utf-8")


class MergerTests(unittest.TestCase):
    def test_merges_in_filename_order_and_removes_footers(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            write_html_xls(
                folder / "claim_002.xls",
                [["94", "B", "2", "ITEM B", "0002", "M2", "-2.5", "2", "EA", "-5"]],
            )
            write_html_xls(
                folder / "claim_001.xls",
                [["94", "A", "1", "ITEM A", "0001", "M1", "-1", "1", "EA", "-1"]],
            )

            files, generated = discover_files(folder)
            self.assertEqual([], generated)
            result = merge_files(files, folder)

            self.assertEqual(2, result.merged_files)
            self.assertEqual(2, result.data_rows)
            book = xlrd.open_workbook(str(result.output_path))
            sheet = book.sheet_by_index(0)
            self.assertEqual(3, sheet.nrows)
            self.assertEqual("A", sheet.cell_value(1, 1))
            self.assertEqual("B", sheet.cell_value(2, 1))
            self.assertNotIn(
                "TOTAL",
                " ".join(
                    str(sheet.cell_value(row, column))
                    for row in range(sheet.nrows)
                    for column in range(sheet.ncols)
                ),
            )

    def test_bad_header_and_incomplete_footer_are_skipped(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            valid = folder / "a.xls"
            bad_header = folder / "b.xls"
            bad_footer = folder / "c.xls"
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            write_html_xls(valid, [row])
            write_html_xls(bad_header, [row], bad_header=True)
            write_html_xls(bad_footer, [row], missing_footer=True)

            result = merge_files([valid, bad_header, bad_footer], folder)

            self.assertEqual(1, result.merged_files)
            self.assertEqual(2, result.skipped_files)
            self.assertTrue(any("表头" in issue.reason for issue in result.issues))
            self.assertTrue(any("五项汇总" in issue.reason for issue in result.issues))

    def test_generated_outputs_are_excluded(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            write_html_xls(folder / "source.xls", [row])
            (folder / "合并结果_20261009_120000.xls").write_bytes(b"old")

            files, generated = discover_files(folder)

            self.assertEqual(["source.xls"], [path.name for path in files])
            self.assertEqual(["合并结果_20261009_120000.xls"], [path.name for path in generated])

    def test_reads_nested_table_shape(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "nested.xls"
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            write_html_xls(path, [row])
            inner = path.read_text(encoding="utf-8")
            path.write_text(f"<table><tr><td>{inner}</td></tr></table>", encoding="utf-8")

            source = read_source(path)

            self.assertEqual(1, len(source.rows))
            self.assertEqual("A", source.rows[0][1])


if __name__ == "__main__":
    unittest.main()

