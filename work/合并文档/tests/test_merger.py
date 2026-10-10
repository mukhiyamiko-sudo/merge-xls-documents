from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import xlrd
from tkinterdnd2 import TkinterDnD

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import MergeApp
from merger import MergeResult, discover_files, merge_files, read_source


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
                folder / "claim_details_000005301102474.xls",
                [["94", "B", "2", "ITEM B", "0002", "M2", "-2.5", "2", "EA", "-5"]],
            )
            write_html_xls(
                folder / "claim_details_000000400699589.xls",
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
            self.assertEqual(11, sheet.ncols)
            self.assertEqual("Invoice number", sheet.cell_value(0, 0))
            self.assertEqual("0400699589", sheet.cell_value(1, 0))
            self.assertEqual("5301102474", sheet.cell_value(2, 0))
            self.assertEqual(["CODE", "DESCRIPTION", "ITEM #", "ITEM DESCRIPTION", "UPC"], sheet.row_values(0)[1:6])
            self.assertEqual([94, "A", 1, "ITEM A", 1, "M1", -1, 1, "EA", -1], sheet.row_values(1)[1:])
            self.assertEqual("B", sheet.cell_value(2, 2))
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
            valid = folder / "claim_details_000001111111111.xls"
            bad_header = folder / "claim_details_000002222222222.xls"
            bad_footer = folder / "claim_details_000003333333333.xls"
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            write_html_xls(valid, [row])
            write_html_xls(bad_header, [row], bad_header=True)
            write_html_xls(bad_footer, [row], missing_footer=True)

            result = merge_files([valid, bad_header, bad_footer], folder)

            self.assertEqual(1, result.merged_files)
            self.assertEqual(2, result.skipped_files)
            self.assertTrue(any("表头" in issue.reason for issue in result.issues))
            self.assertTrue(any("五项汇总" in issue.reason for issue in result.issues))

    def test_invoice_number_fills_every_detail_row_from_its_source_filename(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            six = folder / "claim_details_000005301106261.xls"
            two = folder / "claim_details_000005301111972.xls"
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            write_html_xls(six, [row] * 6)
            write_html_xls(two, [row] * 2)

            result = merge_files([two, six], folder)

            sheet = xlrd.open_workbook(str(result.output_path)).sheet_by_index(0)
            self.assertEqual(8, result.data_rows)
            self.assertEqual(["5301106261"] * 6 + ["5301111972"] * 2, sheet.col_values(0)[1:])

    def test_filename_without_five_zero_prefix_is_skipped_and_logged(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            valid = folder / "claim_details_000000400699589.xls"
            invalid = folder / "claim_details_000040400699589.xls"
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            write_html_xls(valid, [row])
            write_html_xls(invalid, [row])
            messages: list[str] = []

            result = merge_files([invalid, valid], folder, log=messages.append)

            self.assertEqual(1, result.merged_files)
            self.assertEqual(1, result.skipped_files)
            self.assertEqual(invalid, result.issues[0].path)
            self.assertIn("五个 0", result.issues[0].reason)
            self.assertTrue(any("跳过" in message and invalid.name in message for message in messages))

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


class AppInteractionTests(unittest.TestCase):
    def test_dragged_old_output_is_excluded_before_merge(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            source = folder / "claim_details_000005301111972.xls"
            old_output = folder / "合并结果_20261010_100000.xls"
            source.write_bytes(b"sample")
            old_output.write_bytes(b"old")

            root = TkinterDnD.Tk()
            root.withdraw()
            try:
                app = MergeApp(root)
                root.after = Mock()
                with patch("app.messagebox.showwarning"):
                    app._on_drop(SimpleNamespace(data=f"{{{source}}} {{{old_output}}}"))

                self.assertEqual("0 / 1", app.count_var.get())
                self.assertIn("1 个历史合并结果", app.log_text.get("1.0", "end"))
                with patch("app.messagebox.askyesno", return_value=True) as confirm:
                    with patch("app.threading.Thread") as worker:
                        app._start_merge()
                confirm.assert_called_once()
                self.assertEqual([source], worker.call_args.kwargs["args"][0])
            finally:
                root.destroy()

    def test_dragging_multiple_xls_scans_only_dropped_files_until_start(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xls drag ") as temp:
            folder = Path(temp)
            row = ["94", "A", "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1"]
            first = folder / "claim_details_000000400699589.xls"
            second = folder / "claim_details_000005301102474.xls"
            unselected = folder / "claim_details_000005301106261.xls"
            for path in (first, second, unselected):
                write_html_xls(path, [row])

            root = TkinterDnD.Tk()
            root.withdraw()
            try:
                app = MergeApp(root)
                root.after = Mock()
                event = SimpleNamespace(data=f"{{{second}}} {{{first}}}")
                with patch("app.messagebox.showwarning"):
                    app._on_drop(event)

                self.assertFalse(app.busy)
                self.assertEqual("0 / 2", app.count_var.get())
                self.assertIn("已读取 2 个 Excel 文件", app.status_var.get())
                self.assertIn("发现 2 个可处理的 .xls 文件", app.log_text.get("1.0", "end"))
                self.assertEqual("normal", str(app.start_button.cget("state")))
                self.assertFalse(list(folder.glob("合并结果_*.xls")))

                with patch("app.threading.Thread") as worker:
                    app._start_merge()
                self.assertEqual([first, second], worker.call_args.kwargs["args"][0])
            finally:
                root.destroy()

    def test_open_output_selects_the_generated_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "合并结果_20261010_100000.xls"
            output.write_bytes(b"sample")
            root = TkinterDnD.Tk()
            root.withdraw()
            try:
                app = MergeApp(root)
                root.after = Mock()
                app.events.put(("done", MergeResult(output, 1, 1, 0, 1)))
                app._poll_events()
                self.assertEqual("normal", str(app.open_button.cget("state")))
                with patch("app.subprocess.Popen") as launch:
                    app._open_output()
                launch.assert_called_once_with(["explorer.exe", "/select,", str(output)])
            finally:
                root.destroy()

    def test_selecting_folder_only_scans_until_start_is_clicked(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            (folder / "claim_001.xls").write_bytes(b"sample")
            (folder / "claim_002.xls").write_bytes(b"sample")
            (folder / "合并结果_20261009_120000.xls").write_bytes(b"old")

            root = TkinterDnD.Tk()
            root.withdraw()
            try:
                app = MergeApp(root)
                root.after = Mock()
                app._start_merge = Mock()

                app._set_folder(folder)

                root.after.assert_not_called()
                app._start_merge.assert_not_called()
                self.assertEqual("0 / 2", app.count_var.get())
                self.assertIn("已读取 2 个 Excel 文件", app.status_var.get())
                self.assertEqual("normal", str(app.start_button.cget("state")))
            finally:
                root.destroy()


if __name__ == "__main__":
    unittest.main()
