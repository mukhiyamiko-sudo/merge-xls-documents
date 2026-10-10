"""Exercise a source or packaged V5 executable without real claim files."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import xlrd
from openpyxl import load_workbook


HEADER = (
    "CODE", "DESCRIPTION", "ITEM #", "ITEM DESCRIPTION", "UPC",
    "MODEL&nbsp;#", "UNIT COST", "CLAIM QTY", "UOM", "CLAIM AMT",
)
FOOTERS = (
    "TOTAL QUANTITY", "TOTAL FOR CLAIM # 123", "TOTAL FREIGHT FEE",
    "TOTAL HANDLING COST", "TOTAL DISPOSAL FEE",
)


def write_source(path: Path, description: str) -> None:
    row = ("94", description, "1", "ITEM", "0001", "M", "-1", "1", "EA", "-1")
    lines = ["<table>"]
    for values in (HEADER, row):
        lines.append("<tr>" + "".join(f"<td>{value}</td>" for value in values) + "</tr>")
    for label in FOOTERS:
        lines.append(f"<tr><td colspan='5'></td><td colspan='5'>{label}</td></tr>")
    lines.append("</table>")
    path.write_text("\n".join(lines), encoding="utf-8")


def main(command: list[str]) -> int:
    if not command:
        raise SystemExit("Usage: smoke_package.py APP_EXECUTABLE [APP_ARGS ...]")
    with tempfile.TemporaryDirectory(prefix="merge-v5-smoke-") as temporary:
        folder = Path(temporary)
        write_source(folder / "claim_details_000000400699589.xls", "FIRST")
        write_source(folder / "claim_details_000005301102474.xls", "SECOND")
        report = folder / "result.json"
        run = subprocess.run(
            [*command, "--self-test", str(folder), str(report)],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if run.returncode:
            raise AssertionError(f"Packaged self-test failed ({run.returncode}): {run.stdout} {run.stderr}")
        result = json.loads(report.read_text(encoding="utf-8"))
        assert result["ok"] is True, result
        assert (result["merged_files"], result["skipped_files"], result["data_rows"]) == (2, 0, 2), result
        output = Path(result["output"])
        assert output.suffix == (".xlsx" if sys.platform == "darwin" else ".xls"), output
        if output.suffix == ".xlsx":
            book = load_workbook(output, read_only=True, data_only=True)
            try:
                rows = list(book.active.values)
            finally:
                book.close()
        else:
            book = xlrd.open_workbook(str(output))
            sheet = book.sheet_by_index(0)
            rows = [sheet.row_values(index) for index in range(sheet.nrows)]
        assert len(rows) == 3 and all(len(row) == 11 for row in rows), rows
        assert [row[0] for row in rows[1:]] == ["0400699589", "5301102474"]
        assert [row[2] for row in rows[1:]] == ["FIRST", "SECOND"]
        assert not any("TOTAL" in str(value) for row in rows for value in row)

        gui_report = folder / "gui.json"
        gui = subprocess.run(
            [*command, "--gui-self-test", str(gui_report)],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
        if gui.returncode or not gui_report.exists():
            detail = gui_report.read_text(encoding="utf-8") if gui_report.exists() else "no GUI report"
            raise AssertionError(f"Packaged GUI self-test failed ({gui.returncode}): {detail} {gui.stdout} {gui.stderr}")
        gui_result = json.loads(gui_report.read_text(encoding="utf-8"))
        assert gui_result["ok"] is True and gui_result["tkdnd"], gui_result
    print("Packaged V5 smoke test passed: GUI, 2 files, 2 details, 11 columns")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
