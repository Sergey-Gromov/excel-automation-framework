from pathlib import Path

from excel_framework.compare_workbooks import compare_workbooks
from excel_framework.report import write_html_report, write_json_report


def test_no_unexpected_changes_from_input(paths, allowed_changes):
    if not paths["input"].exists() or not paths["actual"].exists():
        return

    diffs = compare_workbooks(paths["input"], paths["actual"], allowed_changes=allowed_changes)

    write_json_report(diffs, Path("reports/input_vs_actual_report.json"))
    write_html_report(diffs, Path("reports/input_vs_actual_report.html"))

    unexpected = [d for d in diffs if not d.allowed]
    assert not unexpected, (
        "Actual workbook contains changes outside allowed ranges. "
        "See reports/input_vs_actual_report.html."
    )
