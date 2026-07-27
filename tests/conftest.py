from pathlib import Path

import pytest

from excel_framework.config import load_allowed_changes
from excel_framework.compare_workbooks import compare_workbooks
from excel_framework.report import write_html_report, write_json_report


ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
REPORTS = ROOT / "reports"


@pytest.fixture(scope="session")
def paths():
    return {
        "input": SAMPLES / "input.xlsx",
        "expected": SAMPLES / "expected_output.xlsx",
        "actual": SAMPLES / "actual_output.xlsx",
        "allowed": ROOT / "docs" / "ALLOWED_CHANGES.yml",
        "json_report": REPORTS / "diff_report.json",
        "html_report": REPORTS / "diff_report.html",
    }


@pytest.fixture(scope="session")
def allowed_changes(paths):
    return load_allowed_changes(paths["allowed"])


@pytest.fixture(scope="session")
def workbook_diffs(paths, allowed_changes):
    if not paths["expected"].exists() or not paths["actual"].exists():
        pytest.skip("Put expected_output.xlsx and actual_output.xlsx into samples/ first.")

    diffs = compare_workbooks(paths["expected"], paths["actual"], allowed_changes=allowed_changes)
    write_json_report(diffs, paths["json_report"])
    write_html_report(diffs, paths["html_report"])
    return diffs
