import shutil
from pathlib import Path

from openpyxl import load_workbook

from run_production import run
import transform

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"


def _make_drifted_input(tmp_path: Path) -> Path:
    source = SAMPLES / "input.xlsx"
    drifted = tmp_path / "input_drifted.xlsx"
    shutil.copy2(source, drifted)

    wb = load_workbook(drifted)
    ws = wb[transform.SHEET_NAME]
    ws["BW9"] = "drift"
    wb.save(drifted)

    return drifted


def test_clean_input_gets_transformed(tmp_path):
    output_path = tmp_path / "actual_output.xlsx"
    needs_review_dir = tmp_path / "needs_review"

    exit_code = run(
        input_path=SAMPLES / "input.xlsx",
        output_path=output_path,
        sample_path=SAMPLES / "input.xlsx",
        report_dir=tmp_path / "reports",
        needs_review_dir=needs_review_dir,
    )

    assert exit_code == 0
    assert output_path.exists()
    assert not needs_review_dir.exists()


def test_drifted_input_is_quarantined_not_transformed(tmp_path):
    drifted_input = _make_drifted_input(tmp_path)
    output_path = tmp_path / "actual_output.xlsx"
    needs_review_dir = tmp_path / "needs_review"

    exit_code = run(
        input_path=drifted_input,
        output_path=output_path,
        sample_path=SAMPLES / "input.xlsx",
        report_dir=tmp_path / "reports",
        needs_review_dir=needs_review_dir,
    )

    assert exit_code == 1
    assert not output_path.exists()

    case_dirs = list(needs_review_dir.iterdir())
    assert len(case_dirs) == 1
    quarantined = case_dirs[0]
    assert (quarantined / "input_drifted.xlsx").exists()
    assert (quarantined / "input_vs_sample.html").exists()
