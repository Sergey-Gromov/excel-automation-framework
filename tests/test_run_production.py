from pathlib import Path

from run_production import run

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"


def test_clean_input_gets_transformed(tmp_path):
    output_path = tmp_path / "actual_output.xlsx"
    needs_review_dir = tmp_path / "needs_review"

    exit_code = run(
        input_path=SAMPLES / "sample.xlsx",
        output_path=output_path,
        sample_path=SAMPLES / "sample.xlsx",
        report_dir=tmp_path / "reports",
        needs_review_dir=needs_review_dir,
    )

    assert exit_code == 0
    assert output_path.exists()
    assert not needs_review_dir.exists()


def test_drifted_input_is_quarantined_not_transformed(tmp_path):
    output_path = tmp_path / "actual_output.xlsx"
    needs_review_dir = tmp_path / "needs_review"

    exit_code = run(
        input_path=SAMPLES / "input_modified.xlsx",
        output_path=output_path,
        sample_path=SAMPLES / "sample.xlsx",
        report_dir=tmp_path / "reports",
        needs_review_dir=needs_review_dir,
    )

    assert exit_code == 1
    assert not output_path.exists()

    case_dirs = list(needs_review_dir.iterdir())
    assert len(case_dirs) == 1
    quarantined = case_dirs[0]
    assert (quarantined / "input_modified.xlsx").exists()
    assert (quarantined / "input_vs_sample.html").exists()
