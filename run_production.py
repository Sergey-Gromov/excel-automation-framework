from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

from excel_framework.validate_input import validate

from transform import transform

DEFAULT_SAMPLE = "samples/input.xlsx"
DEFAULT_OUTPUT = "samples/actual_output.xlsx"
DEFAULT_REPORT_DIR = "reports"
DEFAULT_NEEDS_REVIEW_DIR = "needs_review"


def run(
    input_path: str | Path,
    output_path: str | Path = DEFAULT_OUTPUT,
    sample_path: str | Path = DEFAULT_SAMPLE,
    report_dir: str | Path = DEFAULT_REPORT_DIR,
    needs_review_dir: str | Path = DEFAULT_NEEDS_REVIEW_DIR,
) -> int:
    """
    Gate a production transformation run behind a structural check against
    the known-good sample. If the incoming file's structure has drifted
    (header row moved, columns added/removed, formulas/formats changed),
    quarantine it for human review instead of running the transformation —
    a silent structural mismatch is worse than a loud refusal.

    Returns 0 if the transformation ran, 1 if the file was quarantined.
    """
    input_path = Path(input_path)
    sample_path = Path(sample_path)
    report_dir = Path(report_dir)
    needs_review_dir = Path(needs_review_dir)

    diff_count = validate(sample_path, input_path, report_dir=report_dir)

    if diff_count > 0:
        needs_review_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        case_dir = needs_review_dir / f"{stamp}_{input_path.stem}"
        case_dir.mkdir(parents=True)

        shutil.copy2(input_path, case_dir / input_path.name)
        for report_name in ("input_vs_sample.html", "input_vs_sample.json", "input_vs_sample.xlsx"):
            report_file = report_dir / report_name
            if report_file.exists():
                shutil.copy2(report_file, case_dir / report_name)

        print(
            f"REJECTED: '{input_path}' does not match the expected structure "
            f"({diff_count} difference(s)). Not transformed.\n"
            f"Quarantined to: {case_dir}\n"
            f"Report: {case_dir / 'input_vs_sample.html'}",
            file=sys.stderr,
        )
        return 1

    output_path = Path(output_path)
    transform(input_path, output_path)
    print(f"OK: '{input_path}' matches the expected structure, transformed to '{output_path}'.")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Validate an incoming file against the sample before transforming it. "
            "Quarantines the file for human review instead of transforming it if "
            "the structure has drifted from the sample."
        )
    )
    parser.add_argument("input_path", help="Incoming file to validate and transform")
    parser.add_argument("output_path", nargs="?", default=DEFAULT_OUTPUT)
    parser.add_argument("--sample", default=DEFAULT_SAMPLE, help="Known-good reference file")
    parser.add_argument("--report-dir", default=DEFAULT_REPORT_DIR)
    parser.add_argument("--needs-review-dir", default=DEFAULT_NEEDS_REVIEW_DIR)
    args = parser.parse_args()

    sys.exit(
        run(
            args.input_path,
            args.output_path,
            sample_path=args.sample,
            report_dir=args.report_dir,
            needs_review_dir=args.needs_review_dir,
        )
    )


if __name__ == "__main__":
    main()
