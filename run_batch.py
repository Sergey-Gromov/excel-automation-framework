from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from uuid import uuid4

from run_production import (
    DEFAULT_NEEDS_REVIEW_DIR,
    DEFAULT_REPORT_DIR,
    DEFAULT_SAMPLE,
    run,
)


def discover_workbooks(input_dir: str | Path) -> list[Path]:
    input_dir = Path(input_dir)
    if not input_dir.is_dir():
        return []

    workbooks = [
        path
        for path in input_dir.iterdir()
        if path.is_file()
        and path.suffix == ".xlsx"
        and not path.name.startswith("~$")
    ]
    return sorted(workbooks, key=lambda path: (path.name.casefold(), path.name))


def _unlink_missing(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        pass


def run_batch(
    input_dir: str | Path,
    output_dir: str | Path,
    sample_path: str | Path = DEFAULT_SAMPLE,
    report_dir: str | Path = DEFAULT_REPORT_DIR,
    needs_review_dir: str | Path = DEFAULT_NEEDS_REVIEW_DIR,
    overwrite: bool = False,
) -> int:
    input_resolved = Path(input_dir).resolve()
    output_resolved = Path(output_dir).resolve()

    if input_resolved == output_resolved:
        print(
            f"ERROR: input and output directories must be different: {input_resolved}",
            file=sys.stderr,
        )
        return 1

    if not input_resolved.is_dir():
        print(f"ERROR: input directory not found: {input_resolved}", file=sys.stderr)
        return 1

    files = discover_workbooks(input_resolved)
    if not files:
        print(f"No eligible .xlsx workbooks found in '{input_resolved}'.")
        return 0

    try:
        output_resolved.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        print(f"ERROR: cannot create output directory '{output_resolved}': {exc}", file=sys.stderr)
        return 1

    report_base = Path(report_dir)
    needs_review_base = Path(needs_review_dir)

    transformed: list[Path] = []
    rejected: list[Path] = []
    failed: list[tuple[Path, str]] = []

    for file in files:
        output_path = output_resolved / file.name

        if output_path.exists() or output_path.is_symlink():
            if output_path.is_dir():
                reason = f"output path is a directory: {output_path}"
            elif not overwrite:
                reason = f"output already exists and --overwrite was not supplied: {output_path}"
            else:
                reason = None
            if reason is not None:
                print(f"FAILED: {file}: {reason}", file=sys.stderr)
                failed.append((file, reason))
                continue

        if output_path.resolve() == file.resolve():
            reason = f"output path resolves to the source file: {output_path}"
            print(f"FAILED: {file}: {reason}", file=sys.stderr)
            failed.append((file, reason))
            continue

        file_report_dir = report_base / file.stem
        file_needs_review_dir = needs_review_base / file.stem
        temp_path = output_resolved / f".{file.name}.tmp-{uuid4().hex}"

        try:
            exit_code = run(
                input_path=file,
                output_path=temp_path,
                sample_path=sample_path,
                report_dir=file_report_dir,
                needs_review_dir=file_needs_review_dir,
            )
            if exit_code == 0:
                os.replace(temp_path, output_path)
                transformed.append(file)
            else:
                rejected.append(file)
        except Exception as exc:
            reason = str(exc) or exc.__class__.__name__
            print(f"FAILED: {file}: {reason}", file=sys.stderr)
            failed.append((file, reason))
        finally:
            _unlink_missing(temp_path)

    print(f"Files discovered: {len(files)}")
    print(f"Transformed successfully: {len(transformed)}")
    print(f"Rejected because of input drift: {len(rejected)}")
    print(f"Failed because of execution or output conflicts: {len(failed)}")

    if rejected:
        print("Rejected files:")
        for file in rejected:
            print(f"- {file}")

    if failed:
        print("Failed files:")
        for file, reason in failed:
            print(f"- {file}: {reason}")

    return 0 if not rejected and not failed else 1


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Validate and transform every .xlsx workbook directly inside an input "
            "directory. Output filenames match input filenames. Drifted files are "
            "quarantined; processing continues for remaining files."
        )
    )
    parser.add_argument("input_dir", help="Directory containing incoming .xlsx workbooks")
    parser.add_argument("output_dir", help="Different directory for transformed workbooks")
    parser.add_argument("--sample", default=DEFAULT_SAMPLE, help="Known-good reference file")
    parser.add_argument("--report-dir", default=DEFAULT_REPORT_DIR)
    parser.add_argument("--needs-review-dir", default=DEFAULT_NEEDS_REVIEW_DIR)
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing output files using an atomic temporary-file replacement",
    )
    args = parser.parse_args()

    sys.exit(
        run_batch(
            args.input_dir,
            args.output_dir,
            sample_path=args.sample,
            report_dir=args.report_dir,
            needs_review_dir=args.needs_review_dir,
            overwrite=args.overwrite,
        )
    )


if __name__ == "__main__":
    main()
