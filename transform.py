from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook


def transform(input_path: str | Path, output_path: str | Path) -> None:
    """
    Implement the workbook transformation here.

    Rules:
    - Do not recreate the workbook.
    - Do not use pandas to write the workbook.
    - Change only ranges allowed in docs/ALLOWED_CHANGES.yml.
    - Preserve formulas and formatting outside allowed ranges.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

    wb = load_workbook(input_path)

    # TODO: implement project-specific changes here.
    #
    # Example:
    # ws = wb["Orders"]
    # ws["A1"] = "New Header"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_path", nargs="?", default="samples/input.xlsx")
    parser.add_argument("output_path", nargs="?", default="samples/actual_output.xlsx")
    args = parser.parse_args()

    transform(args.input_path, args.output_path)


if __name__ == "__main__":
    main()
