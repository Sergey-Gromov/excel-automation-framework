"""Focused tests for the deletion of the "Исчезнувшие" (EA / 131) and
"Период" (EB / 132) columns.

The transformation removes source columns 131 (EA) and 132 (EB) and shifts
every column to their right two positions left, so the output is 181 columns
wide instead of 183. These tests pin that behavior down directly rather than
through the generic comparator.
"""
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import transform
from transform import data_row, new_subtotal_formula


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_INPUT = ROOT / "samples" / "input.xlsx"
DATE_SOURCE = ROOT / "sh513.xlsx"


def _require_samples() -> None:
    if not SAMPLE_INPUT.exists() or not DATE_SOURCE.exists():
        pytest.skip("Put samples/input.xlsx and sh513.xlsx into the repository first.")


def _load_sheet(path: Path):
    return load_workbook(path, data_only=False)[transform.SHEET_NAME]


@pytest.fixture(scope="module")
def transformed_output(tmp_path_factory):
    _require_samples()
    tmp_dir = tmp_path_factory.mktemp("deleted_column_transform")
    output_path = tmp_dir / "actual_output.xlsx"
    transform.transform(SAMPLE_INPUT, output_path, DATE_SOURCE)
    return output_path


def test_output_is_181_columns_wide(transformed_output):
    ws = _load_sheet(transformed_output)
    assert ws.max_column == transform.OUTPUT_MAX_COL == 181


def test_deleted_headers_are_absent(transformed_output):
    input_ws = _load_sheet(SAMPLE_INPUT)
    ws = _load_sheet(transformed_output)

    def headers(sheet):
        return [
            sheet.cell(transform.HEADER_ROW, col).value
            for col in range(1, sheet.max_column + 1)
        ]

    # "Исчезнувшие" (source EA / 131) is unique, so it must be gone entirely.
    assert not any(isinstance(h, str) and "счез" in h for h in headers(ws))

    # "Период" (source EB / 132) is not unique, so verify exactly one of the
    # source's "Период" columns was removed and the rest survived the shift.
    source_period = sum(1 for h in headers(input_ws) if h == "Период")
    output_period = sum(1 for h in headers(ws) if h == "Период")
    assert output_period == source_period - 1


def test_columns_after_deletion_shift_left(transformed_output):
    input_ws = _load_sheet(SAMPLE_INPUT)
    actual_ws = _load_sheet(transformed_output)

    # Two source columns (131 EA, 132 EB) are deleted, so every column to
    # their right shifts two positions left. Source DY (129) and DZ (130)
    # land on final EA (131) and EB (132); source EC (133) stays on final EC
    # (133) because the +2 move and the -2 deletion cancel out.
    assert actual_ws.cell(9, 131).value == input_ws.cell(9, 129).value
    assert actual_ws.cell(9, 131).value == "Организация ликвидирована"
    assert actual_ws.cell(9, 132).value == input_ws.cell(9, 130).value
    assert actual_ws.cell(9, 132).value == "Период"
    assert actual_ws.cell(9, 133).value == input_ws.cell(9, 133).value
    assert actual_ws.cell(9, 133).value == "Соглашение о реструктуризации"


def test_final_ea_holds_translated_dy_subtotal(transformed_output):
    input_ws = _load_sheet(SAMPLE_INPUT)
    actual_ws = _load_sheet(transformed_output)
    last_data_row = data_row(input_ws)

    # Final EA (131) <- source DY (129). The deleted source EA (131) leaves no
    # formula behind; the cell now holds source DY's row-7 SUBTOTAL, translated
    # DY -> EA. This pins down that the column landing on EA is correct and that
    # no untranslated leftover from the deleted column remains.
    source_value = input_ws.cell(7, 129).value
    assert isinstance(source_value, str) and "SUBTOTAL" in source_value
    expected = new_subtotal_formula("EA", last_data_row)
    assert actual_ws.cell(7, 131).value == expected


def test_shifted_subtotals_are_translated(transformed_output):
    input_ws = _load_sheet(SAMPLE_INPUT)
    actual_ws = _load_sheet(transformed_output)
    last_data_row = data_row(input_ws)

    # (final_col, source_col) pairs in the shifted block (77:132) whose source
    # row-7 cell holds a SUBTOTAL.
    for final_col, source_col in ((77, 75), (131, 129)):
        source_value = input_ws.cell(7, source_col).value
        assert isinstance(source_value, str) and "SUBTOTAL" in source_value
        final_letter = get_column_letter(final_col)
        expected = new_subtotal_formula(final_letter, last_data_row)
        actual = actual_ws.cell(7, final_col).value
        assert actual == expected, (
            f"row-7 final col {final_col} ({final_letter}) from source "
            f"{source_col}: {actual!r} != {expected!r}"
        )


def test_merged_ranges_shrink_around_deletion(transformed_output):
    ws = _load_sheet(transformed_output)
    assert sorted(str(rng) for rng in ws.merged_cells.ranges) == [
        "DH8:EM8",
        "EN8:EP8",
    ]


def test_column_dimensions_follow_new_mapping(transformed_output):
    input_ws = _load_sheet(SAMPLE_INPUT)
    actual_ws = _load_sheet(transformed_output)
    # (final_col, source_col): shifted (+2), identity, and tail (-2) blocks.
    for final_col, source_col in ((131, 129), (133, 133), (162, 164)):
        final_letter = get_column_letter(final_col)
        source_letter = get_column_letter(source_col)
        for attr in ("width", "bestFit", "hidden", "outlineLevel"):
            assert getattr(actual_ws.column_dimensions[final_letter], attr) == (
                getattr(input_ws.column_dimensions[source_letter], attr)
            ), f"column {attr} mismatch at {final_letter}"
    # The removed tail columns leave no dimension behind.
    assert get_column_letter(182) not in actual_ws.column_dimensions
    assert get_column_letter(183) not in actual_ws.column_dimensions


def test_leftover_tail_columns_are_empty(transformed_output):
    ws = _load_sheet(transformed_output)
    for row in range(1, ws.max_row + 1):
        assert ws.cell(row, 182).value is None
        assert ws.cell(row, 183).value is None


def test_missing_deleted_header_is_rejected(tmp_path):
    _require_samples()
    for expected, col in transform.DELETED_HEADERS:
        input_path = tmp_path / f"input_{col}.xlsx"
        input_path.write_bytes(SAMPLE_INPUT.read_bytes())
        wb = load_workbook(input_path)
        ws = wb[transform.SHEET_NAME]
        ws.cell(transform.HEADER_ROW, col).value = "Другой заголовок"
        wb.save(input_path)
        with pytest.raises(ValueError, match=expected):
            transform.transform(input_path, tmp_path / f"out_{col}.xlsx", DATE_SOURCE)
