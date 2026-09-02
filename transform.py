from __future__ import annotations

import argparse
import re
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


SHEET_NAME = "Ликвидность"
HEADER_ROW = 9
EXPECTED_MAX_ROW = 100
EXPECTED_MAX_COL = 183

DEFAULT_DATE_SOURCE = Path(__file__).resolve().parent / "sh513.xlsx"
DATE_SOURCE_CELL = "B1"
REPLACEMENT_DATE_FORMAT = "%d.%m.%Y"
_DATE_TEXT = r"\d{2}\.\d{2}\.\d{4}"

DATED_HEADER_RULES = (
    (97, "БУ ДЗ всего на"),
    (98, "Сумма просроченной ДЗ с учетом сроков оплаты счетов по состоянию на"),
    (101, "БУ Авансы на"),
)

DATE_COLUMNS = {
    "Дата заключения": 9,
    "Дата расторжения договора": 10,
    "ДЗ начальная дата образования": 19,
    "ДЗ конечная дата образования": 20,
}

ANCHOR_AFTER = (
    "Начислено нарастающим итогом с начала периода с учетом всех перерасчетов",
    74,
)
ANCHOR_BEFORE = ("Оплачено нарастающим итогом с начала периода", 75)

MOVED_HEADERS = (
    ("Сумма годовой корректировки", 162),
    ("Перерасчет суммы годовой корректировки", 163),
)

NEW_FORMULAS = {
    "BB7": "=SUBTOTAL(109,BB10:BB101)",
    "CV7": "=SUBTOTAL(109,CV10:CV101)",
}

DATE_NUMBER_FORMAT = "dd.mm.yyyy"
TOTAL_NUMBER_FORMAT = "#,##0.00"


def source_for_final_column(final_col: int) -> int:
    if final_col <= 74:
        return final_col
    if final_col == 75:
        return 162
    if final_col == 76:
        return 163
    if 77 <= final_col <= 163:
        return final_col - 2
    return final_col


def final_for_source_column(source_col: int) -> int:
    if source_col <= 74:
        return source_col
    if source_col == 162:
        return 75
    if source_col == 163:
        return 76
    if 75 <= source_col <= 161:
        return source_col + 2
    return source_col


_COLUMN_MAP = {
    get_column_letter(source_col): get_column_letter(final_for_source_column(source_col))
    for source_col in range(1, EXPECTED_MAX_COL + 1)
}

_STRING_PART_RE = re.compile(r'("(?:[^"]|"")*")')
_CELL_REF_RE = re.compile(r"(\$?[A-Z]{1,3})(\$?\d+)")


def translate_formula(formula: str) -> str:
    def replace_ref(match: re.Match[str]) -> str:
        col_token = match.group(1)
        row_token = match.group(2)
        col = col_token.lstrip("$")
        mapped = _COLUMN_MAP.get(col)
        if mapped is None:
            return match.group(0)
        prefix = "$" if col_token.startswith("$") else ""
        return f"{prefix}{mapped}{row_token}"

    parts = _STRING_PART_RE.split(formula)
    translated = []
    for index, part in enumerate(parts):
        if index % 2 == 1:
            translated.append(part)
        else:
            translated.append(_CELL_REF_RE.sub(replace_ref, part))
    return "".join(translated)


def _header_at(ws, col: int) -> str | None:
    value = ws.cell(HEADER_ROW, col).value
    if not isinstance(value, str):
        return None
    return value.strip()


def _dated_header_pattern(prefix: str) -> re.Pattern[str]:
    return re.compile(re.escape(prefix) + r"(\s+)(" + _DATE_TEXT + r")")


def _match_dated_header(value, prefix: str) -> re.Match[str] | None:
    if not isinstance(value, str) or not value.lstrip().startswith(prefix):
        return None
    return _dated_header_pattern(prefix).search(value)


def replace_date_in_header(value: str, prefix: str, new_date: str) -> str:
    pattern = _dated_header_pattern(prefix)
    return pattern.sub(lambda m: f"{prefix}{m.group(1)}{new_date}", value, count=1)


def read_replacement_date(date_source_path: str | Path) -> str:
    path = Path(date_source_path)
    if not path.exists():
        raise ValueError(f"Файл с датой '{path}' не найден.")

    wb = load_workbook(path, data_only=False)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Лист '{SHEET_NAME}' не найден в файле с датой '{path}'.")

    value = wb[SHEET_NAME][DATE_SOURCE_CELL].value
    if isinstance(value, str) and value.startswith("="):
        wb = load_workbook(path, data_only=True)
        value = wb[SHEET_NAME][DATE_SOURCE_CELL].value

    if isinstance(value, datetime):
        parsed = value.date()
    elif isinstance(value, date):
        parsed = value
    elif isinstance(value, str):
        text = value.strip()
        for source_format in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, source_format).strftime(REPLACEMENT_DATE_FORMAT)
            except ValueError:
                continue
        raise ValueError(
            f"Не удалось распознать дату в {path}!{DATE_SOURCE_CELL}: '{text}'."
        )
    else:
        raise ValueError(
            f"Не удалось получить дату из {path}!{DATE_SOURCE_CELL}: значение {value!r}."
        )

    return parsed.strftime(REPLACEMENT_DATE_FORMAT)


def validate_dated_headers(ws) -> None:
    expected_by_prefix = {prefix: col for col, prefix in DATED_HEADER_RULES}

    for col, prefix in DATED_HEADER_RULES:
        value = ws.cell(HEADER_ROW, col).value
        if _match_dated_header(value, prefix) is None:
            raise ValueError(
                f"Заголовок '{prefix}' с датой не найден в "
                f"{get_column_letter(col)}{HEADER_ROW}."
            )

    for col in range(1, ws.max_column + 1):
        value = ws.cell(HEADER_ROW, col).value
        if not isinstance(value, str):
            continue
        for prefix, expected_col in expected_by_prefix.items():
            if _match_dated_header(value, prefix) and col != expected_col:
                raise ValueError(
                    f"Заголовок '{prefix}' с датой найден в "
                    f"{get_column_letter(col)}{HEADER_ROW}, ожидается "
                    f"{get_column_letter(expected_col)}{HEADER_ROW}."
                )


def replace_dated_headers(ws, new_date: str) -> None:
    for col, prefix in DATED_HEADER_RULES:
        cell = ws.cell(HEADER_ROW, col)
        value = cell.value
        if _match_dated_header(value, prefix) is None:
            raise ValueError(
                f"Заголовок '{prefix}' с датой не найден в "
                f"{get_column_letter(col)}{HEADER_ROW}."
            )
        cell.value = replace_date_in_header(value, prefix, new_date)


def _sample_dated_header_dates(ws) -> dict[str, str]:
    dates: dict[str, str] = {}
    for col, prefix in DATED_HEADER_RULES:
        value = ws.cell(HEADER_ROW, col).value
        match = _match_dated_header(value, prefix)
        if match is None:
            raise ValueError(
                f"В образце не найден заголовок '{prefix}' с датой в "
                f"{get_column_letter(col)}{HEADER_ROW}."
            )
        dates[prefix] = match.group(2)
    return dates


def prepare_validation_input(
    input_path: str | Path,
    sample_path: str | Path,
    temp_dir: str | Path,
) -> Path:
    input_path = Path(input_path)
    sample_path = Path(sample_path)
    temp_dir = Path(temp_dir)

    sample_wb = load_workbook(sample_path, data_only=False)
    if SHEET_NAME not in sample_wb.sheetnames:
        raise ValueError(f"Лист '{SHEET_NAME}' не найден в образце '{sample_path}'.")
    sample_dates = _sample_dated_header_dates(sample_wb[SHEET_NAME])

    wb = load_workbook(input_path, data_only=False)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Лист '{SHEET_NAME}' не найден во входной книге '{input_path}'.")
    ws = wb[SHEET_NAME]

    changed = False
    for col, prefix in DATED_HEADER_RULES:
        cell = ws.cell(HEADER_ROW, col)
        value = cell.value
        match = _match_dated_header(value, prefix)
        if match is not None and match.group(2) != sample_dates[prefix]:
            cell.value = replace_date_in_header(value, prefix, sample_dates[prefix])
            changed = True

    if not changed:
        return input_path

    normalized_path = temp_dir / input_path.name
    normalized_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(normalized_path)
    return normalized_path


def validate_structure(ws) -> None:
    if ws.max_row != EXPECTED_MAX_ROW:
        raise ValueError(
            f"Ожидалось {EXPECTED_MAX_ROW} строк в листе '{SHEET_NAME}', "
            f"фактически {ws.max_row}."
        )
    if ws.max_column != EXPECTED_MAX_COL:
        raise ValueError(
            f"Ожидалось {EXPECTED_MAX_COL} столбцов в листе '{SHEET_NAME}', "
            f"фактически {ws.max_column}."
        )

    required_headers = (
        *DATE_COLUMNS.items(),
        ANCHOR_AFTER,
        ANCHOR_BEFORE,
        *MOVED_HEADERS,
    )
    for name, col in required_headers:
        actual = _header_at(ws, col)
        if actual != name:
            raise ValueError(
                f"Заголовок в {get_column_letter(col)}{HEADER_ROW} должен быть "
                f"'{name}', фактически '{actual}'."
            )

    validate_dated_headers(ws)


def _snapshot_column_dimensions(ws, max_col: int) -> dict[int, tuple]:
    snapshot = {}
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        dim = ws.column_dimensions.get(letter)
        if dim is None:
            snapshot[col] = None
        else:
            snapshot[col] = (dim.width, dim.bestFit, dim.hidden, dim.outlineLevel)
    return snapshot


def _apply_column_dimensions(ws, max_col: int, snapshot: dict[int, tuple]) -> None:
    for final_col in range(1, max_col + 1):
        final_letter = get_column_letter(final_col)
        if final_letter in ws.column_dimensions:
            del ws.column_dimensions[final_letter]

        source_dim = snapshot[source_for_final_column(final_col)]
        if source_dim is None:
            continue

        dim = ws.column_dimensions[final_letter]
        dim.width, dim.bestFit, dim.hidden, dim.outlineLevel = source_dim


def _remap_merged_ranges(ws, original_ranges: list[str]) -> None:
    for original in original_ranges:
        min_col, min_row, max_col, max_row = range_boundaries(original)
        remapped = (
            f"{get_column_letter(final_for_source_column(min_col))}{min_row}:"
            f"{get_column_letter(final_for_source_column(max_col))}{max_row}"
        )
        ws.merge_cells(remapped)


def transform(
    input_path: str | Path,
    output_path: str | Path,
    date_source_path: str | Path | None = None,
) -> None:
    input_path = Path(input_path)
    output_path = Path(output_path)
    date_source_path = (
        Path(date_source_path) if date_source_path is not None else DEFAULT_DATE_SOURCE
    )

    if input_path.resolve() == output_path.resolve():
        raise ValueError("Input and output paths must be different.")

    wb = load_workbook(input_path)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Лист '{SHEET_NAME}' не найден во входной книге.")

    ws = wb[SHEET_NAME]
    validate_structure(ws)
    replacement_date = read_replacement_date(date_source_path)
    replace_dated_headers(ws, replacement_date)

    max_row = ws.max_row
    max_col = ws.max_column

    cell_snapshot = {}
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            cell_snapshot[(row, col)] = (
                cell.value,
                copy(cell._style),
                copy(cell.hyperlink) if cell.hyperlink else None,
                copy(cell.comment) if cell.comment else None,
            )

    column_snapshot = _snapshot_column_dimensions(ws, max_col)
    merged_ranges = [str(rng) for rng in ws.merged_cells.ranges]
    for merged_range in merged_ranges:
        ws.unmerge_cells(merged_range)

    for final_col in range(1, max_col + 1):
        source_col = source_for_final_column(final_col)

        for row in range(1, max_row + 1):
            value, style, hyperlink, comment = cell_snapshot[(row, source_col)]
            if isinstance(value, str) and value.startswith("="):
                value = translate_formula(value)

            cell = ws.cell(row, final_col)
            cell.value = value
            cell._style = copy(style)
            cell.hyperlink = hyperlink
            if hyperlink is not None:
                hyperlink.ref = cell.coordinate
            cell.comment = comment

    _apply_column_dimensions(ws, max_col, column_snapshot)
    _remap_merged_ranges(ws, merged_ranges)

    for col in DATE_COLUMNS.values():
        for row in range(10, max_row + 1):
            ws.cell(row, col).number_format = DATE_NUMBER_FORMAT

    for coord, formula in NEW_FORMULAS.items():
        cell = ws[coord]
        cell.value = formula
        cell.number_format = TOTAL_NUMBER_FORMAT

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_path", nargs="?", default="samples/input.xlsx")
    parser.add_argument("output_path", nargs="?", default="samples/actual_output.xlsx")
    parser.add_argument(
        "--date-source",
        default=None,
        help=f"Workbook containing the replacement date in {SHEET_NAME}!{DATE_SOURCE_CELL}",
    )
    args = parser.parse_args()

    transform(args.input_path, args.output_path, args.date_source)


if __name__ == "__main__":
    main()
